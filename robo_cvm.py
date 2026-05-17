"""
============================================================
ROB_ CVM v2.0 _ Catalogo de Fundos Caixa
============================================================
Fonte   : dados.cvm.gov.br  (API oficial, gratuita)
GitHub  : eltonprivatebanker/catalogo-fundos-cvm

O que faz:
  1. Descobre fundos Caixa automaticamente via cadastro CVM
     OU le uma lista de CNPJs de um arquivo (JSON/CSV/TXT)
  2. Baixa cotas diarias com cache local (nao rebaixa o que ja tem)
  3. Calcula: variacao dia, acum. mes, acum. ano, acum. 12M
  4. Retry automatico em cada download
  5. Salva: dados_atuais.csv + sipii_caixa_YYYYMMDD.csv/xlsx + HTML

Uso:
  python robo_cvm.py                         # detecta JSON na pasta OU busca Caixa no cadastro
  python robo_cvm.py --lista fundos.json     # usa lista especifica de CNPJs
  python robo_cvm.py --sem-cache             # forca redownload de tudo
  python robo_cvm.py --apenas-abertos        # filtra so fundos abertos para captacao
  python robo_cvm.py --saida minha_pasta     # define pasta de saida
============================================================
"""

import argparse
import glob
import io
import json
import os
import re
import time
import unicodedata
import zipfile
from datetime import date, datetime
from io import BytesIO
from typing import Optional

import pandas as pd
import requests

# ==========================================================
# CONFIGURACAO
# ==========================================================
URL_CAD     = "https://dados.cvm.gov.br/dados/FI/CAD/DADOS/cad_fi.csv"
URL_INF_MES = "https://dados.cvm.gov.br/dados/FI/DOC/INF_DIARIO/DADOS/inf_diario_fi_{yyyymm}.zip"
URL_INF_ZIP = "https://dados.cvm.gov.br/dados/FI/DOC/INF_DIARIO/DADOS/HIST/inf_diario_fi_{ano}.zip"
CACHE_DIR   = ".cache_cvm"
TIMEOUT     = 90
MAX_RETRY   = 3

CLASSE_MAP = {
    "Fundo de Renda Fixa":                "RENDA FIXA",
    "Fundo de Acoes":                     "ACOES",
    "Fundo Multimercado":                 "MULTIMERCADO",
    "Fundo Cambial":                      "CAMBIAL",
    "Fundo de _ndice de Mercado":         "FUNDO DE INDICE",
    "Fundo Mutuo de Privatizacao - FGTS": "PRIVATIZACAO",
    "Fundo de Investimento Imobiliario":  "FII",
}


# ==========================================================
# CLI
# ==========================================================
def parse_args():
    p = argparse.ArgumentParser(description="Robo CVM _ Catalogo de Fundos Caixa v2.0")
    p.add_argument("--lista",         default=None,
                   help="Arquivo com CNPJs: JSON (Caixa), CSV ou TXT. "
                        "Se omitido, detecta automaticamente na pasta ou busca Caixa no cadastro CVM.")
    p.add_argument("--sem-cache",     action="store_true", dest="sem_cache",
                   help="Forca redownload de todos os arquivos (ignora cache)")
    p.add_argument("--apenas-abertos", action="store_true", dest="apenas_abertos",
                   help="Inclui so fundos abertos para captacao (quando a fonte e JSON Caixa)")
    p.add_argument("--saida",         default=".", dest="saida",
                   help="Pasta de saida para os arquivos gerados (default: pasta atual)")
    p.add_argument("--filtro-gestor", default=None, dest="filtro_gestor",
                   help="Filtra fundos cujo GESTOR contenha este texto (ex: 'caixa')")
    p.add_argument("--modo-cvm", action="store_true", dest="modo_cvm",
                   help="Ignora JSON local e busca direto no cadastro CVM (modo GitHub Actions)")
    return p.parse_args()


# ==========================================================
# UTILIT_RIOS
# ==========================================================
def norm_cnpj(c: str) -> str:
    """Remove pontuacao e retorna 14 digitos com zero-padding."""
    return re.sub(r"\D", "", str(c)).zfill(14)

def fmt_cnpj(c: str) -> str:
    c = norm_cnpj(c)
    return f"{c[:2]}.{c[2:5]}.{c[5:8]}/{c[8:12]}-{c[12:14]}"

def rm_accent(s: str) -> str:
    return "".join(ch for ch in unicodedata.normalize("NFD", s)
                   if unicodedata.category(ch) != "Mn")

def fmt_br(v, decimais: int = 2) -> str:
    """Float -> string no padrao brasileiro ('1.234,56')."""
    if v is None or (isinstance(v, float) and pd.isna(v)):
        return ""
    return f"{v:,.{decimais}f}".replace(",", "X").replace(".", ",").replace("X", ".")

def retorno_pct(cota_ini, cota_fim) -> Optional[float]:
    try:
        if cota_ini and cota_fim and float(cota_ini) > 0:
            return (float(cota_fim) / float(cota_ini) - 1) * 100
    except (TypeError, ValueError):
        pass
    return None

def limpar_aba(nome: str) -> str:
    """Remove caracteres proibidos pelo Excel e limita a 31 chars."""
    return re.sub(r'[:\\/*_\[\]]', "", nome)[:31]

def log(msg: str):
    print(f"{datetime.now().strftime('%H:%M:%S')}  {msg}")

def refinar_categoria(classe_cvm: str, nome_fundo: str) -> str:
    nome = rm_accent(nome_fundo.upper())
    if "SIMPLES" in nome:
        return "RENDA FIXA SIMPLES"
    if any(x in nome for x in ("REF DI", "REFERENCIADO", "REFERENC")):
        return "RENDA FIXA REF DI"
    if "CURTO PRAZO" in nome and "Renda Fixa" in classe_cvm:
        return "RENDA FIXA CURTO PRAZO"
    if "DEBENTURE" in nome or "DEB INCENT" in nome:
        return "RENDA FIXA DEBENTURES"
    if "CREDITO" in nome or "CRED PRIV" in nome:
        return "RENDA FIXA CRED PRIV"
    return CLASSE_MAP.get(classe_cvm, classe_cvm.upper() if classe_cvm else "OUTROS")


# ==========================================================
# DOWNLOAD COM CACHE E RETRY
# ==========================================================
def _cache_path(url: str) -> str:
    nome = re.sub(r"[^\w.]", "_", url.split("/")[-1])
    return os.path.join(CACHE_DIR, nome)

def baixar_bytes(url: str, label: str = "", sem_cache: bool = False) -> bytes | None:
    cache = _cache_path(url)
    if not sem_cache and os.path.exists(cache):
        log(f"[cache] Cache: {label or os.path.basename(cache)}")
        with open(cache, "rb") as f:
            return f.read()

    for tentativa in range(MAX_RETRY):
        try:
            log(f"[download]  Baixando: {label or url}")
            r = requests.get(url, timeout=TIMEOUT, headers={"User-Agent": "Mozilla/5.0"})
            r.raise_for_status()
            os.makedirs(CACHE_DIR, exist_ok=True)
            with open(cache, "wb") as f:
                f.write(r.content)
            log(f"[OK]  Salvo ({len(r.content)/1e6:.1f} MB)")
            return r.content
        except Exception as e:
            if tentativa == MAX_RETRY - 1:
                log(f"[ERRO]  Falhou apos {MAX_RETRY} tentativas: {e}")
                return None
            log(f"[AVISO]   Tentativa {tentativa+1} falhou, retry em 3s...")
            time.sleep(3)
    return None

def baixar_csv_bytes(conteudo: bytes) -> pd.DataFrame:
    """Le um CSV a partir de bytes (latin1 ou utf-8)."""
    for enc in ("latin1", "utf-8-sig", "utf-8"):
        try:
            return pd.read_csv(io.StringIO(conteudo.decode(enc)), sep=";", low_memory=False, dtype=str)
        except Exception:
            continue
    return pd.DataFrame()

def baixar_csv_url(url: str, label: str = "", sem_cache: bool = False) -> pd.DataFrame:
    dados = baixar_bytes(url, label, sem_cache)
    if dados is None:
        return pd.DataFrame()
    return baixar_csv_bytes(dados)

def baixar_zip_inf(ano: int, mes: int, sem_cache: bool = False) -> pd.DataFrame:
    """Baixa informe diario mensal em ZIP e retorna DataFrame."""
    yyyymm = f"{ano}{mes:02d}"
    url = URL_INF_MES.format(yyyymm=yyyymm)
    label = f"informe diario {yyyymm}"
    dados = baixar_bytes(url, label, sem_cache)
    if dados is None:
        return pd.DataFrame()
    try:
        with zipfile.ZipFile(BytesIO(dados)) as zf:
            csvs = [n for n in zf.namelist() if n.lower().endswith(".csv")]
            if not csvs:
                return pd.DataFrame()
            with zf.open(csvs[0]) as f:
                return pd.read_csv(f, sep=";", encoding="latin1", low_memory=False, dtype=str)
    except Exception as e:
        log(f"[ERRO]  Erro ao abrir ZIP {yyyymm}: {e}")
        return pd.DataFrame()


# ==========================================================
# LEITURA DA LISTA DE CNPJs
# ==========================================================
def _autodetectar_lista() -> Optional[str]:
    """Retorna o primeiro JSON/CSV/TXT encontrado na pasta atual, ou None."""
    for ext in ("*.json", "*.csv", "*.txt"):
        arquivos = sorted(glob.glob(ext))
        # Ignora arquivos gerados pelo proprio script
        arquivos = [a for a in arquivos if not any(
            x in a.lower() for x in ("dados_atuais", "sipii_caixa", "cnpjs_lista", "requirements")
        )]
        if arquivos:
            log(f"[arquivo] Arquivo detectado automaticamente: {arquivos[0]}")
            return arquivos[0]
    return None

def ler_lista_cnpjs(caminho: str) -> tuple:
    """
    Le CNPJs de: JSON (formato Caixa), CSV ou TXT.
    Retorna (lista_cnpjs, df_extra_ou_None).
    df_extra contem colunas extras do JSON (performance, taxa, etc.)
    """
    ext = os.path.splitext(caminho)[1].lower()
    log(f"[1] Lendo lista: {caminho}")

    if ext == ".json":
        with open(caminho, encoding="utf-8") as f:
            dados = json.load(f)
        if isinstance(dados, dict):
            dados = next(iter(dados.values()))
        df = pd.json_normalize(dados)
        col = next((c for c in df.columns if "cnpj" in c.lower()), None)
        if col is None:
            raise ValueError("Coluna CNPJ nao encontrada no JSON.")
        df["_CNPJ_NORM"] = df[col].apply(norm_cnpj)
        log(f"   [OK] {len(df)} fundos no JSON (coluna '{col}')")
        return df["_CNPJ_NORM"].tolist(), df

    elif ext == ".csv":
        for sep in (";", ",", "\t"):
            try:
                df = pd.read_csv(caminho, sep=sep, dtype=str, encoding="utf-8-sig")
                if len(df.columns) > 1:
                    break
            except Exception:
                continue
        col = next((c for c in df.columns if "cnpj" in c.lower()), df.columns[0])
        df["_CNPJ_NORM"] = df[col].apply(norm_cnpj)
        log(f"   [OK] {len(df)} CNPJs no CSV")
        return df["_CNPJ_NORM"].tolist(), df

    else:  # .txt
        with open(caminho, encoding="utf-8") as f:
            linhas = [l.strip() for l in f if l.strip()]
        cnpjs = [norm_cnpj(l) for l in linhas]
        log(f"   [OK] {len(cnpjs)} CNPJs no TXT")
        return cnpjs, None


# ==========================================================
# CADASTRO CVM
# ==========================================================
def carregar_cadastro(sem_cache: bool = False, filtro_gestor: Optional[str] = None) -> pd.DataFrame:
    """
    Baixa o cadastro nacional de fundos da CVM.
    Se filtro_gestor for informado, filtra por GESTOR ou ADMIN.
    Sem filtro, retorna fundos da Caixa em funcionamento.
    """
    log("[1] Cadastro CVM nacional")
    dados = baixar_bytes(URL_CAD, "cad_fi.csv", sem_cache)
    if dados is None:
        return pd.DataFrame()

    df = baixar_csv_bytes(dados)
    df.columns = df.columns.str.strip()
    df["_CNPJ_NORM"] = df["CNPJ_FUNDO"].apply(norm_cnpj)

    # Filtro por gestor customizado
    if filtro_gestor:
        termo = filtro_gestor.upper()
        mask = (
            df["ADMIN"].fillna("").str.upper().str.contains(termo) |
            df["GESTOR"].fillna("").str.upper().str.contains(termo)
        )
        df = df[mask].copy()
        log(f"   [OK] {len(df)} fundos com gestor/admin contendo '{filtro_gestor}'")
        return df

    # Padrao: Caixa em funcionamento
    mask = (
        df["ADMIN"].fillna("").str.upper().str.contains("CAIXA") |
        df["GESTOR"].fillna("").str.upper().str.contains("CAIXA")
    ) & df["SIT"].fillna("").str.upper().str.contains("FUNCIONAMENTO")

    df = df[mask].copy()
    log(f"   [OK] {len(df)} fundos Caixa em funcionamento no cadastro CVM")
    return df


# ==========================================================
# INFORME DI_RIO
# ==========================================================
def preparar_inf(df_raw: pd.DataFrame, cnpjs: set) -> pd.DataFrame:
    """Filtra pelos CNPJs alvo e normaliza tipos."""
    if df_raw.empty:
        return pd.DataFrame()

    # Detecta coluna CNPJ (varia entre versoes do informe)
    col_cnpj = next((c for c in df_raw.columns if "cnpj" in c.lower()), None)
    if col_cnpj is None:
        return pd.DataFrame()

    df_raw["_CNPJ_NORM"] = df_raw[col_cnpj].apply(norm_cnpj)
    df = df_raw[df_raw["_CNPJ_NORM"].isin(cnpjs)].copy()

    # Detecta coluna de data
    col_dt = next((c for c in df.columns if c.upper() in ("DT_COMPTC", "DT_REFER")), "DT_COMPTC")
    if col_dt in df.columns:
        df["DT_COMPTC"] = pd.to_datetime(df[col_dt], errors="coerce")

    # Detecta coluna de cota
    col_cota = next((c for c in df.columns if "quota" in c.lower() or "cota" in c.lower()), None)
    if col_cota:
        df["VL_QUOTA"] = pd.to_numeric(df[col_cota], errors="coerce")

    # Detecta coluna de PL
    col_pl = next((c for c in df.columns if "patrim" in c.lower()), None)
    if col_pl:
        df["VL_PATRIM_LIQ"] = pd.to_numeric(df[col_pl], errors="coerce")

    return df.sort_values(["_CNPJ_NORM", "DT_COMPTC"])


# ==========================================================
# DATAS DE REFER_NCIA
# ==========================================================
def mes_anterior(ano: int, mes: int) -> tuple:
    return (ano - 1, 12) if mes == 1 else (ano, mes - 1)


# ==========================================================
# C_LCULO DE INDICADORES
# ==========================================================
def calcular_indicadores(
    cnpj: str,
    nome: str,
    cat: str,
    dt_ini: str,
    df_recente: pd.DataFrame,
    df_dez: pd.DataFrame,
    df_12m: pd.DataFrame,
    hoje: date,
    extras: dict,
) -> dict:
    """Calcula todos os indicadores para um fundo."""
    d = df_recente[df_recente["_CNPJ_NORM"] == cnpj]

    base = {
        "Categoria":              cat,
        "Fundo":                  nome,
        "CNPJ":                   fmt_cnpj(cnpj),
        "Data Inicio":            dt_ini,
        "Cota (R$)":              "",
        "Variacao Dia (%)":       "",
        "Acum. Mes (%)":          "",
        "Acum. Ano (%)":          "",
        "Acum. 12M (%)":          "",
        "PL (milhoes R$)":        "",
        "PL Medio 30d (R$ mi)":   "",
        # Extras do JSON Caixa (preenchidos se disponiveis)
        "Taxa Adm %":             extras.get("taxa_adm", ""),
        "Aberto Captacao":        extras.get("aberto", ""),
        "Aplic. Minima (R$)":     extras.get("aplic_min", ""),
        "Benchmark":              extras.get("benchmark", ""),
        "Resgate Conv.":          extras.get("resgate_conv", ""),
        "Resgate Pgto":           extras.get("resgate_pgto", ""),
        "% CDI Ano":              extras.get("pct_cdi_ano", ""),
        "% CDI 12m":              extras.get("pct_cdi_12m", ""),
    }

    if d.empty or "VL_QUOTA" not in d.columns:
        return base

    d = d.dropna(subset=["VL_QUOTA"])
    if d.empty:
        return base

    ultima     = d.iloc[-1]
    cota_atual = ultima["VL_QUOTA"]
    pl_atual   = (ultima["VL_PATRIM_LIQ"] / 1_000_000
                  if "VL_PATRIM_LIQ" in d.columns and pd.notna(ultima.get("VL_PATRIM_LIQ")) else None)
    pl_medio   = (d["VL_PATRIM_LIQ"].tail(30).mean() / 1_000_000
                  if "VL_PATRIM_LIQ" in d.columns else None)

    # Variacao dia
    var_dia = retorno_pct(d.iloc[-2]["VL_QUOTA"], cota_atual) if len(d) >= 2 else None

    # Acum. Mes _ primeira cota do mes atual
    d_mes   = d[d["DT_COMPTC"].dt.month == hoje.month]
    acum_mes = retorno_pct(
        d_mes.iloc[0]["VL_QUOTA"] if not d_mes.empty else None,
        cota_atual,
    )

    # Acum. Ano _ ultima cota de dez do ano anterior
    d_dez_f  = df_dez[df_dez["_CNPJ_NORM"] == cnpj]
    d_dez_f  = d_dez_f.dropna(subset=["VL_QUOTA"]) if not d_dez_f.empty else d_dez_f
    acum_ano = retorno_pct(
        d_dez_f.iloc[-1]["VL_QUOTA"] if not d_dez_f.empty else None,
        cota_atual,
    )

    # Acum. 12M _ ultima cota de ~12 meses atras
    d_12m_f  = df_12m[df_12m["_CNPJ_NORM"] == cnpj]
    d_12m_f  = d_12m_f.dropna(subset=["VL_QUOTA"]) if not d_12m_f.empty else d_12m_f
    acum_12m = retorno_pct(
        d_12m_f.iloc[-1]["VL_QUOTA"] if not d_12m_f.empty else None,
        cota_atual,
    )

    base.update({
        "Cota (R$)":            fmt_br(cota_atual, 8),
        "Variacao Dia (%)":     fmt_br(var_dia, 3),
        "Acum. Mes (%)":        fmt_br(acum_mes, 2),
        "Acum. Ano (%)":        fmt_br(acum_ano, 2),
        "Acum. 12M (%)":        fmt_br(acum_12m, 2),
        "PL (milhoes R$)":      fmt_br(pl_atual, 3),
        "PL Medio 30d (R$ mi)": fmt_br(pl_medio, 3),
    })
    return base


# ==========================================================
# EXTRACAO PRINCIPAL
# ==========================================================
def extrair(args) -> pd.DataFrame:
    hoje    = date.today()
    sem_cac = args.sem_cache

    print(f"\n{'='*60}")
    print(f"  ROB_ CVM v2.0 _ {hoje.strftime('%d/%m/%Y')}")
    print(f"{'='*60}")

    # -- Fonte dos CNPJs -----------------------------------
    df_extra    = None
    cnpjs_json  = None    # extras do JSON Caixa, indexado por CNPJ_NORM
    caminho_lista = args.lista

    if caminho_lista is None and not getattr(args, "modo_cvm", False):
        caminho_lista = _autodetectar_lista()

    if caminho_lista and os.path.exists(caminho_lista):
        cnpjs_lista, df_extra = ler_lista_cnpjs(caminho_lista)
        cnpjs_set = set(cnpjs_lista)
        # Indexa extras do JSON por CNPJ
        if df_extra is not None and "_CNPJ_NORM" in df_extra.columns:
            cnpjs_json = df_extra.set_index("_CNPJ_NORM")
    else:
        cnpjs_set   = None  # sera preenchido pelo cadastro
        cnpjs_lista = None

    # -- Cadastro CVM -------------------------------------
    log("\n[1/4] Cadastro CVM")
    df_cad = carregar_cadastro(sem_cac, args.filtro_gestor)
    if df_cad.empty:
        print("[ERRO] Cadastro indisponivel. Abortando.")
        return pd.DataFrame()

    # Se nao tem lista externa, usa todos os CNPJs do cadastro
    if cnpjs_set is None:
        cnpjs_set   = set(df_cad["_CNPJ_NORM"].unique())
        cnpjs_lista = list(cnpjs_set)

    # Filtra cadastro pelos CNPJs alvo (merge)
    df_cad_filtrado = df_cad[df_cad["_CNPJ_NORM"].isin(cnpjs_set)].copy()
    log(f"   {len(df_cad_filtrado)} fundos no cadastro para os CNPJs da lista")

    # Adiciona CNPJs da lista que NAO estao no cadastro (fundos novos ou externos)
    cnpjs_nao_cadastro = cnpjs_set - set(df_cad_filtrado["_CNPJ_NORM"])
    if cnpjs_nao_cadastro:
        log(f"   [AVISO]  {len(cnpjs_nao_cadastro)} CNPJs da lista nao encontrados no cadastro CVM")

    # -- Datas de referencia -------------------------------
    ano_ant, mes_ant = mes_anterior(hoje.year, hoje.month)
    ano_12m, mes_12m = hoje.year - 1, hoje.month

    # -- Downloads de cotas --------------------------------
    log(f"\n[2/4] Cotas {hoje.year}{hoje.month:02d} (atual) + {ano_ant}{mes_ant:02d} (mes anterior)")
    df_atual = preparar_inf(baixar_zip_inf(hoje.year, hoje.month, sem_cac), cnpjs_set)
    df_ant   = preparar_inf(baixar_zip_inf(ano_ant, mes_ant, sem_cac), cnpjs_set)
    df_recente = pd.concat([df_ant, df_atual], ignore_index=True)

    log(f"\n[3/4] Dezembro {hoje.year - 1} (base Acum. Ano)")
    df_dez = preparar_inf(baixar_zip_inf(hoje.year - 1, 12, sem_cac), cnpjs_set)

    log(f"\n[4/4] {ano_12m}{mes_12m:02d} (base Acum. 12M)")
    df_12m = preparar_inf(baixar_zip_inf(ano_12m, mes_12m, sem_cac), cnpjs_set)

    # -- Processa cada fundo -------------------------------
    log("\n[calc] Calculando indicadores...")
    resultados = []

    # Mapeia CNPJ -> linha do cadastro
    cad_idx = df_cad_filtrado.set_index("_CNPJ_NORM") if not df_cad_filtrado.empty else pd.DataFrame()

    for cnpj in cnpjs_set:
        # Dados do cadastro CVM
        if cnpj in (cad_idx.index if not cad_idx.empty else []):
            row_cad = cad_idx.loc[cnpj]
            if isinstance(row_cad, pd.DataFrame):
                row_cad = row_cad.iloc[0]
            nome    = str(row_cad.get("DENOM_SOCIAL", "")).strip()
            classe  = str(row_cad.get("CLASSE", ""))
            dt_ini  = str(row_cad.get("DT_INI_ATIV", ""))
            cat     = refinar_categoria(classe, nome)
        elif cnpjs_json is not None and cnpj in cnpjs_json.index:
            row_j   = cnpjs_json.loc[cnpj]
            nome    = str(row_j.get("no_fundo", row_j.get("no_razao_social", cnpj))).strip()
            classe  = str(row_j.get("no_classificacao_cvm", ""))
            dt_ini  = str(row_j.get("dt_inicial", ""))
            cat     = refinar_categoria(classe, nome)
        else:
            nome, classe, dt_ini, cat = cnpj, "", "", "OUTROS"

        # Dados extras do JSON Caixa
        extras = {}
        if cnpjs_json is not None and cnpj in cnpjs_json.index:
            row_j = cnpjs_json.loc[cnpj]
            if isinstance(row_j, pd.DataFrame):
                row_j = row_j.iloc[0]
            extras = {
                "taxa_adm":    fmt_br(row_j.get("pc_taxa_adm_cliente"), 2),
                "aberto":      "Sim" if row_j.get("ic_aberto_captacao") else "Nao",
                "aplic_min":   fmt_br(row_j.get("vr_aplicacao_inicial"), 2),
                "benchmark":   str(row_j.get("no_benchmark", "")),
                "resgate_conv": str(row_j.get("de_conversao_resgate", "")),
                "resgate_pgto": str(row_j.get("de_pagamento_resgate", "")),
                "pct_cdi_ano": fmt_br(row_j.get("pc_rentabilidade_cdi_ano"), 2),
                "pct_cdi_12m": fmt_br(row_j.get("pc_rentabilidade_cdi_12m"), 2),
            }

        # Filtra apenas abertos (se flag ativa)
        if args.apenas_abertos and extras.get("aberto") == "Nao":
            continue

        resultados.append(calcular_indicadores(
            cnpj, nome, cat, dt_ini,
            df_recente, df_dez, df_12m, hoje, extras,
        ))

    df_out = pd.DataFrame(resultados)
    if not df_out.empty:
        df_out = df_out.sort_values(["Categoria", "Fundo"])
    log(f"   [OK] {len(df_out)} fundos processados")
    return df_out


# ==========================================================
# VALIDACAO
# ==========================================================
def validar(df: pd.DataFrame):
    print(f"\n{'-'*50}")
    print("[stats] Resumo de qualidade:")
    total     = len(df)
    com_cota  = df["Cota (R$)"].ne("").sum()
    com_12m   = df["Acum. 12M (%)"].ne("").sum()
    com_ano   = df["Acum. Ano (%)"].ne("").sum()
    print(f"   Total fundos    : {total}")
    print(f"   Com cota atual  : {com_cota}  ({com_cota/total*100:.0f}%)")
    print(f"   Com Acum. Ano   : {com_ano}  ({com_ano/total*100:.0f}%)")
    print(f"   Com Acum. 12M   : {com_12m}  ({com_12m/total*100:.0f}%)")
    print(f"\n   Por categoria:")
    for cat, grp in df.groupby("Categoria"):
        c12 = grp["Acum. 12M (%)"].ne("").sum()
        print(f"   {cat:<35}: {len(grp):>3} fundos  |  {c12} com 12M")
    print(f"{'-'*50}")


# ==========================================================
# GERACAO DO HTML
# ==========================================================
def gerar_html(df: pd.DataFrame, caminho: str):
    hoje     = date.today().strftime("%d/%m/%Y")
    total    = len(df)
    com_cota = df["Cota (R$)"].ne("").sum()

    # Colunas a exibir (omite as vazias)
    colunas_mostrar = [c for c in df.columns if c != "CNPJ" and df[c].ne("").any()]
    cabecalhos = "".join(f'<th onclick="sortTable(this)">{c}</th>' for c in colunas_mostrar)

    pct_cols = {"Variacao Dia (%)", "Acum. Mes (%)", "Acum. Ano (%)",
                "Acum. 12M (%)", "Taxa Adm %", "% CDI Ano", "% CDI 12m"}

    def _celula(col, val):
        v = str(val).strip()
        if v == "" or v == "nan":
            return "<td>_</td>"
        if col in pct_cols:
            try:
                fv = float(v.replace(",", ".").replace("%", ""))
                cls = "pos" if fv >= 0 else "neg"
                return f'<td class="{cls}">{v}</td>'
            except Exception:
                pass
        return f"<td>{v}</td>"

    linhas_html = ""
    for _, row in df.iterrows():
        cells = "".join(_celula(c, row[c]) for c in colunas_mostrar)
        linhas_html += f"<tr>{cells}</tr>\n"

    html = f"""<!DOCTYPE html>
<html lang="pt-BR">
<head>
<meta charset="UTF-8">
<title>Catalogo Fundos Caixa _ CVM</title>
<style>
*,*::before,*::after{{box-sizing:border-box;margin:0;padding:0}}
body{{font-family:"Segoe UI",Arial,sans-serif;background:#0f172a;color:#e5e7eb;padding:20px}}
h1{{color:#fff;font-size:1.5rem;margin-bottom:4px}}
.sub{{color:#64748b;font-size:13px;margin-bottom:20px}}
.cards{{display:grid;grid-template-columns:repeat(auto-fit,minmax(150px,1fr));gap:12px;margin-bottom:20px}}
.card{{background:#111827;border:1px solid #1e293b;border-radius:12px;padding:14px 16px}}
.card span{{color:#64748b;font-size:11px;text-transform:uppercase;letter-spacing:.05em}}
.card strong{{display:block;font-size:22px;margin-top:4px;color:#fff}}
.box{{background:#111827;border:1px solid #1e293b;border-radius:12px;padding:18px;margin-bottom:20px;overflow-x:auto}}
.toolbar{{display:flex;gap:10px;align-items:center;margin-bottom:12px;flex-wrap:wrap}}
input[type=text]{{flex:1;min-width:200px;padding:9px 13px;border-radius:9px;border:1px solid #334155;background:#020617;color:#fff;font-size:13px}}
select{{padding:9px 13px;border-radius:9px;border:1px solid #334155;background:#020617;color:#cbd5e1;font-size:13px}}
.btn{{background:#1e293b;color:#cbd5e1;border:1px solid #334155;border-radius:8px;padding:8px 14px;cursor:pointer;font-size:12px}}
.btn:hover{{background:#334155}}
.pag{{color:#64748b;font-size:12px}}
table{{width:100%;border-collapse:collapse;font-size:12px}}
th{{background:#1e293b;color:#94a3b8;text-align:left;padding:8px 10px;cursor:pointer;user-select:none;white-space:nowrap}}
th:hover{{color:#fff}}
th::after{{content:" _";opacity:.35;font-size:10px}}
td{{padding:7px 10px;border-bottom:1px solid #1e293b;white-space:nowrap}}
tr:hover td{{background:#1f2937}}
.pos{{color:#22c55e;font-weight:600}}
.neg{{color:#ef4444;font-weight:600}}
.footer{{color:#475569;font-size:11px;margin-top:20px;text-align:center}}
</style>
</head>
<body>
<h1>Catalogo de Fundos Caixa _ CVM</h1>
<p class="sub">Dados: dados.cvm.gov.br &nbsp;_&nbsp; Gerado em {hoje} &nbsp;_&nbsp; github.com/eltonprivatebanker/catalogo-fundos-cvm</p>

<div class="cards">
  <div class="card"><span>Total de Fundos</span><strong>{total}</strong></div>
  <div class="card"><span>Com Cota Atual</span><strong class="pos">{com_cota}</strong></div>
  <div class="card"><span>Sem Dados</span><strong>{total - com_cota}</strong></div>
  <div class="card"><span>Atualizado</span><strong style="font-size:14px">{hoje}</strong></div>
</div>

<div class="box">
  <div class="toolbar">
    <input type="text" id="busca" placeholder="Buscar por nome, CNPJ, categoria..." oninput="filtrar()">
    <select id="filtroCat" onchange="filtrar()">
      <option value="">Todas as categorias</option>
      {''.join(f'<option value="{c}">{c}</option>' for c in sorted(df["Categoria"].unique()))}
    </select>
    <button class="btn" onclick="mudarPag(-1)">_ Anterior</button>
    <span class="pag" id="pagInfo"></span>
    <button class="btn" onclick="mudarPag(1)">Proxima -></button>
  </div>
  <table id="tbl">
    <thead><tr>{cabecalhos}</tr></thead>
    <tbody id="tbody">{linhas_html}</tbody>
  </table>
</div>

<div class="footer">
  Fonte: CVM Dados Abertos _ Calculo por variacao de cota _ Robo CVM v2.0
</div>

<script>
var POR_PAG = 50, pag = 0, linhasFiltradas = [];

function filtrar() {{
  var t = document.getElementById("busca").value.toLowerCase();
  var cat = document.getElementById("filtroCat").value.toLowerCase();
  var rows = Array.from(document.querySelectorAll("#tbody tr"));
  linhasFiltradas = rows.filter(function(r) {{
    return r.innerText.toLowerCase().includes(t) &&
           (cat === "" || r.innerText.toLowerCase().includes(cat));
  }});
  pag = 0; paginar();
}}

function paginar() {{
  var rows = Array.from(document.querySelectorAll("#tbody tr"));
  var ativas = linhasFiltradas.length || document.getElementById("busca").value || document.getElementById("filtroCat").value
    # linhasFiltradas : rows;
  var total = ativas.length;
  var totalPag = Math.max(1, Math.ceil(total / POR_PAG));
  if (pag >= totalPag) pag = totalPag - 1;
  var set = new Set(ativas.slice(pag * POR_PAG, (pag + 1) * POR_PAG));
  rows.forEach(function(r) {{ r.style.display = set.has(r) _ "" : "none"; }});
  document.getElementById("pagInfo").textContent =
    "Pag " + (pag+1) + "/" + totalPag + " _ " + total.toLocaleString("pt-BR") + " fundos";
}}

function mudarPag(d) {{ pag = Math.max(0, pag + d); paginar(); }}

function sortTable(th) {{
  var idx = Array.from(th.closest("tr").cells).indexOf(th);
  var tbody = document.getElementById("tbody");
  var rows = Array.from(tbody.rows);
  var asc = th.dataset.asc !== "1";
  th.dataset.asc = asc _ "1" : "0";
  rows.sort(function(a,b) {{
    var va = a.cells[idx]_.innerText.trim() || "";
    var vb = b.cells[idx]_.innerText.trim() || "";
    var na = parseFloat(va.replace(/[^0-9.,-]/g,"").replace(",","."));
    var nb = parseFloat(vb.replace(/[^0-9.,-]/g,"").replace(",","."));
    if (!isNaN(na) && !isNaN(nb)) return asc _ na-nb : nb-na;
    return asc _ va.localeCompare(vb,"pt-BR") : vb.localeCompare(va,"pt-BR");
  }});
  rows.forEach(function(r) {{ tbody.appendChild(r); }});
  paginar();
}}

filtrar();
</script>
</body>
</html>"""

    with open(caminho, "w", encoding="utf-8") as f:
        f.write(html)
    log(f"[OK] HTML: {caminho}")


# ==========================================================
# SALVAMENTO
# ==========================================================
def salvar(df: pd.DataFrame, pasta: str):
    if df.empty:
        print("[ERRO] Nenhum dado para salvar.")
        return

    validar(df)
    os.makedirs(pasta, exist_ok=True)
    ts = datetime.now().strftime("%Y%m%d")

    # Remove CNPJ das colunas de saida (e interno)
    cols = [c for c in df.columns if c != "CNPJ"]
    df_out = df[cols]

    # CSV
    p_atual = os.path.join(pasta, "dados_atuais.csv")
    p_hist  = os.path.join(pasta, f"sipii_caixa_{ts}.csv")
    df_out.to_csv(p_atual, index=False, encoding="utf-8-sig")
    df_out.to_csv(p_hist,  index=False, encoding="utf-8-sig")

    # Excel _ aba por categoria
    p_xlsx = os.path.join(pasta, f"sipii_caixa_{ts}.xlsx")
    try:
        with pd.ExcelWriter(p_xlsx, engine="openpyxl") as writer:
            df_out.to_excel(writer, sheet_name="Todos", index=False)
            for cat, grp in df_out.groupby("Categoria"):
                grp.to_excel(writer, sheet_name=limpar_aba(cat), index=False)
        log(f"[OK] Excel: {p_xlsx}")
    except ImportError:
        log("[AVISO]  openpyxl nao instalado. pip install openpyxl")

    # HTML
    p_html = os.path.join(pasta, f"catalogo_caixa_{ts}.html")
    gerar_html(df, p_html)

    print(f"\n{'='*60}")
    print("[OK] CONCLU_DO!")
    print(f"   -> {p_atual}")
    print(f"   -> {p_hist}")
    print(f"   -> {p_xlsx}")
    print(f"   -> {p_html}")
    print(f"{'='*60}\n")


# ==========================================================
# ENTRADA
# ==========================================================
if __name__ == "__main__":
    args = parse_args()
    df   = extrair(args)
    salvar(df, args.saida)
