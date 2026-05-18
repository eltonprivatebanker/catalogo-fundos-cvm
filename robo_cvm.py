"""
============================================================
ROBO CVM v3.0 - Catalogo de Fundos Caixa
============================================================
Fonte   : dados.cvm.gov.br  (API oficial, gratuita)
GitHub  : eltonprivatebanker/catalogo-fundos-cvm

Novidades v3.0:
  - Deteccao automatica do formato do cadastro CVM (v1/v2/v3)
  - Suporte a Resolucao CVM 175 (novo formato 2024+)
  - Busca Caixa por nome E por CNPJ do administrador
  - Salva periodos.json para o HTML saber quais arquivos existem
  - Filtros avancados no HTML (categoria, risco, rentabilidade minima)

Uso:
  python robo_cvm.py                      # auto-detecta formato
  python robo_cvm.py --sem-cache          # forca redownload
  python robo_cvm.py --lista fundos.json  # usa lista de CNPJs
  python robo_cvm.py --saida ./saida      # pasta de saida
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
CACHE_DIR   = ".cache_cvm"
TIMEOUT     = 90
MAX_RETRY   = 3

# CNPJs conhecidos da Caixa como administradora (para busca no formato v3)
CAIXA_CNPJS = [
    "00360305",  # Caixa Economica Federal
    "36113876",  # Caixa DTVM S.A.
    "45756448",  # Caixa Asset Management
]

# Mapeamento de classes CVM para categorias leg-veis
CLASSE_MAP = {
    "Fundo de Renda Fixa":                "RENDA FIXA",
    "Fundo de Acoes":                     "ACOES",
    "Fundo Multimercado":                 "MULTIMERCADO",
    "Fundo Cambial":                      "CAMBIAL",
    "Fundo de Indice de Mercado":         "FUNDO DE INDICE",
    "Fundo Mutuo de Privatizacao - FGTS": "PRIVATIZACAO",
    "Fundo de Investimento Imobiliario":  "FII",
    "Renda Fixa":                         "RENDA FIXA",
    "Acoes":                              "ACOES",
    "Multimercado":                       "MULTIMERCADO",
}

# ==========================================================
# SCHEMAS DO CADASTRO CVM POR VERSAO
# ==========================================================
# A CVM muda o formato do cad_fi.csv a cada resolucao normativa.
# Este mapeamento normaliza os nomes de coluna para o codigo usar
# sempre os mesmos nomes internos.

SCHEMAS = {
    "v1": {
        # Formato ate 2022
        "cnpj":   "CNPJ_FUNDO",
        "nome":   "DENOM_SOCIAL",
        "sit":    "SIT",
        "admin":  "ADMIN",        # texto: nome do administrador
        "gestor": "GESTOR",       # texto: nome do gestor
        "classe": "CLASSE",
        "dt_ini": "DT_INI_ATIV",
    },
    "v2": {
        # Formato 2023-2024 (Resolucao CVM 175 - transicao)
        "cnpj":   "CNPJ_FUNDO",
        "nome":   "DENOM_SOCIAL",
        "sit":    "SIT",
        "admin":  "ADMIN",
        "gestor": "GESTOR",
        "classe": "CLASSE",
        "dt_ini": "DT_INI_ATIV",
    },
    "v3": {
        # Formato 2025+ (Resolucao CVM 175 - completo)
        # ADMIN e GESTOR agora sao CNPJs/flags, nao nomes
        # A busca por "CAIXA" deve ser em DENOM_SOCIAL ou CNPJ_ADMIN
        "cnpj":   "CNPJ_FUNDO",
        "nome":   "DENOM_SOCIAL",
        "sit":    "SIT",
        "admin":  "CNPJ_ADMIN",   # CNPJ numerico, nao nome
        "gestor": "PF_PJ_GESTOR", # S ou N, nao nome
        "classe": "TP_FUNDO",
        "dt_ini": "DT_REG",
    },
}


def detectar_schema(colunas: list) -> dict:
    cols_upper = [c.upper() for c in colunas]
    # v3 tem CNPJ_ADMIN E ainda ADMIN (texto) simultaneamente
    if "CNPJ_ADMIN" in cols_upper and "ADMIN" in cols_upper:
        versao = "v3-compat (2025+ com ADMIN texto)"
    elif "ADMIN" in cols_upper and "GESTOR" in cols_upper:
        versao = "v1/v2 (ate 2024)"
    else:
        versao = "v2-parcial"
    # Sempre usa v1 schema se ADMIN texto existir
    schema = SCHEMAS["v1"] if "ADMIN" in cols_upper else SCHEMAS["v2"]
    log(f"   Schema: {versao}")
    return schema


# ==========================================================
# CLI
# ==========================================================
def parse_args():
    p = argparse.ArgumentParser(description="Robo CVM v3.0 - Catalogo de Fundos Caixa")
    p.add_argument("--lista",          default=None,
                   help="Arquivo com CNPJs: JSON (Caixa), CSV ou TXT")
    p.add_argument("--sem-cache",      action="store_true", dest="sem_cache",
                   help="Forca redownload de todos os arquivos")
    p.add_argument("--apenas-abertos", action="store_true", dest="apenas_abertos",
                   help="Inclui so fundos abertos para captacao")
    p.add_argument("--saida",          default=".", dest="saida",
                   help="Pasta de saida (default: pasta atual)")
    p.add_argument("--filtro-gestor",  default=None, dest="filtro_gestor",
                   help="Busca por gestor/admin (ex: 'caixa', 'btg', 'xp')")
    p.add_argument("--modo-cvm",       action="store_true", dest="modo_cvm",
                   help="Ignora JSON local, busca direto no cadastro CVM")
    return p.parse_args()


# ==========================================================
# UTILITARIOS
# ==========================================================
def norm_cnpj(c: str) -> str:
    return re.sub(r"\D", "", str(c)).zfill(14)

def fmt_cnpj(c: str) -> str:
    c = norm_cnpj(c)
    return f"{c[:2]}.{c[2:5]}.{c[5:8]}/{c[8:12]}-{c[12:14]}"

def rm_accent(s: str) -> str:
    return "".join(ch for ch in unicodedata.normalize("NFD", s)
                   if unicodedata.category(ch) != "Mn")

def fmt_br(v, decimais: int = 2) -> str:
    if v is None or (isinstance(v, float) and pd.isna(v)):
        return ""
    try:
        return f"{float(v):,.{decimais}f}".replace(",", "X").replace(".", ",").replace("X", ".")
    except Exception:
        return ""

def retorno_pct(cota_ini, cota_fim) -> Optional[float]:
    try:
        if cota_ini and cota_fim and float(cota_ini) > 0:
            return (float(cota_fim) / float(cota_ini) - 1) * 100
    except (TypeError, ValueError):
        pass
    return None

def limpar_aba(nome: str) -> str:
    return re.sub(r'[:\\/*-\[\]]', "", nome)[:31]

def log(msg: str):
    print(f"{datetime.now().strftime('%H:%M:%S')}  {msg}")

def refinar_categoria(classe_cvm: str, nome_fundo: str) -> str:
    nome = rm_accent(str(nome_fundo).upper())
    if "SIMPLES" in nome:
        return "RENDA FIXA SIMPLES"
    if any(x in nome for x in ("REF DI", "REFERENCIADO", "REFERENC")):
        return "RENDA FIXA REF DI"
    if "CURTO PRAZO" in nome and "renda" in classe_cvm.lower():
        return "RENDA FIXA CURTO PRAZO"
    if "DEBENTURE" in nome or "DEB INCENT" in nome:
        return "RENDA FIXA DEBENTURES"
    if "CREDITO" in nome or "CRED PRIV" in nome:
        return "RENDA FIXA CRED PRIV"
    if "FGTS" in nome or "PRIVATIZ" in nome:
        return "PRIVATIZACAO"
    if "IMOB" in nome or " FII" in nome:
        return "FII"
    return CLASSE_MAP.get(classe_cvm, rm_accent(classe_cvm.upper()) if classe_cvm else "OUTROS")


# ==========================================================
# DOWNLOAD COM CACHE E RETRY
# ==========================================================
def _cache_path(url: str) -> str:
    nome = re.sub(r"[^\w.]", "_", url.split("/")[-1])
    return os.path.join(CACHE_DIR, nome)

def baixar_bytes(url: str, label: str = "", sem_cache: bool = False) -> Optional[bytes]:
    cache = _cache_path(url)
    if not sem_cache and os.path.exists(cache):
        log(f"[cache] {label or os.path.basename(cache)}")
        with open(cache, "rb") as f:
            return f.read()
    for tentativa in range(MAX_RETRY):
        try:
            log(f"[download] {label or url}")
            r = requests.get(url, timeout=TIMEOUT, headers={"User-Agent": "Mozilla/5.0"})
            r.raise_for_status()
            os.makedirs(CACHE_DIR, exist_ok=True)
            with open(cache, "wb") as f:
                f.write(r.content)
            log(f"[OK] Salvo ({len(r.content)/1e6:.1f} MB)")
            return r.content
        except Exception as e:
            if tentativa == MAX_RETRY - 1:
                log(f"[ERRO] Falhou apos {MAX_RETRY} tentativas: {e}")
                return None
            log(f"[AVISO] Tentativa {tentativa+1} falhou, retry em 3s...")
            time.sleep(3)
    return None

def baixar_csv_bytes(conteudo: bytes) -> pd.DataFrame:
    for enc in ("latin1", "utf-8-sig", "utf-8"):
        try:
            return pd.read_csv(
                io.StringIO(conteudo.decode(enc)),
                sep=";", low_memory=False, dtype=str
            )
        except Exception:
            continue
    return pd.DataFrame()

def baixar_zip_inf(ano: int, mes: int, sem_cache: bool = False) -> pd.DataFrame:
    yyyymm = f"{ano}{mes:02d}"
    url    = URL_INF_MES.format(yyyymm=yyyymm)
    dados  = baixar_bytes(url, f"informe {yyyymm}", sem_cache)
    if dados is None:
        return pd.DataFrame()
    try:
        with zipfile.ZipFile(BytesIO(dados)) as zf:
            csvs = [n for n in zf.namelist() if n.lower().endswith(".csv")]
            if not csvs:
                return pd.DataFrame()
            with zf.open(csvs[0]) as f:
                return pd.read_csv(f, sep=";", encoding="latin1",
                                   low_memory=False, dtype=str)
    except Exception as e:
        log(f"[ERRO] ZIP {yyyymm}: {e}")
        return pd.DataFrame()


# ==========================================================
# LEITURA DA LISTA DE CNPJs
# ==========================================================
def _autodetectar_lista() -> Optional[str]:
    for ext in ("*.json", "*.csv", "*.txt"):
        arquivos = sorted(glob.glob(ext))
        arquivos = [a for a in arquivos if not any(
            x in a.lower() for x in
            ("dados_atuais", "sipii_caixa", "cnpjs_lista", "requirements", "periodos")
        )]
        if arquivos:
            log(f"[arquivo] Detectado: {arquivos[0]}")
            return arquivos[0]
    return None

def ler_lista_cnpjs(caminho: str) -> tuple:
    ext = os.path.splitext(caminho)[1].lower()
    log(f"Lendo lista: {caminho}")
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
        log(f"   [OK] {len(df)} fundos (coluna '{col}')")
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
    else:
        with open(caminho, encoding="utf-8") as f:
            linhas = [l.strip() for l in f if l.strip()]
        cnpjs = [norm_cnpj(l) for l in linhas]
        log(f"   [OK] {len(cnpjs)} CNPJs no TXT")
        return cnpjs, None


# ==========================================================
# CADASTRO CVM - COM DETECCAO DE SCHEMA
# ==========================================================
def carregar_cadastro(sem_cache: bool = False,
                      filtro_gestor: Optional[str] = None) -> pd.DataFrame:
    """
    Baixa o cadastro CVM e filtra pelos fundos da Caixa (ou outro gestor).
    Detecta automaticamente o formato do arquivo (v1/v2/v3).
    """
    log("[1/4] Cadastro CVM nacional")
    dados = baixar_bytes(URL_CAD, "cad_fi.csv", sem_cache)
    if dados is None:
        return pd.DataFrame()

    df = baixar_csv_bytes(dados)
    df.columns = df.columns.str.strip()
    log(f"   Total no cadastro: {len(df)} fundos")
    log(f"   Colunas: {list(df.columns)}")

    # Detecta schema e normaliza nomes de colunas
    schema = detectar_schema(list(df.columns))
    rename = {}
    for interno, externo in schema.items():
        if externo in df.columns and externo != interno.upper():
            rename[externo] = interno.upper()
    # Garante nomes padrao internos
    col_map = {}
    for interno, externo in schema.items():
        src = rename.get(externo, externo)
        col_map[interno] = src if src in df.columns else None

    log(f"   Mapeamento: {col_map}")

    # CNPJ normalizado
    col_cnpj = col_map.get("cnpj") or next(
        (c for c in df.columns if "cnpj" in c.lower() and "fundo" in c.lower()), None
    )
    if col_cnpj is None:
        log("[ERRO] Coluna CNPJ nao encontrada!")
        return pd.DataFrame()
    df["_CNPJ_NORM"] = df[col_cnpj].apply(norm_cnpj)

    # Padroniza colunas essenciais
    for interno, externo in schema.items():
        dst = interno.upper()
        if externo in df.columns and dst not in df.columns:
            df[dst] = df[externo]

    termo = (filtro_gestor or "CAIXA").upper()

    # -- Estrategia 1: busca por nome (DENOM_SOCIAL) ------
    mask = pd.Series([False] * len(df), index=df.index)

    if "DENOM_SOCIAL" in df.columns:
        mask |= df["DENOM_SOCIAL"].fillna("").str.upper().str.contains(termo, regex=False)
        log(f"   Por DENOM_SOCIAL ('{termo}'): {mask.sum()} fundos")

    # Busca em colunas de texto adicionais (formato antigo: ADMIN, GESTOR)
    for col in ("ADMIN", "GESTOR"):
        if col in df.columns and df[col].dtype == object:
            mask |= df[col].fillna("").str.upper().str.contains(termo, regex=False)
            log(f"   Por {col}: {mask.sum()} fundos acumulados")

    # -- Estrategia 2: busca por CNPJ do administrador ---
    # (formato v3: CNPJ_ADMIN contem o CNPJ da instituicao)
    if mask.sum() == 0 or schema == SCHEMAS["v3"]:
        for cnpj_base in CAIXA_CNPJS:
            for col in df.columns:
                if "cnpj" in col.lower() and "fundo" not in col.lower():
                    extra = df[col].fillna("").apply(
                        lambda x: re.sub(r"\D", "", str(x)).startswith(cnpj_base)
                    )
                    mask |= extra
        log(f"   Por CNPJ admin (Caixa): {mask.sum()} fundos acumulados")

    # Filtro de situacao (apenas fundos em funcionamento)
    if "SIT" in df.columns:
        vals_sit = df["SIT"].dropna().unique()
        log(f"   Valores em SIT: {list(vals_sit[:6])}")
        # Exclui apenas fundos explicitamente cancelados/liquidados
        mask_ruim = df["SIT"].fillna("").str.upper().str.contains(
            "CANCEL|LIQUID|ENCERR", regex=True
        )
        if not filtro_gestor:
            antes = mask.sum()
            mask &= ~mask_ruim
            log(f"   Apos excluir cancelados: {mask.sum()} (antes: {antes})")

    df_out = df[mask].copy()
    log(f"   [OK] {len(df_out)} fundos Caixa encontrados")

    if df_out.empty:
        log("[AVISO] Nenhum fundo encontrado! Verifique o formato do cadastro.")
        log(f"   Todas as colunas: {list(df.columns)}")
        amostra = df["DENOM_SOCIAL"].dropna().unique()[:5].tolist() if "DENOM_SOCIAL" in df.columns else []
        log(f"   Amostra DENOM_SOCIAL: {amostra}")

    return df_out


# ==========================================================
# INFORME DIARIO
# ==========================================================
def preparar_inf(df_raw: pd.DataFrame, cnpjs: set) -> pd.DataFrame:
    if df_raw.empty:
        return pd.DataFrame()
    col_cnpj = next((c for c in df_raw.columns if "cnpj" in c.lower()), None)
    if col_cnpj is None:
        return pd.DataFrame()
    df_raw = df_raw.copy()
    df_raw["_CNPJ_NORM"] = df_raw[col_cnpj].apply(norm_cnpj)
    df = df_raw[df_raw["_CNPJ_NORM"].isin(cnpjs)].copy()
    # Normaliza colunas
    for kw, dst in [("quota|cota", "VL_QUOTA"), ("patrim", "VL_PATRIM_LIQ"),
                    ("dt_comptc|dt_refer", "DT_COMPTC"), ("nr_cotst", "NR_COTST")]:
        col = next((c for c in df.columns if re.search(kw, c.lower())), None)
        if col and dst not in df.columns:
            df[dst] = df[col]
    if "DT_COMPTC" in df.columns:
        df["DT_COMPTC"] = pd.to_datetime(df["DT_COMPTC"], errors="coerce")
    if "VL_QUOTA" in df.columns:
        df["VL_QUOTA"] = pd.to_numeric(df["VL_QUOTA"], errors="coerce")
    if "VL_PATRIM_LIQ" in df.columns:
        df["VL_PATRIM_LIQ"] = pd.to_numeric(df["VL_PATRIM_LIQ"], errors="coerce")
    return df.sort_values(["_CNPJ_NORM", "DT_COMPTC"])


# ==========================================================
# DATAS
# ==========================================================
def mes_anterior(ano: int, mes: int) -> tuple:
    return (ano - 1, 12) if mes == 1 else (ano, mes - 1)


# ==========================================================
# CALCULO DE INDICADORES
# ==========================================================
def calcular_indicadores(cnpj, nome, cat, dt_ini,
                         df_recente, df_dez, df_12m,
                         hoje, extras) -> dict:
    d = df_recente[df_recente["_CNPJ_NORM"] == cnpj] if not df_recente.empty else pd.DataFrame()

    base = {
        "Categoria":            cat,
        "Fundo":                nome,
        "CNPJ":                 fmt_cnpj(cnpj),
        "Data Inicio":          dt_ini,
        "Cota (R$)":            "",
        "Variacao Dia (%)":     "",
        "Acum. Mes (%)":        "",
        "Acum. Ano (%)":        "",
        "Acum. 12M (%)":        "",
        "PL (mi R$)":           "",
        "PL Medio 30d (mi R$)": "",
        "Taxa Adm %":           extras.get("taxa_adm", ""),
        "Aberto":               extras.get("aberto", ""),
        "Aplic. Min (R$)":      extras.get("aplic_min", ""),
        "Benchmark":            extras.get("benchmark", ""),
        "Resgate Conv.":        extras.get("resgate_conv", ""),
        "Resgate Pgto":         extras.get("resgate_pgto", ""),
        "% CDI Ano":            extras.get("pct_cdi_ano", ""),
        "% CDI 12m":            extras.get("pct_cdi_12m", ""),
    }

    if d.empty or "VL_QUOTA" not in d.columns:
        return base
    d = d.dropna(subset=["VL_QUOTA"])
    if d.empty:
        return base

    ultima     = d.iloc[-1]
    cota_atual = ultima["VL_QUOTA"]
    pl_atual   = (ultima["VL_PATRIM_LIQ"] / 1e6
                  if "VL_PATRIM_LIQ" in d.columns and pd.notna(ultima.get("VL_PATRIM_LIQ"))
                  else None)
    pl_medio   = (d["VL_PATRIM_LIQ"].tail(30).mean() / 1e6
                  if "VL_PATRIM_LIQ" in d.columns else None)

    var_dia = retorno_pct(d.iloc[-2]["VL_QUOTA"], cota_atual) if len(d) >= 2 else None

    d_mes    = d[d["DT_COMPTC"].dt.month == hoje.month] if "DT_COMPTC" in d.columns else pd.DataFrame()
    acum_mes = retorno_pct(d_mes.iloc[0]["VL_QUOTA"] if not d_mes.empty else None, cota_atual)

    d_dez_f  = df_dez[df_dez["_CNPJ_NORM"] == cnpj].dropna(subset=["VL_QUOTA"]) if not df_dez.empty else pd.DataFrame()
    acum_ano = retorno_pct(d_dez_f.iloc[-1]["VL_QUOTA"] if not d_dez_f.empty else None, cota_atual)

    d_12m_f  = df_12m[df_12m["_CNPJ_NORM"] == cnpj].dropna(subset=["VL_QUOTA"]) if not df_12m.empty else pd.DataFrame()
    acum_12m = retorno_pct(d_12m_f.iloc[-1]["VL_QUOTA"] if not d_12m_f.empty else None, cota_atual)

    base.update({
        "Cota (R$)":            fmt_br(cota_atual, 8),
        "Variacao Dia (%)":     fmt_br(var_dia, 3),
        "Acum. Mes (%)":        fmt_br(acum_mes, 2),
        "Acum. Ano (%)":        fmt_br(acum_ano, 2),
        "Acum. 12M (%)":        fmt_br(acum_12m, 2),
        "PL (mi R$)":           fmt_br(pl_atual, 3),
        "PL Medio 30d (mi R$)": fmt_br(pl_medio, 3),
    })
    return base


# ==========================================================
# EXTRACAO PRINCIPAL
# ==========================================================
def extrair(args) -> pd.DataFrame:
    hoje    = date.today()
    sem_cac = args.sem_cache

    print(f"\n{'='*60}")
    print(f"  ROBO CVM v3.0 - {hoje.strftime('%d/%m/%Y')}")
    print(f"{'='*60}")

    # -- Fonte dos CNPJs -----------------------------------
    df_extra   = None
    cnpjs_json = None
    caminho    = args.lista

    if caminho is None and not getattr(args, "modo_cvm", False):
        caminho = _autodetectar_lista()

    if caminho and os.path.exists(caminho):
        cnpjs_lista, df_extra = ler_lista_cnpjs(caminho)
        cnpjs_set = set(cnpjs_lista)
        if df_extra is not None and "_CNPJ_NORM" in df_extra.columns:
            cnpjs_json = df_extra.set_index("_CNPJ_NORM")
    else:
        cnpjs_set = None

    # -- Cadastro CVM --------------------------------------
    df_cad = carregar_cadastro(sem_cac, args.filtro_gestor)
    if df_cad.empty:
        print("[ERRO] Cadastro vazio. Abortando.")
        return pd.DataFrame()

    if cnpjs_set is None:
        cnpjs_set = set(df_cad["_CNPJ_NORM"].unique())

    df_cad_f = df_cad[df_cad["_CNPJ_NORM"].isin(cnpjs_set)].copy()
    log(f"   {len(df_cad_f)} fundos no cadastro para processar")

    # -- Downloads de cotas --------------------------------
    ano_ant, mes_ant = mes_anterior(hoje.year, hoje.month)
    ano_12m, mes_12m = hoje.year - 1, hoje.month

    log(f"\n[2/4] Cotas {hoje.year}{hoje.month:02d} + {ano_ant}{mes_ant:02d}")
    df_atual   = preparar_inf(baixar_zip_inf(hoje.year, hoje.month, sem_cac), cnpjs_set)
    df_ant     = preparar_inf(baixar_zip_inf(ano_ant, mes_ant, sem_cac), cnpjs_set)
    df_recente = pd.concat([df_ant, df_atual], ignore_index=True)
    log(f"   Linhas carregadas: {len(df_recente)}")

    log(f"\n[3/4] Dez/{hoje.year-1} (Acum. Ano)")
    df_dez = preparar_inf(baixar_zip_inf(hoje.year - 1, 12, sem_cac), cnpjs_set)

    log(f"\n[4/4] {ano_12m}{mes_12m:02d} (Acum. 12M)")
    df_12m = preparar_inf(baixar_zip_inf(ano_12m, mes_12m, sem_cac), cnpjs_set)

    # -- Calcula indicadores -------------------------------
    log("\n[calc] Calculando indicadores...")
    cad_idx = df_cad_f.set_index("_CNPJ_NORM") if not df_cad_f.empty else pd.DataFrame()
    resultados = []

    for cnpj in cnpjs_set:
        # Dados do cadastro
        if not cad_idx.empty and cnpj in cad_idx.index:
            row = cad_idx.loc[cnpj]
            if isinstance(row, pd.DataFrame):
                row = row.iloc[0]
            nome   = str(row.get("DENOM_SOCIAL", row.get("NOME", ""))).strip()
            classe = str(row.get("CLASSE", row.get("TP_FUNDO", "")))
            dt_ini = str(row.get("DT_INI_ATIV", row.get("DT_REG", "")))
            cat    = refinar_categoria(classe, nome)
        elif cnpjs_json is not None and cnpj in cnpjs_json.index:
            row_j  = cnpjs_json.loc[cnpj]
            if isinstance(row_j, pd.DataFrame):
                row_j = row_j.iloc[0]
            nome   = str(row_j.get("no_fundo", cnpj)).strip()
            classe = str(row_j.get("no_classificacao_cvm", ""))
            dt_ini = str(row_j.get("dt_inicial", ""))
            cat    = refinar_categoria(classe, nome)
        else:
            nome, classe, dt_ini, cat = cnpj, "", "", "OUTROS"

        # Extras do JSON Caixa
        extras = {}
        if cnpjs_json is not None and cnpj in cnpjs_json.index:
            row_j = cnpjs_json.loc[cnpj]
            if isinstance(row_j, pd.DataFrame):
                row_j = row_j.iloc[0]
            extras = {
                "taxa_adm":     fmt_br(row_j.get("pc_taxa_adm_cliente"), 2),
                "aberto":       "Sim" if row_j.get("ic_aberto_captacao") else "Nao",
                "aplic_min":    fmt_br(row_j.get("vr_aplicacao_inicial"), 2),
                "benchmark":    str(row_j.get("no_benchmark", "")),
                "resgate_conv": str(row_j.get("de_conversao_resgate", "")),
                "resgate_pgto": str(row_j.get("de_pagamento_resgate", "")),
                "pct_cdi_ano":  fmt_br(row_j.get("pc_rentabilidade_cdi_ano"), 2),
                "pct_cdi_12m":  fmt_br(row_j.get("pc_rentabilidade_cdi_12m"), 2),
            }

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
# SALVAMENTO
# ==========================================================
def validar(df: pd.DataFrame):
    total    = len(df)
    com_cota = df["Cota (R$)"].ne("").sum()
    com_12m  = df["Acum. 12M (%)"].ne("").sum()
    print(f"\n{'-'*50}")
    print(f"[stats] Total: {total}  |  Com cota: {com_cota}  |  Com 12M: {com_12m}")
    for cat, grp in df.groupby("Categoria"):
        print(f"   {cat:<35}: {len(grp):>3} fundos")
    print(f"{'-'*50}")


def salvar_periodos_json(pasta: str):
    """Gera periodos.json listando todos os CSVs disponiveis no historico."""
    padrao = os.path.join(pasta, "sipii_caixa_*.csv")
    arquivos = sorted(glob.glob(padrao), reverse=True)
    periodos = []
    for arq in arquivos:
        nome = os.path.basename(arq)
        data_str = nome.replace("sipii_caixa_", "").replace(".csv", "")
        try:
            dt = datetime.strptime(data_str, "%Y%m%d")
            periodos.append({
                "label": dt.strftime("%d/%m/%Y"),
                "file":  f"saida/{nome}",
                "data":  data_str,
            })
        except ValueError:
            pass
    # Adiciona o atual sempre no topo
    periodos_final = [{"label": "Atual", "file": "saida/dados_atuais.csv", "data": "atual"}]
    periodos_final.extend(periodos)
    caminho = os.path.join(pasta, "periodos.json")
    with open(caminho, "w", encoding="utf-8") as f:
        json.dump(periodos_final, f, ensure_ascii=False, indent=2)
    log(f"[OK] periodos.json: {len(periodos_final)} periodos")


def salvar(df: pd.DataFrame, pasta: str):
    if df.empty:
        print("[ERRO] Nenhum dado para salvar.")
        return

    validar(df)
    os.makedirs(pasta, exist_ok=True)
    ts = datetime.now().strftime("%Y%m%d")

    cols   = [c for c in df.columns if c != "CNPJ"]
    df_out = df[cols]

    p_atual = os.path.join(pasta, "dados_atuais.csv")
    p_hist  = os.path.join(pasta, f"sipii_caixa_{ts}.csv")
    df_out.to_csv(p_atual, index=False, encoding="utf-8-sig", sep=";")
    df_out.to_csv(p_hist,  index=False, encoding="utf-8-sig", sep=";")
    log(f"[OK] CSV: {p_atual}")

    try:
        p_xlsx = os.path.join(pasta, f"sipii_caixa_{ts}.xlsx")
        with pd.ExcelWriter(p_xlsx, engine="openpyxl") as writer:
            df_out.to_excel(writer, sheet_name="Todos", index=False)
            for cat, grp in df_out.groupby("Categoria"):
                grp.to_excel(writer, sheet_name=limpar_aba(cat), index=False)
        log(f"[OK] Excel: {p_xlsx}")
    except ImportError:
        log("[AVISO] openpyxl nao instalado")

    # Gera periodos.json para o HTML
    salvar_periodos_json(pasta)

    print(f"\n{'='*60}")
    print("[OK] CONCLUIDO!")
    print(f"   -> {p_atual}")
    print(f"   -> {p_hist}")
    print(f"{'='*60}\n")


# ==========================================================
# ENTRADA
# ==========================================================
if __name__ == "__main__":
    args = parse_args()
    df   = extrair(args)
    salvar(df, args.saida)
