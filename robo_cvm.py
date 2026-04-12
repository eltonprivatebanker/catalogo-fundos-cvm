"""
============================================================
ROBÔ CVM v1.0 — Catálogo de Fundos Caixa
============================================================
Fonte   : dados.cvm.gov.br  (API oficial, gratuita)
Execução: ~15–30 segundos  (vs ~15 minutos com Selenium)
Sem Chrome, sem Selenium, sem popup, sem quebra de DOM.

O que faz:
  1. Baixa o cadastro de fundos da CVM e filtra pela Caixa
  2. Baixa os arquivos de cotas diárias do mês atual + anterior
  3. Baixa dez/ano anterior (para Acum. Ano)
  4. Baixa o mês equivalente 12 meses atrás (para Acum. 12M)
  5. Calcula todos os indicadores por fundo
  6. Salva dados_atuais.csv + histórico CSV/XLSX
============================================================
"""

import io
import re
import unicodedata
from datetime import date, datetime

import pandas as pd
import requests

# ── URLs da CVM ────────────────────────────────────────────
URL_CAD = "https://dados.cvm.gov.br/dados/FI/CAD/DADOS/cad_fi.csv"
URL_INF = "https://dados.cvm.gov.br/dados/FI/DOC/INF_DIARIO/DADOS/inf_diario_fi_{yyyymm}.csv"

# ── Mapeamento de categorias CVM → SIPII Caixa ─────────────
CLASSE_MAP = {
    "Fundo de Renda Fixa":               "RENDA FIXA",
    "Fundo de Ações":                    "ACOES",
    "Fundo Multimercado":                "MULTIMERCADO",
    "Fundo Cambial":                     "CAMBIAL",
    "Fundo de Índice de Mercado":        "FUNDO DE INDICE",
    "Fundo Mútuo de Privatização - FGTS":"FUNDOS MUTUOS DE PRIVATIZACAO",
}

# ── Helpers gerais ─────────────────────────────────────────
def rm_accent(s: str) -> str:
    return "".join(c for c in unicodedata.normalize("NFD", s)
                   if unicodedata.category(c) != "Mn")

def refinar_categoria(classe_cvm: str, nome_fundo: str) -> str:
    """
    Refina a categoria usando pistas do nome do fundo.
    A CVM agrupa 'Simples', 'Referenciado' e 'Curto Prazo'
    todos como 'Fundo de Renda Fixa' — o nome do fundo
    deixa claro o subtipo.
    """
    nome = rm_accent(nome_fundo.upper())
    if "SIMPLES" in nome:
        return "RENDA FIXA SIMPLES"
    if any(x in nome for x in ("REF DI", "REFERENCIADO", "REFERENC DI", "REFERENC")):
        return "RENDA FIXA REFERENCIADO"
    if "CURTO PRAZO" in nome and "Renda Fixa" in classe_cvm:
        return "RENDA FIXA CURTO PRAZO"
    return CLASSE_MAP.get(classe_cvm, classe_cvm or "OUTROS")

def fmt_br(v, decimais: int = 2) -> str:
    """Float → string no formato brasileiro ('1.234,56')."""
    if v is None or (isinstance(v, float) and pd.isna(v)):
        return ""
    return f"{v:,.{decimais}f}".replace(",", "X").replace(".", ",").replace("X", ".")

def retorno_pct(cota_ini, cota_fim) -> float | None:
    """Calcula retorno % entre duas cotas."""
    try:
        if cota_ini and cota_fim and float(cota_ini) > 0:
            return (float(cota_fim) / float(cota_ini) - 1) * 100
    except (TypeError, ValueError):
        pass
    return None

def limpar_aba(nome: str) -> str:
    """Remove caracteres proibidos pelo Excel e limita a 31 chars."""
    return re.sub(r'[:\\/*?\[\]]', "", nome)[:31]

# ── Download com retry ─────────────────────────────────────
def baixar_csv(url: str, label: str = "") -> pd.DataFrame:
    for tentativa in range(3):
        try:
            print(f"   ⬇️  {label or url}")
            r = requests.get(url, timeout=60)
            r.raise_for_status()
            return pd.read_csv(
                io.StringIO(r.content.decode("latin1")),
                sep=";", low_memory=False
            )
        except Exception as e:
            if tentativa == 2:
                print(f"   ❌ Falhou após 3 tentativas: {e}")
                return pd.DataFrame()
            print(f"   ⚠️  Tentativa {tentativa + 1} falhou, retry em 3s...")
            import time; time.sleep(3)
    return pd.DataFrame()

def baixar_inf_diario(ano: int, mes: int) -> pd.DataFrame:
    yyyymm = f"{ano}{mes:02d}"
    return baixar_csv(URL_INF.format(yyyymm=yyyymm), f"cotas diárias {yyyymm}")

# ── Lógica de datas ────────────────────────────────────────
def mes_anterior(ano: int, mes: int):
    return (ano - 1, 12) if mes == 1 else (ano, mes - 1)

def mes_12m_atras(ano: int, mes: int):
    """Mesmo mês, 1 ano atrás. Usa mês anterior se mes==1 por segurança."""
    return (ano - 1, mes)

# ── Filtra e prepara arquivo INF ───────────────────────────
def preparar_inf(df_raw: pd.DataFrame, cnpjs: set) -> pd.DataFrame:
    if df_raw.empty:
        return pd.DataFrame()
    df = df_raw[df_raw["CNPJ_FUNDO"].isin(cnpjs)].copy()
    df["DT_COMPTC"] = pd.to_datetime(df["DT_COMPTC"], errors="coerce")
    df["VL_QUOTA"] = pd.to_numeric(df["VL_QUOTA"], errors="coerce")
    df["VL_PATRIM_LIQ"] = pd.to_numeric(df["VL_PATRIM_LIQ"], errors="coerce")
    return df.sort_values(["CNPJ_FUNDO", "DT_COMPTC"])

# ── Extração principal ─────────────────────────────────────
def extrair() -> pd.DataFrame:
    hoje = date.today()
    print(f"\n🚀 Extração CVM — {hoje.strftime('%d/%m/%Y')}")
    print("=" * 52)

    # ── 1. Cadastro ─────────────────────────────────────────
    print("\n📋 [1/4] Cadastro de fundos")
    df_cad = baixar_csv(URL_CAD, "cad_fi.csv (cadastro nacional)")
    if df_cad.empty:
        print("❌ Cadastro indisponível. Abortando.")
        return pd.DataFrame()

    mask = (
        df_cad["ADMIN"].fillna("").str.upper().str.contains("CAIXA") |
        df_cad["GESTOR"].fillna("").str.upper().str.contains("CAIXA")
    ) & df_cad["SIT"].fillna("").str.upper().str.contains("FUNCIONAMENTO")

    df_caixa = df_cad[mask].copy()
    cnpjs = set(df_caixa["CNPJ_FUNDO"].unique())
    print(f"   ✅ {len(df_caixa)} fundos Caixa em funcionamento")

    # ── 2. Cotas — mês atual + anterior ────────────────────
    print("\n📈 [2/4] Cotas do mês atual e anterior")
    ano_ant, mes_ant = mes_anterior(hoje.year, hoje.month)
    df_atual = preparar_inf(baixar_inf_diario(hoje.year, hoje.month), cnpjs)
    df_ant   = preparar_inf(baixar_inf_diario(ano_ant, mes_ant), cnpjs)
    df_recente = pd.concat([df_ant, df_atual], ignore_index=True)

    # ── 3. Histórico para Acum. Ano ────────────────────────
    print("\n📅 [3/4] Dezembro do ano anterior (Acum. Ano)")
    df_dez = preparar_inf(baixar_inf_diario(hoje.year - 1, 12), cnpjs)

    # ── 4. Histórico para Acum. 12M ────────────────────────
    print("\n📆 [4/4] Mês equivalente 12 meses atrás (Acum. 12M)")
    ano_12m, mes_12m = mes_12m_atras(hoje.year, hoje.month)
    df_12m = preparar_inf(baixar_inf_diario(ano_12m, mes_12m), cnpjs)

    # ── 5. Calcula indicadores por fundo ───────────────────
    print("\n🔢 Calculando indicadores...")
    resultados = []

    for _, fundo in df_caixa.iterrows():
        cnpj    = fundo["CNPJ_FUNDO"]
        nome    = str(fundo.get("DENOM_SOCIAL", "")).strip()
        classe  = str(fundo.get("CLASSE", ""))
        cat     = refinar_categoria(classe, nome)
        dt_ini  = str(fundo.get("DT_INI_ATIV", ""))

        d = df_recente[df_recente["CNPJ_FUNDO"] == cnpj]

        if d.empty:
            resultados.append({
                "Categoria": cat, "Fundo": nome, "CNPJ": cnpj,
                "Data Inicio": dt_ini, "Aplic. Inicial (R$)": "",
                "Cota (R$)": "", "Variacao Dia (%)": "",
                "Acum. Mes (%)": "", "Acum. Ano (%)": "",
                "Acum. 12M (%)": "", "PL (milhoes R$)": "",
                "PL Medio (milhoes R$)": "", "URL": "",
            })
            continue

        ultima      = d.iloc[-1]
        cota_atual  = ultima["VL_QUOTA"]
        pl_atual    = ultima["VL_PATRIM_LIQ"] / 1_000_000 if pd.notna(ultima["VL_PATRIM_LIQ"]) else None
        pl_medio    = d["VL_PATRIM_LIQ"].tail(30).mean()
        pl_medio    = pl_medio / 1_000_000 if pd.notna(pl_medio) else None

        # Variação dia
        var_dia = retorno_pct(d.iloc[-2]["VL_QUOTA"], cota_atual) if len(d) >= 2 else None

        # Acum. Mês — primeira cota do mês atual
        d_mes = d[d["DT_COMPTC"].dt.month == hoje.month]
        acum_mes = retorno_pct(
            d_mes.iloc[0]["VL_QUOTA"] if not d_mes.empty else None,
            cota_atual
        )

        # Acum. Ano — última cota de dez do ano anterior
        d_dez_f = df_dez[df_dez["CNPJ_FUNDO"] == cnpj]
        acum_ano = retorno_pct(
            d_dez_f.iloc[-1]["VL_QUOTA"] if not d_dez_f.empty else None,
            cota_atual
        )

        # Acum. 12M — última cota de ~12 meses atrás
        d_12m_f = df_12m[df_12m["CNPJ_FUNDO"] == cnpj]
        acum_12m = retorno_pct(
            d_12m_f.iloc[-1]["VL_QUOTA"] if not d_12m_f.empty else None,
            cota_atual
        )

        resultados.append({
            "Categoria":           cat,
            "Fundo":               nome,
            "CNPJ":                cnpj,
            "Data Inicio":         dt_ini,
            "Aplic. Inicial (R$)": "",           # não fornecido pela CVM
            "Cota (R$)":           fmt_br(cota_atual, 8),
            "Variacao Dia (%)":    fmt_br(var_dia, 3),
            "Acum. Mes (%)":       fmt_br(acum_mes, 2),
            "Acum. Ano (%)":       fmt_br(acum_ano, 2),
            "Acum. 12M (%)":       fmt_br(acum_12m, 2),
            "PL (milhoes R$)":     fmt_br(pl_atual, 3),
            "PL Medio (milhoes R$)": fmt_br(pl_medio, 3),
            "URL":                 "",
        })

    df_out = pd.DataFrame(resultados)
    print(f"   ✅ {len(df_out)} fundos processados")
    return df_out


# ── Validação ──────────────────────────────────────────────
def validar(df: pd.DataFrame):
    print("\n📊 Validação de qualidade:")
    print(f"   Total fundos      : {len(df)}")
    com_cota  = df["Cota (R$)"].ne("").sum()
    sem_cota  = df["Cota (R$)"].eq("").sum()
    com_12m   = df["Acum. 12M (%)"].ne("").sum()
    print(f"   Com cota          : {com_cota}")
    print(f"   Sem cota          : {sem_cota}  (novos/sem dados no período)")
    print(f"   Com Acum. 12M     : {com_12m}")
    for cat in sorted(df["Categoria"].unique()):
        n = (df["Categoria"] == cat).sum()
        print(f"   {cat:<40}: {n:>3}")


# ── Salvamento ─────────────────────────────────────────────
def salvar(df: pd.DataFrame):
    if df.empty:
        print("❌ Nenhum dado para salvar.")
        return

    validar(df)

    # Remove coluna CNPJ do CSV de saída (é interna)
    cols = [c for c in df.columns if c != "CNPJ"]
    df_out = df[cols]

    ts = datetime.now().strftime("%Y%m%d")

    df_out.to_csv("dados_atuais.csv",           index=False, encoding="utf-8-sig")
    df_out.to_csv(f"sipii_caixa_{ts}.csv",      index=False, encoding="utf-8-sig")

    with pd.ExcelWriter(f"sipii_caixa_{ts}.xlsx", engine="openpyxl") as writer:
        df_out.to_excel(writer, sheet_name="Todos", index=False)
        for cat in df_out["Categoria"].unique():
            df_out[df_out["Categoria"] == cat].to_excel(
                writer, sheet_name=limpar_aba(cat), index=False)

    print(f"\n{'='*52}")
    print("✅ DADOS SALVOS COM SUCESSO!")
    print(f"   → dados_atuais.csv")
    print(f"   → sipii_caixa_{ts}.csv")
    print(f"   → sipii_caixa_{ts}.xlsx")
    print(f"{'='*52}\n")


# ── Entrada ────────────────────────────────────────────────
if __name__ == "__main__":
    df = extrair()
    salvar(df)
