
# streamlit_app.py
# -*- coding: utf-8 -*-

import io
import re
import datetime as dt
import pandas as pd
import streamlit as st

# =====================================================
# CONFIG
# =====================================================
st.set_page_config(
    page_title="Comparativo de Ciclo - Passo 1",
    layout="wide"
)

st.title("Passo 1 — Comparativo Mensal")
st.caption(
    "Comparativo REQUEST − PLAN com filtros e resumos por "
    "PRODUCT NEED e PRODUCT SERIES."
)

PT_BR_MESES = [
    "jan", "fev", "mar", "abr", "mai", "jun",
    "jul", "ago", "set", "out", "nov", "dez"
]

MES_RE = re.compile(
    r'^(jan|fev|mar|abr|mai|jun|jul|ago|set|out|nov|dez)/\d{2}$',
    re.IGNORECASE
)

# =====================================================
# FUNÇÕES AUXILIARES
# =====================================================
def _normalize_header(col):
    if isinstance(col, (pd.Timestamp, dt.date)):
        return f"{PT_BR_MESES[col.month-1]}/{col.year % 100:02d}"

    s = str(col).replace("\u00a0", " ").strip().lower()
    s = re.sub(r"[-_ ]+", "/", s)

    m = re.match(
        r"^(jan|fev|mar|abr|mai|jun|jul|ago|set|out|nov|dez)/(\d{2,4})$",
        s
    )
    if m:
        yy = m.group(2)[-2:]
        return f"{m.group(1)}/{yy}"

    m = re.match(
        r"^(jan|fev|mar|abr|mai|jun|jul|ago|set|out|nov|dez)(\d{2})$",
        s
    )
    if m:
        return f"{m.group(1)}/{m.group(2)}"

    return s


def detectar_colunas_mes(df):
    cols_mes = []
    debug_map = {}

    for c in df.columns:
        alias = _normalize_header(c)
        debug_map[str(c)] = alias
        if MES_RE.match(alias or ""):
            cols_mes.append(c)

    def ordem(c):
        mm, yy = debug_map[str(c)].split("/")
        return int(yy), PT_BR_MESES.index(mm)

    return sorted(cols_mes, key=ordem), debug_map


def garantir_numerico(df, meses):
    for m in meses:
        if m in df.columns:
            df[m] = pd.to_numeric(df[m], errors="coerce").fillna(0).astype(int)
    return df


def colorir_valores(val):
    if isinstance(val, (int, float)):
        if val < 0:
            return "color: red; font-weight: bold;"
        if val > 0:
            return "color: green; font-weight: bold;"
    return ""

# =====================================================
# FUNÇÃO PRINCIPAL
# =====================================================
def gerar_passo1(xlsx_bytes, show_debug=False):

    with pd.ExcelFile(io.BytesIO(xlsx_bytes), engine="openpyxl") as xls:
        if "PLAN" not in xls.sheet_names or "REQUEST" not in xls.sheet_names:
            raise ValueError("O arquivo precisa conter as abas PLAN e REQUEST.")

        plan = pd.read_excel(xls, "PLAN")
        req = pd.read_excel(xls, "REQUEST")

    # -------------------------------------------------
    # DETECTAR MESES
    # -------------------------------------------------
    meses_plan, map_plan = detectar_colunas_mes(plan)
    meses_req, map_req = detectar_colunas_mes(req)
    meses = list(dict.fromkeys(meses_plan + meses_req))

    if not meses:
        raise ValueError("Nenhuma coluna de mês no padrão pt-BR foi encontrada.")

    plan = garantir_numerico(plan, meses)
    req = garantir_numerico(req, meses)

    if show_debug:
        st.subheader("Diagnóstico de colunas - PLAN")
        st.json(map_plan)
        st.subheader("Diagnóstico de colunas - REQUEST")
        st.json(map_req)

    # -------------------------------------------------
    # FILTROS
    # -------------------------------------------------
    st.subheader("Filtros")

    def filtro(df, col):
        if col not in df.columns:
            return None
        valores = sorted(df[col].dropna().unique())
        return st.multiselect(col, valores, default=valores)

    c1, c2, c3, c4 = st.columns(4)

    with c1:
        f_brand = filtro(plan, "PRODUCT BRAND")
    with c2:
        f_market = filtro(plan, "PRODUCT MARKET")
    with c3:
        f_site = filtro(plan, "SITE")
    with c4:
        f_need = filtro(plan, "PRODUCT NEED")

    def aplicar_filtros(df):
        if f_brand is not None:
            df = df[df["PRODUCT BRAND"].isin(f_brand)]
        if f_market is not None:
            df = df[df["PRODUCT MARKET"].isin(f_market)]
        if f_site is not None:
            df = df[df["SITE"].isin(f_site)]
        if f_need is not None:
            df = df[df["PRODUCT NEED"].isin(f_need)]
        return df

    plan = aplicar_filtros(plan)
    req = aplicar_filtros(req)

    # =================================================
    # TABELA 1 — PRODUCT NEED + PRODUCT SERIES
    # =================================================
    grp_serie = ["SITE", "PRODUCT NEED", "PRODUCT SERIES"]

    plan_s = plan[grp_serie + meses].groupby(grp_serie, dropna=False)[meses].sum().reset_index()
    req_s  = req [grp_serie + meses].groupby(grp_serie, dropna=False)[meses].sum().reset_index()

    comp_s = pd.merge(plan_s, req_s, on=grp_serie, how="outer", suffixes=("_PLAN", "_REQ"))

    for m in meses:
        comp_s[m] = comp_s.get(f"{m}_REQ", 0).fillna(0) - comp_s.get(f"{m}_PLAN", 0).fillna(0)

    step1_serie = comp_s[grp_serie + meses].copy()
    step1_serie["TOTAL"] = step1_serie[meses].sum(axis=1)

    total_s = {c: "TOTAL GERAL" for c in grp_serie}
    for m in meses:
        total_s[m] = int(step1_serie[m].sum())
    total_s["TOTAL"] = int(step1_serie["TOTAL"].sum())

    step1_serie = pd.concat(
        [step1_serie, pd.DataFrame([total_s])],
        ignore_index=True
    )

    # =================================================
    # TABELA 2 — APENAS PRODUCT NEED
    # =================================================
    grp_need = ["SITE", "PRODUCT NEED"]

    plan_n = plan[grp_need + meses].groupby(grp_need, dropna=False)[meses].sum().reset_index()
    req_n  = req [grp_need + meses].groupby(grp_need, dropna=False)[meses].sum().reset_index()

    comp_n = pd.merge(plan_n, req_n, on=grp_need, how="outer", suffixes=("_PLAN", "_REQ"))

    for m in meses:
        comp_n[m] = comp_n.get(f"{m}_REQ", 0).fillna(0) - comp_n.get(f"{m}_PLAN", 0).fillna(0)

    step1_need = comp_n[grp_need + meses].copy()
    step1_need["TOTAL"] = step1_need[meses].sum(axis=1)

    total_n = {c: "TOTAL GERAL" for c in grp_need}
    for m in meses:
        total_n[m] = int(step1_need[m].sum())
    total_n["TOTAL"] = int(step1_need["TOTAL"].sum())

    step1_need = pd.concat(
        [step1_need, pd.DataFrame([total_n])],
        ignore_index=True
    )

    # =================================================
    # EXPORTAR EXCEL
    # =================================================
    
uf_out = io.BytesIO()

# Reabrir o Excel original
xls_in = pd.ExcelFile(io.BytesIO(xlsx_bytes), engine="openpyxl")

with pd.ExcelWriter(buf_out, engine="openpyxl") as writer:
    # Copiar todas as abas originais SEM ALTERAR
    for sheet in xls_in.sheet_names:
        df_original = pd.read_excel(xls_in, sheet_name=sheet)
        df_original.to_excel(
            writer,
            sheet_name=sheet,
            index=False
        )

    # Adicionar as novas abas
    step1_serie.to_excel(
        writer,
        sheet_name="Step1_Comparativo_Serie",
        index=False
    )

    step1_need.to_excel(
        writer,
        sheet_name="Step1_Comparativo_Need",
        index=False
    )



# =====================================================
# UI
# =====================================================
uploaded = st.file_uploader(
    "Envie o Excel (precisa conter abas PLAN e REQUEST)",
    type=["xlsx"]
)

debug = st.checkbox("Exibir diagnóstico de colunas", value=False)

if uploaded:
    try:
        excel_out, df_serie, df_need = gerar_passo1(
            uploaded.read(),
            show_debug=debug
        )

        st.success("Processamento concluído ✅")

        st.subheader("Comparativo por PRODUCT NEED + PRODUCT SERIES")
        st.dataframe(
            df_serie.style.applymap(colorir_valores),
            use_container_width=True
        )

        st.subheader("Resumo por PRODUCT NEED")
        st.dataframe(
            df_need.style.applymap(colorir_valores),
            use_container_width=True
        )

        st.download_button(
            "⬇️ Baixar Excel",
            data=excel_out,
            file_name="saida_step1.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    except Exception as e:
        st.error(f"Erro ao processar o arquivo: {e}")
else:
    st.info("Faça upload do Excel para iniciar.")
