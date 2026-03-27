
# -*- coding: utf-8 -*-
import io
import re
import datetime as dt
import pandas as pd
import streamlit as st

# =====================================================
# CONFIGURAÇÃO DA PÁGINA
# =====================================================
st.set_page_config(
    page_title="Comparativo Request Vs Plan",
    layout="wide"
)

st.title("Comparativo Request Vs Plan")
st.caption("Comparativo REQUEST − PLAN com filtros e resumos")

PT_BR_MESES = ["jan","fev","mar","abr","mai","jun","jul","ago","set","out","nov","dez"]

MES_RE = re.compile(
    r'^(jan|fev|mar|abr|mai|jun|jul|ago|set|out|nov|dez)/\d{2}$',
    re.IGNORECASE
)

# =====================================================
# FUNÇÕES UTILITÁRIAS
# =====================================================
def _normalize_header(col):
    if isinstance(col, (pd.Timestamp, dt.date)):
        return f"{PT_BR_MESES[col.month-1]}/{col.year % 100:02d}"

    s = str(col).replace("\u00a0", " ").strip().lower()
    s = re.sub(r"[-_ ]+", "/", s)

    m = re.match(r"^(jan|fev|mar|abr|mai|jun|jul|ago|set|out|nov|dez)/(\d{2,4})$", s)
    if m:
        return f"{m.group(1)}/{m.group(2)[-2:]}"
    return s


def detectar_colunas_mes(df):
    cols_mes, debug_map = [], {}
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
            df[m] = pd.to_numeric(df[m], errors="coerce")
    return df


def colorir_valores(val):
    if isinstance(val, (int, float)):
        if val < 0:
            return "color:red;font-weight:bold;"
        if val > 0:
            return "color:green;font-weight:bold;"
    return ""


# ✅ FORMATAÇÃO VISUAL
def formatar_tabela(df):
    df = df.fillna(0)
    cols_num = df.select_dtypes(include="number").columns

    return (
        df.style
        .format(lambda x: f"{x:,.0f}".replace(",", "."), subset=cols_num)
        .applymap(colorir_valores, subset=cols_num)
        .set_properties(subset=cols_num, **{"text-align": "center"})
        .set_properties(
            subset=df.columns.difference(cols_num),
            **{"text-align": "left"}
        )
    )

# =====================================================
# FUNÇÃO PRINCIPAL
# =====================================================
def gerar_passo1(xlsx_bytes):

    xls = pd.ExcelFile(io.BytesIO(xlsx_bytes), engine="openpyxl")
    plan = pd.read_excel(xls, "PLAN")
    req  = pd.read_excel(xls, "REQUEST")

    meses_plan, _ = detectar_colunas_mes(plan)
    meses_req,  _ = detectar_colunas_mes(req)
    meses = list(dict.fromkeys(meses_plan + meses_req))

    plan = garantir_numerico(plan, meses)
    req  = garantir_numerico(req, meses)

    # =================================================
    # FILTROS
    # =================================================
    def filtro_mult(df, col):
        vals = sorted(df[col].dropna().unique())
        return st.multiselect(col, vals, default=vals)

    c1, c2, c3, c4 = st.columns(4)
    with c1: f_brand = filtro_mult(plan, "PRODUCT BRAND")
    with c2: f_market = filtro_mult(plan, "PRODUCT MARKET")
    with c3: f_site = filtro_mult(plan, "SITE")
    with c4: f_need = filtro_mult(plan, "PRODUCT NEED")

    def aplicar(df):
        return df[
            df["PRODUCT BRAND"].isin(f_brand) &
            df["PRODUCT MARKET"].isin(f_market) &
            df["SITE"].isin(f_site) &
            df["PRODUCT NEED"].isin(f_need)
        ]

    plan, req = aplicar(plan), aplicar(req)

    # =================================================
    # TABELA 1 — PRODUCT SERIES (REQ - PLAN)
    # =================================================
    grp_serie = ["SITE","PRODUCT NEED","PRODUCT SERIES","PRODUCT BRAND","PRODUCT MARKET"]

    plan_s = plan.groupby(grp_serie)[meses].sum().reset_index()
    req_s  = req.groupby(grp_serie)[meses].sum().reset_index()

    comp_s = plan_s.merge(req_s, on=grp_serie, how="outer", suffixes=("_PLAN","_REQ")).fillna(0)
    for m in meses:
        comp_s[m] = comp_s[f"{m}_REQ"] - comp_s[f"{m}_PLAN"]

    df_serie = comp_s[grp_serie + meses]
    df_serie["TOTAL"] = df_serie[meses].sum(axis=1)

    # =================================================
    # TABELA 2 — PRODUCT NEED (REQ - PLAN)
    # =================================================
    grp_need = ["SITE","PRODUCT NEED"]

    plan_n = plan.groupby(grp_need)[meses].sum().reset_index()
    req_n  = req.groupby(grp_need)[meses].sum().reset_index()

    comp_n = plan_n.merge(req_n, on=grp_need, how="outer", suffixes=("_PLAN","_REQ")).fillna(0)
    for m in meses:
        comp_n[m] = comp_n[f"{m}_REQ"] - comp_n[f"{m}_PLAN"]

    df_need = comp_n[grp_need + meses]
    df_need["TOTAL"] = df_need[meses].sum(axis=1)

    total_need = {c:"TOTAL GERAL" for c in grp_need}
    for m in meses:
        total_need[m] = df_need[m].sum()
    total_need["TOTAL"] = df_need["TOTAL"].sum()

    df_need = pd.concat([df_need, pd.DataFrame([total_need])])

    # =================================================
    # TABELA EXTRA — PRODUCT NEED (REQUEST)
    # =================================================
    df_req_need = req.groupby(grp_need)[meses].sum().reset_index()
    df_req_need["TOTAL"] = df_req_need[meses].sum(axis=1)

    total_req = {c:"TOTAL GERAL" for c in grp_need}
    for m in meses:
        total_req[m] = df_req_need[m].sum()
    total_req["TOTAL"] = df_req_need["TOTAL"].sum()

    df_req_need = pd.concat([df_req_need, pd.DataFrame([total_req])])

    # =================================================
    # EXPORTAÇÃO EXCEL
    # =================================================
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        for sheet in xls.sheet_names:
            pd.read_excel(xls, sheet).to_excel(writer, sheet_name=sheet, index=False)

        df_serie.to_excel(writer, "Step1_Comparativo_Serie", index=False)
        df_need.to_excel(writer, "Step1_Comparativo_Need", index=False)
        df_req_need.to_excel(writer, "Resumo_Request_Product_Need", index=False)

    return buf.getvalue(), df_serie, df_need, df_req_need

# =====================================================
# UI
# =====================================================
uploaded = st.file_uploader("Envie o Excel (PLAN e REQUEST)", type=["xlsx"])

if uploaded:
    excel, df_serie, df_need, df_req_need = gerar_passo1(uploaded.read())

    st.subheader("Comparativo por PRODUCT NEED + PRODUCT SERIES")
    st.dataframe(formatar_tabela(df_serie), use_container_width=True)

    st.subheader("Resumo por PRODUCT NEED (REQ - PLAN)")
    st.dataframe(formatar_tabela(df_need), use_container_width=True)

    st.subheader("Resumo por PRODUCT NEED (REQUEST)")
    st.dataframe(formatar_tabela(df_req_need), use_container_width=True)

    st.download_button(
        "⬇️ Baixar Excel",
        data=excel,
        file_name="saida_step1.xlsx"
    )
else:
    st.info("Faça upload do Excel para iniciar.")

