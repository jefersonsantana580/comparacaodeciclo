
# -*- coding: utf-8 -*-
import io
import re
import datetime as dt
import pandas as pd
import streamlit as st

# =====================================================
# CONFIGURAÇÃO DA PÁGINA
# =====================================================
st.set_page_config(page_title="Comparativo Request Vs Plan", layout="wide")
st.title("Comparativo Request Vs Plan")
st.caption("Comparativo REQUEST − PLAN e resultado final do REQUEST")

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
    s = str(col).strip().lower()
    s = re.sub(r"[-_ ]+", "/", s)
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
            df[m] = pd.to_numeric(df[m], errors="coerce")
    return df


def colorir_valores(val):
    if isinstance(val, (int, float)):
        if val < 0:
            return "color:red;font-weight:bold;"
        if val > 0:
            return "color:green;font-weight:bold;"
    return ""


def formatar_tabela(df):
    df = df.fillna(0)
    cols_num = df.select_dtypes(include="number").columns
    return (
        df.style
        .format("{:,.0f}", subset=cols_num)
        .applymap(colorir_valores, subset=cols_num)
        .set_properties(subset=cols_num, **{"text-align": "center"})
        .set_properties(subset=df.columns.difference(cols_num),
                        **{"text-align": "left"})
    )

# =====================================================
# FUNÇÃO PRINCIPAL
# =====================================================
def gerar_passo1(xlsx_bytes):

    xls = pd.ExcelFile(io.BytesIO(xlsx_bytes), engine="openpyxl")
    plan = pd.read_excel(xls, "PLAN")
    req  = pd.read_excel(xls, "REQUEST")

    meses_p, _ = detectar_colunas_mes(plan)
    meses_r, _ = detectar_colunas_mes(req)
    meses = list(dict.fromkeys(meses_p + meses_r))

    plan = garantir_numerico(plan, meses)
    req  = garantir_numerico(req, meses)

    # =================================================
    # TABELA 1 — COMPARATIVO PRODUCT NEED
    # =================================================
    grp_need = ["SITE", "PRODUCT NEED"]

    plan_n = plan[grp_need + meses].groupby(grp_need)[meses].sum().reset_index()
    req_n  = req [grp_need + meses].groupby(grp_need)[meses].sum().reset_index()

    comp_n = pd.merge(plan_n, req_n, on=grp_need, how="outer",
                      suffixes=("_PLAN", "_REQ")).fillna(0)

    for m in meses:
        comp_n[m] = comp_n[f"{m}_REQ"] - comp_n[f"{m}_PLAN"]

    step1_need = comp_n[grp_need + meses]
    step1_need["TOTAL"] = step1_need[meses].sum(axis=1)

    total_n = {c: "TOTAL GERAL" for c in grp_need}
    for m in meses:
        total_n[m] = step1_need[m].sum()
    total_n["TOTAL"] = step1_need["TOTAL"].sum()

    step1_need = pd.concat([step1_need, pd.DataFrame([total_n])])

    # =================================================
    # ✅ TABELA 2 — COMPARATIVO SERIES
    # =================================================
    grp_serie = ["SITE","PRODUCT NEED","PRODUCT SERIES","PRODUCT BRAND","PRODUCT MARKET"]

    plan_s = plan[grp_serie + meses].groupby(grp_serie)[meses].sum().reset_index()
    req_s  = req [grp_serie + meses].groupby(grp_serie)[meses].sum().reset_index()

    comp_s = pd.merge(plan_s, req_s, on=grp_serie,
                      how="outer", suffixes=("_PLAN","_REQ")).fillna(0)

    for m in meses:
        comp_s[m] = comp_s[f"{m}_REQ"] - comp_s[f"{m}_PLAN"]

    step1_serie = comp_s[grp_serie + meses]
    step1_serie["TOTAL"] = step1_serie[meses].sum(axis=1)

    total_s = {c: "TOTAL GERAL" for c in grp_serie}
    for m in meses:
        total_s[m] = step1_serie[m].sum()
    total_s["TOTAL"] = step1_serie["TOTAL"].sum()

    step1_serie = pd.concat([step1_serie, pd.DataFrame([total_s])])

    # =================================================
    # ✅ TABELA 3 — REQUEST FINAL (SEM COMPARAÇÃO)
    # =================================================
    grp_req = ["SITE", "PRODUCT NEED"]

    req_only = (
        req[grp_req + meses]
        .groupby(grp_req)[meses]
        .sum()
        .reset_index()
        .fillna(0)
    )

    req_only["TOTAL"] = req_only[meses].sum(axis=1)

    total_req = {c: "TOTAL GERAL" for c in grp_req}
    for m in meses:
        total_req[m] = req_only[m].sum()
    total_req["TOTAL"] = req_only["TOTAL"].sum()

    req_only = pd.concat([req_only, pd.DataFrame([total_req])])

    # =================================================
    # EXPORTAÇÃO
    # =================================================
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        plan.to_excel(writer, "PLAN", index=False)
        req.to_excel(writer, "REQUEST", index=False)
        step1_need.to_excel(writer, "Comparativo_Need", index=False)
        step1_serie.to_excel(writer, "Comparativo_Serie", index=False)
        req_only.to_excel(writer, "Resumo_REQUEST_Final", index=False)

    return buf.getvalue(), step1_need, step1_serie, req_only

# =====================================================
# UI
# =====================================================
uploaded = st.file_uploader("Envie o Excel (PLAN e REQUEST)", type=["xlsx"])

if uploaded:
    excel_out, df_need, df_serie, df_req = gerar_passo1(uploaded.read())

    st.subheader("Resumo por PRODUCT NEED — Comparativo")
    st.dataframe(formatar_tabela(df_need), use_container_width=True)

    st.subheader("Comparativo por PRODUCT NEED + PRODUCT SERIES")
    st.dataframe(formatar_tabela(df_serie), use_container_width=True)

    st.subheader("Resumo por PRODUCT NEED — REQUEST Final")
    st.dataframe(formatar_tabela(df_req), use_container_width=True)

    st.download_button(
        "⬇️ Baixar Excel",
        data=excel_out,
        file_name="saida_step1.xlsx"
    )
else:
    st.info("Faça upload do Excel para iniciar.")
