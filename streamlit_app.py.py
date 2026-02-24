
# streamlit_app.py
# -*- coding: utf-8 -*-
import io
import re
from pathlib import Path

import pandas as pd
import streamlit as st

PT_BR_MESES = ['jan','fev','mar','abr','mai','jun','jul','ago','set','out','nov','dez']
MES_RE = re.compile(r'^(jan|fev|mar|abr|mai|jun|jul|ago|set|out|nov|dez)/\\d{2}$', re.IGNORECASE)

st.set_page_config(page_title="Comparativo de Ciclo - Passo 1", layout="wide")
st.title("Passo 1 — Comparativo Mensal por PRODUCT SERIES (REQUEST − PLAN)")
st.caption("Lê PLAN e REQUEST, detecta meses pt-BR, agrega por SITE + PRODUCT SERIES e grava a aba Step1_Comparativo_Serie.")

def detectar_colunas_mes(df: pd.DataFrame):
    cols = [c for c in df.columns if isinstance(c, str) and MES_RE.match(c)]
    ordem_map = {m:i for i, m in enumerate(PT_BR_MESES)}
    cols = sorted(cols, key=lambda x: (int(x[-2:]), ordem_map[x.split('/')[0].lower()]))
    return cols

def garantir_numerico(df: pd.DataFrame, mes_cols):
    for m in mes_cols:
        if m in df.columns:
            df[m] = pd.to_numeric(df[m], errors='coerce').fillna(0).astype(int)
    return df

def gerar_passo1(xlsx_bytes, out_sheet="Step1_Comparativo_Serie"):
    # Ler workbook em memória
    with pd.ExcelFile(io.BytesIO(xlsx_bytes), engine="openpyxl") as xls:
        if "PLAN" not in xls.sheet_names or "REQUEST" not in xls.sheet_names:
            raise ValueError("O arquivo precisa conter as abas 'PLAN' e 'REQUEST'.")
        plan = pd.read_excel(xls, sheet_name="PLAN", engine="openpyxl")
        req  = pd.read_excel(xls, sheet_name="REQUEST", engine="openpyxl")

    mes_cols_plan = detectar_colunas_mes(plan)
    mes_cols_req  = detectar_colunas_mes(req)
    mes_cols = sorted(list(dict.fromkeys(mes_cols_plan + mes_cols_req)),
                      key=lambda x: (int(x[-2:]), PT_BR_MESES.index(x.split('/')[0].lower())))

    if not mes_cols:
        raise ValueError("Não encontrei colunas de meses no padrão pt-BR (ex.: 'jan/26', 'fev/26'...).")

    plan = garantir_numerico(plan, mes_cols)
    req  = garantir_numerico(req,  mes_cols)

    grp = [c for c in ["SITE", "PRODUCT SERIES"] if c in plan.columns and c in req.columns]
    if len(grp) < 2:
        raise ValueError("As colunas 'SITE' e 'PRODUCT SERIES' precisam existir em 'PLAN' e 'REQUEST'.")

    plan_agg = plan[grp + mes_cols].groupby(grp, dropna=False)[mes_cols].sum().reset_index()
    req_agg  = req [grp + mes_cols].groupby(grp, dropna=False)[mes_cols].sum().reset_index()

    comp = pd.merge(req_agg, plan_agg, on=grp, how="outer", suffixes=("_REQ", "_PLAN"))
    for m in mes_cols:
        comp[m] = comp.get(f"{m}_REQ", 0).fillna(0).astype(int) - comp.get(f"{m}_PLAN", 0).fillna(0).astype(int)

    step1 = comp[grp + mes_cols].copy()
    step1["TOTAL"] = step1[mes_cols].sum(axis=1)
    step1 = step1.sort_values(by=["SITE", "TOTAL"], ascending=[True, False])

    # Linha TOTAL GERAL (soma por coluna)
    linha_total = {k: "TOTAL GERAL" for k in grp}
    for m in mes_cols:
        linha_total[m] = int(step1[m].sum())
    linha_total["TOTAL"] = int(step1["TOTAL"].sum())
    step1 = pd.concat([step1, pd.DataFrame([linha_total])], ignore_index=True)

    # Escrever de volta todas as abas originais + Step1
    buf_in = io.BytesIO(xlsx_bytes)
    xls_in = pd.ExcelFile(buf_in, engine="openpyxl")
    buf_out = io.BytesIO()
    with pd.ExcelWriter(buf_out, engine="openpyxl") as writer:
        for sheet in xls_in.sheet_names:
            df_sheet = pd.read_excel(xls_in, sheet_name=sheet, engine="openpyxl")
            df_sheet.to_excel(writer, sheet_name=sheet, index=False)
        step1.to_excel(writer, sheet_name=out_sheet, index=False)

    return buf_out.getvalue(), step1

uploaded = st.file_uploader("Envie o Excel (precisa conter abas PLAN e REQUEST)", type=["xlsx"])
if uploaded is not None:
    try:
        out_bytes, df_preview = gerar_passo1(uploaded.read())
        st.success("Aba 'Step1_Comparativo_Serie' gerada com sucesso.")
        st.dataframe(df_preview, use_container_width=True)
        st.download_button(
            label="⬇️ Baixar Excel com a aba Step1_Comparativo_Serie",
            data=out_bytes,
            file_name="saida_step1.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
    except Exception as e:
        st.error(f"Erro ao processar: {e}")
else:
    st.info("Faça o upload do arquivo Excel para iniciar.")
