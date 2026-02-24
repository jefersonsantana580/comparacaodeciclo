
# streamlit_app.py
# -*- coding: utf-8 -*-
import io
import re
import datetime as dt
from pathlib import Path

import pandas as pd
import streamlit as st

PT_BR_MESES = ['jan','fev','mar','abr','mai','jun','jul','ago','set','out','nov','dez']
MES_RE = re.compile(r'^(jan|fev|mar|abr|mai|jun|jul|ago|set|out|nov|dez)/\d{2}$', re.IGNORECASE)

st.set_page_config(page_title="Comparativo de Ciclo - Passo 1", layout="wide")
st.title("Passo 1 — Comparativo Mensal por PRODUCT SERIES (REQUEST − PLAN)")
st.caption("Lê PLAN e REQUEST, detecta meses pt-BR, agrega por SITE + PRODUCT SERIES e grava a aba Step1_Comparativo_Serie.")

# ---------- Normalização dos headers ----------
def _normalize_header(col):
    """
    Transforma o header em alias 'mmm/aa' (pt-BR) SEM alterar o DataFrame original.
    Aceita:
      - pandas.Timestamp / datetime.date (ex.: 2026-01-01)
      - strings com variações: 'Jan/26', 'JAN-26', ' jan 26 ', 'jan/2026'
    """
    # 1) Datas reais
    if isinstance(col, (pd.Timestamp, dt.date)):
        m = col.month
        y = col.year % 100
        return f"{PT_BR_MESES[m-1]}/{y:02d}"

    # 2) Strings com ruídos
    s = str(col).replace('\u00a0', ' ')  # NBSP -> espaço normal
    s = s.strip().lower()
    # trocar separadores por '/'
    s = re.sub(r'[-_ ]+', '/', s)

    # 'jan/2026' -> 'jan/26'
    m = re.match(r'^(jan|fev|mar|abr|mai|jun|jul|ago|set|out|nov|dez)/(\d{2,4})$', s)
    if m:
        mm, yy = m.group(1), m.group(2)
        if len(yy) == 4:
            yy = yy[-2:]
        return f"{mm}/{yy}"

    # 'jan26' -> 'jan/26'
    m = re.match(r'^(jan|fev|mar|abr|mai|jun|jul|ago|set|out|nov|dez)(\d{2})$', s)
    if m:
        return f"{m.group(1)}/{m.group(2)}"

    return s

def detectar_colunas_mes(df: pd.DataFrame):
    """
    Retorna:
      - cols_meses: lista com os NOMES ORIGINAIS das colunas que são meses
      - idx_sorted: mesma lista, porém ordenada cronologicamente via alias normalizado
      - debug_map: dict {col_original -> alias_normalizado}
    """
    debug_map = {}
    cols_meses = []
    for c in df.columns:
        alias = _normalize_header(c)
        debug_map[str(c)] = alias
        if MES_RE.match(alias or ""):
            cols_meses.append(c)

    # ordenar por ano e mês usando o alias
    def _ord_key(c):
        alias = debug_map[str(c)]
        mm, yy = alias.split('/')
        return (int(yy), PT_BR_MESES.index(mm))
    idx_sorted = sorted(cols_meses, key=_ord_key)
    return cols_meses, idx_sorted, debug_map

def garantir_numerico(df: pd.DataFrame, mes_cols):
    for m in mes_cols:
        # se coluna não existir no DF (caso raro), ignore
        if m in df.columns:
            df[m] = pd.to_numeric(df[m], errors='coerce').fillna(0).astype(int)
    return df

def gerar_passo1(xlsx_bytes, out_sheet="Step1_Comparativo_Serie", show_debug=False):
    # Ler workbook em memória
    with pd.ExcelFile(io.BytesIO(xlsx_bytes), engine="openpyxl") as xls:
        if "PLAN" not in xls.sheet_names or "REQUEST" not in xls.sheet_names:
            raise ValueError("O arquivo precisa conter as abas 'PLAN' e 'REQUEST'.")
        plan = pd.read_excel(xls, sheet_name="PLAN", engine="openpyxl")
        req  = pd.read_excel(xls, sheet_name="REQUEST", engine="openpyxl")

    # Detectar meses de cada aba (usando nome ORIGINAL + alias p/ ordenação)
    plan_cols, plan_sorted, plan_map = detectar_colunas_mes(plan)
    req_cols,  req_sorted,  req_map  = detectar_colunas_mes(req)

    if show_debug:
        st.subheader("Diagnóstico de colunas (PLAN)")
        st.json(plan_map)
        st.subheader("Diagnóstico de colunas (REQUEST)")
        st.json(req_map)

    # União dos meses reconhecidos nas duas abas (preservando nomes originais)
    mes_cols_union = list(dict.fromkeys(plan_sorted + req_sorted))
    if not mes_cols_union:
        raise ValueError("Não encontrei colunas de meses no padrão pt-BR (ex.: 'jan/26', 'fev/26'...).")

    # Coagir numérico nos meses detectados
    plan = garantir_numerico(plan, mes_cols_union)
    req  = garantir_numerico(req,  mes_cols_union)

    # Chaves de agrupamento
    grp = [c for c in ["SITE", "PRODUCT SERIES"] if c in plan.columns and c in req.columns]
    if len(grp) < 2:
        raise ValueError("As colunas 'SITE' e 'PRODUCT SERIES' precisam existir em 'PLAN' e 'REQUEST'.")

    plan_agg = plan[grp + mes_cols_union].groupby(grp, dropna=False)[mes_cols_union].sum().reset_index()
    req_agg  = req [grp + mes_cols_union].groupby(grp, dropna=False)[mes_cols_union].sum().reset_index()

    comp = pd.merge(req_agg, plan_agg, on=grp, how="outer", suffixes=("_REQ", "_PLAN"))
    for m in mes_cols_union:
        comp[m] = comp.get(f"{m}_REQ", 0).fillna(0).astype(int) - comp.get(f"{m}_PLAN", 0).fillna(0).astype(int)

    step1 = comp[grp + mes_cols_union].copy()
    step1["TOTAL"] = step1[mes_cols_union].sum(axis=1)
    step1 = step1.sort_values(by=["SITE", "TOTAL"], ascending=[True, False])

    # Linha TOTAL GERAL (soma por coluna)
    linha_total = {k: "TOTAL GERAL" for k in grp}
    for m in mes_cols_union:
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

# -------- UI --------
uploaded = st.file_uploader("Envie o Excel (precisa conter abas PLAN e REQUEST)", type=["xlsx"])
debug = st.checkbox("Exibir diagnóstico de colunas (headers reconhecidos)", value=False)

if uploaded is not None:
    try:
        out_bytes, df_preview = gerar_passo1(uploaded.read(), show_debug=debug)
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
