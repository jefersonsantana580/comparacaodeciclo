
# -*- coding: utf-8 -*-
import io
import pandas as pd
import streamlit as st

# =====================================================
# CONFIGURAÇÃO DA PÁGINA (sempre executa)
# =====================================================
st.set_page_config(page_title="Comparativo Request Vs Plan", layout="wide")
st.title("Comparativo Request Vs Plan")
st.caption("Comparativo REQUEST − PLAN + REQUEST Final")

# =====================================================
# FUNÇÕES DE VISUAL
# =====================================================
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
# UPLOAD
# =====================================================
uploaded = st.file_uploader(
    "Envie o Excel (abas PLAN e REQUEST)",
    type=["xlsx"]
)

# =====================================================
# PROCESSAMENTO (PROTEGIDO)
# =====================================================
if uploaded is not None:
    try:
        # -----------------------------
        # Leitura segura
        # -----------------------------
        xls = pd.ExcelFile(uploaded)
        plan = pd.read_excel(xls, "PLAN")
        req  = pd.read_excel(xls, "REQUEST")

        # Detectar meses automaticamente (colunas numéricas sequenciais)
        meses = [c for c in plan.columns if "/" in str(c)]

        # Garantir numérico
        for m in meses:
            plan[m] = pd.to_numeric(plan[m], errors="coerce").fillna(0)
            req[m]  = pd.to_numeric(req[m], errors="coerce").fillna(0)

        # =================================================
        # TABELA 1 — COMPARATIVO PRODUCT NEED
        # =================================================
        grp_need = ["SITE", "PRODUCT NEED"]

        plan_n = plan.groupby(grp_need, dropna=False)[meses].sum().reset_index()
        req_n  = req.groupby(grp_need, dropna=False)[meses].sum().reset_index()

        comp_need = (
            plan_n.merge(req_n, on=grp_need, how="outer",
                         suffixes=("_PLAN", "_REQ"))
            .fillna(0)
        )

        for m in meses:
            comp_need[m] = comp_need[f"{m}_REQ"] - comp_need[f"{m}_PLAN"]

        comp_need = comp_need[grp_need + meses]
        comp_need["TOTAL"] = comp_need[meses].sum(axis=1)

        # TOTAL GERAL
        total_row = {c: "TOTAL GERAL" for c in grp_need}
        for m in meses:
            total_row[m] = comp_need[m].sum()
        total_row["TOTAL"] = comp_need["TOTAL"].sum()

        comp_need = pd.concat(
            [comp_need, pd.DataFrame([total_row])],
            ignore_index=True
        )

        # =================================================
        # TABELA 2 — REQUEST FINAL (SEM COMPARAÇÃO)
        # =================================================
        req_only = (
            req.groupby(grp_need, dropna=False)[meses]
            .sum()
            .reset_index()
        )

        req_only["TOTAL"] = req_only[meses].sum(axis=1)

        total_req = {c: "TOTAL GERAL" for c in grp_need}
        for m in meses:
            total_req[m] = req_only[m].sum()
        total_req["TOTAL"] = req_only["TOTAL"].sum()

        req_only = pd.concat(
            [req_only, pd.DataFrame([total_req])],
            ignore_index=True
        )

        # =================================================
        # UI
        # =================================================
        st.subheader("Resumo por PRODUCT NEED — Comparativo (REQ − PLAN)")
        st.dataframe(formatar_tabela(comp_need), use_container_width=True)

        st.subheader("Resumo por PRODUCT NEED — REQUEST Final")
        st.dataframe(formatar_tabela(req_only), use_container_width=True)

        # =================================================
        # EXPORTAÇÃO
        # =================================================
        buf = io.BytesIO()
        with pd.ExcelWriter(buf, engine="openpyxl") as writer:
            plan.to_excel(writer, "PLAN", index=False)
            req.to_excel(writer, "REQUEST", index=False)
            comp_need.to_excel(writer, "Comparativo_Need", index=False)
            req_only.to_excel(writer, "Resumo_REQUEST_Final", index=False)

        st.download_button(
            "⬇️ Baixar Excel",
            data=buf.getvalue(),
            file_name="saida_step1.xlsx"
        )

    except Exception as e:
        st.error("❌ Erro ao processar o arquivo")
        st.exception(e)

else:
    st.info("Faça upload do Excel para iniciar.")
