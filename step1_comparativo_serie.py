
# -*- coding: utf-8 -*-
"""
Step 1 — Comparativo Mensal por PRODUCT SERIES (REQUEST − PLAN)
-------------------------------------------------------------------------------
Lê um Excel com as abas PLAN e REQUEST, detecta automaticamente as colunas de
meses no padrão pt-BR (jan/26 ... dez/26), agrega por SITE + PRODUCT SERIES e
calcula o delta mensal (REQUEST − PLAN), adicionando TOTAL por linha e a linha
TOTAL GERAL (soma por coluna). Sobrescreve/cria a aba
"Step1_Comparativo_Serie" no mesmo arquivo de entrada.

Uso (CLI):
    python step1_comparativo_serie.py --arquivo "testar _ analise de ciclo.xlsx"

Requisitos: pandas, openpyxl
"""
import argparse
import re
from pathlib import Path
import pandas as pd

PT_BR_MESES = ['jan','fev','mar','abr','mai','jun','jul','ago','set','out','nov','dez']

MES_RE = re.compile(r'^(jan|fev|mar|abr|mai|jun|jul|ago|set|out|nov|dez)/\d{2}$', re.IGNORECASE)


def detectar_colunas_mes(df: pd.DataFrame) -> list:
    """Retorna as colunas de meses ordenadas (pt-BR) presentes no DataFrame."""
    cols = [c for c in df.columns if isinstance(c, str) and MES_RE.match(c)]
    ordem_map = {m:i for i, m in enumerate(PT_BR_MESES)}
    cols = sorted(cols, key=lambda x: (int(x[-2:]), ordem_map[x.split('/')[0].lower()]))
    return cols


def garantir_numerico(df: pd.DataFrame, mes_cols: list) -> pd.DataFrame:
    for m in mes_cols:
        if m in df.columns:
            df[m] = pd.to_numeric(df[m], errors='coerce').fillna(0).astype(int)
    return df


def comparativo_request_vs_plan_por_serie(xlsx_path: Path, out_sheet: str = 'Step1_Comparativo_Serie') -> None:
    # Leitura das abas de origem
    plan = pd.read_excel(xlsx_path, sheet_name='PLAN', engine='openpyxl')
    req  = pd.read_excel(xlsx_path, sheet_name='REQUEST', engine='openpyxl')

    # Detectar colunas de meses
    mes_cols_plan = detectar_colunas_mes(plan)
    mes_cols_req  = detectar_colunas_mes(req)
    mes_cols = sorted(list(dict.fromkeys(mes_cols_plan + mes_cols_req)),
                      key=lambda x: (int(x[-2:]), PT_BR_MESES.index(x.split('/')[0].lower())))

    # Coagir números
    plan = garantir_numerico(plan, mes_cols)
    req  = garantir_numerico(req,  mes_cols)

    # Chaves de agrupamento
    grp = [c for c in ['SITE','PRODUCT SERIES'] if c in plan.columns and c in req.columns]
    if len(grp) < 2:
        raise ValueError('As colunas SITE e PRODUCT SERIES precisam existir em PLAN e REQUEST.')

    # Agregações mensais
    plan_agg = plan[grp + mes_cols].groupby(grp, dropna=False)[mes_cols].sum().reset_index()
    req_agg  = req [grp + mes_cols].groupby(grp, dropna=False)[mes_cols].sum().reset_index()

    # Merge e cálculo do delta mensal (REQUEST - PLAN)
    comp = pd.merge(req_agg, plan_agg, on=grp, how='outer', suffixes=('_REQ','_PLAN'))
    for m in mes_cols:
        comp[m] = comp.get(f'{m}_REQ', 0).fillna(0).astype(int) - comp.get(f'{m}_PLAN', 0).fillna(0).astype(int)

    # Manter apenas chaves + meses e calcular TOTAL
    step1 = comp[grp + mes_cols].copy()
    step1['TOTAL'] = step1[mes_cols].sum(axis=1)

    # Ordenar por SITE e TOTAL desc
    step1 = step1.sort_values(by=['SITE','TOTAL'], ascending=[True, False])

    # Linha TOTAL GERAL (soma por coluna)
    linha_total = {k:'TOTAL GERAL' for k in grp}
    for m in mes_cols:
        linha_total[m] = int(step1[m].sum())
    linha_total['TOTAL'] = int(step1['TOTAL'].sum())
    step1 = pd.concat([step1, pd.DataFrame([linha_total])], ignore_index=True)

    # Gravar de volta no mesmo arquivo
    with pd.ExcelWriter(xlsx_path, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
        step1.to_excel(writer, sheet_name=out_sheet, index=False)


if __name__ == '__main__':
    parser = argparse.ArgumentParser(description='Passo 1 — Comparativo Mensal por PRODUCT SERIES (REQUEST − PLAN).')
    parser.add_argument('--arquivo', required=True, help='Caminho do arquivo Excel de entrada (contendo abas PLAN e REQUEST).')
    parser.add_argument('--aba_saida', default='Step1_Comparativo_Serie', help='Nome da aba de saída (default: Step1_Comparativo_Serie).')
    args = parser.parse_args()

    xlsx_path = Path(args.arquivo)
    if not xlsx_path.exists():
        raise FileNotFoundError(f"Arquivo não encontrado: {xlsx_path}")

    comparativo_request_vs_plan_por_serie(xlsx_path, out_sheet=args.aba_saida)
    print(f"Aba '{args.aba_saida}' gerada/atualizada em: {xlsx_path}")
