import streamlit as st
import pandas as pd
from io import BytesIO

st.title("Pré — Fechamento Fiscal")

uploaded_file = st.file_uploader("Faça upload do arquivo Excel", type=["xlsx"])

if uploaded_file:
    try:
        planilha = pd.read_excel(uploaded_file)
    except Exception as e:
        st.error(f"Erro ao carregar o arquivo: {e}")
        st.stop()

    # Verifica se colunas existem
    colunas_necessarias = [
        'CFOP', 'Vlr ICMS Com', 'Desc. Produto', 'Retorno SEFAZ', 'Dt. Canc.',
        'Tp. Mov', 'Chave Doc', 'Icms Ret', 'Difal ICMS'
    ]

    for coluna in colunas_necessarias:
        if coluna not in planilha.columns:
            st.error(f"A coluna obrigatória '{coluna}' não existe no arquivo.")
            st.stop()

    # Coluna Observações
    planilha["Observações"] = ""

    # Regras
    filtro = (planilha['CFOP'] == 2556) & (planilha['Vlr ICMS Com'] == 0)
    planilha.loc[filtro, 'Observações'] = '2556 não pode estar zerado o Vlr ICMS Com'

    filtro1 = (planilha["CFOP"] == 2551) & (planilha["Vlr ICMS Com"] == 0)
    planilha.loc[filtro1, 'Observações'] = '2551 não pode estar zerado o Vlr ICMS Com'

    filtro2 = (planilha['CFOP'] == 2352) & (planilha['Desc. Produto'] == '16.02 - FRETE DIVERSOS') & (planilha["Vlr ICMS Com"] == 0)
    planilha.loc[filtro2, 'Observações'] = '2352 com Frete para Uso e Consumo não pode estar zerado'

    filtro3 = (planilha['Retorno SEFAZ'] > 0) & (planilha['Retorno SEFAZ'] > 100)
    planilha.loc[filtro3, 'Observações'] = 'Verifique se a NF está realmente cancelada'

    filtro4 = (planilha['Dt. Canc.'] != '/  /')
    planilha.loc[filtro4, 'Observações'] = 'Verifique se a NF está cancelada'

    # --- Preparação segura das colunas ---
    # Garante que Tp. Mov seja string e compare em maiúsculas
    tp_mov = planilha['Tp. Mov'].astype(str).str.strip().str.upper()

    # Garante que Chave Doc vire string (vazia quando NaN) e faça strip
    chave_doc = planilha['Chave Doc'].where(planilha['Chave Doc'].notna(), '').astype(str).str.strip()

    # Converte Retorno SEFAZ para número quando possível; valores inválidos viram NaN
    retorno = pd.to_numeric(planilha['Retorno SEFAZ'], errors='coerce')

    # Marca se o retorno é 101 ou 102 (True para esses casos)
    retorno_101_102 = retorno.isin([101, 102])

    # --- Filtro final: SAIDA, sem chave, e RETORNO SEFAZ diferente de 101/102 ---
    filtro5 = (
        (tp_mov == 'SAIDA') &
        (chave_doc == '') &
        (~retorno_101_102)
    )

    # Aplicar observação
    planilha.loc[filtro5, 'Observações'] = 'NF de saída não pode estar sem chave'

    # --- DEBUG: informações para checar o que está marcando ---
    num_marked = filtro5.sum()
    st.write(f"Linhas marcadas pelo filtro5: {num_marked}")

    # Mostra algumas colunas relevantes para inspecionar (local) — no Streamlit use st.write/ st.dataframe
    debug_cols = ['Tp. Mov', 'Chave Doc', 'Retorno SEFAZ']
    st.dataframe(planilha.loc[filtro5, debug_cols].head(30))


    filtro6 = (planilha['CFOP'] == 6101) & (planilha['Icms Ret'] == 0) & (planilha['Difal ICMS'] == 0)
    planilha.loc[filtro6, 'Observações'] = 'ICMS Ret ou Difal ICMS deve ter valor'

    filtro7 = (planilha['CFOP'] == 6107) & (planilha['Icms Ret'] == 0) & (planilha['Difal ICMS'] == 0)
    planilha.loc[filtro7, 'Observações'] = 'ICMS Ret ou Difal ICMS deve ter valor'

    filtro8 = (planilha['CFOP'] == 6108) & (planilha['Icms Ret'] == 0) & (planilha['Difal ICMS'] == 0)
    planilha.loc[filtro8, 'Observações'] = 'ICMS Ret ou Difal ICMS deve ter valor'

    

    st.success("Análise concluída!")

    # Download do arquivo gerado
    output = BytesIO()
    planilha.to_excel(output, index=False)
    output.seek(0)

    st.download_button(
        label="Baixar arquivo processado",
        data=output,
        file_name="resultado_validado.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )




