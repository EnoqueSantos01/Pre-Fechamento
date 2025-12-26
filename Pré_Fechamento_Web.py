import streamlit as st
import pandas as pd
from io import BytesIO

st.title("Pré - Fechamento 🧾")

uploaded_file = st.file_uploader("Faça upload do arquivo Excel", type=["xlsx"])


if uploaded_file:
    try:
        planilha = pd.read_excel(uploaded_file)
    except Exception as e:
        st.error(f"Erro ao carregar o arquivo: {e}")
        st.stop()

    planilha['Retorno SEFAZ'] = pd.to_numeric(
    planilha['Retorno SEFAZ'], 
    errors='coerce'
    )

    # Verifica se colunas existem
    colunas_necessarias = [
        'CFOP', 'Vlr ICMS Com', 'Desc. Produto', 'Retorno SEFAZ', 'Dt. Canc.',
        'Tp. Mov', 'Chave Doc', 'Icms Ret', 'Difal ICMS', 'Documento', 'Especie'
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

    filtro5 = (planilha['Tp. Mov'] == 'SAIDA') & (planilha['Chave Doc'].isna() | (planilha['Chave Doc'].str.strip() == '')) & (planilha['Retorno SEFAZ'] != 102) 
    planilha.loc[filtro5, 'Observações'] = 'NF de saída não pode estar sem chave'

    filtro6 = (planilha['CFOP'] == 6101) & (planilha['Icms Ret'] == 0) & (planilha['Difal ICMS'] == 0)
    planilha.loc[filtro6, 'Observações'] = 'ICMS Ret ou Difal ICMS deve ter valor'

    filtro7 = (planilha['CFOP'] == 6107) & (planilha['Icms Ret'] == 0) & (planilha['Difal ICMS'] == 0)
    planilha.loc[filtro7, 'Observações'] = 'ICMS Ret ou Difal ICMS deve ter valor'

    filtro8 = (planilha['CFOP'] == 6108) & (planilha['Icms Ret'] == 0) & (planilha['Difal ICMS'] == 0)
    planilha.loc[filtro8, 'Observações'] = 'ICMS Ret ou Difal ICMS deve ter valor'
    
    # Selecionar apenas notas de SAIDA
    notas_saida = planilha[planilha['Tp. Mov'] == 'SAIDA'].copy()
    notas_entrada = planilha[planilha['Tp. Mov'] == 'ENTRADA'].copy()

    # Garantir documento numerico
    notas_saida['Documento'] = pd.to_numeric(notas_saida['Documento'], errors='coerce')
    notas_entrada['Documento'] = pd.to_numeric(notas_entrada['Documento'], errors='coerce')
    
    # Ordenar pela numeração
    notas_saida = sorted(notas_saida['Documento'].dropna().astype(int).tolist())
    notas_entradas = set(notas_entrada['Documento'].dropna().astype(int).tolist())

    # Percorrer sequÇencia das SAÍDAS
    
    for i in range(len(notas_saida) - 1):
        atual = notas_saida[i]
        proximo = notas_saida[i + 1]

        # Se houver quebra de sequência
        
        if proximo > atual + 1:
            # Identifica notas faltando
            faltando = range(atual + 1, proximo)

            faltantes = []

            for nf in faltando:
                # Se não existe nem na SAIDA nem na ENTRADA
                if nf not in notas_saida and nf not in notas_entrada:
                    faltantes.append(nf)
            # Se realmente hpuver quebra válida
            if faltantes:
                mensagem = "; ".join(
                    [f"NF número {nf} não encontrada na sequência" for nf in faltantes]
                ) + "; "

                # Aplica observação SOMENTE na linha da NF nde houve a quebra
                filtro_linha = (
                    (planilha['Tp. Mov'] == 'SAIDA') &
                    (planilha['Documento'] == proximo)
                )
                      
                obs_atual = planilha.loc[filtro_linha, 'Observações'].fillna("").astype(str)
    
                planilha.loc[filtro_linha, 'Observações'] = obs_atual + mensagem

        # Mapa de espécies e CFOPs válidos
    cfop_validos = {
        "CTE":  [1352, 2352, 2932, 1932, 1353],
        "NFCEE": [1252, 2252],
        "NFS":  [1933, 2933],
        "NFSC": [1302, 2302],
        "NTST": [1302, 2302],
    }
    
    # Garantir que CFOP está numérico
    planilha['CFOP'] = pd.to_numeric(planilha['CFOP'], errors='coerce')
    
    for especie, lista_cfop in cfop_validos.items():
    
        # Filtrar linhas da espécie específica
        filtro_especie = planilha['Especie'] == especie
    
        # Filtrar CFOP incorreto
        filtro_cfop_errado = ~planilha['CFOP'].isin(lista_cfop)
    
        # Filtro final: especie corresponde, mas CFOP não corresponde
        filtro10 = filtro_especie & filtro_cfop_errado
    
        # Mensagem
        mensagem = f"Espécie {especie} incompatível com o CFOP; "
    
        # Adicionar sem substituir
        obs_atual = planilha.loc[filtro10, 'Observações'].fillna("").astype(str)
        planilha.loc[filtro10, 'Observações'] = obs_atual + mensagem



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













