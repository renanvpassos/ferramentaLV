import streamlit as st
import pandas as pd
import io
import re

def limpar_numero(valor):
    """Remove caracteres não numéricos e converte para float."""
    if pd.isna(valor) or valor == "": 
        return 0.0
    valor_str = str(valor)
    valor_limpo = re.sub(r'[^\d,.]', '', valor_str).replace(',', '.')
    try:
        return float(valor_limpo)
    except ValueError:
        return 0.0

def tratar_planilha(file):
    # Carrega a planilha de origem
    df_origem = pd.read_excel(file, header=0)
    df_origem.columns = [str(col).strip() for col in df_origem.columns]

    # --- DICIONÁRIO DE MAPEAMENTO ---
    C_PARTNUMBER = "CODIGO PRINCIPAL"
    C_QUANTIDADE = "QUANTIDADE"
    C_DESCRICAO = "DESCRICAO PORTUGUES"
    C_PRECO_UNIT = "VALOR UNITARIO ITEM"
    C_PESO_UNIT = "PESO LIQUIDO UNITÁRIO"
    C_FATURA = "FATURA"
    C_ORDEM_COMPRA = "ORDEM DE COMPRA"
    C_CFOP = "CFOP"
    C_UNIDADE_MEDIDA = "UNIDADE MEDIDA"

    colunas_obrigatorias = [C_PARTNUMBER, C_QUANTIDADE, C_DESCRICAO, C_PRECO_UNIT, C_PESO_UNIT, C_FATURA, C_ORDEM_COMPRA, C_CFOP]

    # Verifica colunas
    colunas_faltantes = [col for col in colunas_obrigatorias if col not in df_origem.columns]
    if colunas_faltantes:
        raise ValueError(f"Cabeçalhos faltando: {', '.join(colunas_faltantes)}")

    # Estrutura Final (Adicionado CFOP)
    colunas_finais = ['PARTNUMBER', 'QUANTIDADE', 'UNIDADE', 'PRECOTOTAL', 'PESOTOTAL', 'INCOTERMS', 'MOEDA', 'FATURA', 'CFOP', 'PEDIDO']
    df_final = pd.DataFrame(columns=colunas_finais)

    df_final['PARTNUMBER'] = df_origem[C_PARTNUMBER]
    df_final['QUANTIDADE'] = df_origem[C_QUANTIDADE]

    df_final['UNIDADE'] = df_origem[C_UNIDADE_MEDIDA].apply(limpar_numero)

    # Aplicação da limpeza e cálculo
    qtd_num = df_origem[C_QUANTIDADE].apply(limpar_numero)
    preco_num = df_origem[C_PRECO_UNIT].apply(limpar_numero)
    peso_num = df_origem[C_PESO_UNIT].apply(limpar_numero)

    df_final['PRECOTOTAL'] = preco_num * qtd_num
    df_final['PESOTOTAL'] = peso_num * qtd_num 

    df_final['INCOTERMS'] = 'FCA'
    df_final['MOEDA'] = '790'
    df_final['FATURA'] = df_origem[C_FATURA].apply(limpar_numero)
    
    # Atribuição da coluna CFOP
    df_final['CFOP'] = df_origem[C_CFOP]

    df_final['PEDIDO'] = df_origem[C_ORDEM_COMPRA].apply(limpar_numero)

    return df_final

# --- INTERFACE (STREAMLIT) ---
st.set_page_config(page_title="Tratador de Planilhas", layout="centered")
st.title("📂 Tratamento planilha Louis Vuitton")

uploaded_file = st.file_uploader("Selecione o arquivo Excel de origem (.xlsx)", type=["xlsx"])

if uploaded_file:
    if st.button("Processar Planilha", use_container_width=True):
        try:
            with st.spinner("Processando dados..."):
                resultado = tratar_planilha(uploaded_file)
                
                output = io.BytesIO()
                with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                    resultado.to_excel(writer, index=False, sheet_name='Planilha Tratada')
                
                st.success("Planilha processada com sucesso!")
                st.download_button(
                    label="📥 Clique aqui para baixar a Planilha Tratada",
                    data=output.getvalue(),
                    file_name="planilha_tratada.xlsx",
                    use_container_width=True
                )
        except Exception as e:
            st.error(f"Erro ao processar: {e}")
