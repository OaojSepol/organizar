import streamlit as st
import pandas as pd
from io import BytesIO

st.set_page_config(page_title="Organizador de Inventário", page_icon="📊")

st.title("📊 Organizador de Arquivos")
st.write("Envie o CSV para gerar o Excel sem as colunas Filial, Tipo, Sub Tipo e Complemento.")

uploaded_file = st.file_uploader("Escolha o arquivo CSV", type="csv")

if uploaded_file is not None:
    try:
        # Lendo o arquivo original
        df = pd.read_csv(uploaded_file, sep=';')
        
        st.success("Arquivo carregado!")

        if st.button("Gerar Arquivo Excel"):
            output = BytesIO()
            
            # Colunas que o usuário deseja remover do conteúdo das abas
            colunas_para_remover = ['FILIAL', 'TIPO', 'SUB TIPO', 'COMPLEMENTO']
            
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                # Agrupamos por TIPO para criar as abas
                for categoria, dados in df.groupby('TIPO'):
                    nome_aba = str(categoria)[:31].replace('/', '-').replace('*', '').replace('?', '')
                    
                    # Removemos as colunas indesejadas apenas para salvar na aba
                    # Usamos errors='ignore' caso alguma coluna venha com nome ligeiramente diferente
                    dados_limpos = dados.drop(columns=colunas_para_remover, errors='ignore')
                    
                    dados_limpos.to_excel(writer, sheet_name=nome_aba, index=False)
            
            processed_data = output.getvalue()
            
            st.download_button(
                label="📥 Baixar Excel Organizado",
                data=processed_data,
                file_name="Inventario_Limpo.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
            
    except Exception as e:
        st.error(f"Erro: {e}")
        
