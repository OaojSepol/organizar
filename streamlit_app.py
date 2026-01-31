import streamlit as st
import pandas as pd
from io import BytesIO

st.set_page_config(page_title="Organizador de Inventário", page_icon="📊")

st.title("📊 Organizador de Arquivos (Filial 944)")
st.write("Envie seu arquivo CSV e baixe o Excel organizado por categorias (abas).")

# Upload do arquivo
uploaded_file = st.file_uploader("Escolha o arquivo CSV", type="csv")

if uploaded_file is not None:
    try:
        # Lendo o CSV (usando o separador ';' que está no seu arquivo)
        df = pd.read_csv(uploaded_file, sep=';')
        
        st.success("Arquivo carregado com sucesso!")
        st.write("### Prévia dos dados:", df.head())

        # Botão para processar e baixar
        if st.button("Gerar Arquivo Excel"):
            output = BytesIO()
            
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                # Agrupar por 'TIPO' e criar as abas
                for categoria, dados in df.groupby('TIPO'):
                    nome_aba = str(categoria)[:31].replace('/', '-').replace('*', '').replace('?', '')
                    dados.to_excel(writer, sheet_name=nome_aba, index=False)
            
            processed_data = output.getvalue()
            
            st.download_button(
                label="📥 Baixar Excel Organizado",
                data=processed_data,
                file_name="Inventario_Organizado.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
            
    except Exception as e:
        st.error(f"Erro ao processar o arquivo: {e}")

else:
    st.info("Aguardando upload do arquivo CSV...")
    
