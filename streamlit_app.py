import streamlit as st
import pandas as pd
from io import BytesIO
from openpyxl.styles import Font, PatternFill, Alignment

st.set_page_config(page_title="Organizador Pro", page_icon="🎨")

st.title("🎨 Organizador de Inventário Formatado")
st.write("Gere um Excel com abas, colunas auto-ajustáveis e cabeçalho colorido.")

uploaded_file = st.file_uploader("Envie o CSV", type="csv")

if uploaded_file is not None:
    try:
        df = pd.read_csv(uploaded_file, sep=';')
        colunas_remover = ['FILIAL', 'TIPO', 'SUB TIPO', 'COMPLEMENTO']

        if st.button("🚀 Gerar Excel Formatado"):
            output = BytesIO()
            
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                for categoria, grupo in df.groupby('TIPO'):
                    nome_aba = str(categoria)[:31].replace('/', '-')
                    tabela_limpa = grupo.drop(columns=colunas_remover, errors='ignore')
                    
                    # Salva os dados na aba
                    tabela_limpa.to_excel(writer, sheet_name=nome_aba, index=False)
                    
                    # --- INÍCIO DA FORMATAÇÃO ---
                    worksheet = writer.sheets[nome_aba]
                    
                    # 1. Formatar Cabeçalho (Azul escuro com texto branco e negrito)
                    header_fill = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
                    header_font = Font(color="FFFFFF", bold=True)
                    alignment = Alignment(horizontal="center", vertical="center")

                    for cell in worksheet[1]: # Primeira linha
                        cell.fill = header_fill
                        cell.font = header_font
                        cell.alignment = alignment

                    # 2. Ajustar largura das colunas automaticamente
                    for col in worksheet.columns:
                        max_length = 0
                        column = col[0].column_letter # Letra da coluna
                        
                        for cell in col:
                            try:
                                if len(str(cell.value)) > max_length:
                                    max_length = len(str(cell.value))
                            except:
                                pass
                        
                        adjusted_width = (max_length + 2)
                        worksheet.column_dimensions[column].width = adjusted_width
                    # --- FIM DA FORMATAÇÃO ---

            st.download_button(
                label="📥 Baixar Excel Profissional",
                data=output.getvalue(),
                file_name="Inventario_Formatado.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
            st.success("Arquivo pronto e formatado!")

    except Exception as e:
        st.error(f"Erro: {e}")
