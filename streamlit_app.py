import streamlit as st
import pandas as pd
from io import BytesIO
from openpyxl.styles import Font, PatternFill, Alignment

st.set_page_config(page_title="Organizador de Inventário", page_icon="🖥️")

st.title("🖥️ Organizador de Inventário - Filial 944")
st.write("Agrupa Tapes, Racks, Storages e Servidores na aba 'SERVIDOR' e organiza o restante.")

uploaded_file = st.file_uploader("Envie o arquivo CSV", type="csv")

if uploaded_file is not None:
    try:
        # Lendo o arquivo
        df = pd.read_csv(uploaded_file, sep=';')
        
        # 1. Definir a regra de agrupamento especial
        # Mapeamos os tipos que devem ir para a aba "SERVIDOR"
        tipos_servidor = ['SERVIDOR', 'TAPE', 'RACK', 'STORAGE']
        
        # Criamos uma coluna temporária para definir o nome da aba
        df['ABA_DESTINO'] = df['TIPO'].apply(lambda x: 'SERVIDOR' if str(x).upper() in tipos_servidor else x)
        
        colunas_remover = ['FILIAL', 'TIPO', 'SUB TIPO', 'COMPLEMENTO', 'ABA_DESTINO']

        if st.button("🚀 Gerar Planilha Organizada"):
            output = BytesIO()
            
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                # Agrupamos pela nova coluna 'ABA_DESTINO'
                for nome_aba, grupo in df.groupby('ABA_DESTINO'):
                    # Ordenar os dados internamente (se for servidor, organiza por tipo original)
                    grupo_ordenado = grupo.sort_values(by=['TIPO', 'PIP'])
                    
                    # Nome da aba (limite de 31 caracteres)
                    nome_final_aba = str(nome_aba)[:31].replace('/', '-')
                    
                    # Limpeza das colunas
                    tabela_final = grupo_ordenado.drop(columns=colunas_remover, errors='ignore')
                    
                    # Salva na aba
                    tabela_final.to_excel(writer, sheet_name=nome_final_aba, index=False)
                    
                    # --- Estilização ---
                    ws = writer.sheets[nome_final_aba]
                    header_fill = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
                    header_font = Font(color="FFFFFF", bold=True)
                    header_align = Alignment(horizontal="center", vertical="center")

                    for cell in ws[1]:
                        cell.fill = header_fill
                        cell.font = header_font
                        cell.alignment = header_align

                    # Ajuste automático de colunas
                    for col in ws.columns:
                        max_length = 0
                        column_letter = col[0].column_letter
                        for cell in col:
                            try:
                                if len(str(cell.value)) > max_length:
                                    max_length = len(str(cell.value))
                            except: pass
                        ws.column_dimensions[column_letter].width = max_length + 4

            st.download_button(
                label="📥 Baixar Inventário Atualizado",
                data=output.getvalue(),
                file_name="Inventario_Filial_944.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
            st.success("Pronto! Aba SERVIDOR unificada e formatada.")

    except Exception as e:
        st.error(f"Erro: {e}")
        
