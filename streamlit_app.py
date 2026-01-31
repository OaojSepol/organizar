import streamlit as st
import pandas as pd
from io import BytesIO
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter

st.set_page_config(page_title="Inventário Filial 944", page_icon="📝")

st.title("📝 Gerador de Inventário - Versão Estável")
st.write("Correção do erro de células mescladas aplicada.")

uploaded_file = st.file_uploader("Escolha o arquivo CSV", type="csv")

if uploaded_file is not None:
    try:
        df = pd.read_csv(uploaded_file, sep=';')
        
        def definir_aba(linha):
            tipo = str(linha['TIPO']).upper()
            sub_tipo = str(linha['SUB TIPO']).upper()
            if tipo == 'SCANER' and 'MÃO' in sub_tipo:
                return 'SCANER DE MÃO'
            tipos_servidor = ['SERVIDOR', 'TAPE', 'RACK', 'STORAGE']
            if tipo in tipos_servidor:
                return 'SERVIDOR'
            return tipo

        df['ABA_DESTINO'] = df.apply(definir_aba, axis=1)
        colunas_remover = ['FILIAL', 'TIPO', 'SUB TIPO', 'COMPLEMENTO', 'ABA_DESTINO']

        if st.button("🚀 Gerar Planilha Corrigida"):
            output = BytesIO()
            
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                abas = sorted(df['ABA_DESTINO'].unique())
                
                for nome_aba in abas:
                    grupo = df[df['ABA_DESTINO'] == nome_aba]
                    grupo_ordenado = grupo.sort_values(by=['PIP'], ascending=True)
                    
                    nome_final_aba = str(nome_aba)[:31].replace('/', '-')
                    tabela_final = grupo_ordenado.drop(columns=colunas_remover, errors='ignore')
                    
                    # Escreve a tabela na linha 2
                    tabela_final.to_excel(writer, sheet_name=nome_final_aba, index=False, startrow=1)
                    
                    ws = writer.sheets[nome_final_aba]
                    
                    # Título na linha 1
                    titulo = f"inventario filial 944 - {nome_final_aba}"
                    ws.cell(row=1, column=1).value = titulo
                    num_colunas = len(tabela_final.columns)
                    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=num_colunas)
                    
                    # Estilo do Título
                    titulo_cell = ws.cell(row=1, column=1)
                    titulo_cell.font = Font(size=12, bold=True)
                    titulo_cell.fill = PatternFill(start_color="D9D9D9", end_color="D9D9D9", fill_type="solid")
                    titulo_cell.alignment = Alignment(horizontal="center", vertical="center")
                    
                    # Estilo do Cabeçalho (Linha 2)
                    header_fill = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
                    header_font = Font(color="FFFFFF", bold=True)
                    for cell in ws[2]:
                        cell.fill = header_fill
                        cell.font = header_font
                        cell.alignment = Alignment(horizontal="center")

                    # --- CORREÇÃO DO REDIMENSIONAMENTO ---
                    # Ajustamos a largura baseada apenas nas colunas da tabela de dados
                    for i, col_name in enumerate(tabela_final.columns, 1):
                        column_letter = get_column_letter(i)
                        
                        # Encontra o maior valor na coluna (incluindo o cabeçalho)
                        max_length = len(str(col_name))
                        for row in range(2, ws.max_row + 1):
                            cell_value = ws.cell(row=row, column=i).value
                            if cell_value:
                                length = len(str(cell_value))
                                if length > max_length:
                                    max_length = length
                        
                        ws.column_dimensions[column_letter].width = max_length + 4

            st.download_button(
                label="📥 Baixar Inventário",
                data=output.getvalue(),
                file_name="Inventario_Filial_944_OK.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
            st.success("Planilha gerada sem erros!")

    except Exception as e:
        st.error(f"Erro crítico: {e}")
