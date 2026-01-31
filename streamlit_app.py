import streamlit as st
import pandas as pd
from io import BytesIO
from openpyxl.styles import Font, PatternFill, Alignment

st.set_page_config(page_title="Inventário Filial 944", page_icon="📝")

st.title("📝 Gerador de Inventário com Título")
st.write("Cada aba terá o título: 'inventario filial 944 - [Tipo]'")

uploaded_file = st.file_uploader("Escolha o arquivo CSV", type="csv")

if uploaded_file is not None:
    try:
        df = pd.read_csv(uploaded_file, sep=';')
        
        # --- LÓGICA DE DEFINIÇÃO DAS ABAS ---
        def definir_aba(linha):
            tipo = str(linha['TIPO']).upper()
            sub_tipo = str(linha['SUB TIPO']).upper()
            
            # Scanners de Mão
            if tipo == 'SCANER' and 'MÃO' in sub_tipo:
                return 'SCANER DE MÃO'
            
            # Infraestrutura (Servidor)
            tipos_servidor = ['SERVIDOR', 'TAPE', 'RACK', 'STORAGE']
            if tipo in tipos_servidor:
                return 'SERVIDOR'
            
            return tipo

        df['ABA_DESTINO'] = df.apply(definir_aba, axis=1)
        colunas_remover = ['FILIAL', 'TIPO', 'SUB TIPO', 'COMPLEMENTO', 'ABA_DESTINO']

        if st.button("🚀 Gerar Planilha com Títulos"):
            output = BytesIO()
            
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                abas = sorted(df['ABA_DESTINO'].unique())
                
                for nome_aba in abas:
                    grupo = df[df['ABA_DESTINO'] == nome_aba]
                    grupo_ordenado = grupo.sort_values(by=['PIP'], ascending=True)
                    
                    nome_final_aba = str(nome_aba)[:31].replace('/', '-')
                    tabela_final = grupo_ordenado.drop(columns=colunas_remover, errors='ignore')
                    
                    # 1. Escrever a tabela começando na linha 2 (startrow=1)
                    # O pandas conta a partir de 0, então startrow=1 é a segunda linha do Excel
                    tabela_final.to_excel(writer, sheet_name=nome_final_aba, index=False, startrow=1)
                    
                    ws = writer.sheets[nome_final_aba]
                    
                    # 2. Inserir o Título na primeira linha
                    titulo = f"inventario filial 944 - {nome_final_aba}"
                    ws.cell(row=1, column=1).value = titulo
                    
                    # Mesclar as células do título (da coluna A até a última coluna da tabela)
                    num_colunas = len(tabela_final.columns)
                    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=num_colunas)
                    
                    # Estilo do Título (Cinza claro, negrito, centralizado)
                    titulo_cell = ws.cell(row=1, column=1)
                    titulo_cell.font = Font(size=14, bold=True, color="000000")
                    titulo_cell.fill = PatternFill(start_color="D9D9D9", end_color="D9D9D9", fill_type="solid")
                    titulo_cell.alignment = Alignment(horizontal="center", vertical="center")
                    
                    # 3. Estilo do Cabeçalho da Tabela (Linha 2)
                    header_fill = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
                    header_font = Font(color="FFFFFF", bold=True)
                    
                    for cell in ws[2]: # Agora o cabeçalho está na linha 2
                        cell.fill = header_fill
                        cell.font = header_font
                        cell.alignment = Alignment(horizontal="center")

                    # 4. Ajuste automático de colunas
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
                label="📥 Baixar Inventário com Títulos",
                data=output.getvalue(),
                file_name="Inventario_Filial_944_Titulos.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
            st.success("Tabelas geradas com títulos no topo!")

    except Exception as e:
        st.error(f"Erro: {e}")
        
