import streamlit as st
import pandas as pd
from io import BytesIO
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter

st.set_page_config(page_title="Inventário Filial 944", page_icon="🖥️")

st.title("🖥️ Organizador de Inventário Profissional")
st.write("Separação de CPUs (PDV vs Escritório), Scanners, Impressoras e Infraestrutura.")

uploaded_file = st.file_uploader("Escolha o arquivo CSV", type="csv")

if uploaded_file is not None:
    try:
        # Lendo o arquivo e padronizando cabeçalhos
        df = pd.read_csv(uploaded_file, sep=';')
        df.columns = [c.strip().upper() for c in df.columns]
        
        # --- LÓGICA DE DEFINIÇÃO DAS ABAS ---
        def definir_aba(linha):
            tipo = str(linha.get('TIPO', '')).upper().strip()
            sub_tipo = str(linha.get('SUB TIPO', '')).upper()
            complemento = str(linha.get('COMPLEMENTO', '')).upper()
            
            # 1. Scanners
            if tipo == 'SCANER':
                if 'MÃO' in sub_tipo or 'MÃO' in complemento:
                    return 'SCANER DE MÃO'
                return 'SCANER'
            
            # 2. Infraestrutura (SERVIDOR, TAPE, RACK, STORAGE)
            infra = ['SERVIDOR', 'TAPE', 'RACK', 'STORAGE']
            if tipo in infra:
                return 'SERVIDOR'
            
            # 3. Categorias de IMPRESSORA
            if tipo == 'IMPRESSORA':
                if 'CHEQUE' in sub_tipo or 'CHEQUE' in complemento:
                    return 'IMPRESSORA DE CHEQUE'
                if 'CHECK-IN' in sub_tipo or 'CHECK-IN' in complemento:
                    return 'IMPRESSORA CHECK-IN'
                if 'TERMICA' in sub_tipo or 'TERMICA' in complemento:
                    return 'IMPRESSORA TÉRMICA'
                return 'IMPRESSORA'
            
            # 4. CPUs (PDV vs Escritório)
            if tipo == 'CPU':
                if 'PDV' in sub_tipo or 'PDV' in complemento:
                    return 'CPU PDV'
                return 'CPU ESCRITÓRIO'
            
            # 5. Outros (MONITOR, SAT, etc)
            return tipo if tipo != "" else "OUTROS"

        # Aplicar a classificação
        df['ABA_DESTINO'] = df.apply(definir_aba, axis=1)
        
        colunas_remover = ['FILIAL', 'TIPO', 'SUB TIPO', 'COMPLEMENTO', 'ABA_DESTINO']

        if st.button("🚀 Gerar Inventário Completo"):
            output = BytesIO()
            
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                lista_abas = sorted(df['ABA_DESTINO'].unique())
                
                for nome_aba in lista_abas:
                    grupo = df[df['ABA_DESTINO'] == nome_aba].copy()
                    if grupo.empty: continue
                        
                    grupo = grupo.sort_values(by=['PIP'], ascending=True)
                    nome_final_aba = str(nome_aba)[:31].replace('/', '-')
                    tabela_final = grupo.drop(columns=colunas_remover, errors='ignore')
                    
                    # Salva a tabela na linha 2
                    tabela_final.to_excel(writer, sheet_name=nome_final_aba, index=False, startrow=1)
                    ws = writer.sheets[nome_final_aba]
                    
                    # --- Título na linha 1 ---
                    ws.cell(row=1, column=1).value = f"inventario filial 944 - {nome_final_aba}"
                    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=len(tabela_final.columns))
                    
                    # Estilo do Título (Cinza)
                    cell_titulo = ws.cell(row=1, column=1)
                    cell_titulo.font = Font(size=12, bold=True)
                    cell_titulo.fill = PatternFill(start_color="D9D9D9", end_color="D9D9D9", fill_type="solid")
                    cell_titulo.alignment = Alignment(horizontal="center")
                    
                    # Estilo do Cabeçalho (Linha 2 - Azul)
                    for cell in ws[2]:
                        cell.fill = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
                        cell.font = Font(color="FFFFFF", bold=True)
                        cell.alignment = Alignment(horizontal="center")

                    # Ajuste de largura das colunas
                    for i, col in enumerate(tabela_final.columns, 1):
                        column_letter = get_column_letter(i)
                        # Cálculo do comprimento máximo
                        vals = [len(str(x)) for x in grupo[col].values]
                        max_len = max(vals + [len(col)]) if vals else len(col)
                        ws.column_dimensions[column_letter].width = max_len + 5

            st.download_button(
                label="📥 Baixar Inventário Final",
                data=output.getvalue(),
                file_name="Inventario_Filial_944_Completo.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
            st.success("Planilha gerada com todas as separações!")

    except Exception as e:
        st.error(f"Erro ao processar: {e}")
