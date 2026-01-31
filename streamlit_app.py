import streamlit as st
import pandas as pd
from io import BytesIO
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter

st.set_page_config(page_title="Inventário Filial 944", page_icon="📝")

st.title("📝 Gerador de Inventário - Filial 944")
st.write("Organização por abas: Scaner de Mão vs Scaner de Mesa, e Servidores agrupados.")

uploaded_file = st.file_uploader("Escolha o arquivo CSV", type="csv")

if uploaded_file is not None:
    try:
        # Lendo o arquivo (garantindo que os nomes das colunas fiquem em maiúsculo para evitar erro)
        df = pd.read_csv(uploaded_file, sep=';')
        df.columns = [c.strip().upper() for c in df.columns]
        
        # --- LÓGICA DE DEFINIÇÃO DAS ABAS ---
        def definir_aba(linha):
            # Forçamos a leitura para maiúsculo para comparar
            tipo_original = str(linha.get('TIPO', '')).upper().strip()
            sub_tipo = str(linha.get('SUB TIPO', '')).upper()
            complemento = str(linha.get('COMPLEMENTO', '')).upper()
            
            # 1. Regra para SCANER DE MÃO (Se tiver a palavra MÃO no sub-tipo ou complemento)
            if tipo_original == 'SCANER' and ('MÃO' in sub_tipo or 'MÃO' in complemento):
                return 'SCANER DE MÃO'
            
            # 2. Regra para SCANER NORMAL (Mesa/Outros)
            if tipo_original == 'SCANER':
                return 'SCANER'
            
            # 3. Regra para a aba SERVIDOR (Unificada: Servidor, Tape, Rack, Storage)
            infra = ['SERVIDOR', 'TAPE', 'RACK', 'STORAGE']
            if tipo_original in infra:
                return 'SERVIDOR'
            
            # 4. Outros (MONITOR, CPU, etc)
            return tipo_original if tipo_original != "" else "OUTROS"

        # Criar a coluna de destino
        df['ABA_DESTINO'] = df.apply(definir_aba, axis=1)
        
        # Colunas que serão removidas da visualização final
        colunas_remover = ['FILIAL', 'TIPO', 'SUB TIPO', 'COMPLEMENTO', 'ABA_DESTINO']

        if st.button("🚀 Gerar Planilha"):
            output = BytesIO()
            
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                # Pegar nomes das abas únicos
                lista_abas = sorted(df['ABA_DESTINO'].unique())
                
                for nome_aba in lista_abas:
                    # Filtra o grupo correspondente à aba
                    grupo = df[df['ABA_DESTINO'] == nome_aba].copy()
                    
                    if grupo.empty:
                        continue
                        
                    # Ordenar por PIP
                    grupo = grupo.sort_values(by=['PIP'], ascending=True)
                    
                    # Nome da aba (máximo 31 caracteres)
                    nome_final_aba = str(nome_aba)[:31].replace('/', '-')
                    
                    # Limpa as colunas para o Excel
                    tabela_final = grupo.drop(columns=colunas_remover, errors='ignore')
                    
                    # Salva na linha 2 (startrow=1)
                    tabela_final.to_excel(writer, sheet_name=nome_final_aba, index=False, startrow=1)
                    
                    ws = writer.sheets[nome_final_aba]
                    
                    # --- Título na linha 1 ---
                    ws.cell(row=1, column=1).value = f"inventario filial 944 - {nome_final_aba}"
                    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=len(tabela_final.columns))
                    
                    # Estilo do Título
                    ws.cell(row=1, column=1).font = Font(size=12, bold=True)
                    ws.cell(row=1, column=1).fill = PatternFill(start_color="D9D9D9", end_color="D9D9D9", fill_type="solid")
                    ws.cell(row=1, column=1).alignment = Alignment(horizontal="center")
                    
                    # Estilo do Cabeçalho (Linha 2)
                    for cell in ws[2]:
                        cell.fill = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
                        cell.font = Font(color="FFFFFF", bold=True)
                        cell.alignment = Alignment(horizontal="center")

                    # Ajuste de largura
                    for i, col in enumerate(tabela_final.columns, 1):
                        column_letter = get_column_letter(i)
                        max_len = max([len(str(x)) for x in grupo[col].values] + [len(col)])
                        ws.column_dimensions[column_letter].width = max_len + 5

            st.download_button(
                label="📥 Baixar Inventário",
                data=output.getvalue(),
                file_name="Inventario_Filial_944.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
            st.success("Planilha processada com sucesso!")

    except Exception as e:
        st.error(f"Erro: {e}")
