import streamlit as st
import pandas as pd
from functools import reduce
import io
from openpyxl.styles import Alignment

st.set_page_config(
    page_title="Meu App",
    layout="wide"  # deixa o app em tela cheia
)

st.title("Carregar Planilha")

# Upload do arquivo
arquivos = st.file_uploader("Selecione os arquivos", type=["xlsx"], accept_multiple_files=True)

dataframes = []

if arquivos:
    for arq in arquivos:
        nome_base = arq.name.split("-")[0]
        df = pd.read_excel(arq, sheet_name='Participant Data')
        df["Name"] = df["First Name"].astype(str) + " " + df["Last Name"].astype(str)
        df = df[['Class Name', "Name", 'Accuracy']]
        df['Accuracy'] = df['Accuracy'].astype(str).str.replace('%','').astype(float)
        df = (
            df.groupby(['Class Name', 'Name'])
                .agg(
                    **{f"Acc-{nome_base}": ('Accuracy', 'max'),
                        f"Tentativa-{nome_base}": ('Accuracy', 'count')}
                )
                .reset_index()
            )
        dataframes.append(df)

    dfFinal = reduce(lambda left, right: pd.merge(left, right, on=["Class Name", "Name"], how="outer"), dataframes)
    dfFinal.fillna(0, inplace=True)

    # Calcular ACC Total (média das médias de acurácia de cada arquivo)
    colunas_acc = [col for col in dfFinal.columns if col.startswith('Acc-')]
    # Calcular a média apenas das colunas de acurácia (ignorando valores NaN)
    dfFinal['ACC Total'] = dfFinal[colunas_acc].mean(axis=1, skipna=True).round(2)
    
    # Reordenar colunas para colocar ACC Total após Name
    colunas_ordenadas = ['Class Name', 'Name', 'ACC Total']
    colunas_restantes = [col for col in dfFinal.columns if col not in colunas_ordenadas]
    dfFinal = dfFinal[colunas_ordenadas + colunas_restantes]

    st.write("📊 Visualização da Planilha Completa:")
    st.dataframe(dfFinal)
    
    # Botão de download com células mescladas
    def to_excel_with_merged_cells(df):
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            # Criar uma cópia do dataframe para manipulação
            df_export = df.copy()
            
            # Adicionar coluna vazia para mesclar
            df_export.insert(0, 'Class_Name_Merge', '')
            
            # Preencher apenas a primeira linha de cada turma
            current_class = None
            for idx, row in df_export.iterrows():
                if row['Class Name'] != current_class:
                    df_export.at[idx, 'Class_Name_Merge'] = row['Class Name']
                    current_class = row['Class Name']
            
            # Remover a coluna original Class Name
            df_export = df_export.drop('Class Name', axis=1)
            
            # Renomear a coluna mesclada
            df_export = df_export.rename(columns={'Class_Name_Merge': 'Class Name'})
            
            # Escrever no Excel
            df_export.to_excel(writer, index=False, sheet_name='Dados_Consolidados')
            
            # Obter a planilha para mesclar células
            worksheet = writer.sheets['Dados_Consolidados']
            
            # Mesclar células da coluna Class Name
            current_class = None
            start_row = 2  # Começar da linha 2 (após cabeçalho)
            end_row = 2
            
            for idx, row in df_export.iterrows():
                if row['Class Name'] != current_class and row['Class Name'] != '':
                    # Mesclar células da classe anterior se existir
                    if current_class is not None and end_row > start_row:
                        worksheet.merge_cells(f'A{start_row}:A{end_row}')
                        # Centralizar verticalmente e horizontalmente
                        worksheet[f'A{start_row}'].alignment = Alignment(horizontal='center', vertical='center')
                    
                    # Iniciar nova classe
                    current_class = row['Class Name']
                    start_row = idx + 2  # +2 porque Excel começa em 1 e tem cabeçalho
                    end_row = start_row
                elif row['Class Name'] == current_class:
                    end_row = idx + 2
                else:
                    end_row = idx + 2
            
            # Mesclar última classe
            if current_class is not None and end_row > start_row:
                worksheet.merge_cells(f'A{start_row}:A{end_row}')
                # Centralizar verticalmente e horizontalmente
                worksheet[f'A{start_row}'].alignment = Alignment(horizontal='center', vertical='center')
            
            # Ajustar largura das colunas
            worksheet.column_dimensions['A'].width = 90
            worksheet.column_dimensions['B'].width = 50
            worksheet.column_dimensions['C'].width = 15
            worksheet.column_dimensions['D'].width = 25
            worksheet.column_dimensions['E'].width = 25
            worksheet.column_dimensions['F'].width = 25
            worksheet.column_dimensions['G'].width = 25
            worksheet.column_dimensions['H'].width = 25
            worksheet.column_dimensions['I'].width = 25
            worksheet.column_dimensions['J'].width = 25
            worksheet.column_dimensions['K'].width = 25
            worksheet.column_dimensions['L'].width = 25
            worksheet.column_dimensions['M'].width = 25
            worksheet.column_dimensions['N'].width = 25
            worksheet.column_dimensions['O'].width = 25
            worksheet.column_dimensions['P'].width = 25
            worksheet.column_dimensions['Q'].width = 25
            worksheet.column_dimensions['R'].width = 25
            worksheet.column_dimensions['S'].width = 25

        return output.getvalue()
    
    excel_data = to_excel_with_merged_cells(dfFinal)
    st.download_button(
        label="📥 Baixar Planilha Completa (XLSX)",
        data=excel_data,
        file_name="dados_consolidados.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
    
    # Separar por turma
    st.write("---")
    st.write("📚 Visualização por Turma:")
    
    # Obter lista única de turmas
    turmas = dfFinal['Class Name'].unique()
    
    for turma in turmas:
        st.write(f"### Turma: {turma}")
        df_turma = dfFinal[dfFinal['Class Name'] == turma].copy()
        df_turma_display = df_turma.drop('Class Name', axis=1)  # Remove a coluna Class Name pois já está no título
        st.dataframe(df_turma_display)
        
        # Botão de download para a turma específica
        def to_excel_turma(df, nome_turma):
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                # Criar uma cópia do dataframe para manipulação
                df_export = df.copy()
                
                # Adicionar coluna vazia para mesclar
                df_export.insert(0, 'Class_Name_Merge', '')
                
                # Preencher apenas a primeira linha com o nome da turma
                df_export.at[0, 'Class_Name_Merge'] = nome_turma
                
                # Remover a coluna original Class Name
                df_export = df_export.drop('Class Name', axis=1)
                
                # Renomear a coluna mesclada
                df_export = df_export.rename(columns={'Class_Name_Merge': 'Class Name'})
                
                # Escrever no Excel
                df_export.to_excel(writer, index=False, sheet_name=nome_turma[:31])  # Limite de 31 caracteres para nome da aba
                
                # Obter a planilha para mesclar células
                worksheet = writer.sheets[nome_turma[:31]]
                
                # Mesclar todas as células da coluna Class Name (já que é só uma turma)
                if len(df_export) > 1:
                    worksheet.merge_cells(f'A2:A{len(df_export) + 1}')
                    # Centralizar verticalmente e horizontalmente
                    worksheet['A2'].alignment = Alignment(horizontal='center', vertical='center')
                
                # Ajustar largura das colunas
                worksheet.column_dimensions['A'].width = 90
                worksheet.column_dimensions['B'].width = 50
                worksheet.column_dimensions['C'].width = 15
                worksheet.column_dimensions['D'].width = 15
                worksheet.column_dimensions['E'].width = 15
                worksheet.column_dimensions['F'].width = 15
                worksheet.column_dimensions['G'].width = 15
                worksheet.column_dimensions['H'].width = 15
                worksheet.column_dimensions['I'].width = 15
                worksheet.column_dimensions['J'].width = 15
                worksheet.column_dimensions['K'].width = 15
                worksheet.column_dimensions['L'].width = 15
                worksheet.column_dimensions['M'].width = 15
                worksheet.column_dimensions['N'].width = 15
                worksheet.column_dimensions['O'].width = 15
                worksheet.column_dimensions['P'].width = 15
                worksheet.column_dimensions['Q'].width = 15
                worksheet.column_dimensions['R'].width = 15
                worksheet.column_dimensions['S'].width = 15
                
            return output.getvalue()
        
        # Nome do arquivo baseado na turma (removendo caracteres especiais)
        nome_arquivo = turma.replace(' ', '_').replace('-', '_').replace('º', '').replace('°', '')
        excel_data_turma = to_excel_turma(df_turma, turma)
        
        st.download_button(
            label=f"📥 Baixar {turma} (XLSX)",
            data=excel_data_turma,
            file_name=f"{nome_arquivo}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            key=f"download_{turma}"  # Chave única para cada botão
        )
        st.write("---")
