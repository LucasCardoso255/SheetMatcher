import pandas as pd

def criar_planilha_combinada(planilha1_path='planilha1.xlsx', planilha2_path='planilha2.xlsx', output_path='saida.xlsx'):
    df1 = pd.read_excel(planilha1_path)
    df2 = pd.read_excel(planilha2_path)

    coluna_produto_df1 = 'Descricao do produto'
    coluna_produto_df2 = 'Column2'

    if coluna_produto_df1 not in df1.columns:
        print(f"Erro: Coluna '{coluna_produto_df1}' não encontrada em '{planilha1_path}'.")
        return
    if coluna_produto_df2 not in df2.columns:
        print(f"Erro: Coluna '{coluna_produto_df2}' não encontrada em '{planilha2_path}'.")
        return

    df1[coluna_produto_df1] = df1[coluna_produto_df1].astype(str).str.strip()
    df2[coluna_produto_df2] = df2[coluna_produto_df2].astype(str).str.strip()

    df_final = pd.merge(df1, df2,
                        left_on=coluna_produto_df1,
                        right_on=coluna_produto_df2,
                        how='left',
                        suffixes=('', '_df2'))

    if coluna_produto_df2 + '_df2' in df_final.columns:
        df_final = df_final.drop(columns=[coluna_produto_df2 + '_df2'])

    df_final.to_excel(output_path, index=False)
    print(f"Arquivo '{output_path}' gerado com sucesso!")

criar_planilha_combinada()