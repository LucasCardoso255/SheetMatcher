import pandas as pd

def ordenar_planilha_por_referencia(planilha1_path='planilha1.xlsx', planilha2_path='planilha2.xlsx', output_path='saida_ordenada.xlsx'):
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

    df2['chave_correspondencia'] = df2[coluna_produto_df2].astype(str).str.strip()
    df1['chave_ordenacao'] = df1[coluna_produto_df1].astype(str).str.strip()

    df_ordenado_lista = []
    chaves_encontradas = set()

    for chave_ordenacao in df1['chave_ordenacao']:
        linha_correspondente = df2[df2['chave_correspondencia'] == chave_ordenacao]
        if not linha_correspondente.empty:
            df_ordenado_lista.append(linha_correspondente)
            chaves_encontradas.add(chave_ordenacao)

    df_ordenado = pd.concat(df_ordenado_lista, ignore_index=True) if df_ordenado_lista else pd.DataFrame(columns=df2.columns)

    chaves_nao_encontradas = df2[~df2['chave_correspondencia'].isin(chaves_encontradas)]
    
    df_final = pd.concat([df_ordenado, chaves_nao_encontradas], ignore_index=True)

    df_final = df_final.drop(columns=['chave_correspondencia'])

    df_final.to_excel(output_path, index=False)

ordenar_planilha_por_referencia()