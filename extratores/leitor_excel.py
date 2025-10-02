import pandas as pd
import os


def ler_dados_da_planilha(nome_arquivo: str):
    """
    Lê uma planilha Excel e retorna os dados como uma lista de dicionários.

    Cada dicionário na lista representa uma linha da planilha, onde as chaves
    são os nomes das colunas.

    Args:
        nome_arquivo (str): O nome do arquivo Excel (ex: "dados.xlsx").
                            O arquivo deve estar na mesma pasta do script.

    Returns:
        list: Uma lista de dicionários com os dados das linhas,
              ou uma lista vazia se ocorrer um erro ou o arquivo não for encontrado.
    """
    try:
        # Verifica se o arquivo existe na pasta do script
        caminho_arquivo = os.path.join(os.path.dirname(__file__), nome_arquivo)

        if not os.path.exists(caminho_arquivo):
            print(f"Erro: O arquivo '{nome_arquivo}' não foi encontrado na pasta.")
            return []

        # Usa o pandas para ler o arquivo Excel.
        # O resultado é um DataFrame, que é como uma tabela super poderosa.
        dataframe = pd.read_excel(caminho_arquivo)

        # Converte o DataFrame para uma lista de dicionários.
        # 'records' significa que cada item da lista será um registro (uma linha).
        dados_em_lista = dataframe.to_dict(orient="records")

        print(
            f"Sucesso: {len(dados_em_lista)} linhas lidas do arquivo '{nome_arquivo}'."
        )
        return dados_em_lista

    except Exception as e:
        print(f"Ocorreu um erro inesperado ao ler o arquivo Excel: {e}")
        return []
