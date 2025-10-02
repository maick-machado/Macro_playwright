import fitz  # PyMuPDF
import os  # Para navegar em pastas e ficheiros
import pandas  # Para criar e gerir o ficheiro Excel
import re  # Para ajudar a limpar textos (expressões regulares)


# ==============================================================================
# FUNÇÃO DE EXTRAÇÃO ESPECIALIZADA PARA NOTAS DE BOA VISTA
# ==============================================================================
def extrair_dados_boa_vista(caminho_do_pdf):
    """
    Extrai todos os dados necessários de uma NFS-e do município de Boa Vista.
    Recebe o caminho completo do ficheiro PDF e retorna um dicionário com os dados.
    """
    try:
        pagina = fitz.open(caminho_do_pdf).load_page(0)
    except Exception as e:
        print(
            f"  AVISO: Não foi possível abrir o ficheiro {os.path.basename(caminho_do_pdf)}. Erro: {e}"
        )
        return {"STATUS DA EXECUÇÃO": f"Erro ao abrir PDF: {e}"}

    dados = {}

    # Função auxiliar para extrair texto de uma área
    def extrair_texto(pagina, area):
        texto = pagina.get_text("text", clip=area).strip()
        return texto

    # --- Extração de cada campo ---

    # 1. NÚMERO DA NF
    areas = pagina.search_for("Número da Nota")
    if areas:
        area_clip = fitz.Rect(
            areas[0].x0, areas[0].y1, areas[0].x1 + 60, areas[0].y1 + 15
        )
        dados["NUMERO DA NF"] = extrair_texto(pagina, area_clip)

    # 2. CÓDIGO DE VERIFICAÇÃO
    areas = pagina.search_for("Código de Verificação")
    if areas:
        area_clip = fitz.Rect(
            areas[0].x0, areas[0].y1, areas[0].x1 + 150, areas[0].y1 + 20
        )
        texto = extrair_texto(pagina, area_clip)
        dados["CODIGO DE VERIFICAÇÃO"] = texto.split()[0] if texto else None

    # 3. CNPJ FORNECEDOR (Prestador)
    areas = pagina.search_for("Prestador do(s) Serviço(s)")
    if areas:
        # Procuramos por 'CPF/CNPJ:' na área abaixo de 'Prestador'
        area_busca_cnpj = fitz.Rect(
            areas[0].x0 - 150, areas[0].y1, areas[0].x1 + 100, areas[0].y1 + 100
        )
        areas_cnpj = pagina.search_for("CPF/CNPJ:", clip=area_busca_cnpj)
        if areas_cnpj:
            area_clip = fitz.Rect(
                areas_cnpj[0].x0 + 50,
                areas_cnpj[0].y1 - 10,
                areas_cnpj[0].x1 + 130,
                areas_cnpj[0].y1,
            )
            dados["CNPJ FORNECEDOR"] = extrair_texto(pagina, area_clip)

    # 4. DATA DE EMISSÃO NF
    areas = pagina.search_for("Data e Hora de Emissão")
    if areas:
        area_clip = fitz.Rect(
            areas[0].x0, areas[0].y1, areas[0].x1 - 12, areas[0].y1 + 10
        )
        texto_data = extrair_texto(pagina, area_clip)
        # Extrai apenas a data (primeira parte do texto)
        dados["DATA DE EMISSAO NF"] = texto_data.split()[0] if texto_data else None

    # 5. VALOR BRUTO (Valor do(s) Serviço(s))
    areas = pagina.search_for("Valor do(s) Serviço(s)")
    if areas:
        area_clip = fitz.Rect(
            areas[0].x1 - 50, areas[0].y0 + 8, areas[0].x1 + 50, areas[0].y1 + 10
        )
        dados["VALOR BRUTO"] = extrair_texto(pagina, area_clip)

    # 6. TIPO DE SERVIÇO
    areas = pagina.search_for("Classificação do Serviço")
    if areas:
        area_clip = fitz.Rect(
            areas[0].x0, areas[0].y1, pagina.rect.width, areas[0].y1 + 40
        )
        dados["TIPO DE SERVIÇO"] = extrair_texto(pagina, area_clip).replace("\n", " ")

    # 5. VALOR DA RETENÇÃO (Valor do(s) Serviço(s))
    areas = pagina.search_for("Retenções Federais")
    if areas:
        # Procuramos por 'CPF/CNPJ:' na área abaixo de 'Prestador'
        area_busca_cnpj = fitz.Rect(
            areas[0].x0 - 150, areas[0].y1, areas[0].x1 + 150, areas[0].y1 + 50
        )
        areas_cnpj = pagina.search_for("INSS", clip=area_busca_cnpj)
        if areas_cnpj:
            area_clip = fitz.Rect(
                areas_cnpj[0].x0,
                areas_cnpj[0].y1,
                areas_cnpj[0].x1 + 70,
                areas_cnpj[0].y1 + 10,
            )
            dados["VALOR DA RETENÇÃO"] = extrair_texto(pagina, area_clip)

    # 8. SÉRIE NF - Este campo não foi encontrado no modelo de NF de Boa Vista
    dados["SERIE NF"] = "1"

    # Define o status final da execução para este ficheiro
    dados["STATUS DA EXECUÇÃO"] = "Sucesso"

    return dados


# ==============================================================================
# FUNÇÃO PRINCIPAL (ORQUESTRADOR)
# ==============================================================================
def executar_extracao_pdf():
    """
    Função principal que orquestra todo o processo de leitura e gravação.
    """
    pasta_raiz = "nf"
    if not os.path.isdir(pasta_raiz):
        print(
            f"Erro: A pasta '{pasta_raiz}' não foi encontrada. Por favor, crie-a e organize os PDFs."
        )
        return

    todos_os_dados = []

    print("Iniciando o processamento de Notas Fiscais...")

    # Navega pela estrutura de pastas: nf/cidade/
    for nome_cidade in os.listdir(pasta_raiz):
        pasta_cidade = os.path.join(pasta_raiz, nome_cidade)

        if os.path.isdir(pasta_cidade):
            print(f"\nProcessando pasta da cidade: {nome_cidade}...")

            # Itera sobre cada ficheiro na pasta da cidade
            for nome_arquivo in os.listdir(pasta_cidade):
                if nome_arquivo.lower().endswith(".pdf"):
                    caminho_completo_pdf = os.path.join(pasta_cidade, nome_arquivo)
                    print(f"  Lendo ficheiro: {nome_arquivo}")

                    dados_extraidos = None

                    # --- PONTO DE DECISÃO: CHAMA A FUNÇÃO CORRETA PARA A CIDADE ---
                    if nome_cidade == "boa_vista":
                        dados_extraidos = extrair_dados_boa_vista(caminho_completo_pdf)
                    # elif nome_cidade == "manaus":
                    #     dados_extraidos = extrair_dados_manaus(caminho_completo_pdf)
                    # Adicione outras cidades aqui no futuro
                    else:
                        print(
                            f"  AVISO: Nenhum script de extração definido para a cidade '{nome_cidade}'."
                        )
                        dados_extraidos = {
                            "STATUS DA EXECUÇÃO": "Cidade não configurada"
                        }

                    if dados_extraidos:
                        # Adiciona dados que dependem do contexto (fora do PDF)
                        dados_extraidos["MUNICIPIO DA NF"] = nome_cidade
                        todos_os_dados.append(dados_extraidos)

    # --- EXPORTAÇÃO PARA EXCEL ---
    if not todos_os_dados:
        print("\nNenhum dado foi extraído. O ficheiro Excel não será gerado.")
        return

    print("\nConsolidando dados e gerando o ficheiro Excel...")

    df = pandas.DataFrame(todos_os_dados)

    # Define a ordem exata das colunas que você pediu
    ordem_colunas = [
        "NUMERO DA NF",
        "MUNICIPIO DA NF",
        "CODIGO DE VERIFICAÇÃO",
        "CNPJ FORNECEDOR",
        "SERIE NF",
        "DATA DE EMISSAO NF",
        "VALOR BRUTO",
        "TIPO DE SERVIÇO",
        "VALOR DA RETENÇÃO",
        "STATUS DA EXECUÇÃO",
    ]

    # Garante que todas as colunas existam, preenchendo com "Não Encontrado" se faltar
    for col in ordem_colunas:
        if col not in df.columns:
            df[col] = "Não Encontrado"

    df = df[ordem_colunas]  # Reordena o DataFrame

    try:
        df.to_excel("relatorio_consolidado_nf.xlsx", index=False)
        print(
            "\nSucesso! O ficheiro 'relatorio_consolidado_nf.xlsx' foi criado na pasta principal."
        )
    except Exception as e:
        print(f"\nErro ao salvar o ficheiro Excel: {e}")


# --- PONTO DE ENTRADA DO SCRIPT ---
if __name__ == "__main__":
    executar_extracao_pdf()
