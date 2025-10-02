import fitz  # PyMuPDF
import os

# --- CONFIGURAÇÕES ---
# O caminho para a nota fiscal que você quer usar como modelo.
NOME_DO_PDF = os.path.join("nf", "boa_vista", "sample.pdf")

# --- SCRIPT DE VISUALIZAÇÃO PARA VALOR BRUTO ---

print(f"Abrindo o ficheiro de modelo: {NOME_DO_PDF}")
try:
    doc = fitz.open(NOME_DO_PDF)
    pagina = doc.load_page(0)

    # --- Passo 1: Encontrar o texto âncora "Valor do(s) Serviço(s)" ---
    texto_ancora = "Valor do(s) Serviço(s)"
    areas_encontradas = pagina.search_for(texto_ancora)

    if areas_encontradas:
        print(
            f"Texto '{texto_ancora}' encontrado. A desenhar o retângulo de extração..."
        )
        # Pega a localização da primeira ocorrência do texto
        primeira_ocorrencia = areas_encontradas[0]

        # --- Passo 2: Definir o retângulo de extração com base na âncora ---
        # As coordenadas são ajustadas para capturar o texto que está À DIREITA do rótulo.
        # fitz.Rect(x0, y0, x1, y1)
        area_de_extracao_valor = fitz.Rect(
            primeira_ocorrencia.x1
            - 50,  # x0: Começa 10 pixels à direita do FIM do texto âncora.
            primeira_ocorrencia.y0
            + 8,  # y0: Um pouco acima do texto âncora para garantir a captura da altura.
            primeira_ocorrencia.x1
            + 50,  # x1: Define uma largura de 150 pixels para a área de captura.
            primeira_ocorrencia.y1 + 10,  # y1: Um pouco abaixo do texto âncora.
        )

        # --- Passo 3: Desenhar o retângulo na página ---
        # Cor (1, 0, 0) é vermelho em RGB.
        pagina.draw_rect(area_de_extracao_valor, color=(1, 0, 0), width=1.5)

        # --- Passo 4: Salvar a página como uma imagem ---
        pix = pagina.get_pixmap(dpi=150)  # dpi=150 para uma boa qualidade de imagem
        output_image_path = "debug_area_valor_bruto.png"
        pix.save(output_image_path)

        doc.close()
        print(f"\nSucesso! Imagem de depuração salva como '{output_image_path}'")
        print(
            "O retângulo VERMELHO mostra a área exata de onde o 'VALOR BRUTO' seria extraído."
        )

    else:
        print(f"AVISO: O texto '{texto_ancora}' não foi encontrado no PDF.")

except FileNotFoundError:
    print(
        f"ERRO: O ficheiro '{NOME_DO_PDF}' não foi encontrado. Verifique o caminho e a estrutura de pastas."
    )
except Exception as e:
    print(f"Ocorreu um erro inesperado: {e}")
