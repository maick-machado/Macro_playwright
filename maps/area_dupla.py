import fitz  # PyMuPDF
import os

# --- CONFIGURAÇÕES ---
# O caminho para a nota fiscal que você quer usar como modelo.
NOME_DO_PDF = os.path.join("nf", "boa_vista", "sample.pdf")

# --- SCRIPT DE VISUALIZAÇÃO PARA RETENÇÃO DO INSS ---

print(f"Abrindo o ficheiro de modelo: {NOME_DO_PDF}")
try:
    doc = fitz.open(NOME_DO_PDF)
    pagina = doc.load_page(0)

    # --- Passo 1: Encontrar o texto âncora principal "Retenções Federais" ---
    texto_ancora_principal = "Retenções Federais"
    areas_ancora = pagina.search_for(texto_ancora_principal)

    if areas_ancora:
        print(f"Texto principal '{texto_ancora_principal}' encontrado.")
        primeira_ocorrencia = areas_ancora[0]

        # --- Passo 2: Definir a "matriz" de busca, assim como no seu código ---
        # Este é o retângulo azul (área de busca maior)
        matriz_de_busca = fitz.Rect(
            primeira_ocorrencia.x0 - 150,
            primeira_ocorrencia.y1,
            primeira_ocorrencia.x1 + 150,
            primeira_ocorrencia.y1 + 50,
        )
        pagina.draw_rect(
            matriz_de_busca, color=(0, 0, 1), width=1.5, dashes="[3 1]"
        )  # Linha tracejada azul

        # --- Passo 3: Procurar por "INSS" dentro da matriz azul ---
        texto_ancora_secundario = "INSS"
        areas_inss = pagina.search_for(texto_ancora_secundario, clip=matriz_de_busca)

        if areas_inss:
            print(
                f"Texto secundário '{texto_ancora_secundario}' encontrado dentro da matriz."
            )
            ocorrencia_inss = areas_inss[0]

            # --- Passo 4: Definir a área final de extração, com base na sua lógica ---
            # Este é o retângulo vermelho (área de extração final)
            area_de_extracao_final = fitz.Rect(
                ocorrencia_inss.x0,
                ocorrencia_inss.y1,
                ocorrencia_inss.x1 + 70,
                ocorrencia_inss.y1 + 10,
            )
            pagina.draw_rect(area_de_extracao_final, color=(1, 0, 0), width=1.5)

            # --- Passo 5: Salvar a página como uma imagem ---
            pix = pagina.get_pixmap(dpi=150)
            output_image_path = "debug_area_retencao_inss.png"
            pix.save(output_image_path)

            doc.close()
            print(f"\nSucesso! Imagem de depuração salva como '{output_image_path}'")
            print(" - O retângulo AZUL TRACEJADO mostra a 'matriz' de busca principal.")
            print(
                " - O retângulo VERMELHO mostra a área exata de onde o valor do INSS seria extraído."
            )

        else:
            doc.close()
            print(
                f"AVISO: O texto secundário '{texto_ancora_secundario}' não foi encontrado dentro da área de busca definida."
            )

    else:
        doc.close()
        print(
            f"AVISO: O texto principal '{texto_ancora_principal}' não foi encontrado no PDF."
        )

except FileNotFoundError:
    print(f"ERRO: O ficheiro '{NOME_DO_PDF}' não foi encontrado. Verifique o caminho.")
except Exception as e:
    print(f"Ocorreu um erro inesperado: {e}")
