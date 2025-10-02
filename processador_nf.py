import subprocess
import os
import time
import asyncio
from playwright.async_api import async_playwright
from extratores.leitor_excel import ler_dados_da_planilha
from extratores.extrator_pdf import executar_extracao_pdf

# --- CONFIGURAÇÃO ---
CHROME_EXECUTABLE_PATH = r"C:\Program Files\Google\Chrome\Application\chrome.exe"
REMOTE_DEBUGGING_PORT = "9222"
USER_DATA_DIR = (
    r"C:\ChromeDebugProfile"  # Usar um perfil separado é altamente recomendado
)
LOGIN_URL = "https://cav.receita.fazenda.gov.br/"

NOME_ARQUIVO_EXCEL = "relatorio_consolidado_nf.xlsx"


def start_chrome_with_debugging():
    """
    Verifica e inicia uma nova instância do Chrome com a porta de depuração remota.
    """
    if not os.path.exists(CHROME_EXECUTABLE_PATH):
        print(
            f"ERRO: O executável do Chrome não foi encontrado em '{CHROME_EXECUTABLE_PATH}'"
        )
        raise FileNotFoundError(
            "Verifique o caminho na variável CHROME_EXECUTABLE_PATH."
        )

    command = [
        CHROME_EXECUTABLE_PATH,
        f"--remote-debugging-port={REMOTE_DEBUGGING_PORT}",
        f"--user-data-dir={USER_DATA_DIR}",
    ]

    print(f"Iniciando o Chrome na porta de depuração {REMOTE_DEBUGGING_PORT}...")
    subprocess.Popen(command)
    time.sleep(2)  # Pequena pausa para garantir que o navegador inicie completamente


async def main():
    """
    Função principal assíncrona que controla o fluxo de automação com Playwright.
    """
    # ==============================================================================
    # --- FASE 1: EXECUÇÃO DA EXTRAÇÃO DOS PDFs ---
    # ==============================================================================
    print("--- INICIANDO FASE 1: Extraindo dados dos PDFs para o Excel ---")
    try:
        executar_extracao_pdf()
        print("--- FASE 1 CONCLUÍDA: Ficheiro Excel gerado com sucesso! ---\n")
    except Exception as e:
        print(f"ERRO CRÍTICO na fase de extração de PDFs: {e}")
        print(
            "O programa será encerrado pois os dados de entrada não puderam ser gerados."
        )
        return  # Encerra a execução se a extração falhar

    # ==============================================================================
    # --- FASE 2: EXECUÇÃO DA AUTOMAÇÃO WEB ---
    # ==============================================================================
    print("--- INICIANDO FASE 2: Lendo Excel e automatizando o navegador ---")
    start_chrome_with_debugging()

    browser = None  # Inicializa a variável do browser
    try:
        async with async_playwright() as p:
            # 2. Conecta o Playwright ao Chrome já aberto via Chrome DevTools Protocol (CDP)
            print("Conectando o Playwright ao Chrome...")
            endpoint_url = f"http://127.0.0.1:{REMOTE_DEBUGGING_PORT}"
            browser = await p.chromium.connect_over_cdp(endpoint_url)

            # Obtém o contexto padrão do navegador (a janela que foi aberta)
            context = browser.contexts[0]
            page = context.pages[0]  # Pega a primeira aba/página
            print("Conectado com sucesso!")

            print(f"Abrindo a página: {LOGIN_URL}")
            await page.goto(LOGIN_URL)
            botao_sair_locator = page.locator("#sairSeguranca")
            await botao_sair_locator.wait_for(state="visible", timeout=600000)

            print("\nLogin detectado com sucesso!")
            print(f"URL atual: {page.url}")

            # -----------------------------------------------------
            print("\nIniciando a execução da sua macro...")
            # -----------------------------------------------------
            # ACESSANDO O BOTÃO DE PERFIL DO USUÁRIO:
            botao_perfil = page.locator("#btnPerfil")
            await botao_perfil.click()
            print("Botão clicado. A janela de perfil deve estar aberta.")
            # -----------------------------------------------------
            # ABERTURA DA JANELA DE PREENCHIMENTO DO PERFIL
            dialogo_perfil = page.locator("#perfilAcesso")
            await dialogo_perfil.wait_for(state="visible")
            print("Janela de perfil detectada.")
            await dialogo_perfil.wait_for(state="hidden", timeout=0)
            print("\n--- Janela de perfil fechada pelo usuário! ---")
            print("Continuando a execução do script...")
            # -----------------------------------------------------
            # CLICANDO NO BOTÃO "DECLARAÇÕES E DEMONSTRATIVOS"
            botao_declaracoes = page.locator("#btn214")
            await botao_declaracoes.click()
            print("Botão 'Declarações e Demonstrativos' clicado com sucesso!")
            # -----------------------------------------------------
            # CLICANDO NO LINK "ACESSAR EFD-REINF"
            link_reinf = page.get_by_role("link", name="Acessar EFD-Reinf")
            await link_reinf.click()
            print("Link 'Acessar EFD-Reinf' clicado com sucesso!")
            # -----------------------------------------------------
            # INICIANDO O LOOPING PELOS DADOS DO EXCEL
            lista_de_cadastros = ler_dados_da_planilha(NOME_ARQUIVO_EXCEL)
            for cadastro_atual in lista_de_cadastros:

                # usando as chaves do dicionário (os nomes das colunas)
                num_nf_principal = cadastro_atual.get("NUMERO DA NF")
                cod_verificacao = cadastro_atual.get("CODIGO DE VERIFICAÇÃO")
                cnpj_fornecedor_nf = cadastro_atual.get("CNPJ FORNECEDOR")
                serie_nf = cadastro_atual.get("SERIE NF")
                numero_nf = cadastro_atual.get("NUMERO NF")
                data_emissao_nf = cadastro_atual.get("DATA DE EMISSAO NF")
                valor_bruto_nf = cadastro_atual.get("VALOR BRUTO")
                tipo_servico = cadastro_atual.get("TIPO DE SERVIÇO")
                valor_retencao_nf = cadastro_atual.get("VALOR DA RETENÇÃO")
                status = cadastro_atual.get("STATUS DA EXECUÇÃO")
                # -----------------------------------------------------
                # ABRINDO O MENU PARA ACESSAR "RETENÇÃO DE CONTRIBUIÇÃO PREVIDENCIÁRIA TOMADORES DE SERVIÇOS (R2010)"
                print("Localizando o iframe com id='frmApp'...")
                frame_locator = frame_locator.frame_locator("#frmApp")
                print("Iframe localizado com sucesso!")
                testid_menu_principal = (
                    "menu_retencoes_previdenciarias_series_r2000_e_r3000"
                )
                menu_principal_locator = frame_locator.locator(
                    f'[data-testid="{testid_menu_principal}"]'
                )
                await menu_principal_locator.hover()
                print("Menu aberto com sucesso!")
                testid_item_submenu = "menu_retencao_contribuicao_previdenciaria_tomadores_de_servicos_r2010"
                item_submenu_locator = frame_locator.locator(
                    f'[data-testid="{testid_item_submenu}"]'
                )
                await item_submenu_locator.click()
                print(
                    "Opção 'Retenção Contribuição Previdenciária...' clicada com sucesso!"
                )
                # -----------------------------------------------------
                # ACESSANDO AO BOTÃO DE INCLUIR NOVO EVENTO
                testid_do_botao = "botao_evento_incluir_novo"
                botao_incluir = frame_locator.locator(
                    f'[data-testid="{testid_do_botao}"]'
                )
                await botao_incluir.click()
                print("Botão '+ Incluir novo evento' clicado com sucesso!")
                # -----------------------------------------------------
                # PREENCHIMENTO DO FORMULÁRIO
                # CAMPO "PERÍODO DE APURAÇÃO"
                valor_a_preencher = "09/2025"  # alterar depois para o usuario inserir ou pegar do mes atual
                # data_atual = datetime.date.today()
                # valor_a_preencher_dinamico = data_atual.strftime("%m/%Y") # Ex: "09/2025"
                # await campo_periodo.fill(valor_a_preencher_dinamico)
                testid_do_campo = "periodo_apuracao"
                campo_periodo = frame_locator.locator(
                    f'[data-testid="{testid_do_campo}"]'
                )
                await campo_periodo.fill(valor_a_preencher)
                print("Campo 'Período de Apuração' preenchido com sucesso!")
                # -----------------------------------------------------
                # CAMPO CNPJ PRF
                cnpj_a_preencher = "99.999.999/9999-99"
                testid_do_campo = "numero_inscricao_cnpj"
                campo_cnpj = frame_locator.locator(f'[data-testid="{testid_do_campo}"]')
                await campo_cnpj.fill(cnpj_a_preencher)
                print("Campo 'CNPJ' preenchido com sucesso!")
                # -----------------------------------------------------
                # CAMPO CNPJ FORNECEDOR
                cnpj_prestador_a_preencher = cnpj_fornecedor_nf
                testid_do_campo = "cnpj_prestador"
                campo_cnpj_prestador = frame_locator.locator(
                    f'[data-testid="{testid_do_campo}"]'
                )
                await campo_cnpj_prestador.fill(cnpj_prestador_a_preencher)
                print("Campo 'CNPJ do Prestador' preenchido com sucesso!")
                # -----------------------------------------------------
                # CONFIRMAÇÃO BOTÃO "CONTINUAR"
                testid_do_botao = "botao_continuar"
                botao_continuar = frame_locator.locator(
                    f'[data-testid="{testid_do_botao}"]'
                )
                await botao_continuar.click()
                print("Botão 'Continuar' clicado com sucesso!")
                # -----------------------------------------------------
                # SELECIONANDO O CAMPO "INDICATIVO DE PRESTAÇÃO DE SERVIÇOS"
                testid_do_select = "indicativo_obra"
                select_locator = frame_locator.locator(
                    f'[data-testid="{testid_do_select}"]'
                )
                valor_da_opcao = "0"
                await select_locator.select_option(value=valor_da_opcao)
                print("Opção selecionada com sucesso!")
                # -----------------------------------------------------
                # SELECIONANDO O CAMPO "PRESTADOR É CONTRIBUINTE"
                testid_do_select_cprb = "indicativo_cprb"
                select_cprb_locator = frame_locator.locator(
                    f'[data-testid="{testid_do_select_cprb}"]'
                )
                valor_da_opcao = "0"
                await select_cprb_locator.select_option(value=valor_da_opcao)
                print("Opção do campo 'Indicativo CPRB' selecionada com sucesso!")
                # -----------------------------------------------------
                # CLICANDO EM INCLUIR NOTAS FISCAIS
                testid_do_link_correto = "botao_inclusao_nfs"
                link_incluir_nova_nf = frame_locator.get_by_test_id(
                    testid_do_link_correto
                )
                await link_incluir_nova_nf.click()
                print(
                    "Link '[Incluir Nova]' da seção 'Notas fiscais' clicado com sucesso!"
                )
                # -----------------------------------------------------
                # INCLUINDO DADOS DA NOTA FISCAL
                valor_serie = serie_nf
                valor_numero_doc = num_nf_principal
                valor_data_emissao = data_emissao_nf
                valor_bruto = valor_bruto_nf
                await frame_locator.locator('[data-testid="serie"]').fill(valor_serie)
                await frame_locator.locator('[data-testid="numero_documento"]').fill(
                    valor_numero_doc
                )
                await frame_locator.locator('[data-testid="data_emissao_nf"]').fill(
                    valor_data_emissao
                )
                await frame_locator.locator('[data-testid="valor_bruto"]').fill(
                    valor_bruto
                )
                print("\nTodos os campos da nota fiscal foram preenchidos com sucesso!")
                testid_do_botao_salvar = "botao_salvar_nfs"
                botao_salvar = frame_locator.locator(
                    f'[data-testid="{testid_do_botao_salvar}"]'
                )
                await botao_salvar.click()
                await frame_locator.wait_for_load_state("networkidle")
                print("Dados salvos com sucesso e a página foi atualizada!")
                # -----------------------------------------------------
                # CLICANDO EM INCLUIR NOVO TIPO DE SERVIÇO
                texto_da_secao = "Tipos de serviço"
                container_tipos_servico = page.locator(
                    f'*:has-text("{texto_da_secao}")'
                )
                link_incluir_novo_servico = container_tipos_servico.get_by_text(
                    "[Incluir Novo]"
                )
                await link_incluir_novo_servico.click()
                print(
                    "Link '[Incluir Novo]' da seção 'Tipos de serviço' clicado com sucesso!"
                )
                # -----------------------------------------------------
                # PREENCHENDO DADOS DO TIPO DE SERVIÇO
                valor_padrao_servico = "100000002"
                testid_do_select = "tipo_servico"
                select_locator = page.locator(f'[data-testid="{testid_do_select}"]')
                await select_locator.select_option(value=valor_padrao_servico)
                print("Campo 'Tipo de Serviço' selecionado com sucesso!")
                print(
                    "\nProcesso de automação finalizado. A janela do navegador permanecerá aberta."
                )
                valor_base = valor_bruto_nf
                valor_retido = valor_retencao_nf
                testid_base = "valor_base_ret"
                campo_base_ret = page.locator(f'[data-testid="{testid_base}"]')
                await campo_base_ret.fill(valor_base)
                testid_retencao = "valor_retencao"
                campo_retencao = page.locator(f'[data-testid="{testid_retencao}"]')
                await campo_retencao.fill(valor_retido)
                print("\nCampos de valores preenchidos com sucesso!")
                # -----------------------------------------------------
                # SALVANDO O FORMULÁRIO DO TIPO DE SERVIÇO
                testid_botao = "botao_salvar_info_tpserv"
                botao_salvar_servico = page.locator(f'[data-testid="{testid_botao}"]')
                await botao_salvar_servico.click()
                print("Botão 'Salvar' do serviço foi clicado.")
                # -----------------------------------------------------
                # SALVANDO COMO RASCUNHO
                testid_salvar_rascunho = "botao_salvar_rascunho"
                botao_salvar_rascunho = page.locator(
                    f'[data-testid="{testid_salvar_rascunho}"]'
                )
                await botao_salvar_rascunho.click()
                print("Botão 'Salvar rascunho' apareceu e foi clicado com sucesso!")
                # -----------------------------------------------------

    except Exception as e:
        print(f"Ocorreu um erro: {e}")
    finally:
        # Importante: 'browser.close()' ao conectar via CDP apenas desconecta o script.
        # A janela do navegador que abrimos NÃO será fechada.
        if browser:
            await browser.close()
            print("Conexão do Playwright desconectada.")


# --- Ponto de entrada do script ---
if __name__ == "__main__":
    # Executa a função principal assíncrona
    asyncio.run(main())
