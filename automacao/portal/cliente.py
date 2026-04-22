"""
Módulo responsável pela seleção e vinculação do cliente ao pedido no portal.

A seleção do cliente é o primeiro passo de qualquer pedido. Este módulo
cuida de todo o fluxo: abrir o portal, pesquisar o cliente, selecionar
o resultado correto e confirmar os dados.
"""

from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

from automacao.portal import janela


def selecionar_cliente(navegador, wait, nome_cliente, janela_principal, url_portal):
    """
    Abre o portal, pesquisa e vincula o cliente ao pedido atual.

    Args:
        navegador: instância do WebDriver (Chrome).
        wait: WebDriverWait configurado.
        nome_cliente: nome do cliente a ser selecionado (deve existir no cadastro).
        janela_principal: handle da janela principal do navegador.
        url_portal: URL do portal onde pedidos são criados.
    """
    print(f"  Selecionando cliente: {nome_cliente}")

    # 1. Abre o portal e entra no frame principal
    navegador.get(url_portal)
    janela.entrar_frame_cadastro(navegador, wait)

    # 2. Digita o nome do cliente e clica em "Procurar"
    campo_cliente = wait.until(EC.element_to_be_clickable((By.ID, 'iCliente')))
    campo_cliente.clear()
    campo_cliente.send_keys(nome_cliente)
    navegador.find_element(By.ID, 'botaoProcuraCliente').click()

    # 3. Espera o popup de resultados abrir e troca para ele
    janela.aguardar_abertura_popup(navegador, wait, janela_principal)

    # 4. Clica no resultado que bate com o nome do cliente
    wait.until(EC.element_to_be_clickable(
        (By.XPATH, f"//td[contains(text(), '{nome_cliente}')]")
    )).click()

    # 5. Volta para a janela principal e confirma os dados do cliente
    navegador.switch_to.window(janela_principal)
    janela.entrar_frame_cadastro(navegador, wait)
    wait.until(EC.frame_to_be_available_and_switch_to_it("frameConfirmacaoDadosCliente"))
    wait.until(EC.element_to_be_clickable((By.ID, 'botaoConfirmarCliente'))).click()

    print("  Cliente vinculado.")