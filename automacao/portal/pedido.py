"""
Módulo responsável pelas operações de gestão de pedido no portal:
duplicar, limpar a grade de itens, consultar por número, etc.

Essas operações são os "blocos de construção" usados pelo loop principal
para preparar cada pedido antes de lançar os itens.
"""

import time

from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

from automacao.portal import janela


def obter_numero_atual(navegador) -> str:
    """Retorna o número do pedido que está carregado na tela."""
    return navegador.find_element(By.ID, "iNumeroPedido").get_attribute("value")


def contar_itens_no_pedido(navegador, wait) -> int:
    """Retorna a quantidade de itens que o pedido atual tem na grade."""
    elemento = wait.until(EC.presence_of_element_located((By.ID, "qtdePedido")))
    texto = elemento.text.strip()

    if not texto:
        wait.until(lambda d: d.find_element(By.ID, "qtdePedido").text.strip() != "")
        texto = navegador.find_element(By.ID, "qtdePedido").text.strip()  # <-- usa aqui

    return int(texto)


def duplicar_pedido_atual(navegador, wait) -> str:
    """
    Duplica o pedido atualmente carregado (cria um novo pedido copiando
    cliente e configurações). Retorna o número do novo pedido.

    A duplicação traz junto os itens do pedido antigo — é responsabilidade
    do chamador limpá-los se necessário (ver limpar_itens).
    """
    id_antigo = obter_numero_atual(navegador)
    print(f"  Duplicando pedido {id_antigo}...")

    xpath_dup = "//a[contains(@onclick, 'duplicarPedido()')]"
    btn_dup = wait.until(EC.element_to_be_clickable((By.XPATH, xpath_dup)))
    navegador.execute_script("arguments[0].click();", btn_dup)

    # Espera o ID mudar (sinal de que a duplicação concluiu)
    wait.until(
        lambda d: d.find_element(By.ID, "iNumeroPedido").get_attribute("value") != id_antigo
    )
    return obter_numero_atual(navegador)


def limpar_itens(navegador, wait):
    """
    Remove todos os itens da grade do pedido atual.

    Geralmente usado logo após duplicar_pedido_atual(), para limpar os
    itens herdados do pedido anterior.
    """
    print("🧹 Executando limpeza de itens...")
    janela.entrar_frame_itens(navegador, wait)

    wait.until(EC.element_to_be_clickable((By.ID, 'selecionaTodosItens'))).click()

    navegador.switch_to.parent_frame()
    xpath_remover = "//a[contains(@onclick, 'removeItemSelecionado()')]"
    wait.until(EC.element_to_be_clickable((By.XPATH, xpath_remover))).click()
    time.sleep(5)
    print("Comando de remoção enviado.")


def consultar_por_numero(navegador, wait, numero_pedido: str):
    """
    Limpa o formulário via botão nativo 'limparCampos()' e consulta o
    pedido pelo número informado.

    Útil para forçar atualização da tela após operações que podem deixar
    estado residual no formulário.
    """
    print(f"  Consultando pedido: {numero_pedido}")
    janela.entrar_frame_cadastro(navegador, wait)

    try:
        xpath_limpar = "//a[contains(@onclick, 'limparCampos()')]"
        btn_limpar = wait.until(EC.element_to_be_clickable((By.XPATH, xpath_limpar)))
        navegador.execute_script("arguments[0].click();", btn_limpar)
        print("  Botão limpar clickado.")
        time.sleep(2)
    except Exception:
        print("  Não foi possível clicar no botão limpar, tentando seguir...")

    # Digita o número usando CTRL+A + BACKSPACE (mais robusto que .clear())
    campo_num = wait.until(EC.element_to_be_clickable((By.ID, 'iNumeroPedido')))
    campo_num.send_keys(Keys.CONTROL + "a")
    campo_num.send_keys(Keys.BACKSPACE)
    campo_num.send_keys(numero_pedido)

    wait.until(EC.element_to_be_clickable((By.ID, 'consultaPedido'))).click()

    print("   Aguardando confirmação de itens zerados...")
    wait.until(EC.text_to_be_present_in_element((By.ID, 'qtdePedido'), "0"))
    print(f"  Pedido {numero_pedido} está vazio, iniciando lançamento de itens ...")


def preparar_pedido_duplicado(navegador, wait) -> str:
    """
    Fluxo completo de preparação de um pedido duplicado pronto para receber
    novos itens: duplica, limpa os itens herdados e consulta para refresh.

    Retorna o número do novo pedido.

    Essa função encapsula uma sequência que aparece duas vezes no código
    original (transição entre abas e criação do pedido de relançamento).
    """
    novo_id = duplicar_pedido_atual(navegador, wait)
    limpar_itens(navegador, wait)
    consultar_por_numero(navegador, wait, novo_id)
    return novo_id