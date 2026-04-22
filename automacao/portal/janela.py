"""
Helpers para gerenciamento de janelas e iframes do portal.

O sistema que estamos automatizando usa uma arquitetura com múltiplos
iframes aninhados e popups que abrem em janelas separadas. Este módulo
centraliza toda a lógica de navegação entre esses contextos, evitando
duplicação nos módulos de mais alto nível (cliente, pedido, item).
"""

from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC


def entrar_frame_cadastro(navegador, wait):
    """
    Navega para o frame principal do sistema ('cadastro').

    Todas as operações do portal (buscar cliente, lançar item, etc)
    acontecem dentro desse frame, então praticamente toda função que
    interage com o sistema começa chamando este helper.
    """
    navegador.switch_to.default_content()
    wait.until(EC.frame_to_be_available_and_switch_to_it("cadastro"))


def entrar_frame_itens(navegador, wait):
    """
    Navega até a grade de itens do pedido (frame 'frameItemPedido'),
    que está aninhado DENTRO do frame 'cadastro'.

    Depois de usar, geralmente você quer voltar ao frame pai com
    navegador.switch_to.parent_frame() para acessar botões que ficam
    fora da grade (como 'Excluir').
    """
    entrar_frame_cadastro(navegador, wait)
    iframe_itens = wait.until(
        EC.presence_of_element_located((By.ID, "frameItemPedido"))
    )
    navegador.switch_to.frame(iframe_itens)


def fechar_janelas_orfas(navegador, janela_principal):
    """
    Fecha todas as janelas extras além da principal.

    Usado preventivamente no início de cada operação que abre popup,
    para garantir que popups esquecidos de operações anteriores não
    confundam o Selenium (ele pode pegar a janela errada via
    window_handles[-1]).

    Args:
        navegador: instância do WebDriver.
        janela_principal: handle da janela que deve ser preservada.

    Returns:
        int: quantidade de janelas órfãs que foram fechadas.
    """
    if len(navegador.window_handles) <= 1:
        return 0

    janelas_orfas = [
        h for h in navegador.window_handles if h != janela_principal
    ]

    for handle in janelas_orfas:
        try:
            navegador.switch_to.window(handle)
            navegador.close()
        except Exception as e:
            print(f"      Erro ao fechar janela órfã: {e}")

    navegador.switch_to.window(janela_principal)
    return len(janelas_orfas)


def esperar_popup_fechar(navegador, janela_principal, timeout=3):
    """
    Aguarda o popup fechar sozinho após uma ação (como clicar em OK).

    Se o popup não fechar dentro do timeout, fecha manualmente. Isso
    é importante porque, se deixarmos uma janela órfã aberta, o próximo
    window_handles[-1] pode pegar ela por engano, causando o bug que
    depuramos anteriormente.

    Args:
        navegador: instância do WebDriver.
        janela_principal: handle da janela principal.
        timeout: segundos para aguardar o fechamento automático.
    """
    try:
        WebDriverWait(navegador, timeout).until(
            EC.number_of_windows_to_be(1)
        )
    except Exception:
        print(f"   ⚠ Popup não fechou sozinho. Fechando manualmente...")
        fechar_janelas_orfas(navegador, janela_principal)
        print(f"   ✓ Popup(s) fechado(s) manualmente.")


def aguardar_abertura_popup(navegador, wait, janela_principal):
    """
    Aguarda um novo popup abrir e faz switch para ele.

    Deve ser chamado logo após uma ação que sabidamente abre popup
    (ex.: clicar em 'botaoProcuraProduto').
    """
    wait.until(EC.number_of_windows_to_be(2))
    novo_handle = [
        h for h in navegador.window_handles if h != janela_principal
    ][0]
    navegador.switch_to.window(novo_handle)