"""
Módulo responsável pelo lançamento de itens no pedido atual, incluindo:
- Lançamento de um item individual (lancar_item)
- Verificação pós-lançamento de itens em outro estoque
- Detecção e exclusão automática de itens com ícone 'information.png'

Esta é a parte do sistema com mais lógica de negócio. A função lancar_item
é usada tanto no loop principal quanto no pedido de relançamento, com
comportamento ajustável via parâmetro.
"""

import re
import time

from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

from automacao.portal import janela
from automacao.portal import pedido


# ============================================================
# DETECÇÃO E EXCLUSÃO DE ITENS EM OUTRO ESTOQUE NA GRADE
# ============================================================

XPATH_LINHAS_OUTRO_ESTOQUE = "//tr[.//img[contains(@src, 'information.png')]]"
XPATH_TODAS_LINHAS = (
    "//table[@id='produtosInterna']//tr[@onclick and "
    "contains(@onclick, 'retornaAtualizacao')]"
)


def _extrair_nome_do_onclick(linha) -> str:
    """
    Extrai o nome do produto do atributo onclick da linha.

    O onclick tem o formato:
        retornaAtualizacao('NOME_PRODUTO', '10.0', ...)
    """
    onclick_value = linha.get_attribute("onclick") or ""
    match = re.search(r"retornaAtualizacao\('([^']+)'", onclick_value)
    if match:
        return match.group(1).strip()
    return "(nome não identificado)"


def _marcar_checkbox_da_linha(navegador, linha, nome_produto):
    """
    Marca o checkbox da linha para seleção. Tenta primeiro pelo caminho
    esperado (td.controle); se falhar, procura qualquer checkbox na linha.
    """
    try:
        checkbox = linha.find_element(
            By.XPATH, ".//td[contains(@class,'controle')]//input[@type='checkbox']"
        )
    except Exception:
        checkbox = linha.find_element(By.XPATH, ".//input[@type='checkbox']")

    if not checkbox.is_selected():
        navegador.execute_script("arguments[0].click();", checkbox)


def _contar_linhas_na_grade(navegador) -> int:
    """Conta quantas linhas de produto existem atualmente na grade de itens."""
    janela.entrar_frame_itens(navegador, WebDriverWait(navegador, 5))
    qtd = len(navegador.find_elements(By.XPATH, XPATH_TODAS_LINHAS))
    navegador.switch_to.parent_frame()
    return qtd


def _aguardar_remocao_de_linhas(navegador, linhas_antes, qtd_a_remover):
    """
    Espera ATIVAMENTE a grade refletir a exclusão. Muito mais confiável
    que um time.sleep fixo.
    """
    quantidade_esperada = linhas_antes - qtd_a_remover

    def linhas_reduziram(driver):
        try:
            driver.switch_to.default_content()
            driver.switch_to.frame("cadastro")
            driver.switch_to.frame("frameItemPedido")
            atuais = len(driver.find_elements(By.XPATH, XPATH_TODAS_LINHAS))
            driver.switch_to.default_content()
            driver.switch_to.frame("cadastro")
            return atuais <= quantidade_esperada
        except Exception:
            try:
                driver.switch_to.default_content()
                driver.switch_to.frame("cadastro")
            except Exception:
                pass
            return False

    try:
        WebDriverWait(navegador, 15).until(linhas_reduziram)
        print(f"    Backend confirmou remoção. Grade atualizada.")
    except Exception as e:
        print(f"    Timeout aguardando remoção (fallback para sleep). Detalhe: {e}")
        time.sleep(3)


def _resetar_estado_apos_exclusao(navegador, wait):
    """
    Após excluir linhas, o sistema deixa variáveis internas com estado
    residual do item excluído. A forma confiável de limpar é usar o
    botão nativo limparCampos() e recarregar o pedido.
    """
    try:
        num_pedido_atual = pedido.obter_numero_atual(navegador)
        print(f"    Salvando número do pedido atual: {num_pedido_atual}")

        xpath_limpar = "//a[contains(@onclick, 'limparCampos()')]"
        btn_limpar = wait.until(EC.element_to_be_clickable((By.XPATH, xpath_limpar)))
        navegador.execute_script("arguments[0].click();", btn_limpar)
        print(f"    Botão limparCampos() nativo acionado.")
        time.sleep(2)

        # Recarrega o pedido
        campo_num = wait.until(EC.element_to_be_clickable((By.ID, 'iNumeroPedido')))
        campo_num.click()
        campo_num.send_keys(Keys.CONTROL + "a")
        campo_num.send_keys(Keys.BACKSPACE)
        campo_num.send_keys(num_pedido_atual)
        wait.until(EC.element_to_be_clickable((By.ID, 'consultaPedido'))).click()
        time.sleep(3)

        # Valida que voltamos ao pedido certo
        confirmado = pedido.obter_numero_atual(navegador)
        if confirmado != num_pedido_atual:
            print(f"    ⚠ ATENÇÃO: pedido mudou! esperado {num_pedido_atual}, atual {confirmado}")
        else:
            print(f"    ✓ Pedido {num_pedido_atual} recarregado com sucesso.")

    except Exception as e:
        print(f"    Falha no limparCampos+reconsulta: {e}")


def verificar_e_excluir_outro_estoque(navegador, wait) -> list:
    """
    Verifica na grade de itens se existem linhas marcadas com o ícone
    'information.png' (Item em outro estoque). Se houver, captura o nome
    dos produtos, marca os checkboxes e aciona o botão Excluir.

    Retorna:
        list[str]: Lista com os nomes dos produtos excluídos.
    """
    produtos_excluidos = []

    janela.entrar_frame_itens(navegador, wait)

    linhas = navegador.find_elements(By.XPATH, XPATH_LINHAS_OUTRO_ESTOQUE)

    if not linhas:
        janela.entrar_frame_cadastro(navegador, wait)
        return produtos_excluidos

    print(f"    {len(linhas)} item(ns) detectado(s) em OUTRO ESTOQUE na grade. Excluindo...")

    for linha in linhas:
        try:
            nome_produto = _extrair_nome_do_onclick(linha)
            _marcar_checkbox_da_linha(navegador, linha, nome_produto)
            produtos_excluidos.append(nome_produto)
            print(f"    Marcado para exclusão: {nome_produto}")
        except Exception as e:
            print(f"    Erro ao processar linha em outro estoque: {e}")
            continue

    navegador.switch_to.parent_frame()

    if produtos_excluidos:
        linhas_antes = _contar_linhas_na_grade(navegador)
        print(f"    Linhas na grade antes da exclusão: {linhas_antes}")

        try:
            xpath_remover = "//a[contains(@onclick, 'removeItemSelecionado()')]"
            btn_excluir = wait.until(EC.element_to_be_clickable((By.XPATH, xpath_remover)))
            navegador.execute_script("arguments[0].click();", btn_excluir)
            print(f"    Clique em Excluir enviado. Aguardando backend remover linha(s)...")
        except Exception as e:
            print(f"    Falha ao clicar em Excluir: {e}")
            linhas_antes = None

        if linhas_antes is not None:
            _aguardar_remocao_de_linhas(navegador, linhas_antes, len(produtos_excluidos))

        time.sleep(1)
        print(f"    {len(produtos_excluidos)} item(ns) excluído(s) com sucesso.")

        _resetar_estado_apos_exclusao(navegador, wait)

    janela.entrar_frame_cadastro(navegador, wait)
    return produtos_excluidos


# ============================================================
# LANÇAMENTO DE UM ITEM
# ============================================================

def _digitar_codigo_produto(navegador, wait, codigo: str):
    """Digita o código do produto no campo iProduto."""
    campo_cod = wait.until(EC.element_to_be_clickable((By.ID, "iProduto")))
    campo_cod.click()
    campo_cod.send_keys(Keys.CONTROL + "a")
    campo_cod.send_keys(Keys.BACKSPACE)
    campo_cod.send_keys(codigo)

    # Valida que o campo ficou com exatamente o código que queríamos
    valor_digitado = (campo_cod.get_attribute("value") or "").strip()
    if valor_digitado != str(codigo).strip():
        print(f"   ATENÇÃO: campo iProduto ficou com '{valor_digitado}' mas esperávamos '{codigo}'. Corrigindo...")
        navegador.execute_script("arguments[0].value = '';", campo_cod)
        campo_cod.send_keys(codigo)


def _ler_dados_do_popup(navegador, wait) -> tuple:
    """
    Lê o estoque e a presença do ícone information.png no popup de pesquisa.

    Returns:
        tuple(float, int): (valor_estoque, quantidade_icones_information)
    """
    elemento_estq = wait.until(
        EC.visibility_of_element_located((By.XPATH, "//td[contains(@class, 'estoqueProduto')]"))
    )
    val_estq = float(elemento_estq.text.split('\n')[0].strip().replace(',', '.'))
    tem_info = len(navegador.find_elements(By.XPATH, "//img[contains(@src, 'information.png')]"))
    return val_estq, tem_info


def _confirmar_lancamento(navegador, wait, quantidade: str):
    """Digita a quantidade e clica OK para lançar o item."""
    campo_qtd = wait.until(EC.visibility_of_element_located((By.NAME, "quantidade0")))
    navegador.execute_script("arguments[0].focus();", campo_qtd)
    campo_qtd.clear()
    campo_qtd.send_keys(str(quantidade))

    btn_ok = wait.until(EC.presence_of_element_located((By.XPATH, "//input[@value='OK']")))
    navegador.execute_script("arguments[0].click();", btn_ok)


def lancar_item(navegador, wait, janela_principal,
                codigo, nome_produto, quantidade,
                ignorar_outro_estoque=False):
    """
    Executa o fluxo completo de lançamento de UM item.

    Args:
        ignorar_outro_estoque: se True, lança o item mesmo se o popup indicar
            outro estoque, e NÃO executa a verificação pós-lançamento.
            Usado no pedido de relançamento.

    Returns:
        tuple (status, detalhe). Status pode ser:
            - "lancado"             -> item lançado com sucesso
            - "lancado_e_excluido"  -> lançado mas caiu em outro estoque e foi excluído
            - "sem_estoque"         -> popup indicou estoque zerado
            - "outro_estoque_popup" -> popup já indicava outro estoque
            - "erro"                -> exceção durante o processo
    """
    info_completa = f"{codigo} - {nome_produto}"

    try:
        # Rede de segurança: fecha janelas órfãs de interações anteriores
        qtd_orfas = janela.fechar_janelas_orfas(navegador, janela_principal)
        if qtd_orfas > 0:
            print(f"   ⚠ Fechou {qtd_orfas} janela(s) órfã(s) antes de prosseguir.")
            time.sleep(1)

        janela.entrar_frame_cadastro(navegador, wait)

        # Digita o código e clica em buscar produto
        _digitar_codigo_produto(navegador, wait, codigo)
        navegador.find_element(By.ID, "botaoProcuraProduto").click()

        janela.aguardar_abertura_popup(navegador, wait, janela_principal)

        # Lê os dados do popup
        val_estq, tem_info = _ler_dados_do_popup(navegador, wait)

        # Cenário 1: sem estoque
        if val_estq <= 0:
            print(f"  {info_completa}: Sem estoque.")
            navegador.close()
            navegador.switch_to.window(janela_principal)
            time.sleep(3)
            return ("sem_estoque", info_completa)

        # Cenário 2: outro estoque (só rejeita no modo normal)
        if not ignorar_outro_estoque and tem_info > 1:
            print(f"ℹ  {info_completa}: Outro estoque (popup).")
            navegador.close()
            navegador.switch_to.window(janela_principal)
            time.sleep(3)
            return ("outro_estoque_popup", info_completa)

        if ignorar_outro_estoque and tem_info > 1:
            print(f"  {info_completa}: Outro estoque detectado no popup — LANÇANDO MESMO ASSIM (modo relançamento).")

        # Cenário 3: lança o item
        _confirmar_lancamento(navegador, wait, quantidade)
        print(f"{info_completa}: Lançado {quantidade}")

        # Garante que o popup fechou
        janela.esperar_popup_fechar(navegador, janela_principal, timeout=3)
        navegador.switch_to.window(janela_principal)
        time.sleep(3)

        # Verificação pós-lançamento (só no modo normal)
        if not ignorar_outro_estoque:
            try:
                excluidos = verificar_e_excluir_outro_estoque(navegador, wait)
                if excluidos:
                    return ("lancado_e_excluido", {"info": info_completa, "excluidos_grade": excluidos})
            except Exception as e_verif:
                print(f"   Falha na verificação pós-lançamento: {e_verif}")

        return ("lancado", info_completa)

    except Exception as e:
        print(f"  Erro no item {codigo}: {e}")
        if len(navegador.window_handles) > 1:
            navegador.close()
        navegador.switch_to.window(janela_principal)
        return ("erro", f"{info_completa} :: {e}")