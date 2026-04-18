import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.keys import Keys
import time
import getpass
import configparser
import re


# ==================================================================
# 1. MANUAL DE INSTRUÇÕES (Funções)
# ==================================================================

def selecionar_cliente(navegador, wait, nome_cliente, janela_principal):
    """Realiza a busca e confirmação do cliente no portal."""
    print(f"  Selecionando cliente: {nome_cliente}")
    navegador.get(url_portal)

    wait.until(EC.frame_to_be_available_and_switch_to_it("cadastro"))

    campo_cliente = wait.until(EC.element_to_be_clickable((By.ID, 'iCliente')))
    campo_cliente.clear()
    campo_cliente.send_keys(nome_cliente)
    navegador.find_element(By.ID, 'botaoProcuraCliente').click()

    wait.until(EC.number_of_windows_to_be(2))
    nova_janela = [j for j in navegador.window_handles if j != janela_principal][0]
    navegador.switch_to.window(nova_janela)

    wait.until(EC.element_to_be_clickable((By.XPATH, f"//td[contains(text(), '{nome_cliente}')]"))).click()

    navegador.switch_to.window(janela_principal)
    navegador.switch_to.default_content()
    wait.until(EC.frame_to_be_available_and_switch_to_it("cadastro"))
    wait.until(EC.frame_to_be_available_and_switch_to_it("frameConfirmacaoDadosCliente"))
    wait.until(EC.element_to_be_clickable((By.ID, 'botaoConfirmarCliente'))).click()
    print("  Cliente vinculado.")


def limpar_itens_pedido(navegador, wait):
    """Remove todos os itens da grade atual."""
    print("🧹 Executando limpeza de itens...")
    navegador.switch_to.default_content()
    wait.until(EC.frame_to_be_available_and_switch_to_it("cadastro"))

    iframe_itens = wait.until(EC.presence_of_element_located((By.ID, "frameItemPedido")))
    navegador.switch_to.frame(iframe_itens)
    wait.until(EC.element_to_be_clickable((By.ID, 'selecionaTodosItens'))).click()

    navegador.switch_to.parent_frame()
    xpath_remover = "//a[contains(@onclick, 'removeItemSelecionado()')]"
    wait.until(EC.element_to_be_clickable((By.XPATH, xpath_remover))).click()
    time.sleep(5)
    print("Comando de remoção enviado.")


def consultar_pedido_pelo_numero(navegador, wait, numero_pedido):
    """Limpa os campos usando o botão nativo e consulta o novo ID."""
    print(f"  Consultando pedido: {numero_pedido}")
    navegador.switch_to.default_content()
    wait.until(EC.frame_to_be_available_and_switch_to_it("cadastro"))

    try:
        xpath_limpar = "//a[contains(@onclick, 'limparCampos()')]"
        btn_limpar = wait.until(EC.element_to_be_clickable((By.XPATH, xpath_limpar)))
        navegador.execute_script("arguments[0].click();", btn_limpar)
        print("  Botão limpar clickado.")
        time.sleep(2)
    except:
        print("  Não foi possível clicar no botão limpar, tentando seguir...")

    campo_num = wait.until(EC.element_to_be_clickable((By.ID, 'iNumeroPedido')))
    campo_num.send_keys(Keys.CONTROL + "a")
    campo_num.send_keys(Keys.BACKSPACE)
    campo_num.send_keys(numero_pedido)

    wait.until(EC.element_to_be_clickable((By.ID, 'consultaPedido'))).click()

    print("   Aguardando confirmação de itens zerados...")
    wait.until(EC.text_to_be_present_in_element((By.ID, 'qtdePedido'), "0"))
    print(f"  Pedido {numero_pedido} está vazio, iniciando lançamento de itens ...")


def verificar_e_excluir_itens_outro_estoque(navegador, wait):
    """
    Verifica na grade de itens se existem linhas marcadas com o ícone 'information.png'
    (Item em outro estoque). Se houver, captura o nome dos produtos, marca os checkboxes
    dessas linhas e aciona o botão Excluir.

    Retorna:
        list[str]: Lista com os nomes dos produtos excluídos (para log/relatório).
    """
    produtos_excluidos = []

    navegador.switch_to.default_content()
    wait.until(EC.frame_to_be_available_and_switch_to_it("cadastro"))

    try:
        iframe_itens = wait.until(EC.presence_of_element_located((By.ID, "frameItemPedido")))
        navegador.switch_to.frame(iframe_itens)
    except Exception as e:
        print(f"    Não foi possível acessar o frameItemPedido para verificação: {e}")
        return produtos_excluidos

    xpath_linhas_outro_estoque = "//tr[.//img[contains(@src, 'information.png')]]"
    linhas = navegador.find_elements(By.XPATH, xpath_linhas_outro_estoque)

    if not linhas:
        navegador.switch_to.default_content()
        wait.until(EC.frame_to_be_available_and_switch_to_it("cadastro"))
        return produtos_excluidos

    print(f"    {len(linhas)} item(ns) detectado(s) em OUTRO ESTOQUE na grade. Excluindo...")

    for linha in linhas:
        try:
            onclick_value = linha.get_attribute("onclick") or ""
            match = re.search(r"retornaAtualizacao\('([^']+)'", onclick_value)
            if match:
                nome_produto = match.group(1).strip()
            else:
                nome_produto = "(nome não identificado)"

            try:
                checkbox = linha.find_element(By.XPATH, ".//td[contains(@class,'controle')]//input[@type='checkbox']")
                if not checkbox.is_selected():
                    navegador.execute_script("arguments[0].click();", checkbox)
            except Exception:
                try:
                    checkbox = linha.find_element(By.XPATH, ".//input[@type='checkbox']")
                    if not checkbox.is_selected():
                        navegador.execute_script("arguments[0].click();", checkbox)
                except Exception as e_chk2:
                    print(f"    Não consegui marcar o checkbox do item '{nome_produto}': {e_chk2}")
                    continue

            produtos_excluidos.append(nome_produto)
            print(f"    Marcado para exclusão: {nome_produto}")

        except Exception as e:
            print(f"    Erro ao processar linha em outro estoque: {e}")
            continue

    navegador.switch_to.parent_frame()

    if produtos_excluidos:
        try:
            xpath_remover = "//a[contains(@onclick, 'removeItemSelecionado()')]"
            btn_excluir = wait.until(EC.element_to_be_clickable((By.XPATH, xpath_remover)))
            navegador.execute_script("arguments[0].click();", btn_excluir)
            time.sleep(3)
            print(f"    {len(produtos_excluidos)} item(ns) excluído(s) com sucesso.")
        except Exception as e:
            print(f"    Falha ao clicar em Excluir: {e}")

    navegador.switch_to.default_content()
    wait.until(EC.frame_to_be_available_and_switch_to_it("cadastro"))

    """"

    Adicionar Clear() após realizar a exclusão 
    #####################
    #####################
    #####################

    """

    return produtos_excluidos


def lancar_item(navegador, wait, janela_principal, codigo, nome_produto, quantidade,
                ignorar_outro_estoque=False):
    """
    Executa o fluxo completo de lançamento de UM item.

    Parâmetros:
      - ignorar_outro_estoque (bool): Se True, lança o item MESMO se o popup
        apontar que é de outro estoque (information.png), e também NÃO executa
        a verificação pós-lançamento na grade (que excluiria a linha).
        Use True no pedido dedicado ao relançamento de itens em outro estoque.

    Retorna uma tupla (status, detalhe) onde status pode ser:
      - "lancado"               -> item lançado com sucesso na grade
      - "lancado_e_excluido"    -> foi lançado mas caiu em outro estoque na grade e foi excluído
                                   (só ocorre quando ignorar_outro_estoque=False)
      - "sem_estoque"           -> popup indicou estoque zerado
      - "outro_estoque_popup"   -> popup já indicava outro estoque
                                   (só ocorre quando ignorar_outro_estoque=False)
      - "erro"                  -> exceção durante o processo
    """
    info_completa = f"{codigo} - {nome_produto}"

    try:
        navegador.switch_to.default_content()
        wait.until(EC.frame_to_be_available_and_switch_to_it("cadastro"))

        campo_cod = wait.until(EC.element_to_be_clickable((By.ID, "iProduto")))
        campo_cod.clear()
        campo_cod.send_keys(codigo)
        navegador.find_element(By.ID, "botaoProcuraProduto").click()

        wait.until(EC.number_of_windows_to_be(2))
        navegador.switch_to.window(navegador.window_handles[-1])

        elemento_estq = wait.until(
            EC.visibility_of_element_located((By.XPATH, "//td[contains(@class, 'estoqueProduto')]"))
        )
        val_estq = float(elemento_estq.text.split('\n')[0].strip().replace(',', '.'))
        tem_info = len(navegador.find_elements(By.XPATH, "//img[contains(@src, 'information.png')]"))

        if val_estq <= 0:
            print(f"  {info_completa}: Sem estoque.")
            navegador.close()
            navegador.switch_to.window(janela_principal)
            time.sleep(3)
            return ("sem_estoque", info_completa)

        # Só rejeita por "outro estoque" se NÃO estivermos no modo relançamento
        if not ignorar_outro_estoque and val_estq > 0 and tem_info > 1:
            print(f"ℹ  {info_completa}: Outro estoque (popup).")
            navegador.close()
            navegador.switch_to.window(janela_principal)
            time.sleep(3)
            return ("outro_estoque_popup", info_completa)

        if ignorar_outro_estoque and tem_info > 1:
            print(f"  {info_completa}: Outro estoque detectado no popup — LANÇANDO MESMO ASSIM (modo relançamento).")

        # Lança o item
        campo_qtd = wait.until(EC.visibility_of_element_located((By.NAME, "quantidade0")))
        navegador.execute_script("arguments[0].focus();", campo_qtd)
        campo_qtd.clear()
        campo_qtd.send_keys(str(quantidade))

        btn_ok = wait.until(EC.presence_of_element_located((By.XPATH, "//input[@value='OK']")))
        navegador.execute_script("arguments[0].click();", btn_ok)
        print(f"{info_completa}: Lançado {quantidade}")

        navegador.switch_to.window(janela_principal)
        time.sleep(3)

        # Só faz a verificação pós-lançamento no modo normal.
        # No modo relançamento, queremos MANTER o item mesmo com information.png.
        if not ignorar_outro_estoque:
            try:
                excluidos = verificar_e_excluir_itens_outro_estoque(navegador, wait)
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


# =================================================================
# 2. CONFIGURAÇÃO E LOGIN
# =================================================================

config = configparser.ConfigParser()
config.read("config.ini", encoding="utf-8")
url_login = config["GERAL"]["url_login"]
url_portal = config["GERAL"]["url_portal"]
usuario = input(str("  Digite o nome do Usuário:"))
senha = getpass.getpass("  Digite sua senha (Sua senha não vai aparecer):")

itens_em_falta = []
itens_em_outro_estoque = []
itens_para_relancar = []   # NOVA fila: itens que caíram em outro estoque e vão para um pedido separado
pedidos_gerados = []

navegador = webdriver.Chrome()
navegador.maximize_window()
wait = WebDriverWait(navegador, 20)

navegador.get(url_login)
navegador.find_element(By.ID, "login").send_keys(usuario)
navegador.find_element(By.ID, "senha").send_keys(senha)
navegador.find_element(By.XPATH, "//button[contains(@onclick, 'logar()')]").click()
wait.until(EC.url_changes(navegador.current_url))

janela_principal = navegador.current_window_handle

nome_cliente = config["GERAL"]["nome_cliente"]

# Planilha
sheet_id = config["GERAL"]["planilha_id"]
url = f"https://docs.google.com/spreadsheets/d/{sheet_id}/export?format=xlsx"
planilha_completa = pd.read_excel(url, sheet_name=None)

# =================================================================
# 3. EXECUÇÃO POR ABA
# =================================================================

primeira_aba = True

for nome_aba, dados_do_pedido in planilha_completa.items():
    if dados_do_pedido.empty: continue
    soma_qtd = pd.to_numeric(dados_do_pedido['QTD'], errors='coerce').sum()
    if soma_qtd <= 0:
        print(f"PULANDO ABA: {nome_aba} (Nenhuma quantidade lançada)")
        continue
    print(f"\n   --- INICIANDO ABA: {nome_aba} ---")

    if primeira_aba:
        selecionar_cliente(navegador, wait, nome_cliente, janela_principal)
        navegador.switch_to.default_content()
        wait.until(EC.frame_to_be_available_and_switch_to_it("cadastro"))

        quantidade_texto = wait.until(EC.presence_of_element_located((By.ID, "qtdePedido"))).text
        quantidade_inicial = int(quantidade_texto)

        if quantidade_inicial > 0:
            print(f"  Sheet1 detectou {quantidade_inicial} itens residuais. Limpando...")
            limpar_itens_pedido(navegador, wait)
            num_pedido_atual = navegador.find_element(By.ID, "iNumeroPedido").get_attribute("value")
            consultar_pedido_pelo_numero(navegador, wait, num_pedido_atual)
        else:
            print(f"  {nome_aba} iniciada com pedido limpo.")

        num_pedido = navegador.find_element(By.ID, "iNumeroPedido").get_attribute("value")
        pedidos_gerados.append(f"Aba: {nome_aba} -> Pedido: {num_pedido}")
        primeira_aba = False
    else:
        navegador.switch_to.default_content()
        wait.until(EC.frame_to_be_available_and_switch_to_it("cadastro"))

        id_antigo = navegador.find_element(By.ID, "iNumeroPedido").get_attribute("value")

        print(f"  Duplicando pedido {id_antigo}...")
        xpath_dup = "//a[contains(@onclick, 'duplicarPedido()')]"
        btn_dup = wait.until(EC.element_to_be_clickable((By.XPATH, xpath_dup)))
        navegador.execute_script("arguments[0].click();", btn_dup)

        wait.until(lambda d: d.find_element(By.ID, "iNumeroPedido").get_attribute("value") != id_antigo)
        novo_id = navegador.find_element(By.ID, "iNumeroPedido").get_attribute("value")
        pedidos_gerados.append(f"Aba: {nome_aba} -> Pedido: {novo_id}")
        limpar_itens_pedido(navegador, wait)
        consultar_pedido_pelo_numero(navegador, wait, novo_id)

    # Lançamento dos Itens
    itens_validos = dados_do_pedido.dropna(subset=["CODIGO", "QTD"])
    for _, linha in itens_validos.iterrows():
        codigo = str(linha["CODIGO"])
        nome_produto = str(linha["PRODUTO"])
        info_completa = f"{codigo} - {nome_produto}"

        try:
            quantidade = str(int(linha["QTD"]))
        except Exception as e:
            print(f"  Quantidade inválida para {info_completa}: {e}")
            continue

        status, detalhe = lancar_item(navegador, wait, janela_principal, codigo, nome_produto, quantidade)

        if status == "sem_estoque":
            itens_em_falta.append(info_completa)
        elif status == "outro_estoque_popup":
            itens_em_outro_estoque.append(info_completa)
            itens_para_relancar.append({
                "codigo": codigo,
                "quantidade": quantidade,
                "produto": nome_produto,
                "origem_aba": nome_aba,
                "motivo": "outro_estoque_popup",
            })
        elif status == "lancado_e_excluido":
            nomes_grade = detalhe.get("excluidos_grade", []) if isinstance(detalhe, dict) else []
            marcador_grade = " | ".join(nomes_grade) if nomes_grade else "(sem nome)"
            registro = f"{info_completa} | (na grade: {marcador_grade})"
            itens_em_outro_estoque.append(registro)
            itens_para_relancar.append({
                "codigo": codigo,
                "quantidade": quantidade,
                "produto": nome_produto,
                "origem_aba": nome_aba,
                "motivo": "outro_estoque_grade",
            })
        elif status == "lancado":
            pass
        elif status == "erro":
            pass

# =================================================================
# 4. PEDIDO SEPARADO PARA ITENS EM OUTRO ESTOQUE
# =================================================================

pedido_relancamento_info = None
resultado_relancamento = {
    "lancados": [],
    "ainda_outro_estoque": [],
    "sem_estoque": [],
    "erro": [],
}

if itens_para_relancar:
    print(f"\n   === INICIANDO PEDIDO DE RELANÇAMENTO ({len(itens_para_relancar)} itens) ===")

    try:
        navegador.switch_to.default_content()
        wait.until(EC.frame_to_be_available_and_switch_to_it("cadastro"))

        id_antigo = navegador.find_element(By.ID, "iNumeroPedido").get_attribute("value")

        print(f"  Duplicando pedido {id_antigo} para o relançamento...")
        xpath_dup = "//a[contains(@onclick, 'duplicarPedido()')]"
        btn_dup = wait.until(EC.element_to_be_clickable((By.XPATH, xpath_dup)))
        navegador.execute_script("arguments[0].click();", btn_dup)

        wait.until(lambda d: d.find_element(By.ID, "iNumeroPedido").get_attribute("value") != id_antigo)
        novo_id_relancamento = navegador.find_element(By.ID, "iNumeroPedido").get_attribute("value")

        pedidos_gerados.append(f"Aba: RELANÇAMENTO (outro estoque) -> Pedido: {novo_id_relancamento}")
        pedido_relancamento_info = novo_id_relancamento

        # Limpa grade herdada e consulta novamente para refresh
        limpar_itens_pedido(navegador, wait)
        consultar_pedido_pelo_numero(navegador, wait, novo_id_relancamento)

        # Lança cada item pendente
        for item in itens_para_relancar:
            codigo = item["codigo"]
            qtd = item["quantidade"]
            produto = item["produto"]
            origem = item["origem_aba"]
            info = f"{codigo} - {produto} (origem: {origem})"

            status, detalhe = lancar_item(
                navegador, wait, janela_principal,
                codigo, produto, qtd,
                ignorar_outro_estoque=True,  # força lançamento mesmo com information.png
            )

            if status == "lancado":
                resultado_relancamento["lancados"].append(info)
            elif status == "lancado_e_excluido":
                resultado_relancamento["ainda_outro_estoque"].append(info + " [excluído da grade novamente]")
            elif status == "outro_estoque_popup":
                resultado_relancamento["ainda_outro_estoque"].append(info + " [popup]")
            elif status == "sem_estoque":
                resultado_relancamento["sem_estoque"].append(info)
            else:
                resultado_relancamento["erro"].append(info)

    except Exception as e:
        print(f"   Falha ao iniciar pedido de relançamento: {e}")

# Fim
print(f"\n  Processo Concluído!")

# --- FINALIZAÇÃO E GERAÇÃO DE LOG ---
nome_arquivo_log = f"log_pedidos_{time.strftime('%Y%m%d_%H%M%S')}.txt"

with open(nome_arquivo_log, "w", encoding="utf-8") as f:
    f.write("=== RELATÓRIO DE LANÇAMENTO AUTOMÁTICO ===\n")
    f.write(f"Data/Hora: {time.ctime()}\n\n")

    f.write("  PEDIDOS GERADOS NESTA SESSÃO:\n")
    if pedidos_gerados:
        for p in pedidos_gerados:
            f.write(f"{p}\n")
    else:
        f.write("- Nenhum pedido foi gerado.\n")

    f.write("\n" + "=" * 40 + "\n\n")

    f.write("⚠️ ITENS EM FALTA (ESTOQUE ZERADO):\n")
    if itens_em_falta:
        for info in itens_em_falta:
            f.write(f"- {info}\n")
    else:
        f.write("- Nenhum item em falta.\n")

    f.write("\nITENS EM OUTRO ESTOQUE (DETECTADOS):\n")
    if itens_em_outro_estoque:
        for info in itens_em_outro_estoque:
            f.write(f"- {info}\n")
    else:
        f.write("- Nenhum item em outro estoque.\n")

    f.write("\n" + "=" * 40 + "\n\n")
    f.write("📦 PEDIDO DE RELANÇAMENTO (OUTRO ESTOQUE):\n")
    if pedido_relancamento_info:
        f.write(f"Pedido nº: {pedido_relancamento_info}\n")
        f.write(f"Total de itens enviados para relançamento: {len(itens_para_relancar)}\n\n")

        f.write("  LANÇADOS COM SUCESSO NO PEDIDO DE RELANÇAMENTO:\n")
        if resultado_relancamento["lancados"]:
            for info in resultado_relancamento["lancados"]:
                f.write(f"- {info}\n")
        else:
            f.write("- Nenhum.\n")

        f.write("\n  AINDA EM OUTRO ESTOQUE (NÃO LANÇADOS NO RELANÇAMENTO):\n")
        if resultado_relancamento["ainda_outro_estoque"]:
            for info in resultado_relancamento["ainda_outro_estoque"]:
                f.write(f"- {info}\n")
        else:
            f.write("- Nenhum.\n")

        f.write("\n  SEM ESTOQUE NO RELANÇAMENTO:\n")
        if resultado_relancamento["sem_estoque"]:
            for info in resultado_relancamento["sem_estoque"]:
                f.write(f"- {info}\n")
        else:
            f.write("- Nenhum.\n")

        if resultado_relancamento["erro"]:
            f.write("\n  ERROS NO RELANÇAMENTO:\n")
            for info in resultado_relancamento["erro"]:
                f.write(f"- {info}\n")
    else:
        f.write("- Nenhum item precisou ser relançado (nenhum item caiu em outro estoque).\n")

print(f" Relatório completo gerado: {nome_arquivo_log}")
input("\nProcesso finalizado. Pressione Enter para fechar...")

navegador.quit()