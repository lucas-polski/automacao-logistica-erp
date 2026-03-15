import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.keys import Keys
import time
import getpass
import configparser





# ==================================================================
# 1. MANUAL DE INSTRUÇÕES (Funções)
# ==================================================================

def selecionar_cliente(navegador, wait, nome_cliente, janela_principal):
    """Realiza a busca e confirmação do cliente no portal."""
    print(f"🔄 Selecionando cliente: {nome_cliente}")
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
    print("✅ Cliente vinculado.")


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
    print("✨ Comando de remoção enviado.")


def consultar_pedido_pelo_numero(navegador, wait, numero_pedido):
    """Limpa os campos usando o botão nativo e consulta o novo ID."""
    print(f"🔍 Consultando pedido: {numero_pedido}")
    navegador.switch_to.default_content()
    wait.until(EC.frame_to_be_available_and_switch_to_it("cadastro"))

    try:
        xpath_limpar = "//a[contains(@onclick, 'limparCampos()')]"
        btn_limpar = wait.until(EC.element_to_be_clickable((By.XPATH, xpath_limpar)))
        navegador.execute_script("arguments[0].click();", btn_limpar)
        print("Botão limpar clickado.")
        time.sleep(2)  # pausa para o sistema limpar os campos
    except:
        print("Não foi possível clicar no botão limpar, tentando seguir...")

    # localiza o campo de número e digita
    campo_num = wait.until(EC.element_to_be_clickable((By.ID, 'iNumeroPedido')))

    # caso o .clear() der erro:
    from selenium.webdriver.common.keys import Keys
    campo_num.send_keys(Keys.CONTROL + "a")
    campo_num.send_keys(Keys.BACKSPACE)

    campo_num.send_keys(numero_pedido)

    # clica no botão consultar
    wait.until(EC.element_to_be_clickable((By.ID, 'consultaPedido'))).click()

    # valida se a quantidade itens está zerada
    print("   Aguardando confirmação de itens zerados...")
    wait.until(EC.text_to_be_present_in_element((By.ID, 'qtdePedido'), "0"))
    print(f"✅ Pedido {numero_pedido} está vazio, iniciando lançamento de itens ...")

# =================================================================
# 2. CONFIGURAÇÃO E LOGIN
# =================================================================

config = configparser.ConfigParser()
config.read("config.ini", encoding="utf-8")
url_login = config["GERAL"]["url_login"]
url_portal = config["GERAL"]["url_portal"]
usuario = input(str("Digite o nome do Usuário:"))
senha = getpass.getpass("Digite sua senha (Sua senha não vai aparecer):")

itens_em_falta = []
itens_em_outro_estoque = []
pedidos_gerados = []

navegador = webdriver.Chrome()
navegador.maximize_window()
wait = WebDriverWait(navegador, 10)

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
    print(f"\n🚀 --- INICIANDO ABA: {nome_aba} ---")

    if primeira_aba:
        selecionar_cliente(navegador, wait, nome_cliente, janela_principal)
        navegador.switch_to.default_content()
        wait.until(EC.frame_to_be_available_and_switch_to_it("cadastro"))

        # Verifica a quantidade de itens na tela
        quantidade_texto = wait.until(EC.presence_of_element_located((By.ID, "qtdePedido"))).text
        quantidade_inicial = int(quantidade_texto)

        if quantidade_inicial > 0:
            print(f"⚠️ Sheet1 detectou {quantidade_inicial} itens residuais. Limpando...")
            limpar_itens_pedido(navegador, wait)
            # Re-seleciona ou consulta para garantir o refresh da tela, como você descobriu!
            num_pedido_atual = navegador.find_element(By.ID, "iNumeroPedido").get_attribute("value")
            consultar_pedido_pelo_numero(navegador, wait, num_pedido_atual)
        else:
            print(f"✅ {nome_aba} iniciada com pedido limpo.")



        # Captura o número do primeiro pedido
        num_pedido = navegador.find_element(By.ID, "iNumeroPedido").get_attribute("value")
        pedidos_gerados.append(f"Aba: {nome_aba} -> Pedido: {num_pedido}")
        primeira_aba = False
    else:
        # Ritual de Transição entre Abas
        navegador.switch_to.default_content()
        wait.until(EC.frame_to_be_available_and_switch_to_it("cadastro"))

        id_antigo = navegador.find_element(By.ID, "iNumeroPedido").get_attribute("value")

        # Duplica
        print(f"🔄 Duplicando pedido {id_antigo}...")
        xpath_dup = "//a[contains(@onclick, 'duplicarPedido()')]"
        btn_dup = wait.until(EC.element_to_be_clickable((By.XPATH, xpath_dup)))
        navegador.execute_script("arguments[0].click();", btn_dup)

        # Espera o ID mudar e captura o novo
        wait.until(lambda d: d.find_element(By.ID, "iNumeroPedido").get_attribute("value") != id_antigo)
        novo_id = navegador.find_element(By.ID, "iNumeroPedido").get_attribute("value")
        pedidos_gerados.append(f"Aba: {nome_aba} -> Pedido: {novo_id}")
        # Limpa e Força Atualização
        limpar_itens_pedido(navegador, wait)
        consultar_pedido_pelo_numero(navegador, wait, novo_id)

    # Lançamento dos Itens
    itens_validos = dados_do_pedido.dropna(subset=["CODIGO", "QTD"])
    for _, linha in itens_validos.iterrows():
        codigo = str(linha["CODIGO"])
        nome_produto = str(linha["PRODUTO"])  # Captura a coluna PRODUTO da sua planilha

        info_completa = f"{codigo} - {nome_produto}"

        try:
            quantidade = str(int(linha["QTD"]))

            navegador.switch_to.default_content()
            wait.until(EC.frame_to_be_available_and_switch_to_it("cadastro"))

            campo_cod = wait.until(EC.element_to_be_clickable((By.ID, "iProduto")))
            campo_cod.clear()
            campo_cod.send_keys(codigo)
            navegador.find_element(By.ID, "botaoProcuraProduto").click()

            wait.until(EC.number_of_windows_to_be(2))
            navegador.switch_to.window(navegador.window_handles[-1])

            # Checa estoque
            elemento_estq = wait.until(
                EC.visibility_of_element_located((By.XPATH, "//td[contains(@class, 'estoqueProduto')]")))
            val_estq = float(elemento_estq.text.split('\n')[0].strip().replace(',', '.'))
            tem_info = len(navegador.find_elements(By.XPATH, "//img[contains(@src, 'information.png')]"))

            if val_estq <= 0:
                print(f"⚠️ {info_completa}: Sem estoque.")
                itens_em_falta.append(info_completa) # Salva o nome junto
                navegador.close()
            elif val_estq > 0 and tem_info > 1:
                print(f"ℹ️ {info_completa}: Outro estoque.")
                itens_em_outro_estoque.append(info_completa) # Salva o nome junto
                navegador.close()
            else:
                campo_qtd = wait.until(EC.visibility_of_element_located((By.NAME, "quantidade0")))
                navegador.execute_script("arguments[0].focus();", campo_qtd)
                campo_qtd.clear()
                campo_qtd.send_keys(quantidade)

                btn_ok = wait.until(EC.presence_of_element_located((By.XPATH, "//input[@value='OK']")))
                navegador.execute_script("arguments[0].click();", btn_ok)
                print(f"✅ {info_completa}: Lançado {quantidade}")

            navegador.switch_to.window(janela_principal)
            time.sleep(5)

        except Exception as e:
            print(f"❌ Erro no item {codigo}: {e}")
            if len(navegador.window_handles) > 1: navegador.close()
            navegador.switch_to.window(janela_principal)

# Fim
print(f"\n🏁 Processo Concluído!")

# --- FINALIZAÇÃO E GERAÇÃO DE LOG ---
nome_arquivo_log = f"log_pedidos_{time.strftime('%Y%m%d_%H%M%S')}.txt"

with open(nome_arquivo_log, "w", encoding="utf-8") as f:
    f.write("=== RELATÓRIO DE LANÇAMENTO AUTOMÁTICO ===\n")
    f.write(f"Data/Hora: {time.ctime()}\n\n")

    # NOVA SEÇÃO: Pedidos gerados
    f.write("📋 PEDIDOS GERADOS NESTA SESSÃO:\n")
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

    f.write("\nℹ️ ITENS EM OUTRO ESTOQUE (NÃO LANÇADOS):\n")
    if itens_em_outro_estoque:
        for info in itens_em_outro_estoque:
            f.write(f"- {info}\n")
    else:
        f.write("- Nenhum item em outro estoque.\n")

print(f"📄 Relatório completo gerado: {nome_arquivo_log}")
input("\nProcesso finalizado. Pressione Enter para fechar...")

navegador.quit()