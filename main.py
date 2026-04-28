"""
Ponto de entrada da automação de lançamento de pedidos.

Este arquivo apenas ORQUESTRA o fluxo principal:
1. Carrega a configuração
2. Faz o login no portal
3. Lê a planilha do Google Sheets
4. Processa cada pedido
5. Executa o relançamento dos itens em outro estoque
6. Gera o relatório final

A lógica de cada etapa está nos módulos dentro do pacote `automacao/`.
"""

import time

from automacao.config import carregar_config
from automacao.login import coletar_credenciais, abrir_navegador, fazer_login
from automacao.planilha import ler_pedidos, Pedido
from automacao.relatorio import DadosRelatorio, gerar_relatorio
from automacao.portal import cliente, pedido as pedido_mod, item, janela


# ============================================================
# FUNÇÕES AUXILIARES DO ORQUESTRADOR
# ============================================================

def _preparar_primeira_aba(navegador, wait, cfg, janela_principal, dados: DadosRelatorio, nome_aba: str):
    """
    Fluxo específico da PRIMEIRA aba: seleciona cliente, verifica itens
    residuais e registra o primeiro pedido gerado.
    """
    cliente.selecionar_cliente(
        navegador, wait, cfg.nome_cliente, janela_principal, cfg.url_portal
    )
    janela.entrar_frame_cadastro(navegador, wait)

    # Verifica se o pedido inicial veio com itens residuais
    quantidade_inicial = pedido_mod.contar_itens_no_pedido(navegador)
    if quantidade_inicial > 0:
        print(f"  {nome_aba} detectou {quantidade_inicial} itens residuais. Limpando...")
        pedido_mod.limpar_itens(navegador, wait)
        num_pedido_atual = pedido_mod.obter_numero_atual(navegador)
        pedido_mod.consultar_por_numero(navegador, wait, num_pedido_atual)
    else:
        print(f"  {nome_aba} iniciada com pedido limpo.")

    num_pedido = pedido_mod.obter_numero_atual(navegador)
    dados.pedidos_gerados.append(f"Aba: {nome_aba} -> Pedido: {num_pedido}")


def _preparar_aba_seguinte(navegador, wait, dados: DadosRelatorio, nome_aba: str):
    """Fluxo para abas posteriores: duplica pedido anterior, limpa e consulta."""
    novo_id = pedido_mod.preparar_pedido_duplicado(navegador, wait)
    dados.pedidos_gerados.append(f"Aba: {nome_aba} -> Pedido: {novo_id}")


def _lancar_itens_do_pedido(navegador, wait, janela_principal, p: Pedido, dados: DadosRelatorio):
    """Lança todos os itens de um pedido e classifica os resultados."""
    for it in p.itens:
        info_completa = f"{it.codigo} - {it.produto}"
        status, detalhe = item.lancar_item(
            navegador, wait, janela_principal,
            it.codigo, it.produto, str(it.quantidade),
        )

        if status == "sem_estoque":
            dados.itens_em_falta.append(info_completa)

        elif status == "outro_estoque_popup":
            dados.itens_em_outro_estoque.append(info_completa)
            dados.itens_para_relancar.append({
                "codigo": it.codigo,
                "quantidade": str(it.quantidade),
                "produto": it.produto,
                "origem_aba": p.nome_aba,
                "motivo": "outro_estoque_popup",
            })

        elif status == "lancado_e_excluido":
            nomes_grade = detalhe.get("excluidos_grade", []) if isinstance(detalhe, dict) else []
            marcador = " | ".join(nomes_grade) if nomes_grade else "(sem nome)"
            dados.itens_em_outro_estoque.append(f"{info_completa} | (na grade: {marcador})")
            dados.itens_para_relancar.append({
                "codigo": it.codigo,
                "quantidade": str(it.quantidade),
                "produto": it.produto,
                "origem_aba": p.nome_aba,
                "motivo": "outro_estoque_grade",
            })
        # "lancado" e "erro" não precisam de ação adicional aqui


def _processar_pedido(navegador, wait, cfg, janela_principal, p: Pedido, dados: DadosRelatorio, eh_primeiro: bool):
    """Processa uma aba completa da planilha."""
    print(f"\n   --- INICIANDO ABA: {p.nome_aba} ---")

    if eh_primeiro:
        _preparar_primeira_aba(navegador, wait, cfg, janela_principal, dados, p.nome_aba)
    else:
        _preparar_aba_seguinte(navegador, wait, dados, p.nome_aba)

    _lancar_itens_do_pedido(navegador, wait, janela_principal, p, dados)


def _executar_relancamento(navegador, wait, janela_principal, dados: DadosRelatorio):
    """Cria um pedido separado para os itens que caíram em outro estoque."""
    if not dados.itens_para_relancar:
        return

    print(f"\n   === INICIANDO PEDIDO DE RELANÇAMENTO ({len(dados.itens_para_relancar)} itens) ===")

    try:
        novo_id = pedido_mod.preparar_pedido_duplicado(navegador, wait)
        dados.pedidos_gerados.append(f"Aba: RELANÇAMENTO (outro estoque) -> Pedido: {novo_id}")
        dados.pedido_relancamento_numero = novo_id

        for pendente in dados.itens_para_relancar:
            info = f"{pendente['codigo']} - {pendente['produto']} (origem: {pendente['origem_aba']})"

            status, _ = item.lancar_item(
                navegador, wait, janela_principal,
                pendente["codigo"], pendente["produto"], pendente["quantidade"],
                ignorar_outro_estoque=True,
            )

            if status == "lancado":
                dados.relancamento_lancados.append(info)
            elif status == "lancado_e_excluido":
                dados.relancamento_ainda_outro_estoque.append(info + " [excluído da grade novamente]")
            elif status == "outro_estoque_popup":
                dados.relancamento_ainda_outro_estoque.append(info + " [popup]")
            elif status == "sem_estoque":
                dados.relancamento_sem_estoque.append(info)
            else:
                dados.relancamento_erros.append(info)

    except Exception as e:
        print(f"   Falha ao iniciar pedido de relançamento: {e}")


# ============================================================
# MAIN
# ============================================================

def main():
    """Ponto de entrada do programa."""
    # 1. Carrega configuração e credenciais
    cfg = carregar_config()
    usuario, senha = coletar_credenciais()

    # 2. Abre navegador e faz login
    navegador, wait = abrir_navegador()
    fazer_login(navegador, wait, cfg.url_login, usuario, senha)
    janela_principal = navegador.current_window_handle

    # 3. Lê os pedidos da planilha
    pedidos = ler_pedidos(cfg.planilha_id)

    # 4. Processa cada pedido
    dados = DadosRelatorio()
    for i, p in enumerate(pedidos):
        _processar_pedido(
            navegador, wait, cfg, janela_principal, p, dados,
            eh_primeiro=(i == 0),
        )

    # 5. Pedido especial de relançamento
    _executar_relancamento(navegador, wait, janela_principal, dados)

    # 6. Relatório final
    print(f"\n  Processo Concluído!")
    nome_arquivo = gerar_relatorio(dados)
    print(f" Relatório completo gerado: {nome_arquivo}")

    input("\nProcesso finalizado. Pressione Enter para fechar...")
    navegador.quit()


if __name__ == "__main__":
    main()