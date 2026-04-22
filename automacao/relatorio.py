"""
Módulo responsável pela geração do relatório final de execução.

Ao fim de cada rodada, o script gera um arquivo .txt com:
- Lista de pedidos criados
- Itens em falta (estoque zerado)
- Itens em outro estoque detectados
- Resultado do pedido de relançamento

Este módulo NÃO conhece Selenium nem pandas — é puro I/O de texto.
"""

import time
from dataclasses import dataclass, field


@dataclass
class DadosRelatorio:
    """
    Agrega todas as informações acumuladas durante a execução,
    prontas para serem escritas no arquivo de log.
    """
    pedidos_gerados: list = field(default_factory=list)
    itens_em_falta: list = field(default_factory=list)
    itens_em_outro_estoque: list = field(default_factory=list)
    itens_para_relancar: list = field(default_factory=list)
    pedido_relancamento_numero: str = None
    relancamento_lancados: list = field(default_factory=list)
    relancamento_ainda_outro_estoque: list = field(default_factory=list)
    relancamento_sem_estoque: list = field(default_factory=list)
    relancamento_erros: list = field(default_factory=list)


# ============================================================
# FUNÇÕES AUXILIARES (uma por seção do relatório)
# ============================================================

def _escrever_lista_ou_vazio(f, titulo: str, items: list, mensagem_vazio: str):
    """
    Helper genérico: escreve um título, depois uma lista bullet ou uma
    mensagem dizendo que está vazio.
    """
    f.write(f"{titulo}\n")
    if items:
        for info in items:
            f.write(f"- {info}\n")
    else:
        f.write(f"{mensagem_vazio}\n")


def _escrever_cabecalho(f):
    """Escreve as linhas iniciais do relatório."""
    f.write("=== RELATÓRIO DE LANÇAMENTO AUTOMÁTICO ===\n")
    f.write(f"Data/Hora: {time.ctime()}\n\n")


def _escrever_pedidos_gerados(f, dados: DadosRelatorio):
    f.write("  PEDIDOS GERADOS NESTA SESSÃO:\n")
    if dados.pedidos_gerados:
        for p in dados.pedidos_gerados:
            f.write(f"{p}\n")
    else:
        f.write("- Nenhum pedido foi gerado.\n")
    f.write("\n" + "=" * 40 + "\n\n")


def _escrever_itens_em_falta(f, dados: DadosRelatorio):
    _escrever_lista_ou_vazio(
        f,
        titulo="⚠️ ITENS EM FALTA (ESTOQUE ZERADO):",
        items=dados.itens_em_falta,
        mensagem_vazio="- Nenhum item em falta.",
    )
    f.write("\n")


def _escrever_itens_em_outro_estoque(f, dados: DadosRelatorio):
    _escrever_lista_ou_vazio(
        f,
        titulo="ITENS EM OUTRO ESTOQUE (DETECTADOS):",
        items=dados.itens_em_outro_estoque,
        mensagem_vazio="- Nenhum item em outro estoque.",
    )
    f.write("\n" + "=" * 40 + "\n\n")


def _escrever_secao_relancamento(f, dados: DadosRelatorio):
    """Escreve a seção detalhada do pedido de relançamento."""
    f.write("📦 PEDIDO DE RELANÇAMENTO (OUTRO ESTOQUE):\n")

    if not dados.pedido_relancamento_numero:
        f.write("- Nenhum item precisou ser relançado (nenhum item caiu em outro estoque).\n")
        return

    f.write(f"Pedido nº: {dados.pedido_relancamento_numero}\n")
    f.write(f"Total de itens enviados para relançamento: {len(dados.itens_para_relancar)}\n\n")

    _escrever_lista_ou_vazio(
        f,
        titulo="  LANÇADOS COM SUCESSO NO PEDIDO DE RELANÇAMENTO:",
        items=dados.relancamento_lancados,
        mensagem_vazio="- Nenhum.",
    )
    f.write("\n")

    _escrever_lista_ou_vazio(
        f,
        titulo="  AINDA EM OUTRO ESTOQUE (NÃO LANÇADOS NO RELANÇAMENTO):",
        items=dados.relancamento_ainda_outro_estoque,
        mensagem_vazio="- Nenhum.",
    )
    f.write("\n")

    _escrever_lista_ou_vazio(
        f,
        titulo="  SEM ESTOQUE NO RELANÇAMENTO:",
        items=dados.relancamento_sem_estoque,
        mensagem_vazio="- Nenhum.",
    )

    if dados.relancamento_erros:
        f.write("\n")
        _escrever_lista_ou_vazio(
            f,
            titulo="  ERROS NO RELANÇAMENTO:",
            items=dados.relancamento_erros,
            mensagem_vazio="- Nenhum.",
        )


# ============================================================
# FUNÇÃO PÚBLICA
# ============================================================

def gerar_nome_arquivo() -> str:
    """Gera nome de arquivo com timestamp para evitar colisões."""
    return f"log_pedidos_{time.strftime('%Y%m%d_%H%M%S')}.txt"


def gerar_relatorio(dados: DadosRelatorio) -> str:
    """
    Gera o arquivo de log com o relatório completo da execução.

    Args:
        dados: objeto DadosRelatorio preenchido ao longo da execução.

    Returns:
        str: caminho do arquivo gerado.
    """
    nome_arquivo = gerar_nome_arquivo()

    with open(nome_arquivo, "w", encoding="utf-8") as f:
        _escrever_cabecalho(f)
        _escrever_pedidos_gerados(f, dados)
        _escrever_itens_em_falta(f, dados)
        _escrever_itens_em_outro_estoque(f, dados)
        _escrever_secao_relancamento(f, dados)

    return nome_arquivo