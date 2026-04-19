"""
Módulo responsável por ler a planilha do Google Sheets e transformá-la
em estruturas de dados que o resto do projeto consome.

Esse módulo é o único lugar que sabe sobre pandas e sobre o formato
específico das abas da planilha. Se amanhã os dados vierem de uma API,
apenas esse módulo precisa mudar.
"""

from dataclasses import dataclass
import pandas as pd


@dataclass
class Item:
    """Um item a ser lançado num pedido (representa uma linha da planilha)."""
    codigo: str
    produto: str
    quantidade: int


@dataclass
class Pedido:
    """Um pedido a ser criado (representa uma aba da planilha)."""
    nome_aba: str
    itens: list[Item]


def _baixar_planilha(sheet_id: str) -> dict[str, pd.DataFrame]:
    """
    Baixa a planilha do Google Sheets como Excel e retorna um dicionário
    {nome_aba: DataFrame}.

    Prefixo _ indica que é uma função interna do módulo, não deve ser
    usada diretamente por fora.
    """
    url = f"https://docs.google.com/spreadsheets/d/{sheet_id}/export?format=xlsx"
    return pd.read_excel(url, sheet_name=None)


def _extrair_itens_validos(df: pd.DataFrame) -> list[Item]:
    """
    Pega um DataFrame de uma aba e retorna apenas os itens válidos,
    convertidos para objetos Item. Ignora linhas com código ou quantidade
    faltando.
    """
    df_limpo = df.dropna(subset=["CODIGO", "QTD"])

    itens = []
    for _, linha in df_limpo.iterrows():
        try:
            item = Item(
                codigo=str(linha["CODIGO"]),
                produto=str(linha["PRODUTO"]),
                quantidade=int(linha["QTD"]),
            )
            itens.append(item)
        except (ValueError, KeyError) as e:
            # Ignora linhas com dados inválidos (ex.: quantidade não numérica)
            print(f"  Linha inválida ignorada: {e}")
            continue

    return itens


def _aba_tem_quantidade(df: pd.DataFrame) -> bool:
    """Retorna True se a aba tem pelo menos uma quantidade lançada."""
    if df.empty:
        return False
    soma = pd.to_numeric(df["QTD"], errors="coerce").sum()
    return soma > 0


def ler_pedidos(sheet_id: str) -> list[Pedido]:
    """
    Lê a planilha do Google Sheets e retorna uma lista de Pedido,
    pulando abas vazias ou sem quantidade lançada.

    Args:
        sheet_id: ID da planilha do Google Sheets (obtido do config).

    Returns:
        list[Pedido]: lista ordenada de pedidos prontos para serem processados.
    """
    planilha = _baixar_planilha(sheet_id)

    pedidos = []
    for nome_aba, df in planilha.items():
        if not _aba_tem_quantidade(df):
            print(f"PULANDO ABA: {nome_aba} (Nenhuma quantidade lançada)")
            continue

        itens = _extrair_itens_validos(df)
        if not itens:
            print(f"PULANDO ABA: {nome_aba} (Nenhum item válido)")
            continue

        pedidos.append(Pedido(nome_aba=nome_aba, itens=itens))

    return pedidos