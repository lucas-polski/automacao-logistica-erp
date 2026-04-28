"""
Módulo responsável por ler pedidos de uma planilha (Google Sheets ou
arquivo local .xlsx) e transformá-los em estruturas de dados que o
resto do projeto consome.

Suporta duas fontes:
- Google Sheets (via ID público da planilha)
- Arquivo local .xlsx

Em ambos os casos, o formato esperado é o mesmo: cada aba é um pedido,
com colunas CODIGO, PRODUTO e QTD.
"""

from dataclasses import dataclass
from pathlib import Path

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


# ============================================================
# FUNÇÕES INTERNAS DE PROCESSAMENTO
# ============================================================

def _baixar_do_google_sheets(sheet_id: str) -> dict[str, pd.DataFrame]:
    """Baixa a planilha do Google Sheets como Excel."""
    url = f"https://docs.google.com/spreadsheets/d/{sheet_id}/export?format=xlsx"
    return pd.read_excel(url, sheet_name=None)


def _ler_arquivo_local(caminho: str) -> dict[str, pd.DataFrame]:
    """Lê uma planilha local .xlsx."""
    if not Path(caminho).exists():
        raise FileNotFoundError(f"Arquivo '{caminho}' não encontrado.")
    return pd.read_excel(caminho, sheet_name=None)


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
            print(f"  Linha inválida ignorada: {e}")
            continue

    return itens


def _aba_tem_quantidade(df: pd.DataFrame) -> bool:
    """Retorna True se a aba tem pelo menos uma quantidade lançada."""
    if df.empty:
        return False
    soma = pd.to_numeric(df["QTD"], errors="coerce").sum()
    return soma > 0


def _processar_planilha(planilha: dict[str, pd.DataFrame]) -> list[Pedido]:
    """
    Lógica comum a Google Sheets e arquivo local: filtra abas vazias,
    extrai itens válidos e monta a lista de Pedidos.
    """
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


# ============================================================
# FUNÇÕES PÚBLICAS
# ============================================================

def ler_pedidos_do_google_sheets(sheet_id: str) -> list[Pedido]:
    """
    Lê a planilha do Google Sheets e retorna uma lista de Pedido,
    pulando abas vazias ou sem quantidade lançada.
    """
    planilha = _baixar_do_google_sheets(sheet_id)
    return _processar_planilha(planilha)


def ler_pedidos_do_arquivo(caminho: str) -> list[Pedido]:
    """
    Lê um arquivo .xlsx local e retorna uma lista de Pedido,
    pulando abas vazias ou sem quantidade lançada.
    """
    planilha = _ler_arquivo_local(caminho)
    return _processar_planilha(planilha)