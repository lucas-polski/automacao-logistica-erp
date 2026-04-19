"""
Módulo responsável por carregar as configurações do arquivo config.ini.

Esse módulo centraliza a leitura de configuração para que o resto do
projeto não precise saber que existe um config.ini nem como ele é lido.
Se amanhã quisermos trocar para variáveis de ambiente ou um arquivo YAML,
só esse módulo precisa mudar.
"""

import configparser
from dataclasses import dataclass
from pathlib import Path


@dataclass
class Config:
    """Representa as configurações do projeto, carregadas do config.ini."""
    url_login: str
    url_portal: str
    nome_cliente: str
    planilha_id: str


def carregar_config(caminho: str = "config.ini") -> Config:
    """
    Lê o arquivo config.ini e retorna um objeto Config.

    Args:
        caminho: caminho para o arquivo .ini. Por padrão busca 'config.ini'
                 na pasta atual onde o script é executado.

    Returns:
        Config: objeto com todos os valores de configuração tipados.

    Raises:
        FileNotFoundError: se o arquivo config.ini não existir.
        KeyError: se alguma chave obrigatória estiver faltando no .ini.
    """
    if not Path(caminho).exists():
        raise FileNotFoundError(
            f"Arquivo de configuração '{caminho}' não encontrado. "
            f"Copie o 'config.ini.exemplo' para 'config.ini' e preencha os valores."
        )

    parser = configparser.ConfigParser()
    parser.read(caminho, encoding="utf-8")

    return Config(
        url_login=parser["GERAL"]["url_login"],
        url_portal=parser["GERAL"]["url_portal"],
        nome_cliente=parser["GERAL"]["nome_cliente"],
        planilha_id=parser["GERAL"]["planilha_id"],
    )