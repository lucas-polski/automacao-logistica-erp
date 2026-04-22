"""
Módulo responsável pela autenticação do usuário no portal.

Separa três responsabilidades:
1. Coletar credenciais do usuário (terminal)
2. Criar o navegador Chrome pronto para uso
3. Executar o login no portal

Cada uma é uma função independente para facilitar manutenção e teste.
"""

import getpass

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC


def coletar_credenciais() -> tuple:
    """
    Pede usuário e senha no terminal.

    A senha é lida via getpass para não aparecer na tela enquanto é digitada.

    Returns:
        tuple(str, str): (usuario, senha)
    """
    usuario = input("  Digite o nome do Usuário:")
    senha = getpass.getpass("  Digite sua senha (Sua senha não vai aparecer):")
    return usuario, senha


def abrir_navegador(timeout: int = 20) -> tuple:
    """
    Cria uma instância do Chrome maximizada e um WebDriverWait configurado.

    Args:
        timeout: segundos padrão para o WebDriverWait.

    Returns:
        tuple(WebDriver, WebDriverWait): (navegador, wait)
    """
    navegador = webdriver.Chrome()
    navegador.maximize_window()
    wait = WebDriverWait(navegador, timeout)
    return navegador, wait


def fazer_login(navegador, wait, url_login: str, usuario: str, senha: str):
    """
    Executa o fluxo de login no portal.

    Preenche usuário e senha na tela de login, clica no botão de entrar,
    e aguarda a URL mudar (indicando login bem-sucedido).

    Args:
        navegador: instância do WebDriver.
        wait: WebDriverWait configurado.
        url_login: URL da página de login.
        usuario: nome de usuário.
        senha: senha do usuário.
    """
    navegador.get(url_login)
    navegador.find_element(By.ID, "login").send_keys(usuario)
    navegador.find_element(By.ID, "senha").send_keys(senha)
    navegador.find_element(
        By.XPATH, "//button[contains(@onclick, 'logar()')]"
    ).click()
    wait.until(EC.url_changes(navegador.current_url))