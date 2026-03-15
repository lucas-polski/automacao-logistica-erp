# Automação de Lançamento de Pedidos com Python

Este projeto automatiza o lançamento de pedidos em sistemas ERP, acessando remotamente planilhas do Google Planilhas e extraindo dados para execução.

## Soluções
- **Gestão de Sessão**: Automação completa desde o login até a navegação para lançamento de pedidos dentro do sistema ERP.
- **Manipulação de Iframes**: Navegação em múltiplas camadas de frames para interagir com campos de ID e quantidades.
- **Resiliência de Dados**: Validação de estado do DOM para garantir que a grade de itens esteja limpa antes de novos lançamentos.
- **Relatórios**: Geração de relatórios de cada sessão, com informações sobre status de estoque, números de pedidos foram gerados e rupturas de estoque.

## Tecnologias Utilizadas
- **Python:** Pandas, Selenium, ConfigParser
- **Google Sheets API:** via exportação direta
- **Git:** para versionamento.

## Como Executar
1. Clone o repositório.
2. Renomeie o `config.example.ini` para `config.ini` e preencha com seus dados.
3. Execute o script `main.py` ou o executável gerado.
