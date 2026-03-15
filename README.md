# 🤖 Automação de Lançamento de Pedidos - Selenium & Python

Este projeto automatiza o faturamento e lançamento de pedidos em sistemas ERP, integrando dados diretamente do Google Sheets.

## 🚀 Desafios Superados
- **Gestão de Sessão**: Automação completa do login e navegação em sistemas ERP.
- **Manipulação de Iframes**: Navegação profunda em múltiplas camadas de frames para interação com campos de ID e quantidades.
- **Resiliência de Dados**: Validação de estado do DOM para garantir que a grade de itens esteja limpa antes de novos lançamentos.
- **Relatórios**: Geração de logs de execução detalhados com status de estoque e números de pedidos gerados.

## 🛠️ Tecnologias
- **Python** (Pandas, Selenium, ConfigParser)
- **Google Sheets API** (via exportação direta)
- **Git** para versionamento profissional.

## 📦 Como Executar
1. Clone o repositório.
2. Renomeie o `config.ini.example` para `config.ini` e preencha com seus dados.
3. Execute o script principal ou o executável gerado.
