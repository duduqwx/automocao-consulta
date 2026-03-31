# Automação de Consulta de FGTS (Selenium & Openpyxl)

## Sobre o projeto
Este projeto foi desenvolvido com o foco em automatizar o processo de consulta de saldo no sistema V8, subistituindo uma tarefa manual repetitiva por um script eficiênte em Python.

O sistema lê uma base de CPFs a partir de uma planilha Excel, acessa o site web, realiza as consultas de forma automatizada e preenche a planilha original com os status retornados (Ex: "Com saldo", "Sem FGTS", "Não Autorizado", etc.), gerando um relatório final pronto para a equipe de vendas.

> **Aviso de Legado (Archived):** Este repositório serve como um portfólio de demonstração de habilidades e não deve ser aplicado. O Script foi desenvolvido para uma versão específica do site, como as interfaces web recebem atualizações periódicamentes, os XPATHs e seletores devem estar desatualizados, impossibilitando a execução do mesmo.

## Tecnologias e Bibliotecas Utilizadas
* **Python 3**
* **Selenium WebDriver:** Automação de navegação web, preenchimento de formulários e extração de dados da tela.
* **Openpyxl:** Leitura, manipulação e escrita de dados em planilhas Excel (`.xlsx`).
* **Tratamento de Esperas (Waits):** Uso de `WebDriverWait` e `Expected Conditions` para garantir a sincronia entre o carregamento das páginas e as ações do robô.

## Principais Funcionalidades Estruturadas
- Autenticação com credenciais externalizadas (por segurança).
- Interação com elementos dinâmicos da página (menus dropdown, campos de texto, botões).
- Captura de pop-ups (toasts) do sistema para classificar o status do cliente.
- Validação cruzada entre o CPF buscado no Excel e o exibido na interface web para garantir a integridade dos dados.