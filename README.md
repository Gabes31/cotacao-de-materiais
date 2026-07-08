# 🛒 Automação de Cotações e Supply Chain (Bidding System)

Uma aplicação web desenvolvida para a área de Controladoria e Supply Chain, focada na automação do processo de cotação de materiais. O sistema gerencia fornecedores, dispara solicitações de cotação em massa (anexando planilhas geradas dinamicamente) e processa os retornos para apoiar a tomada de decisão através de um dashboard gerencial.

## ✨ Principais Funcionalidades

* **Gestão Base (CRUD):** Cadastro e manutenção do banco de dados de Fornecedores e Materiais (Especificações e Quantidades) integrados ao Google Sheets.
* **Disparo Automatizado de E-mails:**
  * O sistema compila os itens correspondentes ao tipo de cada fornecedor selecionado.
  * Cria planilhas temporárias no Google Drive, converte automaticamente para o formato Excel (`.xlsx`) e anexa em um e-mail padronizado enviado ao fornecedor.
  * O lixo eletrônico (arquivos temporários) é deletado automaticamente do Drive.
* **Processamento de Retornos (Upload de CSV):** * O usuário faz upload do arquivo de resposta do fornecedor no portal.
  * O sistema faz o _parsing_ local e injeta os dados reais diretamente na base de dados de retornos.
* **Dashboard de Decisão Interativo:**
  * Indicadores de Menor Preço Ofertado e Melhor Fornecedor.
  * Geração de gráficos evolutivos (Google Charts) acompanhando as rodadas de negociação.
  * Tabela comparativa que calcula automaticamente a % de redução necessária para que os outros fornecedores atinjam o preço do vencedor da rodada.

## 🛠️ Tecnologias Utilizadas

* **Front-end:** HTML5, CSS3, Vanilla JavaScript (ES6+).
* **Back-end:** Google Apps Script (Serviço `doGet`).
* **Integração e APIs Google:** * `SpreadsheetApp` para manipulação nativa de dados.
  * `DriveApp` e `UrlFetchApp` para criação, conversão (Google Sheets para Excel) e exclusão de arquivos.
  * `MailApp` para disparo automatizado de e-mails com anexos.
* **Visualização de Dados:** Google Charts API.

## ⚙️ Nota Técnica

Este sistema foi desenvolvido nativamente no ecossistema Google Workspace. O arquivo `Codigo.gs` atua como o back-end da aplicação (processando regras de negócio e integrações via `SpreadsheetApp`), enquanto o `index.html` consome esses dados via chamadas assíncronas. O código-fonte está disponibilizado neste repositório para demonstração de arquitetura e lógica (com dados de empresa e credenciais sanitizados/removidos), devendo ser executado dentro de um ambiente Google Apps Script.