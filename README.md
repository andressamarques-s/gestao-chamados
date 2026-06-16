# Sistema de Chamados — Suporte SOC

Sistema web de registro e acompanhamento de chamados de suporte de TI, construído com Google Apps Script e integrado ao Google Sheets. 
## Sobre o projeto

O sistema substitui o preenchimento manual da planilha de chamados por um formulário guiado, evitando erros de digitação e perda de informação. 
Toda a equipe de suporte registra, acompanha e conclui chamados pela mesma interface, com os dados gravados automaticamente na planilha já usada pelo time.

## Funcionalidades

- **Registro de chamados** com data e hora de abertura preenchidas automaticamente
- **Campos com autocomplete** (suporte, instituição, demanda, categoria) navegáveis pelo teclado
- **Registro de chamados já concluídos**, sem precisar passar pela etapa de "em andamento"
- **Painel de acompanhamento** com:
  - Lista de chamados em aberto, com tempo de espera
  - Filtro por período (datas inicial e final)
  - Cards com totais (em andamento, concluídos, total do período)
  - Gráfico de chamados por dia/semana/mês
  - Ranking de atendimentos por pessoa
- **Conclusão de chamados** direto pelo painel, atualizando a planilha automaticamente
- **Configurações gerenciáveis pela interface**: instituições, categorias, equipe de suporte e tipos de demanda podem ser adicionados ou removidos sem tocar no código
- Listas armazenadas nas Propriedades do Script, sem sobrecarregar a planilha

## Stack

- **Google Apps Script** (JavaScript no back-end)
- **HTML, CSS e JavaScript** puro na interface (`HtmlService`)
- **Google Sheets** como banco de dados

## Estrutura de arquivos

| Arquivo | Função |
|---|---|
| `Codigo.gs` | Lógica do servidor: leitura/escrita na planilha, cálculo do painel, persistência das listas |
| `Index.html` | Estrutura visual da interface (formulário, painel, configurações, modal) |
| `Interface.html` | Lógica do cliente: autocomplete, gráficos, comunicação com o servidor |

## Como usar

1. Crie um projeto do Apps Script vinculado à sua planilha (Extensões → Apps Script)
2. Crie os arquivos `Codigo.gs`, `Index.html` e `Interface.html` com o conteúdo deste repositório
3. Ajuste a constante `NOME_ABA` em `Codigo.gs` para o nome da aba onde os chamados são registrados
4. Garanta que a planilha tenha as colunas na ordem esperada pelo código (status, suporte inicial, datas/horas, solicitante, instituição, demanda, descrição, resolução, categoria, suporte final, observações)
5. Execute a função `onOpen` uma vez para criar o menu "Suporte de Sistemas"
6. Reabra a planilha — o sistema fica disponível em **Suporte de Sistemas → Abrir sistema de chamados**

