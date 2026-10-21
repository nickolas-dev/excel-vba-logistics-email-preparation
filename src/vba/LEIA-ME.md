# VBA para leitura no GitHub

Extração em UTF-8 e quebras de linha LF do projeto de `app/Comunicados_Logisticos.xlsm`. Os atributos e o texto dos módulos foram preservados. O XLSM é a aplicação executável; os fontes permitem ler e revisar a implementação no GitHub.

O [índice](INDICE.json) registra o arquivo de origem, o caminho e o SHA-256 de cada fonte. Há **seis arquivos**, com três módulos padrão e três módulos de documento. Os dois auxiliares não contêm projeto VBA.

| Fonte | Responsabilidade |
| --- | --- |
| [Módulo1.bas](<principal/Módulo1.bas>) | Importação dos relatórios, montagem dos comunicados e abertura no Outlook; também conserva rotinas adicionais fora do caso demonstrado |
| [modPortfolio.bas](principal/modPortfolio.bas) | Caminhos relativos, verificação de auxiliares, contatos fictícios compartilhados e busca de cabeçalhos |
| [ModGlobais.bas](principal/ModGlobais.bas) | Caminho do último arquivo carregado e indicador de origem |
| [EstaPastaDeTrabalho.cls](principal/EstaPastaDeTrabalho.cls) | Módulo do arquivo, sem rotina de abertura ativa |
| [Planilha1.cls](principal/Planilha1.cls) | Módulo da planilha `Export`, sem eventos próprios |
| [Planilha2.cls](principal/Planilha2.cls) | Módulo da planilha `Motivos`, sem eventos próprios |

## Por onde começar

1. `BtnEnviarTudo` chama a importação e, quando há registros, a montagem dos e-mails.
2. `CarregarArquivo` reconhece o nome do relatório, encontra os cabeçalhos, aplica os filtros e prepara a área de trabalho de `Export`.
3. `EnviarEmailsOutlook` consulta as bases e monta os destinatários, cópias e o corpo HTML. **Usa `.Display` para abrir as mensagens. Não há chamada a `.Send`.**
4. `CarregarEmailsRCS` consulta o CD e a nomenclatura para selecionar o destinatário do aviso de mobile.
5. `BuscarEmailResponsavelTD` e `EmailSupervisorPortfolio` fornecem os contatos em cópia nos exemplos de morosidade e atraso.
6. `ObterMotivoExibicao` consulta a planilha `Motivos`: a classificação `REAL` mantém o motivo; `MB` e motivos não encontrados produzem a descrição genérica `Problemas operacionais`.

O argumento opcional de `CarregarArquivo` permite informar o caminho por código nos testes. A chamada normal do botão abre o seletor na pasta `arquivos/`.

## Horários na morosidade

Na importação, a coluna auxiliar **N** recebe `DTH ENTR PROG ATRASO − HR ATRASO`. A coluna **O** recebe `CH ENTREG`. A montagem do e-mail lê o programado de N e a chegada real de O, sem descontar o atraso uma segunda vez.

## Rotinas adicionais

`btnMorIniDesc`, `btnMorIniDesc2`, `AjustarAssuntoPosEnvioMorosidade` e `RetornoCDWhatsApp` foram mantidas como parte do código existente. Não compõem os três fluxos validados do pacote. As rotinas de programação dependem de arquivos adicionais em `programacoes/`, e o retorno ao CD depende de outro relatório. Não foram fornecidos exemplos para essas funções.

Esta extração não inclui módulos de teste ou simulações de Outlook. Também não é um mecanismo automático de reconstrução dos arquivos Excel: planilhas, referências, botão e demais recursos continuam no XLSM.
