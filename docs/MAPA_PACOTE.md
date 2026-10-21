# Mapa do pacote

## Aplicação e entradas

| Arquivo | Função |
| --- | --- |
| [Comunicados_Logisticos.xlsm](../app/Comunicados_Logisticos.xlsm) | Aplicação principal; planilhas `Export` e `Motivos` |
| [BD_McD.xlsx](../app/BD_McD.xlsx) | Seis lojas em `BD_Restaurantes` e doze registros em `BD_Contatos` |
| [CD e NOMENCLATURAS.xlsm](<../app/CD e NOMENCLATURAS.xlsm>) | Três combinações de CD/nomenclatura e respectivos contatos fictícios; duas planilhas, sem VBA |
| [LOJAS COM MOROSIDADE NA DESCARGA.xlsx](<../app/arquivos/LOJAS COM MOROSIDADE NA DESCARGA.xlsx>) | Quatro registros para demonstrar morosidade e exclusões |
| [ATRASO FUTURO.xlsx](<../app/arquivos/ATRASO FUTURO.xlsx>) | Cinco registros para demonstrar previsões de atraso e exclusões |
| [TABELA GERAL.xlsx](<../app/arquivos/TABELA GERAL.xlsx>) | Sete registros para demonstrar mobile/RCS, agrupamento por rota e exclusões |

A aplicação principal foi apresentada com o nome `Comunicados_Logisticos.xlsm`. Seu nome anterior continha a versão `v2_4_5`. A renomeação da cópia não muda a lógica: os caminhos são calculados a partir da pasta do arquivo aberto, e o botão continua associado a `BtnEnviarTudo`.

O nome `BD_McD.xlsx` e a extensão XLSM do auxiliar foram conservados para preservar as referências já validadas. Ambos são abertos para leitura no fluxo de e-mails. A pasta `arquivos/` é o ponto inicial do seletor.

## Código e documentação

- [README.md](../README.md): apresentação em inglês britânico.
- [README.pt-BR.md](../README.pt-BR.md): apresentação em português brasileiro.
- [src/vba/](../src/vba/LEIA-ME.md): seis arquivos de código e índice de origem.
- [examples/expected-output/](../examples/expected-output/README.md): seis HTMLs de referência e o JSON do conteúdo.
- [docs/previews/](previews/): três imagens de referências HTML; não são capturas do Outlook.
- [VALIDACAO.md](VALIDACAO.md) e [VALIDACAO.json](VALIDACAO.json): escopo dos testes.
- [REVISAO_CONTEUDO_PUBLICO.md](REVISAO_CONTEUDO_PUBLICO.md) e [CONTEUDO_PUBLICO.json](CONTEUDO_PUBLICO.json): revisão do que integra o pacote.
- [MANIFESTO.json](MANIFESTO.json): caminhos relativos e hashes dos arquivos publicados no pacote, exceto do próprio manifesto.

## Fluxo principal

`BtnEnviarTudo` chama `CarregarArquivo` e `EnviarEmailsOutlook`. A importação reconhece o relatório pelo nome, aplica suas regras e prepara a planilha `Export`. A montagem consulta os auxiliares e abre os comunicados no Outlook por `.Display`.

O usuário confere destinatários, cópias e conteúdo e decide pelo envio manual. **O código não contém chamada a `.Send`.** Os exemplos de relatório não são atualizados com uma marcação de envio por esse fluxo, e o log persistente original está inativo.

As rotinas adicionais para início de descarga e retorno ao CD não foram incluídas na demonstração validada. Não há uma pasta `programacoes/` preenchida no pacote.
