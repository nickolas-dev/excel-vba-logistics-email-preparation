# Validação da demonstração

Registro da preparação local em **21 de outubro de 2026**, com Excel desktop e Outlook clássico no Windows, em ambiente em português. A data dos relatórios de exemplo é **19 de outubro de 2026**.

## Casos e critérios

| Entrada | Registros | Casos selecionados | Casos ignorados |
| --- | ---: | --- | --- |
| Morosidade na descarga | 4 | Z00 e Z01 | Comunicação já marcada e duração de descarga de exatamente 10 minutos |
| Atraso futuro | 5 | Z02 e Z03 | Comunicação já marcada, atraso zerado e motivo `-` |
| Tabela geral | 7 | Rotas AL005M e NP006M | Rota repetida, RCS utilizado, comunicação já marcada, aguardando carregamento e em carregamento |

As seis lojas têm cadastro de e-mails e telefones. Todos os CDs/prefixos das rotas de exemplo encontram correspondência no auxiliar. Há duas entradas de telefone por loja, seis restaurantes e três combinações de CD/nomenclatura.

## Montagem isolada e Outlook

1. Uma simulação do objeto de Outlook exercitou a montagem dos seis comunicados pelo VBA, verificando assuntos, destinatários, cópias, conteúdo e os contatos fictícios.
2. Os seis comunicados foram abertos no **Outlook clássico real**. Foram comparados assuntos, destinatários e cópias com as referências, e conferidos os horários corrigidos.
3. Um representante de cada fluxo foi inspecionado visualmente no Outlook.
4. As janelas de teste foram fechadas sem salvar. **Nenhum e-mail foi enviado.** A aplicação usa `.Display`; a decisão de enviar fica com o usuário.

Os módulos temporários usados nos testes não fazem parte do XLSM nem do código exportado. Os exemplos HTML vieram da montagem isolada, sem assinatura do Outlook. As imagens foram renderizadas a partir desses exemplos, com apresentação própria da documentação.

## Horário programado e chegada real

O caso de morosidade da Z01 utiliza:

| Campo | Valor |
| --- | --- |
| `DTH ENTR PROG ATRASO` | 19/10/2026 10:30 |
| `HR ATRASO` | 00:30 |
| Programado: diferença entre os dois campos | 19/10/2026 10:00 |
| Chegada real: `CH ENTREG` | 19/10/2026 10:30 |

A importação grava o programado na coluna auxiliar N e a chegada real na O. A montagem da mensagem lê N sem descontar o atraso novamente. Um teste separado mudou apenas a chegada para **12:15**; o programado permaneceu em **10:00**.

## Portabilidade e preservação

Antes do empacotamento, os três relatórios também foram importados depois de mover a aplicação para uma pasta com espaços e caracteres acentuados. Os auxiliares foram encontrados por caminhos relativos, e cada relatório produziu dois comunicados.

Depois da correção, o arquivo principal abriu normalmente no Excel, com saída inicial vazia, VBA editável e o botão preservado. A correção do horário modificou apenas o componente binário do VBA; planilhas, estilos e botão ficaram intactos. Os cinco arquivos auxiliares e de entrada não foram alterados por essa correção.

A conferência da cópia pública, incluindo o novo nome do executável e a localização em `app/`, está registrada em `VALIDACAO.json` no campo `publication_package_location_test`.

## Limites

- A validação cobre os três relatórios fornecidos, totalizando seis mensagens por conjunto de exemplos.
- Não houve teste de envio, confirmação de recebimento ou entrega de e-mail. Os endereços são fictícios.
- Não foi medido desempenho em grandes volumes, economia de tempo ou resultado operacional.
- Não foi validado um relatório separado de `ATRASO ATUAL`.
- Atualização do início da descarga, ajuste de assunto pós-envio e retorno ao CD por WhatsApp não foram exercitados.
- Outras versões e plataformas de Excel/Outlook e configurações regionais não foram validadas.
- A repetição pode abrir novamente os mesmos comunicados: o fluxo demonstrado não salva marcações no relatório e não tem log persistente de envio ativo.
