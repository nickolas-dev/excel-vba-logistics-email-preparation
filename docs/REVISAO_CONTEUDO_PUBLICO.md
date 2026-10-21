# Revisão do conteúdo público

O pacote contém apenas a aplicação preparada para o portfólio, seus auxiliares fictícios, fontes extraídos, exemplos estáticos e documentação. Cópias anteriores, ferramentas locais de edição, sessões de teste e registros privados ficaram fora dele.

## Dados dos arquivos

- As seis lojas dos exemplos usam os códigos Z00–Z05.
- Os cadastros de nomes, e-mails, telefones, transportadoras, placas e links de rastreamento foram substituídos por dados fictícios.
- E-mails de exemplo usam `example.com`. Telefones são preenchimentos demonstrativos, e os links de rastreamento apontam para páginas fictícias desse domínio.
- Os XMLs e relacionamentos dos seis arquivos Excel foram verificados quanto a e-mails fora do domínio demonstrativo, marcadores pessoais conhecidos, caminhos antigos de rede e destinos reais de WhatsApp/rastreamento.
- O código VBA foi verificado quanto aos contatos e caminhos antigos. A cópia principal está editável e não exige uma senha para leitura ou alteração do código.

Identificadores técnicos originais, rótulos de relatórios e algumas classificações de motivos permanecem na aplicação. Isso inclui `BD_McD.xlsx`, `MB`/`REAL`, códigos de CD e nomes de rotinas antigas. A presença desses rótulos não significa que a base real de contatos ou dados da operação esteja incluída.

## Código e envio manual

Os seis fontes foram extraídos do XLSM fornecido em `app/`. O índice relaciona cada fonte ao executável e registra seu hash. Não há módulos de teste nem simulações de Outlook no pacote executável.

A conferência dos fontes encontrou **quatro chamadas a `.Display` e nenhuma chamada a `.Send`**. O botão `Enviar` prepara e abre mensagens para revisão; o envio final é manual no Outlook. Esse comportamento foi mantido de forma intencional.

## Exemplos e imagens

As referências HTML e o JSON usam somente o conteúdo fictício da validação isolada. Não incluem perfil, conta, assinatura pessoal, identificadores internos de mensagens ou registros de caixa postal do Outlook.

As três imagens foram renderizadas a partir desses HTMLs com requisições externas bloqueadas. São ilustrações do conteúdo para a documentação, não capturas do Outlook nem evidência visual de uma mensagem enviada.

## Integridade

O pacote de trabalho de origem foi preservado. A cópia pública do principal recebeu o nome `Comunicados_Logisticos.xlsm`; o conteúdo do arquivo foi copiado sem alteração durante essa renomeação. Os auxiliares e relatórios também foram copiados sem alteração.

`CONTEUDO_PUBLICO.json` reúne os resultados estruturados da revisão. `MANIFESTO.json` registra os caminhos relativos e hashes do pacote. O manifesto não inclui seu próprio hash.

Esta revisão descreve as verificações realizadas sobre esta cópia. Os limites funcionais da demonstração estão em [VALIDACAO.md](VALIDACAO.md).
