# Comunicados Logísticos — guia rápido

A apresentação principal está no [README em inglês britânico](README.md), com uma [versão em português brasileiro](README.pt-BR.md). Aplicação, comentários e nomes técnicos permanecem em português.

1. Use Excel desktop e Outlook clássico no Windows, com perfil de e-mail configurado.
2. Abra [app/Comunicados_Logisticos.xlsm](app/Comunicados_Logisticos.xlsm), preservando os auxiliares e a pasta `app/arquivos/`.
3. Na planilha `Export`, clique em **Enviar** e escolha um dos três relatórios fictícios. Cada arquivo prepara duas mensagens.
4. Confira as mensagens abertas no Outlook. **A macro usa `.Display`; o usuário decide pelo envio manual. Não há chamada a `.Send`.**
5. Ao terminar uma demonstração, feche as mensagens fictícias sem salvar nem enviar.

Os exemplos podem ser repetidos: o fluxo demonstrado não altera os relatórios de entrada nem mantém um log persistente de envio ativo. Repetir o relatório pode abrir novamente os mesmos comunicados.

Para consultar os resultados sem executar macros, veja os [exemplos estáticos](examples/expected-output/README.md). Para ler o código, comece pelo [guia do VBA](src/vba/LEIA-ME.md). O [mapa do pacote](docs/MAPA_PACOTE.md) relaciona arquivos e responsabilidades.

Este pacote está preparado para revisão e publicação no GitHub. A preparação local, por si só, não cria nem publica um repositório.
