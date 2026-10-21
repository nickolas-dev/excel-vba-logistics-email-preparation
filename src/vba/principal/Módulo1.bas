Attribute VB_Name = "Módulo1"
Sub EnviarEmailsOutlook()

    Dim ws As Worksheet
    Dim ultimaLinha As Long
    Dim i As Long
    
    Dim OutlookApp As Object
    Dim OutlookMail As Object
    
    Dim wbBD As Workbook
    Dim wsBD As Worksheet
    Dim linhaBD As Long
    Dim ultimaBD As Long
    
    Dim bdAberto As Boolean
    Dim respAberto As Boolean
    
    Dim wbResp As Workbook
    Dim wsResp As Worksheet
    
    Dim assinatura As String
    
    Set ws = ThisWorkbook.ActiveSheet
    ultimaLinha = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    
    bdAberto = False
    respAberto = False

    On Error Resume Next
    Set OutlookApp = GetObject(, "Outlook.Application")
    If OutlookApp Is Nothing Then Set OutlookApp = CreateObject("Outlook.Application")
    On Error GoTo 0
    
    'Dim wbLog As Workbook
    'Dim wsLog As Worksheet
    'Dim ultimaLog As Long
    'Dim logAberto As Boolean
    'logAberto = False
    
'Set wbLog = Workbooks.Open( _


'Set wsLog = wbLog.Sheets(1)
'logAberto = True

'ultimaLog = wsLog.Cells(wsLog.Rows.Count, "A").End(xlUp).Row

    'Dim dictLogExecucao As Object
    'Set dictLogExecucao = CreateObject("Scripting.Dictionary")
    'dictLogExecucao.CompareMode = vbTextCompare

    For i = 2 To ultimaLinha
        
        If Trim(ws.Cells(i, "A").Value) <> "" Then

            Dim assuntoAtual As String
            assuntoAtual = ws.Cells(i, "A").Value
        
            ' ? SE FOR CASO 3, valida antes de TUDO
            If Left(assuntoAtual, 19) = "Previsão de entrega" Then
                
                'Dim linhaLog As Long
                'Dim jaExiste As Boolean
                
                'jaExiste = False
        
                ' verifica LOG
                'For linhaLog = 2 To ultimaLog
                 '   If wsLog.Cells(linhaLog, "A").Value = assuntoAtual Then
                  '      jaExiste = True
                   '     Exit For
                    'End If
                'Next linhaLog
        
                ' verifica execução atual
                'If dictLogExecucao.Exists(assuntoAtual) Then
                 '   jaExiste = True
                'End If
        
                ' ? SE JÁ EXISTE ? PULA ANTES DE CRIAR EMAIL
                'If jaExiste Then GoTo ProximaLinhaCaso3
            End If


            If Left(ws.Cells(i, "A").Value, 17) = "RCS NÃO UTILIZADO" Then
                If Trim(ws.Cells(i, "B").Value) = "" Then GoTo ProximaLinhaCaso3
            End If
            
            Set OutlookMail = OutlookApp.CreateItem(0)
            
            With OutlookMail
                
                .Display
                assinatura = .HTMLBody
                
                .Subject = ws.Cells(i, "A").Value
                
                '======================================
                ' CASO 2 – RCS NÃO UTILIZADO
                '======================================
                If Left(ws.Cells(i, "A").Value, 17) = "RCS NÃO UTILIZADO" Then
                    
                    .To = ws.Cells(i, "B").Value
                    .CC = EMAIL_TORRE
                    
                    .HTMLBody = _
                        "<p>Prezados,</p>" & _
                        "<p>Para conhecimento, identificamos que o RCS não está sendo utilizado no mobile.</p>" & _
                        "<p>Rota: " & ExtrairRotaDoAssunto(ws.Cells(i, "A").Value) & "</p>" & _
                        assinatura
                    
                    .Display
                
                '======================================
                ' CASO 3 – ATRASOS
                '======================================
                ElseIf Left(ws.Cells(i, "A").Value, 19) = "Previsão de entrega" Then
                
                    'assuntoAtual = ws.Cells(i, "A").Value
                    'jaExiste = False
                    
                    ' ? 1. Verifica no LOG (arquivo)
                    'ultimaLog = wsLog.Cells(wsLog.Rows.Count, "A").End(xlUp).Row
                    
                    'If ultimaLog = 1 Then ultimaLog = ult
                    
                    'For linhaLog = 2 To ultimaLog
                        'If wsLog.Cells(linhaLog, "A").Value = assuntoAtual Then
                           ' jaExiste = True
                          '  Exit For
                       ' End If
                    'Next linhaLog
                    
                    ' ? 2. Verifica na execução atual
                    'If dictLogExecucao.Exists(assuntoAtual) Then
                     '   jaExiste = True
                    'End If
                    
                    'If jaExiste Then GoTo ProximaLinhaCaso3
                    
                    Dim loja3 As String, emailsTo3 As String
                    Dim supervisao3 As String
                    Dim emailCCSupervisor3 As String
                    Dim emailResponsavelTD3 As String
                    Dim ccFinal3 As String
                    
                    Dim atraso3 As String
                    Dim prog3 As String, chegada3 As String
                    Dim motivo3 As String, link3 As String
                    Dim htmlBody3 As String
                    
                    loja3 = ws.Cells(i, "L").Value
                    emailsTo3 = ""
                    emailCCSupervisor3 = ""
                    
                    ' abrir base RESPONSÁVEL
                    If respAberto = False Then
                        Set wbResp = Workbooks.Open(ArquivoProjeto("CD e NOMENCLATURAS.xlsm"), UpdateLinks:=0, ReadOnly:=True)
                        Set wsResp = wbResp.Sheets("Planilha2")
                        respAberto = True
                    End If
                    
                    emailResponsavelTD3 = BuscarEmailResponsavelTD(ws.Cells(i, "A").Value, wsResp)
                    
                    ' abrir base BD
                    If bdAberto = False Then
                        Set wbBD = Workbooks.Open(ArquivoProjeto("BD_McD.xlsx"), UpdateLinks:=0, ReadOnly:=True)
                        Set wsBD = wbBD.Sheets("BD_Restaurantes")
                        ultimaBD = wsBD.Cells(wsBD.Rows.Count, "A").End(xlUp).Row
                        bdAberto = True
                    End If
                    
                    ' buscar emails
                    For linhaBD = 2 To ultimaBD
                        If wsBD.Cells(linhaBD, "A").Value = loja3 Then
                            
                            supervisao3 = wsBD.Cells(linhaBD, "B").Value
                            
                            emailCCSupervisor3 = EmailSupervisorPortfolio(supervisao3)
                            
                            If wsBD.Cells(linhaBD, "D").Value <> "" Then emailsTo3 = emailsTo3 & wsBD.Cells(linhaBD, "D").Value & ";"
                            If wsBD.Cells(linhaBD, "F").Value <> "" Then emailsTo3 = emailsTo3 & wsBD.Cells(linhaBD, "F").Value & ";"
                            If wsBD.Cells(linhaBD, "G").Value <> "" Then emailsTo3 = emailsTo3 & wsBD.Cells(linhaBD, "G").Value
                            
                            Exit For
                        End If
                    Next linhaBD
                    
                    .To = emailsTo3
                    
                    ' montar CC
                    ccFinal3 = emailCCSupervisor3
                    
                    ' ? entra o responsável TD (igual caso 1)
                    If emailResponsavelTD3 <> "" Then
                        ccFinal3 = ccFinal3 & emailResponsavelTD3 & ";"
                    End If
                    
                    If origemOpenTech = True Then
                        ccFinal3 = ccFinal3 & EMAIL_TORRE & "; " & EMAIL_JORNADA & "; " & EMAIL_RELACIONAMENTO
                    Else
                        ccFinal3 = ccFinal3 & EMAIL_TORRE & "; " & EMAIL_RELACIONAMENTO
                    End If
                    
                    .CC = ccFinal3
                    
                    ' dados
                    atraso3 = Format(ws.Cells(i, "M").Value, "hh:mm:ss")
                    prog3 = Format(ws.Cells(i, "N").Value - ws.Cells(i, "M").Value, "dd/mm/yyyy hh:mm")
                    chegada3 = Format(ws.Cells(i, "N").Value, "dd/mm/yyyy hh:mm")
                    
                    '============================
                    'MOTIVO ATRASO MOROSIDADE
                    '============================
                    motivo3 = ObterMotivoExibicao(ws.Cells(i, "S").Value)
                    
                    link3 = ws.Cells(i, "R").Value
                    
                    Dim wsContatos As Worksheet
                    Dim ultimaContatos As Long
                    Dim numeros As String
                    Dim numeroAtual As String
                    Dim contadorNumeros As Long
                    Dim linhaContato As Long
                    
                    Set wsContatos = wbBD.Sheets("BD_Contatos")
                    
                    numeros = ""
                    contadorNumeros = 0
                    
                    ' -------- BUSCAR NA BD_CONTATOS (vários números)
                    ultimaContatos = wsContatos.Cells(wsContatos.Rows.Count, "A").End(xlUp).Row
                    
                    For linhaContato = 2 To ultimaContatos
                        If wsContatos.Cells(linhaContato, "A").Value = loja3 Then
                                                   
                            numeroAtual = Trim(wsContatos.Cells(linhaContato, "D").Value)
                            
                            If numeroAtual <> "" Then
                                numeros = numeros & numeroAtual & ", "
                                contadorNumeros = contadorNumeros + 1
                            End If
                            
                        End If
                    Next linhaContato
                    
                    ' -------- BUSCAR NA BD_RESTAURANTES (coluna H)
                    numeroAtual = Trim(wsBD.Cells(linhaBD, "H").Value)
                    
                    If numeroAtual <> "" Then
                        numeros = numeros & numeroAtual & ", "
                        contadorNumeros = contadorNumeros + 1
                    End If
                    
                    ' -------- REMOVER ÚLTIMA VÍRGULA
                    If Right(numeros, 2) = ", " Then
                        numeros = Left(numeros, Len(numeros) - 2)
                    End If
                    
                    Dim textoContato As String

                    If contadorNumeros = 1 Then
                        textoContato = "Realizamos tentativas de contato pelo número " & numeros & ", porém não obtivemos retorno até o momento."
                        
                    ElseIf contadorNumeros > 1 Then
                        textoContato = "Realizamos tentativas de contato pelos números " & numeros & ", porém não obtivemos retorno até o momento."
                        
                    Else
                        textoContato = ""
                    End If
                    
                    ' HTML
                    htmlBody3 = "<p>Prezados, " & Saudacao() & "!</p>"
                    htmlBody3 = htmlBody3 & "<p>Informamos que o veículo que realizará a sua entrega não chegará dentro do horário programado.</p>"
                    htmlBody3 = htmlBody3 & "<p>Seguem mais detalhes e a previsão atualizada na tabela abaixo:</p>"
                    
                    htmlBody3 = htmlBody3 & "<table border='1' style='border-collapse:collapse; text-align:center;'>"
                    htmlBody3 = htmlBody3 & "<tr style='background-color:#E26A23; font-weight:bold;'>"
                    htmlBody3 = htmlBody3 & "<td>LOJA</td><td>HORÁRIO PROGRAMADO</td><td>HORAS DE ATRASO</td><td>PREVISÃO ATUALIZADA</td><td>MOTIVO</td></tr>"
                    
                    htmlBody3 = htmlBody3 & "<tr style='font-weight:bold;'>"
                    htmlBody3 = htmlBody3 & "<td>" & loja3 & "</td>"
                    htmlBody3 = htmlBody3 & "<td>" & prog3 & "</td>"
                    htmlBody3 = htmlBody3 & "<td>" & atraso3 & "</td>"
                    htmlBody3 = htmlBody3 & "<td>" & chegada3 & "</td>"
                    htmlBody3 = htmlBody3 & "<td>" & motivo3 & "</td></tr></table>"
                    
                    ' ? NOVO PARÁGRAFO AQUI
                    If textoContato <> "" Then
                        htmlBody3 = htmlBody3 & "<p>" & textoContato & "</p>"
                    End If
                    
                    htmlBody3 = htmlBody3 & "<p>Para acompanhar o veículo em tempo real, <a href='" & link3 & "'>clique aqui</a></p>"
                    
                    'TEXTO AUXILIAR
                    htmlBody3 = htmlBody3 & TextoSuportePortfolio()
                    
                    'ASSINATURA
                    .HTMLBody = htmlBody3 & assinatura
                    .Display
                    
                    'ultimaLog = wsLog.Cells(wsLog.Rows.Count, "A").End(xlUp).Row + 1
                    'wsLog.Cells(ultimaLog, "A").Value = assuntoAtual
                    'wsLog.Cells(ultimaLog, "B").Value = "ENVIADO"
            
                    ' ? adiciona na execução atual (MUITO IMPORTANTE)
                    'dictLogExecucao.Add assuntoAtual, True
                
                Else
                
                '======================================
                ' CASO 1 – MOROSIDADE
                '======================================
                    
                    'Dim loja As String, emailsTo As String
                    'Dim atraso As String, motivo As String
                    'Dim dthProg As String, chegada As String
                    'Dim tempProg As String, tempAtual As String
                    'Dim link As String, htmlBody As String
                    'Dim supervisao As String
                    
                    'Dim emailCCSupervisor As String
                    'Dim emailResponsavelTD As String
                    'Dim ccFinal As String
                    
                    ' abrir base RESPONSÁVEL
                    If respAberto = False Then
                        Set wbResp = Workbooks.Open(ArquivoProjeto("CD e NOMENCLATURAS.xlsm"), UpdateLinks:=0, ReadOnly:=True)
                        Set wsResp = wbResp.Sheets("Planilha2")
                        respAberto = True
                    End If
                    
                    emailResponsavelTD = BuscarEmailResponsavelTD(ws.Cells(i, "A").Value, wsResp)
                    
                    ' abrir base BD
                    If bdAberto = False Then
                        Set wbBD = Workbooks.Open(ArquivoProjeto("BD_McD.xlsx"), UpdateLinks:=0, ReadOnly:=True)
                        Set wsBD = wbBD.Sheets("BD_Restaurantes")
                        ultimaBD = wsBD.Cells(wsBD.Rows.Count, "A").End(xlUp).Row
                        bdAberto = True
                    End If
                    
                    loja = ws.Cells(i, "L").Value
                    emailsTo = ""
                    emailCCSupervisor = ""
                    
                    For linhaBD = 2 To ultimaBD
                        If wsBD.Cells(linhaBD, "A").Value = loja Then
                            
                            supervisao = wsBD.Cells(linhaBD, "B").Value
                            
                            emailCCSupervisor = EmailSupervisorPortfolio(supervisao)
                            
                            'E-mail Gerente Mercado
                            'If wsBD.Cells(linhaBD, "D").Value <> "" Then emailsTo = emailsTo & wsBD.Cells(linhaBD, "D").Value & ";"
                            
                            'E-mail Consultor
                            If wsBD.Cells(linhaBD, "F").Value <> "" Then emailsTo = emailsTo & wsBD.Cells(linhaBD, "F").Value & ";"
                            
                            'E-mail Restaurantes
                            If wsBD.Cells(linhaBD, "G").Value <> "" Then emailsTo = emailsTo & wsBD.Cells(linhaBD, "G").Value
                            
                            Exit For
                        End If
                    Next linhaBD
                    
                    .To = emailsTo
                    
                    ccFinal = emailCCSupervisor
                    If emailResponsavelTD <> "" Then ccFinal = ccFinal & emailResponsavelTD & ";"
                    
                    If origemOpenTech = True Then
                        ccFinal = ccFinal & EMAIL_TORRE & "; " & EMAIL_JORNADA & "; " & EMAIL_RELACIONAMENTO
                    Else
                        ccFinal = ccFinal & EMAIL_TORRE & "; " & EMAIL_RELACIONAMENTO
                    End If
                    
                    .CC = ccFinal
                    
                    If ws.Cells(i, "T").Value = "ON TIME" Then
                        atraso = "dentro do horário programado"
                        dthProg = Format(ws.Cells(i, "N").Value, "dd/mm/yyyy hh:mm")
                    Else
                        atraso = "com atraso devido "
                        
                        'MOTIVO DO ATRASO
                        motivo = ObterMotivoExibicao(ws.Cells(i, "S").Value)
                        
                        dthProg = Format(ws.Cells(i, "N").Value, "dd/mm/yyyy hh:mm")
                    End If
                    
                    chegada = Format(ws.Cells(i, "O").Value, "dd/mm/yyyy hh:mm")
                    tempProg = Format(ws.Cells(i, "P").Value, "hh:mm:ss")
                    tempAtual = Format(ws.Cells(i, "Q").Value, "hh:mm:ss")
                    caixas = ws.Cells(i, "U").Value
                    horInicioDescarga = Format(ws.Cells(i, "V").Value, "dd/mm/yyyy hh:mm")
                    link = ws.Cells(i, "R").Value
                    
                    'Dim wsContatos As Worksheet
                    'Dim ultimaContatos As Long
                    'Dim numeros As String
                    'Dim numeroAtual As String
                    'Dim contadorNumeros As Long
                    
                    'Set wsContatos = wbBD.Sheets("BD_Contatos")
                    
                    'numeros = ""
                    'contadorNumeros = 0
                    
                    ' -------- BUSCAR NA BD_CONTATOS (vários números)
                    'ultimaContatos = wsContatos.Cells(wsContatos.Rows.Count, "A").End(xlUp).Row
                    
                    'For linhaBD = 2 To ultimaContatos
                        'If wsContatos.Cells(linhaBD, "A").Value = loja Then
                                                   
                           ' numeroAtual = Trim(wsContatos.Cells(linhaBD, "D").Value)
                            
                            'If numeroAtual <> "" Then
                                'numeros = numeros & numeroAtual & ", "
                                'contadorNumeros = contadorNumeros + 1
                            'End If
                            
                        'End If
                    'Next linhaBD
                    
                    ' -------- BUSCAR NA BD_RESTAURANTES (coluna H)
                    'numeroAtual = Trim(wsBD.Cells(linhaBD, "H").Value)
                    
                    'If numeroAtual <> "" Then
                     '   numeros = numeros & numeroAtual & ", "
                       ' contadorNumeros = contadorNumeros + 1
                    'End If
                    
                    ' -------- REMOVER ÚLTIMA VÍRGULA
                    'If Right(numeros, 2) = ", " Then
                    '    numeros = Left(numeros, Len(numeros) - 2)
                    'End If
                    
                    'Dim textoContato As String

                    'If contadorNumeros = 1 Then
                        'textoContato = "Realizamos tentativas de contato pelo número " & numeros & ", porém não obtivemos retorno até o momento."
                        
                   ' ElseIf contadorNumeros > 1 Then
                      '  textoContato = "Realizamos tentativas de contato pelos números " & numeros & ", porém não obtivemos retorno até o momento."
                        
                    'Else
                        'textoContato = ""
                    'End If
                    
                    HTMLBody = "<p>Prezados, " & Saudacao() & "!</p>"
                    
                    If ws.Cells(i, "T").Value = "ON TIME" Then
                    
                        HTMLBody = HTMLBody & "<p>Nossa equipe chegou à loja <b>" & loja & "</b> " & atraso & ", porém a descarga ainda não foi finalizada.</p>"
                    
                    Else
                    
                        HTMLBody = HTMLBody & "<p>Nossa equipe chegou à loja <b>" & loja & "</b> " & atraso & motivo & ", porém a descarga ainda não foi finalizada.</p>"
                        
                    End If
                    
                    HTMLBody = HTMLBody & "<p>Por gentileza, poderiam verificar?</p>"
                    
                    HTMLBody = HTMLBody & "<p>" & _
                        "• Loja: " & loja & "<br>" & _
                        "• Data e hora programados: " & dthProg & "<br>" & _
                        "• Data e hora da chegada do veículo: " & chegada & "<br>" & _
                        "• Data e hora do início da descarga: " & horInicioDescarga & "<br>" & _
                        "• Total de volumes: " & caixas & " caixas" & "<br>" & _
                        "• Tempo de descarga programado: " & tempProg & "<br>" & _
                        "• Tempo de descarga atual: <b>" & tempAtual & " – MOROSIDADE</b></p>"
                        
                    ' ? NOVO PARÁGRAFO AQUI
                    'If textoContato <> "" Then
                        'htmlBody = htmlBody & "<p>" & textoContato & "</p>"
                    'End If
                    
                    HTMLBody = HTMLBody & "<p>Para acompanhar o veículo em tempo real, <a href='" & link & "'>clique aqui</a></p>"
                    HTMLBody = HTMLBody & "<p>Desde já, agradecemos a atenção e a colaboração.</p>"
                    
                    'TEXTO AUXILIAR
                    htmlBody3 = HTMLBody & TextoSuportePortfolio()
                    
                    'ASSINATURA
                    .HTMLBody = htmlBody3 & assinatura
                    .Display
                
                End If
            
            End With
            
            Set OutlookMail = Nothing
        End If
        
ProximaLinhaCaso3:
    Next i

    If respAberto Then wbResp.Close False
    If bdAberto Then wbBD.Close False
    
    ThisWorkbook.Activate
    
    Application.Wait Now + TimeValue("00:00:01")
    
    AppActivate Application.Caption
    
    MsgBox "E-mails carregados para envio!", vbInformation

End Sub

Sub btnMorIniDesc()

    Dim caminho As String
    Dim arquivo As String
    Dim wb As Workbook
    Dim ws As Worksheet
    
    Dim ultimaLinha As Long
    Dim i As Long
    Dim dataAtual As Date
    Dim valorN As Variant
    Dim confirmacao As Integer
    
    confirmacao = MsgBox("Confirma que deseja atualizar TODOS os casos de morosidade no início da descarga?" & vbNewLine & vbNewLine & "Obs. Ao continuar, TODOS os casos anteriores ao horário atual serão afetados.", vbYesNo + vbQuestion, "Confirmação")
    
    If confirmacao = 7 Then
    
    Exit Sub
    
    Else
    
    End If
    
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    
    dataAtual = Now
    caminho = PastaProgramacoesProjeto()
    If Len(caminho) = 0 Then GoTo FinalizarMorInicio
    
    arquivo = Dir(caminho & "*.xlsx")
    
    Do While arquivo <> ""
        
        ' Somente arquivos que começam com Tbl_DistrMCD
        If Left(arquivo, 12) = "Tbl_DistrMCD" Then
            
            Set wb = Workbooks.Open(caminho & arquivo)
            Set ws = wb.Sheets("Sheet1") ' ajuste aqui se não for a primeira aba
            
            ultimaLinha = ws.Cells(ws.Rows.Count, "N").End(xlUp).Row
            
            For i = 2 To ultimaLinha
            
                ' Se AD já estiver preenchida, pula a linha
                If ws.Cells(i, "AD").Value <> "" Then GoTo ProximaLinha
            
                valorN = ws.Cells(i, "N").Value
            
                If IsDate(valorN) Then
                    If CDate(valorN) <= dataAtual Then
                        ws.Cells(i, "AD").Value = CDate(valorN)
                        ws.Cells(i, "AD").NumberFormat = "dd/mm/yyyy hh:mm"
                    End If
                End If
            
ProximaLinha:
            Next i
            
            wb.Close SaveChanges:=True
            Set wb = Nothing
            
        End If
        
        arquivo = Dir
    Loop
    
FinalizarMorInicio:
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    
    MsgBox "Atualização de Morosidade Início Descarga concluída.", vbInformation

End Sub

Sub CarregarArquivo(Optional ByVal arquivoSelecionado As String = "")
    Dim fd As FileDialog
    Dim caminhoArquivo As String
    Dim nomeArquivo As String
    
    Dim wbOrigem As Workbook
    Dim wsOrigem As Worksheet
    Dim wsDestino As Worksheet
    
    Dim ultimaLinha As Long
    Dim linhaDestino As Long
    Dim i As Long
    
    Dim rota As String
    'Dim loja As String
    Dim embarque As String
    Dim cdFAT As String
    Dim dtFAT As String
    
    Dim dictRotas As Object
    Dim emailDestino As String
    emailDestino = EMAIL_TORRE

    Set wsDestino = ThisWorkbook.ActiveSheet

    '=========================================' LIMPAR A:B(mantendo cabeçalho)'=========================================
    ultimaLinha = wsDestino.Cells(wsDestino.Rows.Count, "A").End(xlUp).Row
    If ultimaLinha >= 2 Then
        wsDestino.Range("A2:B" & ultimaLinha).ClearContents
    End If
    
    ' ? limpar também L:V
    ultimaLinha = wsDestino.Cells(wsDestino.Rows.Count, "L").End(xlUp).Row
    If ultimaLinha >= 2 Then
        wsDestino.Range("L2:V" & ultimaLinha).ClearContents
    End If

    linhaDestino = 2

    '=========================================' Selecionar arquivo do projeto'=========================================
    If Len(arquivoSelecionado) > 0 Then
        caminhoArquivo = arquivoSelecionado
    Else
    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    With fd
        .Title = "Selecione o arquivo"
        .InitialFileName = PortfolioPath("arquivos\")
        .AllowMultiSelect = False
        .Filters.Clear
        .Filters.Add "Arquivos Excel", "*.xlsx"

        If .Show <> -1 Then Exit Sub
        caminhoArquivo = .SelectedItems(1)
    End With

    End If
    CaminhoUltimoArquivoCarregado = caminhoArquivo

    nomeArquivo = UCase(Dir(caminhoArquivo))
    Set wbOrigem = Workbooks.Open(caminhoArquivo, UpdateLinks:=0, ReadOnly:=True)
    Set wsOrigem = wbOrigem.Sheets(1)
    '=========================================' CASO 1 – MOROSIDADE'=========================================
    If InStr(nomeArquivo, "MOROSIDADE") > 0 Then

        ultimaLinha = wsOrigem.Cells(wsOrigem.Rows.Count, "G").End(xlUp).Row
        
        'Dim colEmailEnv As Long
        'Dim colTempDescAtu As Long
        
        colEmailEnv = ColunaCabecalho(wsOrigem, "EMAIL ENV")
        colTempDescAtu = ColunaCabecalho(wsOrigem, "TEMP DESC ATU")

        For i = 2 To ultimaLinha
        
            If UCase(Trim(wsOrigem.Cells(i, colEmailEnv).Value)) <> "SIM" _
               And wsOrigem.Cells(i, colTempDescAtu).Value <> TimeSerial(0, 10, 0) Then
        
                'Dim colRota As Long
                'Dim colLoja As Long
                'Dim colEmbarque As Long
                'Dim colOrigem As Long
                
                colRota = ColunaCabecalho(wsOrigem, "ROTA")
                colLoja = ColunaCabecalho(wsOrigem, "LOJA")
                colEmbarque = ColunaCabecalho(wsOrigem, "EMBARQUE")
                colOrigem = ColunaCabecalho(wsOrigem, "ORIGEM")
                colInicioDesc = ColunaCabecalho(wsOrigem, "INI DESC RCS")
                colCaixas = ColunaCabecalho(wsOrigem, "CX")
                
                    rota = wsOrigem.Cells(i, colRota).Value
                    loja = wsOrigem.Cells(i, colLoja).Value
                    embarque = wsOrigem.Cells(i, colEmbarque).Value
                    origem = wsOrigem.Cells(i, colOrigem).Value

                
                    If origem = "CDFT" Then
                    
                        origemOpenTech = True
                    
                    Else
                    
                        origemOpenTech = False
                    
                    End If
        
                If Len(embarque) >= 12 Then
                    cdFAT = Left(embarque, 2)
                    dtFAT = Right(embarque, 10)
                    wsDestino.Cells(linhaDestino, "A").Value = "MOROSIDADE - " & cdFAT & _
                        " - " & rota & _
                        " - " & loja & _
                        " - " & dtFAT
        
                    wsDestino.Cells(linhaDestino, "B").Value = emailDestino
                    
                    'Dim colHrAtraso As Long
                    'Dim colDthEntProgAtraso As Long
                    'Dim colChEntreg As Long
                    'Dim colTempDescPro As Long
                    'Dim colTempDescAtu As Long
                    'Dim colLocalizacaoInfolog As Long
                    'Dim colOcLoja As Long
                    'Dim colStEntreg As Long
                    
                    colHrAtraso = ColunaCabecalho(wsOrigem, "HR ATRASO")
                    colDthEntProgAtraso = ColunaCabecalho(wsOrigem, "DTH ENTR PROG ATRASO")
                    colChEntreg = ColunaCabecalho(wsOrigem, "CH ENTREG")
                    colTempDescPro = ColunaCabecalho(wsOrigem, "TEMP DESC PRO")
                    colTempDescAtu = ColunaCabecalho(wsOrigem, "TEMP DESC ATU")
                    colLocalizacaoInfolog = ColunaCabecalho(wsOrigem, "LOCALIZAÇÃO INFOLOG")
                    colOcLoja = ColunaCabecalho(wsOrigem, "OC LOJA")
                    colStEntreg = ColunaCabecalho(wsOrigem, "ST ENTREG")

                    ' ? dados auxiliares(L em diante)
                    'Loja
                    wsDestino.Cells(linhaDestino, "L").Value = loja
                    
                    'M = valor da coluna HR ATRASO
                    ' -> usado para saber se houve atraso (se estiver vazio = dentro do horário)
                    wsDestino.Cells(linhaDestino, "M").Value = wsOrigem.Cells(i, colHrAtraso).Value
                    
                    'DTH ENT PROG + ATRASO - HR ATRASO
                    dthENTPROG = wsOrigem.Cells(i, colDthEntProgAtraso).Value - _
                                 wsOrigem.Cells(i, colHrAtraso).Value
                    
                    'DTH ENT PROG
                    wsDestino.Cells(linhaDestino, "N").Value = dthENTPROG
                    
                    'O = CH ENTREG (Data e hora real de chegada do veículo)
                    wsDestino.Cells(linhaDestino, "O").Value = wsOrigem.Cells(i, colChEntreg).Value
                    
                    'P = TEMP DESC PRO (Tempo de descarga programado)
                    wsDestino.Cells(linhaDestino, "P").Value = wsOrigem.Cells(i, colTempDescPro).Value
                    
                    'Q = TEMP DESC ATU (Tempo de descarga atual)
                    wsDestino.Cells(linhaDestino, "Q").Value = wsOrigem.Cells(i, colTempDescAtu).Value
                    
                    'R = LOCALIZAÇÃO INFOLOG (link de rastreamento)
                    wsDestino.Cells(linhaDestino, "R").Value = wsOrigem.Cells(i, colLocalizacaoInfolog).Value
                    
                    'S = OC LOJA (motivo do atraso)
                    wsDestino.Cells(linhaDestino, "S").Value = wsOrigem.Cells(i, colOcLoja).Value
                    
                    'T = ST ENTREG (Status Entrega)
                    wsDestino.Cells(linhaDestino, "T").Value = wsOrigem.Cells(i, colStEntreg).Value
                    
                    'U = CX (Caixas)
                    wsDestino.Cells(linhaDestino, "U").Value = wsOrigem.Cells(i, colCaixas).Value
                    
                    'V = INI DESC RCS (Inicio da Descarga)
                    wsDestino.Cells(linhaDestino, "V").Value = wsOrigem.Cells(i, colInicioDesc).Value
                    
                    linhaDestino = linhaDestino + 1
                End If
        
            End If
        
        Next i

    '=========================================' CASO 2 – TABELA GERAL(SEM ROTAS DUPLICADAS)'=========================================
    
    ElseIf InStr(nomeArquivo, "TABELA GERAL") > 0 Then

        Dim dictEmails As Object
        Dim chaveEmail As String
        Dim nomenclatura As String

        Set dictRotas = CreateObject("Scripting.Dictionary")
        dictRotas.CompareMode = vbTextCompare

        Set dictEmails = CarregarEmailsRCS()
        ultimaLinha = wsOrigem.Cells(wsOrigem.Rows.Count, "A").End(xlUp).Row

        
        'Dim colDataFAT As Long
        'Dim colCdFAT As Long
        'Dim colRota As Long
        'Dim colEtapa As Long
        'Dim colRcsUtil As Long
        'Dim colEmailEnvRcs As Long
        'Dim colEmEnv As Long
        
        colDataFAT = ColunaCabecalho(wsOrigem, "DATA FAT")
        colCdFAT = ColunaCabecalho(wsOrigem, "CD FAT")
        colRota = ColunaCabecalho(wsOrigem, "ROTA")
        colEtapa = ColunaCabecalho(wsOrigem, "ETAPA")
        colRcsUtil = ColunaCabecalho(wsOrigem, "RCS UTIL")
        colEmailEnvRcs = ColunaCabecalho(wsOrigem, "EMAIL ENV RCS")
        colEmEnv = ColunaCabecalho(wsOrigem, "EM ENV LJ ATRAS")
     
        For i = 2 To ultimaLinha

            If UCase(Trim(wsOrigem.Cells(i, colEmailEnvRcs).Value)) = "NÃO" _
               And UCase(Trim(wsOrigem.Cells(i, colRcsUtil).Value)) <> "UTILIZADO" _
               And UCase(Trim(wsOrigem.Cells(i, colEtapa).Value)) <> "AG CARREGAMENTO" _
               And UCase(Trim(wsOrigem.Cells(i, colEtapa).Value)) <> "EM CARREGAMENTO" Then
        
                rota = wsOrigem.Cells(i, colRota).Value
        
                If Not dictRotas.Exists(rota) Then
                    dictRotas.Add rota, True
        
                    cdFAT = wsOrigem.Cells(i, colCdFAT).Value
                    nomenclatura = Left(rota, 2)
                    chaveEmail = cdFAT & nomenclatura
        
                    If dictEmails.Exists(chaveEmail) Then
        
                        If IsDate(wsOrigem.Cells(i, colDataFAT).Value) Then
                            dtFAT = Format(wsOrigem.Cells(i, colDataFAT).Value, "dd/mm/yyyy")
                        End If
        
                        wsDestino.Cells(linhaDestino, "A").Value = "RCS NÃO UTILIZADO - " & cdFAT & _
                            " - " & rota & _
                            " - " & dtFAT
        
                        wsDestino.Cells(linhaDestino, "B").Value = dictEmails(chaveEmail)
                        linhaDestino = linhaDestino + 1
        
                    End If
                End If
            End If
        
        Next i
    
'=========================================' CASO 3 – ATRASOS'=========================================

        ElseIf InStr(nomeArquivo, "ATRASO FUTURO") > 0 _
        Or InStr(nomeArquivo, "ATRASO ATUAL") > 0 Then
    
        'Dim colLoja As Long, colProg As Long, colAtraso As Long, colMotivo As Long, colRotograma As Long
        'Dim j As Long
    
        ' mapear cabeçalhos dinamicamente
        For j = 1 To wsOrigem.Cells(1, Columns.Count).End(xlToLeft).Column
        
            Select Case UCase(Trim(wsOrigem.Cells(1, j).Value))
                Case "LOJA": colLoja = j
                Case "DTH ENTREG PROG ATRASO", "DTH ENTR PROG ATRASO"
    colProg = j
                Case "HR ATRASO": colAtraso = j
                Case "OC LOJA": colMotivo = j
                Case "ROTOGRAMA": colRotograma = j
                Case "EMBARQUE": colEmbarque = j
                Case "EM ENV LJ ATRAS": colEmEnv = j
            End Select
    
        Next j
    
        ultimaLinha = wsOrigem.Cells(wsOrigem.Rows.Count, colLoja).End(xlUp).Row
    
        For i = 2 To ultimaLinha
        
            If wsOrigem.Cells(i, colEmEnv).Value = "SIM" Then GoTo Desconsiderar
    
            If wsOrigem.Cells(i, colAtraso).Value <> "" _
               And wsOrigem.Cells(i, colAtraso).Value <> 0 _
               And wsOrigem.Cells(i, colMotivo).Value <> "" _
               And wsOrigem.Cells(i, colMotivo).Value <> "-" Then
    
                loja = wsOrigem.Cells(i, colLoja).Value
                If Len(loja) <> 3 And Len(loja) <> 7 Then GoTo Desconsiderar
                rota = wsOrigem.Cells(i, colLoja).Offset(0, -2).Value ' ajuste se necessário
                'Dim embarque As String
                embarque = wsOrigem.Cells(i, colEmbarque).Value
                
                If Len(embarque) >= 2 Then
                    cdFAT = Left(embarque, 2)
                End If
                
                If Len(embarque) >= 10 Then
                    dtFAT = Right(embarque, 10)
                End If

                wsDestino.Cells(linhaDestino, "A").Value = _
                    "Previsão de entrega - " & cdFAT & " - " & rota & " - " & loja & " - " & dtFAT
    
                ' dados auxiliares (reaproveitando estrutura L:T)
                wsDestino.Cells(linhaDestino, "L").Value = loja
                wsDestino.Cells(linhaDestino, "M").Value = wsOrigem.Cells(i, colAtraso).Value
                wsDestino.Cells(linhaDestino, "N").Value = wsOrigem.Cells(i, colProg).Value
                wsDestino.Cells(linhaDestino, "R").Value = wsOrigem.Cells(i, colRotograma).Value
                wsDestino.Cells(linhaDestino, "S").Value = wsOrigem.Cells(i, colMotivo).Value
                'wsDestino.Cells(linhaDestino, "T").Value = wsOrigem.Cells(i, colEmEnv).Value
    
                linhaDestino = linhaDestino + 1
    
            End If
Desconsiderar:
        Next i
    
    Else
        MsgBox "O arquivo selecionado não atende aos critérios esperados.", vbExclamation
    End If

    wbOrigem.Close SaveChanges:=False

End Sub

Sub AjustarAssuntoPosEnvioMorosidade()

    Dim wbOrigem As Workbook
    Dim wsOrigem As Worksheet
    Dim wsDestino As Worksheet

    Dim caminhoArquivo As String
    Dim nomeArquivo As String

    Dim ultimaLinhaOrigem As Long
    Dim ultimaLinhaDestino As Long
    Dim i As Long, j As Long

    Dim rota As String
    'Dim loja As String
    Dim embarque As String
    Dim cdFAT As String
    Dim dtFAT As String
    Dim statusEntrega As String

    Dim assuntoBase As String
    Dim complemento As String

    ' ===============================
    ' Validar arquivo carregado
    ' ===============================
    caminhoArquivo = CaminhoUltimoArquivoCarregado
    If caminhoArquivo = "" Then
        MsgBox "Nenhum arquivo foi carregado anteriormente.", vbExclamation
        Exit Sub
    End If

    nomeArquivo = UCase(Dir(caminhoArquivo))
    If InStr(nomeArquivo, "MOROSIDADE") = 0 Then
        Exit Sub ' só roda para MOROSIDADE
    End If

    ' ===============================
    ' Abrir arquivo origem
    ' ===============================
    Set wsDestino = ThisWorkbook.ActiveSheet
    Set wbOrigem = Workbooks.Open(caminhoArquivo, UpdateLinks:=0, ReadOnly:=True)
    Set wsOrigem = wbOrigem.Sheets(1) ' ajuste se necessário

    ultimaLinhaOrigem = wsOrigem.Cells(wsOrigem.Rows.Count, "G").End(xlUp).Row
    ultimaLinhaDestino = wsDestino.Cells(wsDestino.Rows.Count, "A").End(xlUp).Row

    ' ===============================
    ' Ajustar assunto pós-envio
    ' ===============================
    For i = 2 To ultimaLinhaOrigem

        ' Apenas linhas ainda NÃO marcadas como enviadas
        If UCase(Trim(wsOrigem.Cells(i, "N").Value)) = "NÃO" Then

            rota = wsOrigem.Cells(i, "G").Value
            loja = wsOrigem.Cells(i, "I").Value
            embarque = wsOrigem.Cells(i, "AC").Value
            statusEntrega = UCase(Trim(wsOrigem.Cells(i, "J").Value))

            If Len(embarque) >= 12 Then
                cdFAT = Left(embarque, 2)
                dtFAT = Right(embarque, 10)

                assuntoBase = "MOROSIDADE - " & cdFAT & _
                              " - " & rota & _
                              " - " & loja & _
                              " - " & dtFAT

                Select Case statusEntrega
                    Case "ON TIME"
                        complemento = " chegou dentro do horário"
                    Case "LATE"
                        complemento = " chegou com atraso devido morosidade na entrega anterior"
                    Case Else
                        complemento = ""
                End Select

                ' Procurar o assunto base no arquivo da macro
                For j = 2 To ultimaLinhaDestino
                    If wsDestino.Cells(j, "A").Value = assuntoBase Then
                        wsDestino.Cells(j, "A").Value = assuntoBase & complemento
                        Exit For
                    End If
                Next j

            End If

        End If

    Next i

    wbOrigem.Close SaveChanges:=False

    'MsgBox "Assuntos de MOROSIDADE ajustados pós-envio com sucesso.", vbInformation

End Sub
Sub BtnEnviarTudo()

    Call CarregarArquivo
    
    ' ? verifica se carregou algo (coluna A)
    If ThisWorkbook.ActiveSheet.Cells(2, "A").Value = "" Then Exit Sub
    
    Call EnviarEmailsOutlook
    'Call AjustarAssuntoPosEnvioMorosidade
End Sub

Sub btnMorIniDesc2()

    Dim fd As FileDialog
    Dim caminhoInicio As String
    Dim nomeArquivoInicio As String

    Dim wbInicio As Workbook, wsInicio As Worksheet
    Dim wbProg As Workbook, wsProg As Worksheet

    Dim dictGrupos As Object          ' grupo -> dictionary de embarques
    Dim dictEmbarques As Object       ' embarques por grupo
    Dim grupo As String

    Dim ultimaLinha As Long
    Dim i As Long
    Dim embarque As String

    Dim caminhoProg As String
    Dim arquivoProg As String

    Dim dataAtual As Date
    Dim valorN As Variant

    Dim chaveGrupo As Variant
    Dim chaveEmbarque As Variant

    Dim celEncontrada As Range
    Dim primeiraCel As String

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    dataAtual = Now
    caminhoProg = PastaProgramacoesProjeto()
    If Len(caminhoProg) = 0 Then GoTo Finalizar

    ' ==================================================
    ' Selecionar arquivo INICIO
    ' ==================================================
    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    With fd
        .Title = "Selecione o arquivo INICIO"
        .InitialFileName = PortfolioPath("arquivos\")
        .AllowMultiSelect = False
        .Filters.Clear
        .Filters.Add "Arquivos Excel", "*.xlsx"

        If .Show <> -1 Then GoTo Finalizar
        caminhoInicio = .SelectedItems(1)
    End With

    nomeArquivoInicio = UCase(Dir(caminhoInicio))
    If InStr(nomeArquivoInicio, "INICIO") = 0 Then
        MsgBox "O arquivo selecionado não contém 'INICIO' no nome.", vbExclamation
        GoTo Finalizar
    End If

    ' ==================================================
    ' Ler embarques e classificar por grupo
    ' ==================================================
    Set dictGrupos = CreateObject("Scripting.Dictionary")

    Set wbInicio = Workbooks.Open(caminhoInicio)
    Set wsInicio = wbInicio.Sheets(1)

    ultimaLinha = wsInicio.Cells(wsInicio.Rows.Count, "A").End(xlUp).Row
    
    Dim colEmbarque As Long
    Dim celCabecalho As Range
    
    Set celCabecalho = wsInicio.Rows(1).Find(What:="EMBARQUE", LookAt:=xlWhole)
    
    If celCabecalho Is Nothing Then
        MsgBox "Cabeçalho 'EMBARQUE' não encontrado.", vbExclamation
        Exit Sub
    End If
    
    colEmbarque = celCabecalho.Column

    For i = 2 To ultimaLinha

        embarque = Trim(wsInicio.Cells(i, colEmbarque).Value)
        If embarque = "" Then GoTo ProximoEmbarque

        Select Case True
            Case Left(embarque, 4) = "FTON"
                grupo = "Tbl_DistrMCD_FT_*OnePallet.xlsx"
        
            Case Left(embarque, 2) = "FT"
                grupo = "Tbl_DistrMCD_FT_*.xlsx"
        
            Case Left(embarque, 4) = "PRRS"
                grupo = "Tbl_DistrMCD_RS_*.xlsx"
        
            Case Left(embarque, 2) = "PR"
                grupo = "Tbl_DistrMCD_PR_*.xlsx"
        
            Case Left(embarque, 2) = "NE"
                grupo = "Tbl_DistrMCD_NE_*.xlsx"
        
            ' RJ
            Case Left(embarque, 4) = "JCNF", Left(embarque, 4) = "JCRJ", _
                 Left(embarque, 4) = "JCRL", Left(embarque, 4) = "JCSF", _
                 Left(embarque, 4) = "JCSR", Left(embarque, 4) = "JCJF"
                grupo = "Tbl_DistrMCD_ES_RJ_*.xlsx"
        
            ' BA / BS / MG
            Case Left(embarque, 4) = "JCBA", Left(embarque, 4) = "JCBS", _
                 Left(embarque, 4) = "JCIG", Left(embarque, 4) = "JCMG", _
                 Left(embarque, 4) = "JCCS", Left(embarque, 4) = "JCMC"
                grupo = "Tbl_DistrMCD_BA_BS_MG_*.xlsx"
        
            ' DF / GO / MT / TO
            Case Left(embarque, 4) = "JCDF", Left(embarque, 4) = "JCGO", _
                 Left(embarque, 4) = "JCMT", Left(embarque, 4) = "JCTO"
                grupo = "Tbl_DistrMCD_DF_GO_MT_TO_*.xlsx"
        
            ' DIRETAS
            Case Left(embarque, 4) = "JCSM", Left(embarque, 4) = "JCTM", _
                 Left(embarque, 4) = "JCRO", Left(embarque, 4) = "JCMS", _
                 Left(embarque, 4) = "JCMD"
                grupo = "Tbl_DistrMCD_DIRETAS_*.xlsx"
        
            ' ES
            Case Left(embarque, 4) = "JCES"
                grupo = "Tbl_DistrMCD_ES_RJ_*.xlsx"
        
            ' JC genéricos
            Case Left(embarque, 4) = "JCRE", Left(embarque, 4) = "JCAB", _
                 Left(embarque, 4) = "JCVP", Left(embarque, 4) = "JCZL", _
                 Left(embarque, 4) = "JCLN"
                grupo = "Tbl_DistrMCD_JC_*.xlsx"
        
            ' Se começar com JC mas não for nenhum dos casos acima
            Case Left(embarque, 2) = "JC"
                GoTo ProximoEmbarque
        
            Case Else
                GoTo ProximoEmbarque
        End Select

        If Not dictGrupos.Exists(grupo) Then
            Set dictEmbarques = CreateObject("Scripting.Dictionary")
            dictEmbarques.CompareMode = vbTextCompare
            dictGrupos.Add grupo, dictEmbarques
        End If

        dictGrupos(grupo)(embarque) = True

ProximoEmbarque:
    Next i

    wbInicio.Close SaveChanges:=False

    ' ==================================================
    ' Processar apenas os arquivos necessários
    ' ==================================================
    For Each chaveGrupo In dictGrupos.Keys

        arquivoProg = Dir(caminhoProg & chaveGrupo)

        Do While arquivoProg <> ""

            Set wbProg = Workbooks.Open(caminhoProg & arquivoProg)
            Set wsProg = wbProg.Sheets(1)

            For Each chaveEmbarque In dictGrupos(chaveGrupo).Keys

                Set celEncontrada = wsProg.Columns("O").Find( _
                                        What:=chaveEmbarque, _
                                        LookIn:=xlValues, _
                                        LookAt:=xlWhole, _
                                        MatchCase:=False)

                If Not celEncontrada Is Nothing Then

                    primeiraCel = celEncontrada.Address

                    Do
                        ' Preenche somente se AD estiver vazia
                        If celEncontrada.EntireRow.Cells(1, "AD").Value = "" Then

                            valorN = celEncontrada.EntireRow.Cells(1, "N").Value

                            If IsDate(valorN) Then
                                celEncontrada.EntireRow.Cells(1, "AD").Value = CDate(valorN)
                                celEncontrada.EntireRow.Cells(1, "AD").NumberFormat = "dd/mm/yyyy hh:mm"
                            End If

                        End If

                        Set celEncontrada = wsProg.Columns("O").FindNext(celEncontrada)

                    Loop While Not celEncontrada Is Nothing _
                           And celEncontrada.Address <> primeiraCel

                End If

            Next chaveEmbarque

            wbProg.Close SaveChanges:=True
            arquivoProg = Dir
        Loop

    Next chaveGrupo

    MsgBox "Atualização de Morosidade Início (otimizada) concluída.", vbInformation

Finalizar:
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True

End Sub

Function CarregarEmailsRCS() As Object

    Dim wb As Workbook
    Dim ws As Worksheet
    Dim dict As Object
    Dim ultimaLinha As Long
    Dim i As Long
    Dim chave As String

    Set dict = CreateObject("Scripting.Dictionary")
    dict.CompareMode = vbTextCompare

    Set wb = Workbooks.Open( _
        ArquivoProjeto("CD e NOMENCLATURAS.xlsm"), UpdateLinks:=0, ReadOnly:=True)
    Set ws = wb.Sheets("Planilha2")

    ultimaLinha = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row

    For i = 2 To ultimaLinha
        chave = Trim(ws.Cells(i, "A").Value) & Trim(ws.Cells(i, "B").Value)
        dict(chave) = ws.Cells(i, "D").Value   ' e-mail
    Next i

    wb.Close SaveChanges:=False
    Set CarregarEmailsRCS = dict

End Function

Function ExtrairRotaDoAssunto(assunto As String) As String
    ' Ex: RCS NÃO UTILIZADO - JC - ZL603M - 17/04/2026
    Dim partes() As String
    partes = Split(assunto, " - ")
    
    If UBound(partes) >= 2 Then
        ExtrairRotaDoAssunto = partes(2)
    Else
        ExtrairRotaDoAssunto = ""
    End If
End Function

Function URLEncode(str As String) As String
    Dim i As Long
    Dim charCode As Integer
    Dim result As String
    
    For i = 1 To Len(str)
        charCode = Asc(Mid(str, i, 1))
        Select Case charCode
            Case 48 To 57, 65 To 90, 97 To 122 ' 0-9, A-Z, a-z
                result = result & Chr(charCode)
            Case 32
                result = result & "%20"
            Case 10, 13 ' Quebra de linha
                result = result & "%0A"
            Case Else
                result = result & "%" & Hex(charCode)
        End Select
    Next i
    
    URLEncode = result
End Function

Sub RetornoCDWhatsApp()

    Dim wbImport As Workbook
    Dim wsImport As Worksheet
    Dim fd As FileDialog
    Dim fPath As String
    Dim fName As String
    Dim lastRow As Long
    Dim i As Long
    Dim logRow As Long
    Dim embarque As String
    Dim status As String
    Dim rota As String
    Dim retornoPrev As String
    Dim placa As String
    Dim rotograma As String
    Dim msg As String
    Dim waURL As String
    
    ' Caixa de diálogo na pasta de exemplos do projeto
    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    fd.InitialFileName = PortfolioPath("arquivos\")
    fd.Filters.Clear
    fd.Filters.Add "Arquivos Excel", "*.xlsx"
    
    If fd.Show = -1 Then
        fPath = fd.SelectedItems(1)
        fName = Dir(fPath)
    Else
        Exit Sub
    End If
    
    ' Verifica nome do arquivo
    If fName <> "AG RETORNO AO CD.xlsx" Then
        MsgBox "Arquivo inválido. Selecione o 'AG RETORNO AO CD'.", vbExclamation
        Exit Sub
    End If
    
    ' Abre o arquivo
    Set wbImport = Workbooks.Open(fPath)
    Set wsImport = wbImport.Sheets(1)
    
    lastRow = wsImport.Cells(wsImport.Rows.Count, "C").End(xlUp).Row
    
    ' Garante que o Edge está aberto com uma aba neutra
    Shell "cmd /c start microsoft-edge:about:blank", vbHide
    
    ' Loop em todos os casos válidos
    For i = 2 To lastRow
        status = wsImport.Cells(i, "M").Value
        embarque = wsImport.Cells(i, "O").Value
        
        ' Verifica se já foi enviado
        logRow = ThisWorkbook.Sheets(1).Cells(ThisWorkbook.Sheets(1).Rows.Count, "J").End(xlUp).Row
        If Application.WorksheetFunction.CountIf(ThisWorkbook.Sheets(1).Range("J3:J" & logRow), embarque) > 0 Then
            GoTo Proximo
        End If
        
        If status = "RETORNANDO AO CD" Then
            rota = wsImport.Cells(i, "C").Value
            retornoPrev = Format(wsImport.Cells(i, "F").Value, "dd/mm/yyyy hh:mm")
            placa = wsImport.Cells(i, "I").Value
            Dim motorista As String
            motorista = wsImport.Cells(i, "J").Value
            rotograma = wsImport.Cells(i, "K").Value
            If rotograma = "Não encontrado" Then rotograma = "Sem rotograma"
            
            If motorista = "" Then
            
                ' Monta mensagem com Motorista vazio
                msg = "Motorista: " & vbLf & _
                      "Status: Retornando ao CD" & vbLf & _
                      "Retorno previsto: " & retornoPrev & vbLf & _
                      "Rota: " & rota & vbLf & _
                      "Link Rotograma: " & rotograma
                  
            Else
            
                ' Monta mensagem com Motorista
                msg = "Motorista: " & motorista & vbLf & _
                      "Status: Retornando ao CD" & vbLf & _
                      "Retorno previsto: " & retornoPrev & vbLf & _
                      "Rota: " & rota & vbLf & _
                      "Link Rotograma: " & rotograma
            
            End If
                  
            ' Abre WhatsApp Web com mensagem pré-preenchida
            waURL = "https://web.whatsapp.com/send?text=" & URLEncode(msg)
            Shell "cmd /c start microsoft-edge:" & waURL, vbHide
            
            ' Log
            logRow = logRow + 1
            With ThisWorkbook.Sheets(1)
                .Cells(logRow, "J").Value = embarque
                .Cells(logRow, "K").Value = "Pré-preenchido"
            End With
            
            ' Espera 25 segundos para você ajustar e enviar
            Application.Wait (Now + TimeValue("0:00:25"))
            
            ' Fecha a aba atual (Ctrl+W)
            AppActivate "Edge"
            SendKeys "^w"
        End If
        
Proximo:
    Next i
    
    wbImport.Close False
    
End Sub

Function Saudacao() As String
    Dim horaAtual As Integer
    horaAtual = Hour(Now)

    If horaAtual < 12 Then
        Saudacao = "bom dia"
    ElseIf horaAtual < 19 Then
        Saudacao = "boa tarde"
    Else
        Saudacao = "boa noite"
    End If
End Function

Function BuscarEmailResponsavelTD(assunto As String, wsResp As Worksheet) As String

    Dim cdFatResp As String
    Dim rotaResp As String
    Dim nomenclaturaResp As String
    Dim responsavelTD As String
    Dim linhaResp As Long
    Dim ultimaResp As Long

    BuscarEmailResponsavelTD = ""

    ' extrair dados do assunto
    'cdFatResp = Mid(assunto, 14, 2)
    'rotaResp = Mid(assunto, 19, InStr(19, assunto, " -") - 19)
    'nomenclaturaResp = Left(rotaResp, 2)
    
    Dim partes As Variant
    partes = Split(assunto, " - ")
    
    If UBound(partes) >= 3 Then
        cdFatResp = Trim(partes(1))
        rotaResp = Trim(partes(2))
        nomenclaturaResp = Left(rotaResp, 2)
    End If

    ultimaResp = wsResp.Cells(wsResp.Rows.Count, "A").End(xlUp).Row

    ' -----------------------------
    ' 1. descobrir responsável
    ' -----------------------------
    For linhaResp = 2 To ultimaResp

        If wsResp.Cells(linhaResp, "A").Value = cdFatResp And _
           wsResp.Cells(linhaResp, "B").Value = nomenclaturaResp Then
            
            responsavelTD = wsResp.Cells(linhaResp, "C").Value
            Exit For
            
        End If

    Next linhaResp

    ' -----------------------------
    ' 2. buscar o e-mail da transportadora cadastrada
    ' -----------------------------
    If Len(Trim$(responsavelTD)) > 0 Then

        For linhaResp = 2 To ultimaResp
        
            If StrComp(Trim$(wsResp.Cells(linhaResp, "F").Value), Trim$(responsavelTD), vbTextCompare) = 0 Then
                
                BuscarEmailResponsavelTD = wsResp.Cells(linhaResp, "G").Value
                Exit Function
                
            End If
        
        Next linhaResp

    End If

End Function

Public Function ObterMotivoExibicao(MotivoOriginal As String) As String

    Dim wsMotivos As Worksheet
    Dim ultimaLinha As Long
    Dim i As Long
    Dim Classificacao As String

    Set wsMotivos = ThisWorkbook.Sheets("Motivos")

    ultimaLinha = wsMotivos.Cells(wsMotivos.Rows.Count, "A").End(xlUp).Row

    For i = 2 To ultimaLinha

        If Trim(wsMotivos.Cells(i, "A").Value) = Trim(MotivoOriginal) Then

            Classificacao = UCase(Trim(wsMotivos.Cells(i, "B").Value))

            Select Case Classificacao

                Case "REAL"
                    ObterMotivoExibicao = MotivoOriginal

                Case "MB"
                    ObterMotivoExibicao = "Problemas operacionais"

                Case Else
                    ObterMotivoExibicao = "Problemas operacionais"

            End Select

            Exit Function

        End If

    Next i

    ' Se não encontrar o motivo na planilha
    ObterMotivoExibicao = "Problemas operacionais"

End Function



