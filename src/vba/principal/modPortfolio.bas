Attribute VB_Name = "modPortfolio"
Option Explicit

Public Const EMAIL_TORRE As String = "operacao@example.com"
Public Const EMAIL_JORNADA As String = "jornada@example.com"
Public Const EMAIL_RELACIONAMENTO As String = "relacionamento@example.com"

Public Function PortfolioPath(ByVal caminhoRelativo As String) As String
    Dim pasta As String, raizOneDrive As String
    Dim fimConta As Long, fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    pasta = ThisWorkbook.Path
    If LCase$(Left$(pasta, 24)) = "https://d.docs.live.net/" Then
        raizOneDrive = Environ$("OneDriveConsumer")
        If Len(raizOneDrive) = 0 Then raizOneDrive = Environ$("OneDrive")
        fimConta = InStr(25, pasta, "/")
        If Len(raizOneDrive) > 0 And fimConta > 0 Then
            pasta = fso.BuildPath(raizOneDrive, Replace(Mid$(pasta, fimConta + 1), "/", "\"))
            pasta = Replace(pasta, "%20", " ")
        End If
    End If
    If Len(pasta) = 0 Or InStr(1, pasta, "://", vbTextCompare) > 0 Then
        Err.Raise vbObjectError + 2101, "PortfolioPath", _
            "Abra o arquivo em uma pasta local ou sincronizada e mantenha os arquivos do projeto juntos."
    End If
    PortfolioPath = fso.GetAbsolutePathName(fso.BuildPath(pasta, caminhoRelativo))
    If Right$(caminhoRelativo, 1) = "\" And Right$(PortfolioPath, 1) <> "\" Then
        PortfolioPath = PortfolioPath & "\"
    End If
End Function

Public Function ArquivoProjeto(ByVal nomeArquivo As String) As String
    Dim fso As Object, caminho As String
    Set fso = CreateObject("Scripting.FileSystemObject")
    caminho = PortfolioPath(nomeArquivo)
    If Not fso.FileExists(caminho) Then
        Err.Raise vbObjectError + 2102, "ArquivoProjeto", _
            "Arquivo auxiliar não encontrado: " & nomeArquivo & vbCrLf & _
            "Mantenha-o na mesma pasta do arquivo principal."
    End If
    ArquivoProjeto = caminho
End Function

Public Function PastaProgramacoesProjeto() As String
    Dim fso As Object, pasta As String
    Set fso = CreateObject("Scripting.FileSystemObject")
    pasta = PortfolioPath("programacoes\")
    If Not fso.FolderExists(pasta) Then
        MsgBox "Esta rotina adicional precisa de arquivos na pasta programacoes do projeto. " & _
            "Os exemplos de e-mail usam apenas a pasta arquivos.", vbInformation
        Exit Function
    End If
    PastaProgramacoesProjeto = pasta
End Function

Public Function EmailSupervisorPortfolio(ByVal nome As String) As String
    Select Case UCase$(Trim$(nome))
        Case "SUPERVISOR DEMO 01"
            EmailSupervisorPortfolio = "supervisor.demo01@example.com;"
        Case "SUPERVISOR DEMO 02"
            EmailSupervisorPortfolio = "supervisor.demo02@example.com;"
    End Select
End Function

Public Function TextoSuportePortfolio() As String
    TextoSuportePortfolio = "<p>Para dúvidas sobre a demonstração, entre em contato com " & _
        "<b>suporte@example.com</b> e informe a sigla da unidade. " & _
        "Todos os dados deste exemplo são fictícios.</p>"
End Function

Public Function ColunaCabecalho(ByVal planilha As Worksheet, ByVal nome As String) As Long
    Dim coluna As Long, ultimaColuna As Long
    ultimaColuna = planilha.Cells(1, planilha.Columns.Count).End(xlToLeft).Column
    For coluna = 1 To ultimaColuna
        If StrComp(Trim$(CStr(planilha.Cells(1, coluna).Value2)), Trim$(nome), vbTextCompare) = 0 Then
            ColunaCabecalho = coluna
            Exit Function
        End If
    Next coluna
    Err.Raise vbObjectError + 2103, "ColunaCabecalho", "Cabeçalho obrigatório não encontrado: " & Trim$(nome)
End Function


