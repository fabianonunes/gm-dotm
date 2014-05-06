Attribute VB_Name = "Reference"
Option Explicit

Private Const ACORDAOS_FOLDER As String = "K:\TRT\"

Sub esij()
    
On Error GoTo try

    Dim Id  As Identifier
    Dim URL As String
    
    Helpers.waitApplication
        
    Id = ParseIdentifier(ActiveDocument.name)
    
    URL = "https://aplicacao6.tst.jus.br/esij/ConsultarProcesso.do?consultarNumeracao=Consultar" _
    & "&numProc=" & Id.Numero & "&digito=" & Id.Digito & "&anoProc=" & Id.Ano & "&justica=" & Id.Justica _
    & "&numTribunal=" & Id.Tribunal & " &numVara=" & Id.Vara & "&codigoBarra="
    
    Navigate URL

finally: On Error Resume Next
    Helpers.resumeApplication
    Exit Sub

try: Catch Err
    Resume finally
    Resume
        
End Sub


Sub openAcordaoFolder()

On Error GoTo try

    Dim Id       As Identifier
    Dim folder   As String
    Dim filename As String
    
    Helpers.waitApplication

    Id = ParseIdentifier(ActiveDocument.name)

    folder = ACORDAOS_FOLDER & "TRT" & Format(Id.Tribunal, "00")

    filename = folder & "\" & Id.Formatado
    
    If Dir(filename, vbDirectory) <> "" Then
        Explore filename
    Else
        MsgBox "Não há acórdão para o processo especificado"
    End If


finally: On Error Resume Next
   Helpers.resumeApplication
   Exit Sub

try: Catch Err
    Resume finally
    Resume
    
End Sub

Sub openMemorialFolder()

On Error GoTo try

    Dim Id       As Identifier
    Dim folder   As String
    Dim filename As String
    
    Helpers.waitApplication

    Id = ParseIdentifier(ActiveDocument.name)

    folder = Constants.MEMORIAIS_PATH & Id.Formatado
    
    If Dir(folder, vbDirectory) <> "" Then
        Explore folder
    Else
        MsgBox "Não há memoriais para o processo especificado"
    End If


finally: On Error Resume Next
   Helpers.resumeApplication
   Exit Sub

try: Catch Err
    Resume finally
    Resume
    
End Sub


Sub importUltimoDespacho()
    
On Error GoTo try

    Dim Id       As Identifier
    Dim undo     As UndoRecord
    Dim request  As WinHttp.WinHttpRequest
    Dim oDoc     As MSHTML.HTMLDocument
    Dim pk()     As String
    Dim URL      As String
    Dim htmlText As String

    Helpers.waitApplication

    Id = ParseIdentifier(ActiveDocument.name)
    pk = getPK(Id)
    
    Set request = New WinHttp.WinHttpRequest
    Set oDoc = New MSHTML.HTMLDocument
    
    URL = "http://aplicacao5.tst.jus.br/decisoes/consultas/ultimoDespachoTRT/" & pk(1) & "/" & pk(0)
    
    request.Open "GET", URL, True
    request.Send
    request.WaitForResponse
    
    htmlText = request.ResponseText
    oDoc.body.innerHTML = htmlText
    
    Set undo = Application.UndoRecord
    undo.StartCustomRecord ("Importar Despacho")
    
    Selection.Style = ActiveDocument.Styles("Transcrição")
    Selection.InsertAfter oDoc.body.innerText
    
    With Selection.Find
        .ClearFormatting
        .Replacement.ClearFormatting
        .Forward = True
        .Wrap = wdFindStop
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchAllWordForms = False
        .MatchSoundsLike = False
        .MatchWildcards = True

        .text = "^13 {1;}^13"
        .Replacement.text = "^p"
        .Execute Replace:=wdReplaceAll

        .text = "^13{1;}"
        .Replacement.text = "^p"
        .Execute Replace:=wdReplaceAll

    End With
    
finally: On Error Resume Next
    Helpers.resumeApplication
    undo.EndCustomRecord
    Exit Sub

try: Catch Err
    Resume finally
    Resume
    
End Sub

Sub openUltimoDespacho()
    
On Error GoTo try

    Dim Id   As Identifier
    Dim pk() As String
    
    System.Cursor = wdCursorWait
    Application.ScreenUpdating = False

    Id = ParseIdentifier(ActiveDocument.name)
    pk = getPK(Id)
    
    Navigate "http://aplicacao5.tst.jus.br/decisoes/consultas/ultimoDespachoTRT/" & pk(1) & "/" & pk(0)

finally: On Error Resume Next
    Application.ScreenUpdating = True
    Exit Sub

try: Catch Err
    Resume finally
    Resume
    
End Sub


Sub openAllPDFs()

On Error GoTo try

    Dim Id As Identifier
    Dim pk() As String
    
    Helpers.waitApplication
    
    Id = ParseIdentifier(ActiveDocument.name)
    pk = getPK(Id)
    
    Navigate _
        ("https://aplicacao6.tst.jus.br/esij/VisualizarPecas.do?visualizarTodos=1&anoProcInt=" _
        & pk(1) & "&numProcInt=" & pk(0))

finally: On Error Resume Next
    Helpers.resumeApplication
    Exit Sub

try: Catch Err
    Resume finally
    Resume
    
End Sub

Public Function openAllCallback(control As IRibbonControl)
    openAllPDFs
End Function

Public Function esijCallback(control As IRibbonControl)
    esij
End Function

Public Function openAcordaosFolderCallback(control As IRibbonControl)
    openAcordaoFolder
End Function

Public Function openMemorialFolderCallback(control As IRibbonControl)
    openMemorialFolder
End Function

Public Function importUltimoDespachoCallback(control As IRibbonControl)
    importUltimoDespacho
End Function

Public Function openUltimoDespachoCallback(control As IRibbonControl)
    openUltimoDespacho
End Function

