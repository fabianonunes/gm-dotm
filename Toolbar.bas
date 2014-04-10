Attribute VB_Name = "Toolbar"
Option Explicit

Sub JoinLines()

   On Error GoTo try

    System.Cursor = wdCursorWait
    Application.ScreenUpdating = False
    
    Dim selBkUp As Range
    Set selBkUp = ActiveDocument.Range(Selection.Range.Start, Selection.Range.End - 1)
    
    With selBkUp.Find
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
                
        .text = " {1;}^13"
        .Replacement.text = "^p"
        .Execute Replace:=wdReplaceAll
        
        .text = " {1;}"
        .Replacement.text = " "
        .Execute Replace:=wdReplaceAll

        .text = "([!.])^13"
        .Replacement.text = "\1 "
        .Execute Replace:=wdReplaceAll
       
    End With
    
finally:
   On Error GoTo 0
    Application.ScreenUpdating = True
    Exit Sub

try:
    If Err.Number = 4608 Then
        MsgBox "Não há texto selecionado"
    Else
        Catch Err
    End If
    
    GoTo finally
        
End Sub

Sub destacarParagrafo()
    Selection.Paragraphs(1).Range.Select
    With Selection.Borders(wdBorderRight)
        .LineStyle = wdLineStyleSingle
        .LineWidth = wdLineWidth150pt
    End With
End Sub

Sub removerDestaque()
    Selection.Paragraphs(1).Range.Select
    Selection.Borders(wdBorderRight).LineStyle = wdLineStyleNone
End Sub

Sub esij()
    
On Error GoTo try

    System.Cursor = wdCursorWait
    Application.ScreenUpdating = False
        
    Dim Id As Identifier
    Dim URL As String
   
    Id = ParseIdentifier(ActiveDocument.Name)
    
    URL = "https://aplicacao6.tst.jus.br/esij/ConsultarProcesso.do?consultarNumeracao=Consultar" _
    & "&numProc=" & Id.Numero & "&digito=" & Id.Digito & "&anoProc=" & Id.Ano & "&justica=" & Id.Justica _
    & "&numTribunal=" & Id.Tribunal & " &numVara=" & Id.Vara & "&codigoBarra="
    
    Navigate URL

finally:
   On Error GoTo 0
    Application.ScreenUpdating = True
    Exit Sub

try:
    Catch Err
    GoTo finally
        
End Sub


Sub openAcordaoFolder()

On Error GoTo try

    System.Cursor = wdCursorWait
    Application.ScreenUpdating = False

    Dim Id As Identifier, folder As String, filename As String
   
    Id = ParseIdentifier(ActiveDocument.Name)
       
    folder = "K:\TRT\TRT" & Format(Id.Tribunal, "00")
        
    filename = folder & "\" & Id.Formatado
    
    If Dir(filename, vbDirectory) <> "" Then
        Explore filename
    Else
        MsgBox "Não há acórdão para o processo especificado"
    End If

finally:
   On Error GoTo 0
    Application.ScreenUpdating = True
    Exit Sub

try:
    Catch Err
    GoTo finally
    
End Sub

Sub importUltimoDespacho()
    
On Error GoTo try

    System.Cursor = wdCursorWait
    Application.ScreenUpdating = False
    
    Dim Id As Identifier
    Dim undo As UndoRecord
        
    Id = ParseIdentifier(ActiveDocument.Name)

    Dim pk
    pk = getPK(Id)
    
    Dim request As WinHttp.WinHttpRequest
    Dim oDoc As MSHTML.HTMLDocument
    
    Dim URL As String
    Dim htmlText As String
    
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
    
finally:
   On Error GoTo 0
    Application.ScreenUpdating = True
    undo.EndCustomRecord
    Exit Sub

try:
    Catch Err
    GoTo finally
    
End Sub


Sub openUltimoDespacho()
    
On Error GoTo try

    System.Cursor = wdCursorWait
    Application.ScreenUpdating = False

    Dim Id As Identifier
    Id = ParseIdentifier(ActiveDocument.Name)

    Dim pk
    pk = getPK(Id)
    
    Navigate ("http://aplicacao5.tst.jus.br/decisoes/consultas/ultimoDespachoTRT/" & pk(1) & "/" & pk(0))

finally:
   On Error GoTo 0
    Application.ScreenUpdating = True
    Exit Sub

try:
    Catch Err
    GoTo finally
    
End Sub


Sub openAllPDFs()

On Error GoTo try

    System.Cursor = wdCursorWait
    Application.ScreenUpdating = False
    
    Dim Id As Identifier
    
    Id = ParseIdentifier(ActiveDocument.Name)

    openAll Id

finally:
   On Error GoTo 0
    Application.ScreenUpdating = True
    Exit Sub

try:
    Catch Err
    GoTo finally
    
End Sub

