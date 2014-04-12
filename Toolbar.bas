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
        
        .MatchWildcards = False
        
        .text = "^w"
        .Replacement.text = " "
        .Execute Replace:=wdReplaceAll
        
        .MatchWildcards = True
                
        .text = " {1;}^13"
        .Replacement.text = "^p"
        .Execute Replace:=wdReplaceAll
        
        .text = "([!.])^13"
        .Replacement.text = "\1 "
        .Execute Replace:=wdReplaceAll
        
        .text = " {1;}"
        .Replacement.text = " "
        .Execute Replace:=wdReplaceAll
       
    End With
    
finally: On Error Resume Next
    Application.ScreenUpdating = True
    Exit Sub

try:
    If Err.Number = 4608 Then
        MsgBox "Não há texto selecionado"
    Else
        Catch Err
    End If
    
    Resume finally
    Resume
        
End Sub

Sub esij()
    
    Dim Id  As Identifier
    Dim URL As String
   
On Error GoTo try

    System.Cursor = wdCursorWait
    Application.ScreenUpdating = False
        
    Id = ParseIdentifier(ActiveDocument.Name)
    
    URL = "https://aplicacao6.tst.jus.br/esij/ConsultarProcesso.do?consultarNumeracao=Consultar" _
    & "&numProc=" & Id.Numero & "&digito=" & Id.Digito & "&anoProc=" & Id.Ano & "&justica=" & Id.Justica _
    & "&numTribunal=" & Id.Tribunal & " &numVara=" & Id.Vara & "&codigoBarra="
    
    Navigate URL

finally: On Error Resume Next
    Application.ScreenUpdating = True
    Exit Sub

try: Catch Err
    Resume finally
    Resume
        
End Sub


Sub openAcordaoFolder()

    Dim Id       As Identifier
    Dim folder   As String
    Dim filename As String

On Error GoTo try

    System.Cursor = wdCursorWait
    Application.ScreenUpdating = False

    Id = ParseIdentifier(ActiveDocument.Name)
       
    folder = "K:\TRT\TRT" & Format(Id.Tribunal, "00")
        
    filename = folder & "\" & Id.Formatado
    
    If Dir(filename, vbDirectory) <> "" Then
        Explore filename
    Else
        MsgBox "Não há acórdão para o processo especificado"
    End If


finally: On Error Resume Next
   Application.ScreenUpdating = False
   Exit Sub

try: Catch Err
    Resume finally
    Resume
    
End Sub

Sub importUltimoDespacho()
    
    Dim Id       As Identifier
    Dim undo     As UndoRecord
    Dim request  As WinHttp.WinHttpRequest
    Dim oDoc     As MSHTML.HTMLDocument
    Dim pk()     As String
    Dim URL      As String
    Dim htmlText As String
    
On Error GoTo try

    System.Cursor = wdCursorWait
    Application.ScreenUpdating = False
        
    Id = ParseIdentifier(ActiveDocument.Name)

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
    Application.ScreenUpdating = True
    undo.EndCustomRecord
    Exit Sub

try: Catch Err
    Resume finally
    Resume
    
End Sub


Sub openUltimoDespacho()
    
    Dim Id   As Identifier
    Dim pk() As String

On Error GoTo try

    System.Cursor = wdCursorWait
    Application.ScreenUpdating = False

    Id = ParseIdentifier(ActiveDocument.Name)
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

    System.Cursor = wdCursorWait
    Application.ScreenUpdating = False
    
    Dim Id As Identifier
    
    Id = ParseIdentifier(ActiveDocument.Name)

    openAll Id

finally: On Error Resume Next
    Application.ScreenUpdating = True
    Exit Sub

try: Catch Err
    Resume finally
    Resume
    
End Sub


Public Sub loadStyles()
   
On Error GoTo try

    ActiveDocument.ApplyQuickStyleSet2 "GMJD"

finally: On Error Resume Next
    Exit Sub

try: Catch Err
    Resume finally
    Resume

End Sub

Public Function loadStylesCallback(control As IRibbonControl)
    loadStyles
End Function

