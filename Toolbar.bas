Attribute VB_Name = "Toolbar"
Option Explicit

Sub JoinLines()

   On Error GoTo JoinLines_Error

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
        
        .text = " {1;}"
        .Replacement.text = " "
        .Execute Replace:=wdReplaceAll

        .text = " {1;}^13"
        .Replacement.text = ""
        .Execute Replace:=wdReplaceAll

        .text = "([!.])^13"
        .Replacement.text = "\1 "
        .Execute Replace:=wdReplaceAll
       
    End With
    
    Application.ScreenUpdating = True

   On Error GoTo 0
   Exit Sub

JoinLines_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure JoinLines of Módulo Toolbar"
        
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
    
   On Error GoTo esij_Error

    System.Cursor = wdCursorWait
        
    Dim Id As Identifier, URL As String
   
    If Not ParseIdentifier(ActiveDocument.Name, Id) Then
        MsgBox "O nome do arquivo não se parece com um processo."
        Exit Sub
    End If
    
    URL = "https://aplicacao6.tst.jus.br/esij/ConsultarProcesso.do?consultarNumeracao=Consultar" _
    & "&numProc=" & Id.Numero & "&digito=" & Id.Digito & "&anoProc=" & Id.Ano & "&justica=" & Id.Justica _
    & "&numTribunal=" & Id.Tribunal & " &numVara=" & Id.Vara & "&codigoBarra="
    
    Navigate URL

   On Error GoTo 0
   Exit Sub

esij_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure esij of Módulo Toolbar"
    
End Sub


Sub openAcordaoFolder()

   On Error GoTo openAcordaoFolder_Error

    System.Cursor = wdCursorWait

    Dim Id As Identifier, folder As String, filename As String
      
   
    If Not ParseIdentifier(ActiveDocument.Name, Id) Then
        MsgBox "O nome do arquivo não se parece com um processo."
        Exit Sub
    End If
       
    folder = "K:\TRT\TRT" & Format(Id.Tribunal, "00")
        
    filename = folder & "\" & Id.Formatado
    
    If Dir(filename, vbDirectory) <> "" Then
        Explore filename
    Else
        MsgBox "Não há acórdão para o processo especificado"
    End If

   On Error GoTo 0
   Exit Sub

openAcordaoFolder_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure openAcordaoFolder of Módulo Toolbar"
    
End Sub

Sub importUltimoDespacho()
    
   On Error GoTo importUltimoDespacho_Error

    System.Cursor = wdCursorWait

    Dim Id As Identifier
        
    If Not ParseIdentifier(ActiveDocument.Name, Id) Then
        MsgBox "O nome do arquivo não se parece com um processo."
        Exit Sub
    End If

    Dim pk
    pk = getPK(Id)
    
    
    Dim request As New WinHttpRequest

    Dim URL As String
    Dim htmlText As String
    Dim oDoc As New HTMLDocument
    
    URL = "http://aplicacao5.tst.jus.br/decisoes/consultas/ultimoDespachoTRT/" & pk(1) & "/" & pk(0)
    
    request.Open "GET", URL, True
    request.Send
    request.WaitForResponse
    
    htmlText = request.ResponseText
    oDoc.body.innerHTML = htmlText
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
        
        .text = " {1;}"
        .Replacement.text = " "
        .Execute Replace:=wdReplaceAll

        .text = "^13{1;}"
        .Replacement.text = "^13"
        .Execute Replace:=wdReplaceAll
    
    End With
    
    Application.ScreenUpdating = True

   On Error GoTo 0
   Exit Sub

importUltimoDespacho_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure importUltimoDespacho of Módulo Toolbar"
    
End Sub

Sub openUltimoDespacho()
    
   On Error GoTo openUltimoDespacho_Error

    System.Cursor = wdCursorWait

    Dim Id As Identifier
    
    If Not ParseIdentifier(ActiveDocument.Name, Id) Then
        MsgBox "O nome do arquivo não se parece com um processo."
        Exit Sub
    End If

    Dim pk
    pk = getPK(Id)
    
    Navigate ("http://aplicacao5.tst.jus.br/decisoes/consultas/ultimoDespachoTRT/" & pk(1) & "/" & pk(0))

   On Error GoTo 0
   Exit Sub

openUltimoDespacho_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure openUltimoDespacho of Módulo Toolbar"
    
End Sub


Sub openAllPDFs()

   On Error GoTo openAllPDFs_Error

    System.Cursor = wdCursorWait

    Dim Id As Identifier
    
    If Not ParseIdentifier(ActiveDocument.Name, Id) Then
        MsgBox "O nome do arquivo não se parece com um processo."
        Exit Sub
    End If

    openAll Id

   On Error GoTo 0
   Exit Sub

openAllPDFs_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure openAllPDFs of Módulo Toolbar"
    
End Sub

Public Sub loadStyles()
   
   On Error GoTo loadStyles_Error

    ActiveDocument.ApplyQuickStyleSet2 "GMJD"

   On Error GoTo 0
   Exit Sub

loadStyles_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure loadStyles of Módulo Toolbar"

End Sub

