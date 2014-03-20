Attribute VB_Name = "Toolbar"
Option Explicit

Sub JoinLines()

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
    
End Sub


Sub openAcordaoFolder()

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
    
End Sub

Sub importUltimoDespacho()
    
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
    'Selection.TypeText oDoc.body.innerText
    
End Sub

Private Function FunctionReadyStateChange()
    
End Function

Sub openUltimoDespacho()
    
    System.Cursor = wdCursorWait

    Dim Id As Identifier
    
    If Not ParseIdentifier(ActiveDocument.Name, Id) Then
        MsgBox "O nome do arquivo não se parece com um processo."
        Exit Sub
    End If

    Dim pk
    pk = getPK(Id)
    
    Navigate ("http://aplicacao5.tst.jus.br/decisoes/consultas/ultimoDespachoTRT/" & pk(1) & "/" & pk(0))
    
End Sub


Sub openAllPDFs()

    System.Cursor = wdCursorWait

    Dim Id As Identifier
    
    If Not ParseIdentifier(ActiveDocument.Name, Id) Then
        MsgBox "O nome do arquivo não se parece com um processo."
        Exit Sub
    End If

    openAll Id
    
End Sub
