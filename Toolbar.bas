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
        
    Dim id As Identifier, url As String
   
    If Not ParseIdentifier(ActiveDocument.Name, id) Then
        MsgBox "O nome do arquivo não se parece com um processo."
        Exit Sub
    End If
    
    url = "https://aplicacao6.tst.jus.br/esij/ConsultarProcesso.do?consultarNumeracao=Consultar" _
    & "&numProc=" & id.Numero & "&digito=" & id.Digito & "&anoProc=" & id.Ano & "&justica=" & id.Justica _
    & "&numTribunal=" & id.Tribunal & " &numVara=" & id.Vara & "&codigoBarra="
    
    Navigate url
    
End Sub


Sub openAcordaoFolder()

    System.Cursor = wdCursorWait

    Dim id As Identifier, folder As String, filename As String
      
    folder = "K:\001 - JOD - GMJOD (2013)\005 - DIVERSOS\TRT"
    
    If Not ParseIdentifier(ActiveDocument.Name, id) Then
        MsgBox "O nome do arquivo não se parece com um processo."
        Exit Sub
    End If
       
    filename = folder & "\" & id.Formatado
    
    If Dir(filename, vbDirectory) <> "" Then
        Explore filename
    Else
        MsgBox "Não há acórdão para o processo especificado"
    End If
    
End Sub

Sub openUltimoDespacho()
    
    System.Cursor = wdCursorWait

    Dim id As Identifier
    
    If Not ParseIdentifier(ActiveDocument.Name, id) Then
        MsgBox "O nome do arquivo não se parece com um processo."
        Exit Sub
    End If

    Dim pk
    pk = getPK(id)
    
    Navigate ("http://aplicacao5.tst.jus.br/decisoes/consultas/ultimoDespachoTRT/" & pk(1) & "/" & pk(0))
    
    
    
End Sub

Sub openAllPDFs()

    System.Cursor = wdCursorWait

    Dim id As Identifier
    
    If Not ParseIdentifier(ActiveDocument.Name, id) Then
        MsgBox "O nome do arquivo não se parece com um processo."
        Exit Sub
    End If

    openAll id
    
End Sub
