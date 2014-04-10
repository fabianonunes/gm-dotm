Attribute VB_Name = "DocumentContents"

Sub italicsLatin()

    Dim fs As FileSystemObject
    Dim stream As TextStream
    Dim text As String
    Dim undo As UndoRecord
    Dim oRng As Word.Range
    Dim counter As Integer
    
   On Error GoTo italicsLatin_Error

    Set fs = New FileSystemObject
    Set stream = fs.OpenTextFile(LATIN_FILEPATH)
    Set undo = Application.UndoRecord
    counter = 0
    
    undo.EndCustomRecord
    undo.StartCustomRecord ("Destacar palavras em latim")
    
    Do While Not stream.AtEndOfStream
        
        text = stream.ReadLine
        
        Set oRng = ActiveDocument.Range
        With oRng.Find
            .ClearFormatting
            .MatchWholeWord = True
            .text = text
            While .Execute
                counter = counter + 1
                With oRng
                    .text = text
                    .Font.Italic = True
                End With
            Wend
        End With
    
    Loop
    
    undo.EndCustomRecord
 
    stream.Close
    
    Application.ScreenUpdating = True
    
    If counter = 0 Then
        MsgBox "Nenhuma expressão foi encontrada."
    ElseIf counter = 1 Then
        MsgBox "Uma expressão foi destacada."
    ElseIf counter > 1 Then
        MsgBox counter & " expressões foram destacadas."
    End If

   On Error GoTo 0
   Exit Sub

italicsLatin_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure italicsLatin of Módulo DocumentContents"

End Sub
Sub comment()
         
    Dim excel_app As Excel.Application
    Dim workbook As Excel.workbook
    Dim sheet As Excel.Worksheet
    Dim table As Excel.ListObject
    Dim table_rng As Excel.Range
    Dim doc_rng As Range
    Dim undo As UndoRecord
    Dim size As Integer

   On Error GoTo comment_Error
   
    AutoExec.AutoExec

    SendKeys "%v%"
    removeComments

    Set excel_app = New Excel.Application
    Set workbook = excel_app.Workbooks.Open(filename:=DICX_FILEPATH, ReadOnly:=True)
    Set sheet = workbook.Sheets("Dicionario")
    Set table = sheet.ListObjects("TabelaDicionario")
    Set undo = Application.UndoRecord
    
    undo.StartCustomRecord "Destacar Expressões"

    For Each table_rng In table.DataBodyRange.rows

        Set doc_rng = ActiveDocument.Range

        With doc_rng.Find
            .ClearFormatting
            .MatchWholeWord = True
            .text = table_rng.Cells(1).Value

            While .Execute

                With doc_rng

                    If .Style Like "Transcrição*" Then

                        GoTo NextIteration
                    End If

                    If table_rng.Cells(3) <> "" And .Style <> table_rng.Cells(3) Then
                       GoTo NextIteration
                    End If

                    .Comments.Add Range:=doc_rng, text:=table_rng.Cells(2).Value


                End With

NextIteration:

            Wend

        End With

    Next

    workbook.Close
    Set workbook = Nothing

    undo.EndCustomRecord
    
    Set doc_rng = ActiveDocument.Range
    
    If doc_rng.Comments.Count > 0 Then
        doc_rng.Comments.Item(1).Reference.Select
    Else
        MsgBox "Nenhuma expressão foi encontrada."
    End If
    
    Application.ScreenUpdating = True

   On Error GoTo 0
   Exit Sub

comment_Error:

    workbook.Close
    Catch Err

End Sub

Sub comment_filebased()
         
    Dim fs As FileSystemObject
    Dim stream As TextStream
    Dim text As String
    Dim undo As UndoRecord
    Dim splitted() As String
    Dim oRng As Word.Range
    Dim size As Integer

   On Error GoTo comment_Error

    Set fs = New FileSystemObject
    Set stream = fs.OpenTextFile(DIC_FILEPATH)
    Set undo = Application.UndoRecord
    
    SendKeys "%v%"
    
    removeComments
    
    undo.EndCustomRecord
    undo.StartCustomRecord "Destacar Expressões"
    
    Do While Not stream.AtEndOfStream
        
        text = stream.ReadLine
        splitted = Split(text, "|")
        size = UBound(splitted) + 1
        
        Set oRng = ActiveDocument.Range
        
        With oRng.Find
            .ClearFormatting
            .MatchWholeWord = True
            .text = splitted(0)
            
            While .Execute
            
                With oRng
                
                    If .Style Like "Transcrição*" Then
                        GoTo NextIteration
                    End If
                
                    If size > 1 Then
                        
                        If (size = 3) Then
                            If .Style <> splitted(2) Then
                               GoTo NextIteration
                            End If
                        End If
                        
                        .Comments.Add Range:=oRng, text:=splitted(1)
                        
                    End If
                    
                End With
                
NextIteration:

            Wend
            
        End With

    Loop
    
    undo.EndCustomRecord
    
    Set oRng = ActiveDocument.Range
    
    If oRng.Comments.Count > 0 Then
        oRng.Comments.Item(1).Reference.Select
    Else
        MsgBox "Nenhuma expressão foi encontrada."
    End If
    
    stream.Close
    Application.ScreenUpdating = True

   On Error GoTo 0
   Exit Sub

comment_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure comment of Módulo DocumentContents"

End Sub

Private Sub removeComments()
   
   Dim oRng As Word.Range, i As Integer
   
  On Error GoTo removeComments_Error

    Set oRng = ActiveDocument.Range

    With oRng.Comments
      For i = .Count To 1 Step -1
           .Item(i).Delete
      Next i
    End With

   On Error GoTo 0
   Exit Sub

removeComments_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure removeComments of Módulo DocumentContents"

End Sub

