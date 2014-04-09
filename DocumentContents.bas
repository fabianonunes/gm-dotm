Attribute VB_Name = "DocumentContents"

Sub italicsLatin()

    Dim fs As FileSystemObject
    Dim stream As TextStream
    Dim text As String
    Dim undo As UndoRecord
    
   On Error GoTo italicsLatin_Error

    Set fs = New FileSystemObject
    Set stream = fs.OpenTextFile(LATIN_FILEPATH)
    Set undo = Application.UndoRecord
    
    undo.EndCustomRecord
    undo.StartCustomRecord ("Destacar palavras em latim")
    
    Do While Not stream.AtEndOfStream
        
        text = stream.ReadLine
        
        Dim oRng As Word.Range
        Set oRng = ActiveDocument.Range
        With oRng.Find
            .ClearFormatting
            .MatchWholeWord = True
            .text = text
            While .Execute
                With oRng
                Str
                    If .Style Like "Transcrição*" Then
                       .text = text
                       .Font.Italic = True
                    End If
                End With
            Wend
        End With
    
    Loop
    
    undo.EndCustomRecord
 
    stream.Close
    
    Application.ScreenUpdating = True

   On Error GoTo 0
   Exit Sub

italicsLatin_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure italicsLatin of Módulo DocumentContents"

End Sub

Sub comment()
         
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

