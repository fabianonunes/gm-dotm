Attribute VB_Name = "Latin"
Option Explicit

Sub italicsLatin()

    Dim fs      As FileSystemObject
    Dim stream  As TextStream
    Dim text    As String
    Dim undo    As UndoRecord
    Dim oRng    As Word.Range
    Dim counter As Integer
    
On Error GoTo try
   
    System.Cursor = wdCursorWait
    Application.ScreenUpdating = False

    Set fs = New FileSystemObject
    Set stream = fs.OpenTextFile(LATIN_FILEPATH)
    Set undo = Application.UndoRecord
    counter = 0
    
    undo.StartCustomRecord ("Destacar palavras em latim")
    
    Do While Not stream.AtEndOfStream
        
        text = stream.ReadLine
        
        Set oRng = ActiveDocument.Range
        With oRng.Find
            .ClearFormatting
            .MatchWholeWord = True
            .MatchCase = False
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
    
    
    If counter = 0 Then
        MsgBox "Nenhuma expressão foi encontrada."
    ElseIf counter = 1 Then
        MsgBox "Uma expressão foi destacada."
    ElseIf counter > 1 Then
        MsgBox counter & " expressões foram destacadas."
    End If

finally:
    On Error Resume Next
    Application.ScreenUpdating = True
    undo.EndCustomRecord
    stream.Close
    Exit Sub

try:
    Catch Err
    Resume finally
    Resume

End Sub

Public Function italicsLatinCallback(control As IRibbonControl)
    italicsLatin
End Function

