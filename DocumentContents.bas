Attribute VB_Name = "DocumentContents"

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

Sub comment()
         
    Dim excel_app   As Excel.Application
    Dim workbook    As Excel.workbook
    Dim sheet       As Excel.Worksheet
    Dim table       As Excel.ListObject
    Dim table_rng   As Excel.Range
    Dim doc_rng     As Range
    Dim undo        As UndoRecord
    
On Error GoTo try
    
    SendKeys "%v%"
    DoEvents ' confirma a ativação do ribbon
   
    System.Cursor = wdCursorWait
    Application.ScreenUpdating = False
    
    AutoExec.AutoExec ' liga os eventos do oApp, caso necessário

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
            .MatchCase = False
            .text = table_rng.Cells(1).Value

            While .Execute

                With doc_rng

                    If (table_rng.Cells(3) = "" Or .Style = table_rng.Cells(3)) _
                        And (Not .Style Like "Transcrição*") Then
                        
                        .Comments.Add Range:=doc_rng, text:=table_rng.Cells(2).Value
                        
                    End If

                End With

            Wend

        End With

    Next

    Set doc_rng = ActiveDocument.Range
    
    If doc_rng.Comments.Count > 0 Then
        doc_rng.Comments.Item(1).Reference.Select
    Else
        MsgBox "Nenhuma expressão foi encontrada."
    End If
    

finally:
    On Error Resume Next
    
    Application.ScreenUpdating = True
    undo.EndCustomRecord
    workbook.Close
    Set workbook = Nothing

    Exit Sub

try:

    Catch Err
    Resume finally
    Resume

End Sub



