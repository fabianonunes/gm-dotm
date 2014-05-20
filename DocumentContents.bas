Attribute VB_Name = "DocumentContents"
Option Explicit

Sub comment()
On Error GoTo try

    Dim excel_app   As Excel.Application
    Dim workbook    As Excel.workbook
    Dim sheet       As Excel.Worksheet
    Dim table       As Excel.ListObject
    Dim table_rng   As Excel.Range
    Dim doc_rng     As Range
    Dim undo        As UndoRecord
    Dim textToFind  As String
    
    waitApplication
    
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
        
        textToFind = table_rng.Cells(1).Value

        With doc_rng.Find
            .ClearFormatting
            
            If InStr(textToFind, "*") Then
                .Forward = True
                .MatchWholeWord = False
                textToFind = Replace(textToFind, "*", "")
            Else
                .MatchWholeWord = True
                .Forward = False
            End If
            
            .MatchCase = False
            .text = textToFind

            While .Execute

                With doc_rng

                    If (table_rng.Cells(3) = "" Or .Style = table_rng.Cells(3)) _
                        And (.ParagraphFormat.LeftIndent < 120) Then
                        
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
    
    resumeApplication
    undo.EndCustomRecord
    workbook.Close
    Set workbook = Nothing

    Exit Sub

try:
    Catch Err
    Resume finally
    Resume

End Sub

Function commentAction(control As IRibbonControl, pressed As Boolean)
On Error GoTo try

    If Not pressed Then
        Helpers.removeComments
        Exit Function
    End If
    
    DocumentContents.comment

finally: On Error Resume Next
    AutoExec.uiRibbon.InvalidateControl control.Id
    Exit Function

try: Catch Err, "commentAction"
    Resume finally
    Resume

End Function

Public Function commentPressed(control As IRibbonControl, ByRef toggleState)

On Error GoTo try

    toggleState = ActiveDocument.Comments.Count > 0

finally: On Error Resume Next
   Exit Function

try: Catch Err, "commentPressed"
    Resume finally
    Resume
End Function
