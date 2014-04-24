Attribute VB_Name = "Tools"
Option Explicit

Sub qq()
    
On Error GoTo try

    Dim clipboard As DataObject
    Dim text      As String
    Dim mask      As RegExp
    Dim result    As MatchCollection
    Dim mt        As Match
    Dim undo      As UndoRecord
    
    Set undo = Application.UndoRecord

    Set clipboard = New DataObject
    clipboard.GetFromClipboard
    
    If Not clipboard.GetFormat(1) Then
        Exit Sub
    End If
       
    text = clipboard.GetText(1)

    Set mask = New RegExp
    mask.Global = True
    mask.IgnoreCase = True
    mask.Pattern = PROCESSO_PATTERN
      
    Set result = mask.Execute(text)
    
    undo.StartCustomRecord ("qq")
    
    For Each mt In result
        ActiveDocument.Range.InsertBefore mt.Value & vbCrLf
    Next
    
    undo.EndCustomRecord
    
finally: On Error Resume Next
    Exit Sub

try: Catch Err
    Resume finally
    Resume
    
End Sub

