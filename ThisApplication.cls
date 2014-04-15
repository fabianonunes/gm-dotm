VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisApplication"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public WithEvents oApp As Word.Application
Attribute oApp.VB_VarHelpID = -1

Private Sub oApp_DocumentBeforeSave(ByVal Doc As Document, SaveAsUI As Boolean, Cancel As Boolean)

On Error GoTo try

    Dim oRng As Word.Range
    Dim i As Integer
    
    Set oRng = Doc.Range
    
    If oRng.Comments.Count = 0 Then
        Exit Sub
    End If
    
    Select Case _
        MsgBox("Ainda existem expressões não recomendadas no documento." & vbCrLf & "Deseja continuar e excluir os comentários?", _
        vbYesNo Or vbExclamation Or vbDefaultButton2, "Expressões")
    
        Case vbYes
            
            With oRng.Comments
                For i = .Count To 1 Step -1
                    .Item(i).Delete
                Next i
            End With
            
            Doc.Save
            
        Case vbNo
        
            Cancel = True
    
    End Select

finally: On Error Resume Next 'ou [Goto 0]
   Exit Sub

try: Catch Err
    Resume finally
    Resume

End Sub

