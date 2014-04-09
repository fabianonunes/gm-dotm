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

End Sub

Private Sub oApp_DocumentBeforeClose(ByVal Doc As Document, Cancel As Boolean)
    
    Dim oRng As Word.Range, i As Integer
   
   On Error GoTo oApp_DocumentBeforeSave_Error

    Set oRng = Doc.Range
    
    If oRng.Comments.Count = 0 Then
        Exit Sub
    End If
    
    Select Case _
        MsgBox("Ainda existem express�es n�o recomendadas no documento." & vbCrLf & "Deseja continuar e excluir os coment�rios?", _
        vbYesNo Or vbExclamation Or vbDefaultButton2, "Express�es")
    
        Case vbYes
            
            With oRng.Comments
                For i = .Count To 1 Step -1
                    .Item(i).Delete
                Next i
            End With
    
        Case vbNo
        
            Cancel = True
    
    End Select

   On Error GoTo 0
   Exit Sub

oApp_DocumentBeforeSave_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & _
        ") in procedure oApp_DocumentBeforeSave of M�dulo de classe ThisApplication"
End Sub

Private Sub oApp_DocumentOpen(ByVal Doc As Document)
   On Error GoTo oApp_DocumentOpen_Error

    Doc.ApplyQuickStyleSet2 "GMJD"

   On Error GoTo 0
   Exit Sub

oApp_DocumentOpen_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure oApp_DocumentOpen of M�dulo de classe ThisApplication"
End Sub

Private Sub oApp_Quit()
    'MsgBox "quit"
End Sub