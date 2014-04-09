Attribute VB_Name = "AutoExec"
Option Explicit

Dim oAppClass As New ThisApplication

Public Sub AutoExec()
    
    On Error GoTo ErrorHandler:
    
    ' habilitar eventos do Word.Application
    ' http://word.mvps.org/faqs/macrosvba/appclassevents.htm
    Set oAppClass.oApp = Word.Application
    
    Exit Sub

ErrorHandler:
    Catch Err

End Sub
