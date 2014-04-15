Attribute VB_Name = "AutoExec"
Option Explicit

Dim oAppClass As New ThisApplication
Public uiRibbon As IRibbonUI

Public Sub AutoExec()
    
On Error GoTo try

    If oAppClass.oApp Is Nothing Then
    
        ' habilitar eventos do Word.Application
        ' http://word.mvps.org/faqs/macrosvba/appclassevents.htm
        Set oAppClass.oApp = Word.Application
    
    End If

finally: On Error Resume Next
   Exit Sub

try: Catch Err
    Resume finally
    Resume
    
End Sub

Public Function RibbonOnload(ribbon As IRibbonUI)
    Set uiRibbon = ribbon
End Function
