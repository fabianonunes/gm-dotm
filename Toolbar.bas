Attribute VB_Name = "Toolbar"
Option Explicit

Sub JoinLines()

On Error GoTo try
   
    Dim selBkUp As Range
    
    Helpers.waitApplication
    
    Set selBkUp = ActiveDocument.Range(Selection.Range.Start, Selection.Range.End - 1)
    
    With selBkUp.Find
        .ClearFormatting
        .Replacement.ClearFormatting
        .Forward = True
        .Wrap = wdFindStop
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchAllWordForms = False
        .MatchSoundsLike = False
        
        .MatchWildcards = False
        
        .text = "^w"
        .Replacement.text = " "
        .Execute Replace:=wdReplaceAll
        
        .MatchWildcards = True
                
        .text = " {1;}^13"
        .Replacement.text = "^p"
        .Execute Replace:=wdReplaceAll
        
        .text = "([!.])^13"
        .Replacement.text = "\1 "
        .Execute Replace:=wdReplaceAll
        
        .text = " {1;}"
        .Replacement.text = " "
        .Execute Replace:=wdReplaceAll
       
    End With
    
finally: On Error Resume Next
    Helpers.resumeApplication
    Exit Sub

try:
    If Err.Number = 4608 Then
        MsgBox "Não há texto selecionado"
    Else
        Catch Err
    End If
    
    Resume finally
    Resume
        
End Sub

Public Sub loadStyles()
   
On Error GoTo try

    ActiveDocument.ApplyQuickStyleSet2 "GMJD"
    AutoExec.uiRibbon.ActivateTabMso "TabHome"

finally: On Error Resume Next
    Exit Sub

try: Catch Err
    Resume finally
    Resume

End Sub

Public Function loadStylesCallback(control As IRibbonControl)
    loadStyles
End Function

Public Function joinLinesCallback(control As IRibbonControl)
    JoinLines
End Function
