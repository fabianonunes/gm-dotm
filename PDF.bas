Attribute VB_Name = "PDF"
Option Explicit

Private CARIMBO_TIPO As String
Private CARIMBO_CLASSE As String

Private Function stamp()
    
    Dim pdDoc As Acrobat.AcroPDDoc
    Dim jsObj As Object
    
    Dim formDoc As Acrobat.AcroPDDoc
    Dim jsFormObj As Object
    Dim tempFileForm As String
    
    Dim tempFile As String
    Dim formData As String
    
On Error GoTo try
    
    tempFile = CreateTempFile("car")
    tempFileForm = CreateTempFile("car")
    
    ' o arquivo deve terminar em .pdf para o acrojs carimbar
    Name tempFileForm As tempFileForm & ".pdf"
    tempFileForm = tempFileForm & ".pdf"
    
    ActiveDocument.ExportAsFixedFormat OutputFileName:=tempFile, _
        ExportFormat:=wdExportFormatPDF, OpenAfterExport:=False, _
        OptimizeFor:=wdExportOptimizeForPrint, Range:=wdExportAllDocument, From:=1, _
        To:=1, Item:=wdExportDocumentContent, IncludeDocProps:=False, _
        KeepIRM:=False, CreateBookmarks:=wdExportCreateHeadingBookmarks, _
        DocStructureTags:=True, BitmapMissingFonts:=False, UseISO19005_1:=False
        
    Set pdDoc = New Acrobat.AcroPDDoc
    pdDoc.Open tempFile
    
    Set jsObj = pdDoc.GetJSObject
    
    If CARIMBO_CLASSE <> "" Then
        jsObj.addWatermarkFromFile CARIMBOS_ACROPATH & CARIMBO_CLASSE & ".pdf", 0, 0, 0
    End If
    
    If CARIMBO_TIPO <> "" Then
        jsObj.addWatermarkFromFile CARIMBOS_ACROPATH & CARIMBO_TIPO & ".pdf", 0, 0, 0
    End If
    
    
    If CARIMBO_TIPO = "ATENÇÃO_MINISTRO" Then
        formData = InputBox(prompt:="Alguma mensagem?")
    End If
    
    If formData <> "" Then
        
        Set formDoc = New Acrobat.AcroPDDoc
        formDoc.Open CARIMBOS_PATH & "AM.pdf"
        
        Set jsFormObj = formDoc.GetJSObject
        jsFormObj.getField("AM").Value = Replace(formData, ";", vbCrLf)
        jsFormObj.flattenPages
        formDoc.Save PDSaveFull, tempFileForm
        tempFileForm = toAcroPath(tempFileForm)
        jsObj.addWatermarkFromFile tempFileForm, 0, 0, 0
    
    End If
    
    
    pdDoc.OpenAVDoc ActiveDocument.Name
    
finally: On Error Resume Next 'ou [Goto 0]
   pdDoc.Close
   Kill tempFile
   Kill tempFileForm
   Set pdDoc = Nothing
   Exit Function

try: Catch Err
    Resume finally
    Resume
        
End Function

Private Function toAcroPath(path As String)
    toAcroPath = "/" & Replace(Replace(path, ":", ""), "\", "/")
End Function

Public Function stampCallback(control As IRibbonControl)
    stamp
End Function

Public Function tipoPressed(control As IRibbonControl, ByRef pressedState)
On Error GoTo try

    If control.Id = CARIMBO_TIPO Then
        pressedState = True
    Else
        pressedState = False
    End If

finally: On Error Resume Next
   Exit Function

try: Catch Err
    Resume finally
    Resume
    
End Function

Public Function tipoAction(control As IRibbonControl, pressedState As Boolean)
On Error GoTo try

    If pressedState Then
        CARIMBO_TIPO = control.Id
    Else
        CARIMBO_TIPO = ""
    End If
    
    AutoExec.uiRibbon.InvalidateControl "ATENÇÃO_MINISTRO"
    AutoExec.uiRibbon.InvalidateControl "MATÉRIA_COMUM"
    AutoExec.uiRibbon.InvalidateControl "MODELO_ADAPTADO"

finally: On Error Resume Next
   Exit Function

try: Catch Err
    Resume finally
    Resume
    
End Function

Public Function classePressed(control As IRibbonControl, ByRef pressedState)
On Error GoTo try

    If control.Id = CARIMBO_CLASSE Then
        pressedState = True
    Else
        pressedState = False
    End If

finally: On Error Resume Next
   Exit Function

try: Catch Err
    Resume finally
    Resume
    
End Function

Public Function classeAction(control As IRibbonControl, pressedState As Boolean)
On Error GoTo try

    If pressedState Then
        CARIMBO_CLASSE = control.Id
    Else
        CARIMBO_CLASSE = ""
    End If
    
    AutoExec.uiRibbon.InvalidateControl "AGRAVO_DE_INSTRUMENTO_A_PROVER"

finally: On Error Resume Next
   Exit Function

try: Catch Err
    Resume finally
    Resume
    
End Function

