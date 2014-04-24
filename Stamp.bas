Attribute VB_Name = "Stamp"
Option Explicit

Private CARIMBO_TIPO As String
Private CARIMBO_CLASSE As String

Private Function Stamp()
   
On Error GoTo try
    
    Dim pdDoc        As Acrobat.AcroPDDoc
    Dim formDoc      As Acrobat.AcroPDDoc
    
    Dim jsObj        As Object
    Dim jsFormObj    As Object

    Dim tempFile     As String
    Dim tempFileForm As String
    Dim formData     As String
    
    Helpers.waitApplication
    
    tempFile = CreateTempFile("car")
    exportToPdf (tempFile)
        
    Set pdDoc = New Acrobat.AcroPDDoc
    pdDoc.Open tempFile
    
    Set jsObj = pdDoc.GetJSObject
    
    jsObj.addWatermarkFromFile toAcroPath(CARIMBOS_PATH) & "TIMBRE.pdf"
    
    If CARIMBO_CLASSE <> "" Then
        jsObj.addWatermarkFromFile toAcroPath(CARIMBOS_PATH) & CARIMBO_CLASSE & ".pdf", 0, 0
    End If
    
    If CARIMBO_TIPO <> "" Then
        
        jsObj.addWatermarkFromFile toAcroPath(CARIMBOS_PATH) & CARIMBO_TIPO & ".pdf", 0, 0
        
        If CARIMBO_TIPO = "ATEN��O_MINISTRO" Then
            formData = InputBox(prompt:="Alguma mensagem?")
        End If
        
    End If
    
    If formData <> "" Then
    
        tempFileForm = CreateTempFile("car")
        
        ' o arquivo deve terminar em .pdf para o acrojs carimbar
        Name tempFileForm As tempFileForm & ".pdf"
        tempFileForm = tempFileForm & ".pdf"
    
        formData = Replace(Replace(formData, ";", vbCrLf), vbCrLf & " ", vbCrLf)
        
        Set formDoc = New Acrobat.AcroPDDoc
        formDoc.Open CARIMBOS_PATH & "AM.pdf"
        
        Set jsFormObj = formDoc.GetJSObject
        jsFormObj.getField("AM").Value = formData
        jsFormObj.flattenPages
        
        formDoc.Save PDSaveFull, tempFileForm
        
        tempFileForm = toAcroPath(tempFileForm)
        jsObj.addWatermarkFromFile tempFileForm, 0, 0
    
    End If
    
    pdDoc.OpenAVDoc ActiveDocument.name
    pdDoc.ClearFlags PDDocNeedsSave
    
finally: On Error Resume Next
    
    Helpers.resumeApplication
    pdDoc.Close
    
    Kill tempFile
    Kill tempFileForm
    
    Set pdDoc = Nothing
    
    Exit Function

try: Catch Err
    Resume finally
    Resume
        
End Function

Private Function exportToPdf(tempFile As String)
    ActiveDocument.ExportAsFixedFormat OutputFileName:=tempFile, _
        ExportFormat:=wdExportFormatPDF, OpenAfterExport:=False, _
        OptimizeFor:=wdExportOptimizeForPrint, Range:=wdExportAllDocument, From:=1, _
        To:=1, Item:=wdExportDocumentContent, IncludeDocProps:=False, _
        KeepIRM:=False, CreateBookmarks:=wdExportCreateHeadingBookmarks, _
        DocStructureTags:=True, BitmapMissingFonts:=False, UseISO19005_1:=False
End Function

Private Function toAcroPath(path As String)
    toAcroPath = "/" & Replace(Replace(path, ":", ""), "\", "/")
End Function

Public Function stampCallback(control As IRibbonControl)
    Stamp
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

try: Catch Err, "tipoPressed"
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
    
    AutoExec.uiRibbon.InvalidateControl "ATEN��O_MINISTRO"
    AutoExec.uiRibbon.InvalidateControl "MAT�RIA_COMUM"
    AutoExec.uiRibbon.InvalidateControl "MODELO_ADAPTADO"

finally: On Error Resume Next
   Exit Function

try: Catch Err, "tipoAction"
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

try: Catch Err, "classePressed"
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

try: Catch Err, "classeAction"
    Resume finally
    Resume
    
End Function

