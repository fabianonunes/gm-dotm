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
    Dim tempBarcode  As String
    Dim formData     As String
    Dim Id           As Identifier
        
    Helpers.waitApplication
    
    Id = ParseIdentifier(ActiveDocument.name, True)
    
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
        
        If CARIMBO_TIPO = "ATENÇÃO_MINISTRO" Then
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
    
    If Id.Formatado <> "" Then
    
        tempBarcode = CreateTempFile("car")
        
        ' o arquivo deve terminar em .pdf para o acrojs carimbar
        Name tempBarcode As tempBarcode & ".pdf"
        tempBarcode = tempBarcode & ".pdf"
    
        formData = spartanEncode128C(Id.Padded)
        
        Set formDoc = New Acrobat.AcroPDDoc
        formDoc.Open CARIMBOS_PATH & "BARCODE.pdf"
        
        Set jsFormObj = formDoc.GetJSObject
        jsFormObj.getField("barcode").Value = formData
        jsFormObj.flattenPages
        
        formDoc.Save PDSaveFull, tempBarcode
        
        tempFileForm = toAcroPath(tempBarcode)
        jsObj.addWatermarkFromFile tempBarcode, 0, 0
    
    End If
    
    pdDoc.OpenAVDoc ActiveDocument.name
    pdDoc.ClearFlags PDDocNeedsSave
    
finally: On Error Resume Next
    
    Helpers.resumeApplication
    pdDoc.Close
    
    Kill tempFile
    Kill tempFileForm
    Kill tempBarcode
    
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


Private Function spartanEncode128C(code As String)

   ' http://en.wikipedia.org/wiki/Code_128
   ' http://grandzebu.net/informatique/codbar-en/code128.htm
   
    Dim c     As String
    Dim check As Integer
    Dim pair  As Integer
    Dim i     As Integer
    Dim regex As RegExp
   
    c = Chr(210)
    check = 105
    
    Set regex = New RegExp
    regex.Pattern = "[^0-9]"
    regex.Global = True
    
    code = regex.Replace(code, "") ' this is sparta
    
    If Len(code) Mod 2 > 0 Then
       code = "0" & code ' this is spaaaaarta
    End If
    
    For i = 1 To Len(code) Step 2
       pair = 0 + Mid(code, i, 2)
       check = check + pair * ((i - 1) / 2 + 1)
       c = c & Chr(pair + IIf(pair < 95, 32, 105))
    Next
    
    check = check Mod 103
    c = c & Chr(check + IIf(check < 95, 32, 105)) & Chr(211)
    
    spartanEncode128C = c

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
    
    AutoExec.uiRibbon.InvalidateControl "ATENÇÃO_MINISTRO"
    AutoExec.uiRibbon.InvalidateControl "MATÉRIA_COMUM"
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

