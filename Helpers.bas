Attribute VB_Name = "Helpers"
Option Explicit

Private Declare Function ShellExecute _
  Lib "shell32.dll" Alias "ShellExecuteA" ( _
  ByVal hWnd As Long, _
  ByVal Operation As String, _
  ByVal filename As String, _
  Optional ByVal Parameters As String, _
  Optional ByVal Directory As String, _
  Optional ByVal WindowStyle As Long = vbMinimizedFocus _
  ) As Long
  
  
Public Type Identifier
    Numero As String
    Digito As String
    Ano As String
    Justica As String
    Tribunal As String
    Vara As String
    Formatado As String
End Type

Public Function ParseIdentifier(text As String) As Identifier

    Dim mask As New RegExp, result As MatchCollection
    Dim firstMatch As Match
    
   On Error GoTo ParseIdentifier_Error

    mask.Global = True
    mask.IgnoreCase = True
    mask.Pattern = PROCESSO_PATTERN
      
    Set result = mask.Execute(text)
    
    If (result.Count > 0) Then
        Set firstMatch = result.Item(0)
        ParseIdentifier.Numero = firstMatch.SubMatches(0)
        ParseIdentifier.Digito = firstMatch.SubMatches(1)
        ParseIdentifier.Ano = firstMatch.SubMatches(2)
        ParseIdentifier.Justica = firstMatch.SubMatches(3)
        ParseIdentifier.Tribunal = firstMatch.SubMatches(4)
        ParseIdentifier.Vara = firstMatch.SubMatches(5)
        ParseIdentifier.Formatado = firstMatch.Value
    Else
        Err.Raise 600, "ParseIdentifier"
    End If

   On Error GoTo 0
   Exit Function

ParseIdentifier_Error:

    If Err.Number = 600 Then
        Err.Raise 600, "ParseIdentifier"
        Exit Function
    End If

    Catch Err
        
End Function

Public Function Navigate(URL As String)
    ShellExecute hWnd:=0, Operation:="open", filename:=URL, WindowStyle:=5
End Function

Public Function Explore(folder As String)
    ShellExecute hWnd:=0, Operation:="explore", filename:=folder, WindowStyle:=5
End Function

Public Function getPK(Id As Identifier)
    
    Dim URL As String
    Dim headers As String
    Dim request As New WinHttpRequest
    
    Dim retval(2) As String
        
    Dim mask As New RegExp, result As MatchCollection, firstMatch As Match
   On Error GoTo getPK_Error

    mask.Global = True
    mask.IgnoreCase = True
    mask.Pattern = "num_int=([0-9]*)&ano_int=([0-9]{4})"
     
    URL = _
        "http://ext02.tst.jus.br/pls/ap01/ap_proc100.dados_processos?num_proc=" & _
        Id.Numero & "&dig_proc=" & Id.Digito & "&ano_proc=" & Id.Ano & _
        "&num_orgao=" & Id.Justica & "&TRT_proc=" & Id.Tribunal & "&vara_proc=" & _
        Id.Vara
        
    request.Open "GET", URL
    request.Option(WinHttpRequestOption_EnableRedirects) = False
    request.Send
        
    headers = request.GetAllResponseHeaders()
        
    Set result = mask.Execute(headers)
        
    If (result.Count > 0) Then
        
        Set firstMatch = result.Item(0)
        retval(0) = firstMatch.SubMatches(0)
        retval(1) = firstMatch.SubMatches(1)
        getPK = retval
            
    End If
       

   On Error GoTo 0
   Exit Function

getPK_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & _
        ") in procedure getPK of Módulo Helpers"
    
End Function

Public Function openAll(Id As Identifier)

    Dim pk
   
   On Error GoTo openAll_Error

    pk = getPK(Id)
    
    Navigate _
        ("https://aplicacao6.tst.jus.br/esij/VisualizarPecas.do?visualizarTodos=1&anoProcInt=" _
        & pk(1) & "&numProcInt=" & pk(0))

   On Error GoTo 0
    Exit Function

openAll_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & _
        ") in procedure openAll of Módulo Helpers"

End Function


Sub qq()
    
    Dim clipboard As DataObject
    Dim text As String
    Dim mask As RegExp
    Dim result As MatchCollection
    Dim mt As Match
    Dim undo As UndoRecord
        
   On Error GoTo qq_Error

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

   On Error GoTo 0
   Exit Sub

qq_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure qq of Módulo Helpers"
    
End Sub

Sub WaitFor(NumOfSeconds As Long)
    
    Dim SngSec As Long
    SngSec = Timer + NumOfSeconds

    Do While Timer < SngSec
        DoEvents
    Loop

End Sub

Public Function Catch(error As ErrObject)

    Application.ScreenUpdating = True
    
    If (error.Number = 600) Then
        MsgBox "O nome do arquivo não se parece com um processo"
    Else
        MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure ..."
    End If

End Function

Public Sub loadStyles()
   
   On Error GoTo loadStyles_Error

    ActiveDocument.ApplyQuickStyleSet2 "GMJD"

   On Error GoTo 0
   Exit Sub

loadStyles_Error:

    Catch Err

End Sub
