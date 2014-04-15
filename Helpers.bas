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

Private Declare Function GetTempPath Lib "kernel32" _
   Alias "GetTempPathA" (ByVal nBufferLength As Long, _
   ByVal lpBuffer As String) As Long

Private Declare Function GetTempFileName Lib "kernel32" _
   Alias "GetTempFileNameA" (ByVal lpszPath As String, _
   ByVal lpPrefixString As String, ByVal wUnique As Long, _
   ByVal lpTempFileName As String) As Long
  
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

End Function

Public Function Navigate(URL As String)
    ShellExecute hWnd:=0, Operation:="open", filename:=URL, WindowStyle:=5
End Function

Public Function Explore(folder As String)
    ShellExecute hWnd:=0, Operation:="explore", filename:=folder, WindowStyle:=5
End Function

Public Function getPK(Id As Identifier)
    
    Dim URL        As String
    Dim headers    As String
    Dim retval(2)  As String
    Dim result     As MatchCollection
    Dim firstMatch As Match
    
    Dim mask       As New RegExp
    Dim request    As New WinHttpRequest

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
       
End Function

Public Function WaitFor(NumOfSeconds As Long)
    
    Dim SngSec As Long
    SngSec = Timer + NumOfSeconds

    Do While Timer < SngSec
        DoEvents
    Loop

End Function

Public Function Catch(error As ErrObject, Optional name As String = "...")

    Application.ScreenUpdating = True
    
    If (error.Number = 600) Then
        MsgBox "O nome do arquivo não se parece com um processo"
    Else
        MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure " & name
    End If

End Function

Public Function removeComments()
   
    Dim oRng As Word.Range, i As Integer
   
    Set oRng = ActiveDocument.Range

    With oRng.Comments
      For i = .Count To 1 Step -1
           .Item(i).Delete
      Next i
    End With

End Function


Public Function CreateTempFile(sPrefix As String) As String

   Dim sTmpPath As String * 512
   Dim sTmpName As String * 576
   Dim nRet As Long

   nRet = GetTempPath(512, sTmpPath)
   If (nRet > 0 And nRet < 512) Then
      nRet = GetTempFileName(sTmpPath, sPrefix, 0, sTmpName)
      If nRet <> 0 Then
         CreateTempFile = Left$(sTmpName, _
            InStr(sTmpName, vbNullChar) - 1)
      End If
   End If
End Function



