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

Public Function ParseIdentifier(text As String, ByRef Id As Identifier) As Boolean

    Dim mask As New RegExp, result As MatchCollection
    Dim firstMatch As Match
    
    mask.Global = True
    mask.IgnoreCase = True
    mask.Pattern = "([1-9][0-9]{0,6})-([0-9]{1,2})[-.]([0-9]{4})[-.]([0-9])[-.]([0-9]{2})[-.]([0-9]{4})"
      
    Set result = mask.Execute(text)
    
    ParseIdentifier = False
    
    If (result.Count > 0) Then
        Set firstMatch = result.Item(0)
        Id.Numero = firstMatch.SubMatches(0)
        Id.Digito = firstMatch.SubMatches(1)
        Id.Ano = firstMatch.SubMatches(2)
        Id.Justica = firstMatch.SubMatches(3)
        Id.Tribunal = firstMatch.SubMatches(4)
        Id.Vara = firstMatch.SubMatches(5)
        Id.Formatado = firstMatch.Value
        ParseIdentifier = True
    End If
        
End Function

Public Function Navigate(URL As String)
    ShellExecute hWnd:=0, Operation:="open", filename:=URL, WindowStyle:=5
End Function

Public Function Explore(folder As String)
    ShellExecute hWnd:=0, Operation:="explore", filename:=folder, WindowStyle:=5
End Function

Public Function getPK(Id As Identifier)
    
    Dim processo As String
    Dim URL As String
    Dim headers As String
    Dim request As New WinHttpRequest
    
    Dim retval(2) As String
        
    Dim mask As New RegExp, result As MatchCollection, firstMatch As Match
    mask.Global = True
    mask.IgnoreCase = True
    mask.Pattern = "num_int=([0-9]*)&ano_int=([0-9]{4})"
     
    URL = "http://ext02.tst.jus.br/pls/ap01/ap_proc100.dados_processos?num_proc=" & Id.Numero _
    & "&dig_proc=" & Id.Digito & "&ano_proc=" & Id.Ano & "&num_orgao=" & Id.Justica _
    & "&TRT_proc=" & Id.Tribunal & "&vara_proc=" & Id.Vara
        
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

Public Function openAll(Id As Identifier)

          Dim pk
    pk = getPK(Id)
    
    Navigate ("https://aplicacao6.tst.jus.br/esij/VisualizarPecas.do?visualizarTodos=1&anoProcInt=" & pk(1) & "&numProcInt=" & pk(0))

End Function
    

'    Dim clipboard As DataObject
'    Set clipboard = New DataObject
'    clipboard.SetText "A string value"
'    clipboard.PutInClipboard


