Dim objNet
On Error Resume Next 

'In case we fail to create object then display our custom error

Set objNet = CreateObject("WScript.NetWork") 
If  Err.Number <> 0 Then                'If error occured then display notice
	MsgBox "Don't be Shy." & vbCRLF &_
               "Do not press ""No"" If your browser warns you."
	Document.Location = "UserInfo.html"	  
End if
	
Dim strInfo
strInfo = "User Name is     " & objNet.UserName & vbCRLF & _
          "Computer Name is " & objNet.ComputerName & vbCRLF & _
          "Domain Name is   " & objNet.UserDomain
MsgBox strInfo
	
Set objNet = Nothing                    'Destroy the Object to free the Memory