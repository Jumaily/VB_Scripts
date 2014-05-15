Set wshNetwork = CreateObject( "WScript.Network" )
On Error Resume Next

wshnetwork.AddWindowsPrinterConnection "\\print_server\Printer 1" 'Description of Printer 1
wshnetwork.AddWindowsPrinterConnection "\\print_server\Printer 2" 'Description of Printer 2
wshnetwork.AddWindowsPrinterConnection "\\print_server\Printer 3" 'Description of Printer 3

On Error Goto 0
Set wshNetwork = Nothing