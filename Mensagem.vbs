On Error Resume Next
Dim objSysInfo, objUser

Set objSysInfo = CreateObject("ADSystemInfo")
Set WshShell = WScript.CreateObject("WScript.Shell")
Set MSIE = CreateObject("InternetExplorer.Application")

' Lista as propriedades do usu�rio conforme o ADUC
Set objUser = GetObject("LDAP://" & objSysInfo.UserName)


strTitulo   =   "Mensagem de ..."	
strMensagem =	"   Ol� " & objUser.givenName & "."  & chr(10) & chr(10) & _
				"	Mensagem" & chr(10) & _
				
WshShell.popup strMensagem, 240, strTitulo
