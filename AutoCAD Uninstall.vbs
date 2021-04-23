'////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'/////////////////////////////////////////////////////// ESCRITO POR: BRUNO BRANCO BICUDO ///////////////////////////////////////////////////////
'////////////////////////////////////////////////////////////////// 15/07/2016 //////////////////////////////////////////////////////////////////
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

'####################################################################VERSÃO 1.0##################################################################



'################################################################################################################################################

'Declaração de Variáveis
Dim strSystem32, strSystem64, objWMIService, colSoftware, objFileFolder, objNetwork, objShell, strComputer

On Error Resume Next

strComputer = "." 

'Referenciando Objetos
Set objWMIService = GetObject("winmgmts:" & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2") 
Set	objFileFolder = CreateObject("Scripting.FileSystemObject")
Set objNetwork = CreateObject("WScript.Network")          
Set objShell = WScript.CreateObject ("WScript.Shell")

'Valores estáticos de variáveis de checagem e Log
strSystem32 = objShell.ExpandEnvironmentStrings("%ProgramFiles(x86)%")
strSystem64 = objShell.ExpandEnvironmentStrings("%ProgramFiles%")
strOffice = ""
strFolder = "C:\\Admin\\"
strSubFolder = "Logs\\"
strLogFile = "4u70C4D.log"

'Se o arquivo "Aut0C4D.log" existir, termina o script
If(objFileFolder.FileExists(strFolder&strSubFolder&strLogFile)) Then

	'Wscritp.Echo "Ja foi amigao!"
	Wscript.Quit
	
End If	

'Se a estrutura c:\admin não existir, cria a estrutura
If Not(objFileFolder.FolderExists(strFolder))Then

	objFileFolder.CreateFolder(strFolder)
	objFileFolder.CreateFolder(strFolder&strSubFolder)
	
	Else If Not(objFileFolder.FolderExists(strFolder&strSubFolder)) Then
	
		objFileFolder.CreateFolder(strFolder&strSubFolder)
	
	End If
End If

'Executa a desinstalação
Dim strTarefas
	strTarefas = Array("%Auto%CAD%","%Autodesk%")
	
	For Each Tarefa in strTarefas	
	'Wscript.Echo "Select * from Win32_Product Where Name LIKE '"&Tarefa&"'"
		Set colSoftware = objWMIService.ExecQuery("Select * from Win32_Product Where Name LIKE '"&Tarefa&"'") 
		
			For Each objSoftware in colSoftware 	
				objSoftware.Uninstall() 
			Next
	Next		
	
	objFileFolder.CreateTextFile(strFolder&strSubFolder&strLogFile)
	
	