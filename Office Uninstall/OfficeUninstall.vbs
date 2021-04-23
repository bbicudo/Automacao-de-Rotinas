'////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'/////////////////////////////////////////////////////// ESCRITO POR: BRUNO BRANCO BICUDO ///////////////////////////////////////////////////////
'////////////////////////////////////////////////////////////////// 24/05/2016 //////////////////////////////////////////////////////////////////
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

'####################################################################VERSÃO 1.0##################################################################

'Este script aplica-se às versões do Office 2003, 2007, 2010, 2013, 2013 Click to Run, 2016 e 2016 Click to Run.
'Nas versões Click to Run, deve-se utilizar também o script para usuários, que remove os atalhos
'Nas versões Click to Run, o programa continua listado em "Add e Remover Programas", será trabalhado nas próximas versões

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
strLogFile = "0FFICE.log"

'Se o arquivo "0FF1C3.log" existir, termina o script
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

'Verifica qual a versão do Office instalada
If (objFileFolder.FolderExists(strSystem32 & "\Microsoft Office\Office11") OR objFileFolder.FolderExists(strSystem64 & "\Microsoft Office\Office11")) Then
		
		strOffice = "Office 2003"

	Else If(objFileFolder.FolderExists(strSystem32 & "\Microsoft Office\Office12") OR objFileFolder.FolderExists(strSystem64 & "\Microsoft Office\Office12")) Then
		
		strOffice = "Office 2007"
	
	Else If(objFileFolder.FolderExists(strSystem32 & "\Microsoft Office\Office14") OR objFileFolder.FolderExists(strSystem64 & "\Microsoft Office\Office14")) Then
		
		strOffice = "Office 2010"
			
	Else If (objFileFolder.FolderExists(strSystem32 & "\Microsoft Office\Office15") OR objFileFolder.FolderExists(strSystem64 & "\Microsoft Office\Office15") OR objFileFolder.FolderExists(strSystem32 & "\Microsoft Office\Office16") OR objFileFolder.FolderExists(strSystem64 & "\Microsoft Office\Office16") OR (strSystem32 & "\Common Files\Microsoft Shared\Office15")OR objFileFolder.FolderExists(strSystem64 & "\Common Files\Microsoft Shared\Office15")) Then
		
		strOffice = "Office 2016"
		
	Else
	
		'Wscript.Echo "Office Não Encontrado!!"
		
	End If
	End If
	End If
End If	

'Wscript.Echo strOffice

'Executa a desinstalação de acordo com a versão encontrada
Select Case strOffice

	Case "Office 2003"
	
	Set colSoftware = objWMIService.ExecQuery("Select * from Win32_Product Where Name LIKE 'Microsoft%Office%2003'") 
	
	For Each objSoftware in colSoftware 	
		objSoftware.Uninstall() 
	Next 

	objFileFolder.CreateTextFile(strFolder&strSubFolder&strLogFile)
	
	Case "Office 2007"
	
	objNetwork.RemoveNetworkDrive "X:", True, True
	objNetwork.MapNetworkDrive "X:", "\\addc\scripts$\Scripts GPO\Office Uninstall"
	objShell.Run "cmd /C copy X:\uninstall2007.xml C:\Admin\ /Y /Z > C:\Admin\LogOffice.txt",0,True
	objNetwork.RemoveNetworkDrive "X:", True, True
	
	objShell = Nothing
	
	If(objFileFolder.FolderExists(strSystem32)) Then 
		objShell.CurrentDirectory="C:\Program Files (x86)\Common Files\Microsoft Shared\OFFICE12\Office Setup Controller"
	Else
		objShell.CurrentDirectory="C:\Program Files\Common Files\Microsoft Shared\OFFICE12\Office Setup Controller"
	End If
	
	objShell.Run "cmd /C setup.exe /uninstall Enterprise /dll OSETUP.DLL /config C:\Admin\uninstall2007.xml >> C:\Admin\LogOffice.txt",0,True
	
	objFileFolder.CreateTextFile(strFolder&strSubFolder&strLogFile)
	
	Case "Office 2010"
	
	objNetwork.RemoveNetworkDrive "X:", True, True
	objNetwork.MapNetworkDrive "X:", "\\addc\scripts$\Scripts GPO\Office Uninstall"
	objShell.Run "cmd /C copy X:\uninstall2010.xml C:\Admin\ /Y /Z > C:\Admin\LogOffice.txt",0,True
	objNetwork.RemoveNetworkDrive "X:", True, True
	
	objShell = Nothing
	
	objShell.CurrentDirectory = "C:\Program Files\Common Files\Microsoft Shared\OFFICE14\Office Setup Controller"

	objShell.Run "cmd /C setup.exe /uninstall PROPLUS /dll OSETUP.DLL /config C:\Admin\uninstall2010.xml >> C:\Admin\LogOffice.txt",0,True
	
	objFileFolder.CreateTextFile(strFolder&strSubFolder&strLogFile)
	
	Case "Office 2016"
	
	objNetwork.RemoveNetworkDrive "X:", True, True
	objNetwork.MapNetworkDrive "X:", "\\addc\scripts$\Scripts GPO\Office Uninstall"
	objShell.Run "cmd /C copy X:\uninstall2013.xml C:\Admin\ /Y /Z > C:\Admin\LogOffice.txt",0,True
	bjNetwork.RemoveNetworkDrive "X:", True, True
	
	objShell = Nothing
	
	objShell.CurrentDirectory = "C:\Program Files\Common Files\Microsoft Shared\OFFICE15\Office Setup Controller"
	
	objShell.Run "cmd /C setup.exe /uninstall PROPLUS /dll OSETUP.DLL /config C:\Admin\uninstall2013.xml >> C:\Admin\LogOffice.txt",0,True
	
	'Fechar os Aplicativos e Matar os Processos
	Dim strNomeProcesso
	strNomeProcesso = Array("winword.exe","excel.exe","powerpnt.exe","onenote.exe","outlook.exe","mspub.exe","msaccess.exe","infopath.exe","groove.exe","lync.exe","officeclicktorun.exe","officeondemand.exe","officec2rclient.exe","appvshnotify.exe","firstrun.exe","setup.exe","integratedoffice.exe","integrator.exe","communicator.exe","msosync.exe","onenotem.exe","iexplore.exe","mavinject32.exe","werfault.exe","perfboost.exe","roamingoffice.exe","msiexec.exe","ose.exe")
	
	For Each Processo In strNomeProcesso
		objShell.Run "cmd /C taskkill /f /im " & Processo & " /t >> C:\Admin\LogOffice.txt",0,True
	Next
	
	'Remover Arquivos do Office
	Dim strComando
	strComando = Array("takeown /f ""%PROGRAMFILES%\Microsof Office 15""","takeown /f ""%PROGRAMFILES%\Microsof Office\Office15\Root""","takeown /f ""%PROGRAMFILES%\Microsof Office\Office16\Root""","attrib -r -s -h /s /d  ""%PROGRAMFILES%\Microsoft Office 15""","attrib -r -s -h /s /d  ""%PROGRAMFILES%\Microsoft Office\Office15\Root""","attrib -r -s -h /s /d  ""%PROGRAMFILES%\Microsoft Office\Office16\Root""","rmdir /s /q  ""%PROGRAMFILES%\Microsoft Office 15""","rmdir /s /q  ""%PROGRAMFILES%\Microsoft Office\Office15\Root""","rmdir /s /q  ""%PROGRAMFILES%\Microsoft Office\Office16\Root""","takeown /f  ""%PROGRAMFILES%\Microsoft Office""","attrib -r -s -h /s /d  ""%PROGRAMFILES%\Microsoft Office""", "rmdir /s /q  ""%PROGRAMFILES%\Microsoft Office""","takeown /f  ""%PROGRAMFILES(x86)%\Microsoft Office""","attrib -r -s -h /s /d  ""%PROGRAMFILES(x86)%\Microsoft Office""","rmdir /s /q ""%PROGRAMFILES(x86)%\Microsoft Office""")
	For Each Comando In strComando
		objShell.Run "cmd /C " & Comando &" >> C:\Admin\LogOffice.txt",0,True
	Next
	
	'Remover Tarefas Agendadas
	Dim strTarefas
	strTarefas = Array("""FF_INTEGRATEDstreamSchedule""","""FF_INTEGRATEDUPDATEDETECTION""","""C2RAppVLoggingStart""","""Office 15 Subscription Heartbeat""","""\Microsoft\Office\Office 15 Subscription Heartbeat""","""Microsoft Office 15 Sync Maintenance for {d068b555-9700-40b8-992c-f866287b06c1}""","""\Microsoft\Office\OfficeInventoryAgentFallBack""","""\Microsoft\Office\OfficeTelemetryAgentFallBack2016""","""\Microsoft\Office\OfficeInventoryAgentLogOn2016""","""\Microsoft\Office\OfficeTelemetryAgentLogOn2016""","""Office Background Streaming""","""\Microsoft\Office\Office Automatic Updates""","""\Microsoft\Office\Office ClickToRun Service Monitor""","""Office Subscription Maintenance""","""\Microsoft\Office\Office Subscription Maintenance""")
	For Each Tarefa In strTarefas
		objShell.Run "cmd /C schtasks.exe /delete /tn " & Tarefa & " /f >> C:\Admin\LogOffice.txt",0,True
	Next
	
	'Remover o serviço do Click To Run
	objShell.Run "cmd /C sc delete Clicktorunsvc >> C:\Admin\LogOffice.txt",0,True
	
	'Deletar Chaves do registro
	Dim strRegistro
	strRegistro = Array("HKCU\Software\Microsoft\Office","HKLM\SOFTWARE\Microsoft\Office")
	For Each Registro In strRegistro
		objShell.Run "cmd /C reg delete  " & Registro & " /f >> C:\Admin\LogOffice.txt",0,True
	Next	
	
	'Remover Atalhos
	Dim strAtalhos
	strAtalhos = Array("takeown /f  ""%PROGRAMDATA%\Microsoft\Windows\Start Menu\Programs\Ferramentas do Microsoft Office 2016""","attrib -r -s -h /s /d  ""%PROGRAMDATA%\Microsoft\Windows\Start Menu\Programs\Ferramentas do Microsoft Office 2016""","rmdir /s /q  ""%PROGRAMDATA%\Microsoft\Windows\Start Menu\Programs\Ferramentas do Microsoft Office 2016""","del /F /Q ""%PROGRAMDATA%\Microsoft\Windows\Start Menu\Programs\Access 2016.lnk""","del /F /Q ""%PROGRAMDATA%\Microsoft\Windows\Start Menu\Programs\Excel 2016.lnk""","del /F /Q ""%PROGRAMDATA%\Microsoft\Windows\Start Menu\Programs\OneNote 2016.lnk""","del /F /Q ""%PROGRAMDATA%\Microsoft\Windows\Start Menu\Programs\Outlook 2016.lnk""","del /F /Q ""%PROGRAMDATA%\Microsoft\Windows\Start Menu\Programs\PowerPoint 2016.lnk""","del /F /Q ""%PROGRAMDATA%\Microsoft\Windows\Start Menu\Programs\Publisher 2016.lnk""","del /F /Q ""%PROGRAMDATA%\Microsoft\Windows\Start Menu\Programs\Skype For Business 2016.lnk""","del /F /Q ""%PROGRAMDATA%\Microsoft\Windows\Start Menu\Programs\Word 2016.lnk""")
	For Each Atalho In strAtalhos
		objShell.Run "cmd /C " & Atalho & " >> C:\Admin\LogOffice.txt",0,True
	Next
	
	'Remover Arquivos Temporarios
	objShell.Run "cmd /C del /s /f /q %TEMP%\*.* >> C:\Admin\LogOffice.txt",0,True
	objShell.Run "cmd /C del /F /Q ""%PROGRAMFILES(x86)%\Common Files\Microsoft Shared\*.*"" >> C:\Admin\LogOffice.txt",0,True
	For Each objSoftware in colSoftware 	
			objSoftware.Uninstall() 
	Next 
	
	'Remove os Serviços do Ativador KMS
	Dim strServicos
	strServicos = Array("sc stop Service KMSELDI","sc delete Service KMSELDI","sc stop KMSServerService","sc delete KMSServerService")
	For Each Servico In strServicos
		objShell.Run "cmd /C " & strServico & " >> C:\Admin\LogOffice.txt",0,True
	Next
	
	objFileFolder.CreateTextFile(strFolder&strSubFolder&strLogFile)
	
End Select