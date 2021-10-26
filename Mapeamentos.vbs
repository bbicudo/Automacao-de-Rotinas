'////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'/////////////////////////////////////////////////////// ESCRITO POR: BRUNO BRANCO BICUDO ///////////////////////////////////////////////////////
'////////////////////////////////////////////////////////////////// 12/01/2016 //////////////////////////////////////////////////////////////////
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

'################################################################################################################################################
'Este script cria os mapeamentos baseado nos grupos onde o usuário está incluído,
'visando esconder o nome do caminho do servidor nos compartilhamentos.

'Execução: 
'A estrutura de pastas deve estar criada e com as devidas permissões setadas. O usuário deve estar incluso no grupo ao qual será mapeada a pasta
'Mapeia as pastas no login, tornando o mapeamento um processo descentralizado, tendo o usuário acesso à pasta de seu setor em qualquer 
'máquina do parque operacional.

'Os mapeamentos nas letras S e P são mapeamentos presentes para todos os usuários no domínio
'Os mapeamentos na letra Q são mapeamentos presentes para usuários em grupos especiais ou que tem acesso à pastas em outro servidor

'################################################################################################################################################

'Continua a execução do Script mesmo que ocorram falhas
On Error Resume Next

'Referenciando (instanciando) objetos	
set objNetwork= CreateObject("WScript.Network")
strDom = objNetwork.UserDomain
strUser = objNetwork.UserName
strServer = "\\SERVIDOR"
strServer2 = "\\SERVIDOR2"
Set objUser = GetObject("WinNT://" & strDom & "/" & strUser &  ",user")
Set objShell = CreateObject("Shell.Application")

objNetwork.RemoveNetworkDrive "B:", True, True
objNetwork.RemoveNetworkDrive "O:", True, True
objNetwork.RemoveNetworkDrive "R:", True, True
objNetwork.RemoveNetworkDrive "T:", True, True
objNetwork.RemoveNetworkDrive "U:", True, True
objNetwork.RemoveNetworkDrive "V:", True, True

	'Remove o mapeamento anterior por segurança
	objNetwork.RemoveNetworkDrive "S:", True, True

	'Cria o mapemento na letra S:, referente à pasta shareName$ no servidor strServer
	objNetwork.MapNetworkDrive "S:", strServer+"shareName$"

	'Renomeia o compartilhamento, tornando-o simplesmente Nome e removendo o caminho do servidor
	objShell.NameSpace("S:").Self.Name = "Nome"
			
	objNetwork.RemoveNetworkDrive "P:", True, True
	objNetwork.MapNetworkDrive "P:", strServer+"shareName2$"
	objShell.NameSpace("P:").Self.Name = "Nome2"

        'Cria um array da coluna Grupos para o usuario que esta executando o script
	For Each objGroup In objUser.Groups
		'WScript.Echo objGroup.Name 
   	 
	'Seleciona o grupo do usuário
	Select Case objGroup.Name
				 
		'Se o usuário estiver no grupo, inicia os mapeamentos.		
		'#################### Q: ##################################################
		Case "Grupo1" 
			objNetwork.RemoveNetworkDrive "Q:", True, True
			objNetwork.MapNetworkDrive "Q:", strServer2+"pasta1$"	
			objShell.NameSpace("Q:").Self.Name = "Pasta1"
			
		Case "Grupo2" 
			objNetwork.RemoveNetworkDrive "Q:", True, True
			objNetwork.MapNetworkDrive "Q:", strServer2+"pasta2$"	
			objShell.NameSpace("Q:").Self.Name = "Nome2"
			
		Case "Grupo3" 
			objNetwork.RemoveNetworkDrive "Q:", True, True
			objNetwork.MapNetworkDrive "Q:", strServer2+"pasta3$"
			objShell.NameSpace("Q:").Self.Name = "Nome3"
End Select
	Next	
wscript.Quit
