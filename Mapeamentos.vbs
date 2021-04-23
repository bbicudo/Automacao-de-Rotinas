'////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'/////////////////////////////////////////////////////// ESCRITO POR: BRUNO BRANCO BICUDO ///////////////////////////////////////////////////////
'////////////////////////////////////////////////////////////////// 12/01/2016 //////////////////////////////////////////////////////////////////
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

'################################################################################################################################################
'Este script cria os mapeamentos baseado nos grupos onde o usu�rio est� inclu�do,
'visando esconder o nome do caminho do servidor nos compartilhamentos.

'Execu��o: 
'A estrutura de pastas deve estar criada e com as devidas permiss�es setadas. O usu�rio deve estar incluso no grupo ao qual ser� mapeada a pasta
'Mapeia as pastas no login, tornando o mapeamento um processo descentralizado, tendo o usu�rio acesso � pasta de seu setor em qualquer 
'm�quina do parque operacional.

'Os mapeamentos nas letras S e P s�o mapeamentos presentes para todos os usu�rios no dom�nio
'Os mapeamentos na letra Q s�o mapeamentso presentes para usu�rios em grupos esepciais que tem acesso � pastas em outro servidor

'################################################################################################################################################

'Continua a execu��o do Script mesmo que ocorram falhas
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

	'Remove o mapeamento anterior por seguran�a
	objNetwork.RemoveNetworkDrive "S:", True, True

	'Cria o mapemento na letra S:, referente � pasta shareName$ no servidor strServer
	objNetwork.MapNetworkDrive "S:", strServer+"shareName$"

	'Renomeia o compartilhamento, tornando-o simplesmente Nome e removendo o caminho do servidor
	objShell.NameSpace("S:").Self.Name = "Nome"
			
	objNetwork.RemoveNetworkDrive "P:", True, True
	objNetwork.MapNetworkDrive "P:", strServer+"shareName2$"
	objShell.NameSpace("P:").Self.Name = "Nome2"

        'Cria um array da coluna Grupos para o usuario que esta executando o script
	For Each objGroup In objUser.Groups
		'WScript.Echo objGroup.Name 
   	 
	'Seleciona o grupo do usu�rio
	Select Case objGroup.Name
				 
		'Se o usu�rio estiver no grupo, inicia os mapeamentos.		
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