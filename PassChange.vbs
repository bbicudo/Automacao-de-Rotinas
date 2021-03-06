'////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'/////////////////////////////////////////////////////// ESCRITO POR: BRUNO BRANCO BICUDO ///////////////////////////////////////////////////////
'////////////////////////////////////////////////////////////////// 17/10/2015 //////////////////////////////////////////////////////////////////
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

'################################################################################################################################################
'Este script tem a finalidade de excluir quaisquer contas locais das máquinas clientes, tenham estas privilégios administrativos ou não,
'e promover a criação de um usuário local com privilégios para uso restrito da Divisão de Informática.

'Execução: 
'Procura por usuários locais, membros de quaisquer grupos nas máquinas clientes e os elimina.
'Cria um usuário local com as configurações especificadas, membro do grupo Administradores.
'################################################################################################################################################

Dim verConta, strUsuario, strSenha, objShell, objNetwork, objUser, objGroup, colAccounts

verConta = 0 'Declaração de Variável de verificação
strUsuario = "user" 'Nome do usuário à ser criado
strSenha = "password" 'Senha do usuário à ser criado, respeitando as regras de senha do seu domínio
'Obs.: Atentar-se à quantidade de caracteres e critérios ao definir a senha, caso contrário ela não será alterada

Set objShell = CreateObject("Wscript.Shell") 'Instancia o objeto Wscript.Shell
Set objNetwork = CreateObject("Wscript.Network") 'Instancia o objeto Wscript.Network
 
strComputer = objNetwork.ComputerName 'Recupera o nome do computador local
 
Set colAccounts = GetObject("WinNT://" & strComputer & "") 'Utiliza o provedor WinNT para coletar as informações do computador local 
 
colAccounts.Filter = Array("user") 'Filtra apenas as informações ligadas aos usuários
 
For Each objUser In colAccounts 'Para cada objeto retornado
 
    If UCase(objUser.Name) <> "ADMINISTRATOR" AND UCase(objUser.Name) <> strUsuario Then 'Verifica de o nome do usuário é diferente de Administrator e a conta definida
           
            If objUser.AccountDisabled = TRUE then 'Se as duas condições forem atendidas, verifica se a conta está desativada
               
            Else 
                
                objUser.AccountDisabled = True  'Se a conta não estiver desativada, desativa a mesma.
                objUser.SetInfo 
            End if 
    End If 
 
Next 
 	
 	For Each objUser In colAccounts 
 		If UCase(objUser.Name) = strUsuario Then verConta=1 End If 'Se a conta definida já existir, seta a variável de verificação para 1
	Next
		
		
		
		If verConta = 1 Then 'Se a variável de verificação for = 1, ou seja, se a conta já existir
		
		Set objUser = GetObject("WinNT://" & strComputer & "/"&strUsuario&", user") 'Acessa o objeto do usuário
		objUser.SetPassword strSenha       'Modifica a senha para o usuário conforme definido
		objUser.Put "UserFlags", 65600     'Senha nunca expira E usuário não pode alterar a senha
		objUser.AccountDisabled = False    'Ativa a conta, se ela estiver desativada
		objUser.SetInfo 
		
		WScript.Quit
			
		Else
		
		
		Set colAccounts = GetObject("WinNT://"& strComputer) 'Acessa as configurações da máquina local
		Set objUser = colAccounts.Create("user", strUsuario) 'Cria o usuário definido
		objUser.SetPassword strSenha      'Cria a senha definida
		objUser.Put "UserFlags", 65600    'Senha nunca expira E usuário não pode alterar a senha
		objUser.AccountDisabled = False   'Conta ativa
		objUser.SetInfo
		Set objGroup = GetObject("WinNT://" & strComputer & "/Administradores,group") 'Insere o usuário no grupo Administradores
		Set objUser = GetObject("WinNT://" & strComputer & "/"&strUsuario&",user")    'Confirma as informações
		objGroup.Add(objUser.ADsPath)
		
	End If
	
