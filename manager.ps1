<#
.SYNOPSIS
    Módulo PowerShell Definitivo para administração do Microsoft 365.
.DESCRIPTION
    Reúne um vasto conjunto de funções para Microsoft Graph e Exchange Online.
    Automatiza a verificação e instalação dos módulos necessários ('ExchangeOnlineManagement', 'Microsoft.Graph').
    Projetado para ser o mais auto-contido possível.
.NOTES
    Autor: Pedro Moura
    Versão: 3.2 (Bug Fix - ShouldProcess e String Parsing)
    Data: 24/05/2025
    Requerimentos: PowerShell 5.1+, Conexão com a Internet.
    Pré-Requisito Manual: A Política de Execução do PowerShell deve ser 'RemoteSigned' ou menos restritiva.
                       (Execute 'Set-ExecutionPolicy RemoteSigned -Force' como Administrador uma vez).
    Aviso: USE COM CUIDADO. TESTE EM AMBIENTE SEGURO ANTES DE USAR EM PRODUÇÃO.
           Muitas funções requerem permissões de Administrador no M365.
#>

#region Configurações Globais e Tratamento de Erros

Set-StrictMode -Version Latest
$Global:ErrorActionPreference = "Stop" # Tenta forçar 'Stop' para capturar erros com try/catch

# Função auxiliar para instalação de módulos
Function Install-RequiredModule {
    Param (
        [String]$ModuleName
    )
    if (-not (Get-Module -ListAvailable -Name $ModuleName)) {
        Write-Host "Módulo '$ModuleName' não encontrado. Tentando instalar via PSGallery..." -ForegroundColor Yellow
        try {
            # Tenta instalar para o usuário atual primeiro
            Install-Module $ModuleName -AllowClobber -Force -Scope CurrentUser -Repository PSGallery -ErrorAction Stop
            Write-Host "Módulo '$ModuleName' instalado para o usuário atual com sucesso." -ForegroundColor Green
        } catch {
            Write-Warning "Falha ao instalar para o usuário atual. Tentando como Administrador (pode exigir elevação)..."
            try {
                Install-Module $ModuleName -AllowClobber -Force -Scope AllUsers -Repository PSGallery -ErrorAction Stop
                Write-Host "Módulo '$ModuleName' instalado para todos os usuários com sucesso." -ForegroundColor Green
            } catch {
                Write-Error "FALHA CRÍTICA ao instalar '$ModuleName': $_. Verifique sua conexão e permissões, e tente 'Install-Module $ModuleName -Force' manualmente em um PowerShell como Administrador."
                throw "Instalação de Módulo Essencial Falhou: $ModuleName"
            }
        }
    } else {
        Write-Host "Módulo '$ModuleName' já está disponível." -ForegroundColor Green
    }
    # Garante que o módulo seja importado na sessão atual
    try {
        Import-Module $ModuleName -ErrorAction Stop
    } catch {
         Write-Error "Falha ao importar o módulo '$ModuleName' após a instalação/verificação: $_"
         throw "Importação de Módulo Essencial Falhou: $ModuleName"
    }
}

#endregion

#region Funções de Conexão e Desconexão (Núcleo)

Function Connect-Super365Services {
    [CmdletBinding()]
    Param (
        [Parameter(Mandatory=$false)]
        [String] $UserPrincipalName,
        [Parameter(Mandatory=$false)]
        [String[]] $GraphScopes = @(
            # Permissões ABRANGENTES - Reduza se necessário para o Princípio do Menor Privilégio
            "User.ReadWrite.All", "Group.ReadWrite.All", "Directory.ReadWrite.All",
            "MailboxSettings.ReadWrite", "Mail.ReadWrite", "Mail.Send",
            "Reports.Read.All", "AuditLog.Read.All", "Policy.ReadWrite.All",
            "DeviceManagementManagedDevices.ReadWrite.All", "Team.ReadBasic.All", "TeamMember.ReadWrite.All",
            "TeamSettings.ReadWrite.All", "Channel.ReadWrite.All", "TeamsApp.ReadWrite.All",
            "RoleManagement.ReadWrite.Directory", "Application.ReadWrite.All", "ServiceHealth.Read.All",
            "LicenseAssignment.ReadWrite.All", "Chat.ReadWrite.All", "Sites.ReadWrite.All",
            "UserAuthenticationMethod.ReadWrite.All" # Para MFA
        )
    )

    Write-Host "--- Iniciando Conexão SuperAdmin365 ---" -ForegroundColor Cyan

    # Instalação/Verificação de Módulos
    Write-Host "[Passo 1/3] Verificando e Instalando Módulos Essenciais..." -ForegroundColor Yellow
    Install-RequiredModule -ModuleName "ExchangeOnlineManagement"
    Install-RequiredModule -ModuleName "Microsoft.Graph"

    # Conexão com Exchange Online
    Write-Host "[Passo 2/3] Conectando ao Exchange Online..." -ForegroundColor Yellow
    try {
        Get-PSSession | Where-Object { $_.ConfigurationName -eq 'Microsoft.Exchange' } | Remove-PSSession -Confirm:$false -ErrorAction SilentlyContinue

        if ($PSBoundParameters.ContainsKey('UserPrincipalName')) {
            Connect-ExchangeOnline -UserPrincipalName $UserPrincipalName -ShowProgress $true
        } else {
            Connect-ExchangeOnline -ShowProgress $true
        }
        Write-Host "=> Conectado ao Exchange Online com sucesso." -ForegroundColor Green
    } catch {
        Write-Error "Falha ao conectar ao Exchange Online: $_"
        return
    }

    # Conexão com Microsoft Graph
    Write-Host "[Passo 3/3] Conectando ao Microsoft Graph..." -ForegroundColor Yellow
    try {
        Disconnect-MgGraph -ErrorAction SilentlyContinue
        Connect-MgGraph -Scopes $GraphScopes
        $context = Get-MgContext
        Write-Host "=> Conectado ao Microsoft Graph como $($context.Account)." -ForegroundColor Green
    } catch {
        Write-Error "Falha ao conectar ao Microsoft Graph: $_"
        return
    }
    Write-Host "--- Conexão aos Serviços M365 Concluída ---" -ForegroundColor Cyan
}

Function Disconnect-Super365Services {
    [CmdletBinding()]
    Param ()
    Write-Host "Desconectando dos serviços M365..." -ForegroundColor Yellow
    try {
        Get-PSSession | Where-Object { $_.ConfigurationName -eq 'Microsoft.Exchange' } | Remove-PSSession -Confirm:$false -ErrorAction SilentlyContinue
        Write-Host "- Sessão Exchange Online encerrada." -ForegroundColor Gray
    } catch { Write-Warning "Não foi possível desconectar do Exchange Online (talvez já estivesse desconectado)." }
    try {
        Disconnect-MgGraph -ErrorAction SilentlyContinue
        Write-Host "- Sessão Microsoft Graph encerrada." -ForegroundColor Gray
    } catch { Write-Warning "Não foi possível desconectar do Microsoft Graph (talvez já estivesse desconectado)." }
    Write-Host "Desconexão concluída." -ForegroundColor Green
}

#endregion

#region Microsoft Graph - Usuários (Expandido)

Function Get-S365User {
    [CmdletBinding()]
    Param (
        [String] $Identity,
        [Switch] $All,
        [String] $Filter,
        [String[]] $Select = @("Id", "DisplayName", "UserPrincipalName", "Mail", "JobTitle", "Department", "AccountEnabled", "CreatedDateTime", "LastPasswordChangeDateTime", "SignInActivity", "UsageLocation", "Manager")
    )
    try {
        $params = @{ Select = $Select }
        if ($All) { Get-MgUser @params -All }
        elseif ($Filter) { Get-MgUser @params -Filter $Filter -ConsistencyLevel eventual -CountVariable countVar } # Adicionado ConsistencyLevel para filtros avançados
        elseif ($Identity) { Get-MgUser @params -UserId $Identity }
        else { Write-Warning "Especifique -Identity, -Filter ou -All." }
    } catch { Write-Error "Erro ao buscar usuário(s): $_" }
}

Function New-S365User {
    [CmdletBinding()]
    Param (
        [Parameter(Mandatory=$true)][String] $UserPrincipalName,
        [Parameter(Mandatory=$true)][String] $DisplayName,
        [Parameter(Mandatory=$true)][String] $MailNickname,
        [Parameter(Mandatory=$true)][System.Security.SecureString] $Password,
        [Parameter(Mandatory=$true)][String] $UsageLocation, # Ex: "BR", "US"
        [String] $GivenName,
        [String] $Surname,
        [String] $JobTitle,
        [String] $Department,
        [Switch] $ForceChangePasswordNextSignIn = $true,
        [Switch] $AccountEnabled = $true
    )
    $params = @{
        UserPrincipalName = $UserPrincipalName
        DisplayName = $DisplayName
        MailNickname = $MailNickname
        PasswordProfile = @{ ForceChangePasswordNextSignIn = $ForceChangePasswordNextSignIn; Password = $Password }
        AccountEnabled = $AccountEnabled
        UsageLocation = $UsageLocation
    }
    if ($GivenName) { $params.Add("GivenName", $GivenName) }
    if ($Surname) { $params.Add("Surname", $Surname) }
    if ($JobTitle) { $params.Add("JobTitle", $JobTitle) }
    if ($Department) { $params.Add("Department", $Department) }

    try {
        $newUser = New-MgUser -BodyParameter $params
        Write-Host "Usuário '$UserPrincipalName' (ID: $($newUser.Id)) criado com sucesso." -ForegroundColor Green
        return $newUser
    } catch { Write-Error "Erro ao criar usuário: $_" }
}

Function Set-S365User {
    [CmdletBinding()]
    Param (
        [Parameter(Mandatory=$true)][String] $UserPrincipalName,
        [String] $JobTitle,
        [String] $Department,
        [String] $OfficeLocation,
        [String] $MobilePhone,
        [String] $ManagerUPN,
        [String] $UsageLocation,
        [Switch] $EnableAccount,
        [Switch] $DisableAccount
    )
    $params = @{}
    if ($JobTitle) { $params.Add("JobTitle", $JobTitle) }
    if ($Department) { $params.Add("Department", $Department) }
    if ($OfficeLocation) { $params.Add("OfficeLocation", $OfficeLocation) }
    if ($MobilePhone) { $params.Add("MobilePhone", $MobilePhone) }
    if ($UsageLocation) { $params.Add("UsageLocation", $UsageLocation) }
    if ($PSBoundParameters.ContainsKey('EnableAccount')) { $params.Add("AccountEnabled", $true) }
    if ($PSBoundParameters.ContainsKey('DisableAccount')) { $params.Add("AccountEnabled", $false) }

    try {
        if ($params.Count -gt 0) {
            Update-MgUser -UserId $UserPrincipalName -BodyParameter $params
            Write-Host "Propriedades atualizadas para $UserPrincipalName." -ForegroundColor Green
        }
        if ($ManagerUPN) {
            $manager = Get-S365User -Identity $ManagerUPN
            if ($manager) {
                Set-MgUserManagerByRef -UserId $UserPrincipalName -AdditionalProperties @{"@odata.id" = "https://graph.microsoft.com/v1.0/users/$($manager.Id)"}
                Write-Host "Gerente '$ManagerUPN' definido para $UserPrincipalName." -ForegroundColor Green
            } else {
                 Write-Warning "Gerente '$ManagerUPN' não encontrado."
            }
        }
    } catch { Write-Error "Erro ao atualizar usuário '$UserPrincipalName': $_" }
}

Function Remove-S365User {
    [CmdletBinding(SupportsShouldProcess=$true, ConfirmImpact='High')]
    Param (
        [Parameter(Mandatory=$true)][String] $UserPrincipalName
    )
    try {
        if ($PSCmdlet.ShouldProcess($UserPrincipalName, "Remover Usuário (será movido para Lixeira)")) {
            Remove-MgUser -UserId $UserPrincipalName
            Write-Host "Usuário '$UserPrincipalName' movido para a lixeira." -ForegroundColor Green
        }
    } catch { Write-Error "Erro ao remover usuário '$UserPrincipalName': $_" }
}

Function Restore-S365DeletedUser {
     [CmdletBinding()]
    Param (
        [Parameter(Mandatory=$true)][String] $UserId # ID do usuário, não UPN
    )
    try {
        Restore-MgDirectoryDeletedItem -DirectoryObjectId $UserId
        Write-Host "Usuário com ID '$UserId' restaurado com sucesso." -ForegroundColor Green
    } catch { Write-Error "Erro ao restaurar usuário: $_" }
}

Function Get-S365DeletedUser {
     [CmdletBinding()]
    Param ()
    try {
        Get-MgDirectoryDeletedItem -DirectoryObjectId "microsoft.graph.user" | Select-Object Id, DisplayName, UserPrincipalName, DeletedDateTime
    } catch { Write-Error "Erro ao buscar usuários excluídos: $_" }
}

Function Reset-S365UserPassword {
    [CmdletBinding()]
    Param (
        [Parameter(Mandatory=$true)][String] $UserPrincipalName,
        [Parameter(Mandatory=$true)][System.Security.SecureString] $NewPassword,
        [Switch] $ForceChangePasswordNextSignIn = $true
    )
    $passwordProfile = @{
        ForceChangePasswordNextSignIn = $ForceChangePasswordNextSignIn
        Password = $NewPassword
    }
    try {
        Update-MgUser -UserId $UserPrincipalName -PasswordProfile $passwordProfile
        Write-Host "Senha redefinida para '$UserPrincipalName'." -ForegroundColor Green
    } catch { Write-Error "Erro ao redefinir senha: $_" }
}

Function Get-S365UserManager {
    [CmdletBinding()]
    Param ([Parameter(Mandatory=$true)][String] $UserPrincipalName)
    try { Get-MgUserManager -UserId $UserPrincipalName | Select-Object Id, DisplayName, UserPrincipalName }
    catch { Write-Error "Erro ao buscar gerente para '$UserPrincipalName': $_" }
}

Function Get-S365UserDirectReports {
    [CmdletBinding()]
    Param ([Parameter(Mandatory=$true)][String] $UserPrincipalName)
    try { Get-MgUserDirectReport -UserId $UserPrincipalName | Select-Object Id, DisplayName, UserPrincipalName }
    catch { Write-Error "Erro ao buscar subordinados diretos para '$UserPrincipalName': $_" }
}

Function Get-S365MfaStatus {
    [CmdletBinding()]
    Param ([Parameter(Mandatory=$true)][String] $UserPrincipalName)
    Write-Host "Verificando métodos de autenticação para $UserPrincipalName..." -ForegroundColor Yellow
    try {
        $methods = Get-MgUserAuthenticationMethod -UserId $UserPrincipalName
        if ($methods) {
            Write-Host "Métodos registrados:" -ForegroundColor Green
            $methods | ForEach-Object {
                $type = $_.AdditionalProperties.'@odata.type'.Split('.')[-1]
                Write-Host "- Tipo: $type"
            }
            if (($methods | Where-Object { $_.AdditionalProperties.'@odata.type' -like "*PhoneAuthenticationMethod*" -or $_.AdditionalProperties.'@odata.type' -like "*MicrosoftAuthenticatorAuthenticationMethod*" })) {
                 Write-Host "=> Status: MFA provavelmente ATIVO (Método forte registrado)." -ForegroundColor Green
            } else {
                 Write-Host "=> Status: MFA provavelmente INATIVO (Nenhum método forte registrado)." -ForegroundColor Yellow
            }
        } else {
            Write-Host "=> Status: NENHUM método de autenticação registrado." -ForegroundColor Red
        }
    } catch { Write-Error "Erro ao buscar status MFA para '$UserPrincipalName': $_" }
}

#endregion

#region Microsoft Graph - Licenças

Function Get-S365AvailableSkus {
    [CmdletBinding()]
    Param ()
    try {
        Get-MgSubscribedSku | Select-Object SkuId, SkuPartNumber, ConsumedUnits, PrepaidUnits
    } catch { Write-Error "Erro ao buscar SKUs: $_" }
}

Function Get-S365UserLicense {
    [CmdletBinding()]
    Param (
        [Parameter(Mandatory=$true)][String] $UserPrincipalName
    )
    try {
        Get-MgUserLicenseDetail -UserId $UserPrincipalName | Select-Object SkuPartNumber, ServicePlans
    } catch { Write-Error "Erro ao buscar licenças de '$UserPrincipalName': $_" }
}

Function Set-S365UserLicense {
    [CmdletBinding()]
    Param (
        [Parameter(Mandatory=$true)][String] $UserPrincipalName,
        [String[]] $AddSkuIds,
        [String[]] $RemoveSkuIds
    )
    try {
        $user = Get-S365User -Identity $UserPrincipalName -Select UsageLocation
        if (-not $user.UsageLocation) {
             Write-Error "Usuário '$UserPrincipalName' não tem UsageLocation. Defina-o antes de atribuir licenças usando 'Set-S365User'."
             return
        }

        $addLicensesObject = @()
        if ($AddSkuIds) { $addLicensesObject = $AddSkuIds | ForEach-Object { @{ SkuId = $_ } } }

        $params = @{
            UserId = $UserPrincipalName
            AddLicenses = $addLicensesObject
            RemoveLicenses = @($RemoveSkuIds)
        }

        Set-MgUserLicense @params
        Write-Host "Licenças atualizadas para '$UserPrincipalName'." -ForegroundColor Green
    } catch { Write-Error "Erro ao atualizar licenças para '$UserPrincipalName': $_" }
}

#endregion

#region Microsoft Graph - Grupos (Expandido)

Function Get-S365Group {
    [CmdletBinding()]
    Param (
        [String] $Identity,
        [Switch] $All,
        [String] $Filter,
        [String[]] $Select = @("Id", "DisplayName", "Mail", "GroupTypes", "SecurityEnabled", "MailEnabled", "Visibility", "Description")
    )
    try {
        $params = @{ Select = $Select }
        if ($All) { Get-MgGroup @params -All }
        elseif ($Filter) { Get-MgGroup @params -Filter $Filter -ConsistencyLevel eventual -CountVariable countVar }
        elseif ($Identity) { Get-MgGroup @params -GroupId $Identity }
        else { Write-Warning "Especifique -Identity, -Filter ou -All." }
    } catch { Write-Error "Erro ao buscar grupo(s): $_" }
}

Function Get-S365GroupMember {
    [CmdletBinding()]
    Param (
        [Parameter(Mandatory=$true)][String] $GroupId
    )
    try {
        Get-MgGroupMember -GroupId $GroupId -All | Select-Object Id, DisplayName, UserPrincipalName, Mail
    } catch { Write-Error "Erro ao buscar membros do grupo '$GroupId': $_" }
}

Function Add-S365GroupMember {
    [CmdletBinding()]
    Param (
        [Parameter(Mandatory=$true)][String] $GroupId,
        [Parameter(Mandatory=$true)][String] $UserPrincipalName
    )
    try {
        $user = Get-S365User -Identity $UserPrincipalName
        New-MgGroupMember -GroupId $GroupId -DirectoryObjectId $user.Id
        Write-Host "Usuário '$UserPrincipalName' adicionado ao grupo '$GroupId'." -ForegroundColor Green
    } catch { Write-Error "Erro ao adicionar membro '$UserPrincipalName' ao grupo '$GroupId': $_" }
}

Function Remove-S365GroupMember {
    [CmdletBinding(SupportsShouldProcess=$true)]
    Param (
        [Parameter(Mandatory=$true)][String] $GroupId,
        [Parameter(Mandatory=$true)][String] $UserPrincipalName
    )
    try {
        $user = Get-S365User -Identity $UserPrincipalName
        if ($PSCmdlet.ShouldProcess("$UserPrincipalName from $GroupId", "Remover Membro do Grupo")) {
            Remove-MgGroupMemberByRef -GroupId $GroupId -DirectoryObjectId $user.Id
            Write-Host "Usuário '$UserPrincipalName' removido do grupo '$GroupId'." -ForegroundColor Green
        }
    } catch { Write-Error "Erro ao remover membro '$UserPrincipalName' do grupo '$GroupId': $_" }
}

Function New-S365Group {
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory=$true)][String] $DisplayName,
        [Parameter(Mandatory=$true)][String] $MailNickname,
        [Parameter(Mandatory=$true)][ValidateSet("M365", "Security")][String] $GroupType,
        [String] $Description,
        [String[]] $OwnersUPN,
        [String[]] $MembersUPN,
        [Switch] $MailEnabled = ($GroupType -eq "M365"),
        [Switch] $SecurityEnabled = ($GroupType -eq "Security" -or $GroupType -eq "M365"),
        [ValidateSet("Public", "Private")][String] $Visibility = "Private" # Para grupos M365
    )
    $params = @{
        DisplayName = $DisplayName
        MailNickname = $MailNickname
        MailEnabled = $MailEnabled
        SecurityEnabled = $SecurityEnabled
        Description = $Description
        GroupTypes = @(if ($GroupType -eq "M365") { "Unified" } else { })
    }
    if ($GroupType -eq "M365") { $params.Visibility = $Visibility }

    try {
        $newGroup = New-MgGroup -BodyParameter $params
        Write-Host "Grupo '$DisplayName' (ID: $($newGroup.Id)) criado com sucesso." -ForegroundColor Green

        if ($OwnersUPN) {
            $OwnersUPN | ForEach-Object {
                try { Add-S365GroupOwner -GroupId $newGroup.Id -UserPrincipalName $_ }
                catch { Write-Warning "Não foi possível adicionar '$_' como proprietário: $($_.Exception.Message)" }
            }
        }
        if ($MembersUPN) {
             $MembersUPN | ForEach-Object {
                 try { Add-S365GroupMember -GroupId $newGroup.Id -UserPrincipalName $_ }
                 catch { Write-Warning "Não foi possível adicionar '$_' como membro: $($_.Exception.Message)" }
             }
        }
        return $newGroup
    } catch { Write-Error "Erro ao criar grupo: $_" }
}

Function Add-S365GroupOwner {
    [CmdletBinding()]
    Param (
        [Parameter(Mandatory=$true)][String] $GroupId,
        [Parameter(Mandatory=$true)][String] $UserPrincipalName
    )
    try {
        $user = Get-S365User -Identity $UserPrincipalName
        New-MgGroupOwnerByRef -GroupId $GroupId -OdataId "https://graph.microsoft.com/v1.0/users/$($user.Id)"
        Write-Host "Usuário '$UserPrincipalName' adicionado como proprietário do grupo '$GroupId'." -ForegroundColor Green
    } catch { Write-Error "Erro ao adicionar proprietário '$UserPrincipalName' ao grupo '$GroupId': $_" }
}

Function Get-S365GroupOwner {
    [CmdletBinding()]
    Param (
        [Parameter(Mandatory=$true)][String] $GroupId
    )
    try {
        Get-MgGroupOwner -GroupId $GroupId -All | Select-Object Id, DisplayName, UserPrincipalName
    } catch { Write-Error "Erro ao buscar proprietários do grupo '$GroupId': $_" }
}

#endregion

#region Microsoft Graph - Teams

Function Get-S365Team {
    [CmdletBinding()]
    Param (
        [String] $DisplayName,
        [Switch] $All
    )
    try {
        if ($DisplayName) { Get-MgTeam -Filter "DisplayName eq '$DisplayName'" }
        elseif ($All) { Get-MgTeam -All }
        else { Get-MgTeam }
    } catch { Write-Error "Erro ao buscar Teams: $_" }
}

Function New-S365Team {
    [CmdletBinding()]
    Param (
        [Parameter(Mandatory=$true)][String] $DisplayName,
        [String] $Description,
        [String] $OwnerUPN # O dono que está criando (UPN)
    )
    try {
        $template = "https://graph.microsoft.com/v1.0/teamsTemplates('standard')"
        $params = @{
            "template@odata.bind" = $template
            DisplayName = $DisplayName
            Description = $Description
        }

        $newTeam = New-MgTeam -BodyParameter $params
        Write-Host "Team '$DisplayName' (ID: $($newTeam.Id)) criado. Aguarde a provisionação..." -ForegroundColor Yellow

        # Aguarda um pouco e tenta adicionar o dono se fornecido
        Start-Sleep -Seconds 15

        if ($OwnerUPN) {
            try {
                 Add-S365TeamMember -TeamId $newTeam.Id -UserPrincipalName $OwnerUPN -IsOwner
            } catch {
                Write-Warning "Team criado, mas falha ao adicionar '$OwnerUPN' como proprietário inicial. Tente adicioná-lo manualmente: $_"
            }
        } else {
             Write-Warning "Team criado sem proprietário inicial. Adicione um proprietário o mais rápido possível."
        }

        Write-Host "Team '$DisplayName' provisionado (verifique se o proprietário foi adicionado)." -ForegroundColor Green
        return $newTeam
    } catch { Write-Error "Erro ao criar Team: $_" }
}

Function Get-S365TeamChannel {
    [CmdletBinding()]
    Param ([Parameter(Mandatory=$true)][String] $TeamId)
    try { Get-MgTeamChannel -TeamId $TeamId -All | Select-Object Id, DisplayName, WebUrl, MembershipType }
    catch { Write-Error "Erro ao buscar canais para Team '$TeamId': $_" }
}

Function Add-S365TeamMember {
    [CmdletBinding()]
    Param (
        [Parameter(Mandatory=$true)][String] $TeamId,
        [Parameter(Mandatory=$true)][String] $UserPrincipalName,
        [Switch] $IsOwner
    )
    try {
        $user = Get-S365User -Identity $UserPrincipalName
        $params = @{
            "@odata.type" = "#microsoft.graph.aadUserConversationMember"
            "user@odata.bind" = "https://graph.microsoft.com/v1.0/users('$($user.Id)')"
            Roles = @(if ($IsOwner) { "owner" } else { "member" })
        }
        New-MgTeamMember -TeamId $TeamId -BodyParameter $params
        $role = if ($IsOwner) { "Proprietário" } else { "Membro" }
        Write-Host "Usuário '$UserPrincipalName' adicionado ao Team '$TeamId' como $role." -ForegroundColor Green
    } catch { Write-Error "Erro ao adicionar membro '$UserPrincipalName' ao Team '$TeamId': $_" }
}

Function Remove-S365TeamMember {
    [CmdletBinding(SupportsShouldProcess=$true)]
    Param (
        [Parameter(Mandatory=$true)][String] $TeamId,
        [Parameter(Mandatory=$true)][String] $UserPrincipalName
    )
    try {
        $user = Get-S365User -Identity $UserPrincipalName
        $member = Get-MgTeamMember -TeamId $TeamId -Filter "UserId eq '$($user.Id)'"
        if ($member) {
            if ($PSCmdlet.ShouldProcess("$UserPrincipalName from $TeamId", "Remover Membro do Team")) {
                Remove-MgTeamMember -TeamId $TeamId -ConversationMemberId $member.Id
                Write-Host "Usuário '$UserPrincipalName' removido do Team '$TeamId'." -ForegroundColor Green
            }
        } else {
            Write-Warning "Usuário '$UserPrincipalName' não encontrado como membro do Team '$TeamId'."
        }
    } catch { Write-Error "Erro ao remover membro '$UserPrincipalName' do Team '$TeamId': $_" }
}


#endregion

#region Exchange Online - Caixas de Correio

Function Get-S365Mailbox {
    [CmdletBinding()]
    Param (
        [String] $Identity,
        [Switch] $All,
        [String] $Filter,
        [Switch] $Shared,
        [Switch] $Room,
        [Switch] $Equipment,
        [Switch] $Archive
    )
    try {
        $params = @{}
        if ($Shared) { $params.RecipientTypeDetails = "SharedMailbox" }
        elseif ($Room) { $params.RecipientTypeDetails = "RoomMailbox" }
        elseif ($Equipment) { $params.RecipientTypeDetails = "EquipmentMailbox" }
        elseif ($Archive) { $params.Archive = $true } # Busca caixas com arquivo morto

        if ($All) { Get-Mailbox @params -ResultSize Unlimited }
        elseif ($Filter) { Get-Mailbox @params -Filter $Filter -ResultSize Unlimited }
        elseif ($Identity) { Get-Mailbox @params -Identity $Identity }
        else { Get-Mailbox @params -ResultSize 500 } # Default
    } catch { Write-Error "Erro ao buscar caixa(s) de correio: $_" }
}

Function Set-S365MailboxQuota {
     [CmdletBinding()]
    Param (
        [Parameter(Mandatory=$true)][String] $Identity,
        [Parameter(Mandatory=$true)][String] $Quota # Ex: "50GB", "100GB"
    )
    try {
        $quotaValue = [Microsoft.Exchange.Data.ByteQuantifiedSize]::Parse($Quota)
        $warningQuota = ([Microsoft.Exchange.Data.ByteQuantifiedSize]($quotaValue.ToBytes() * 0.9)).ToString()
        Set-Mailbox -Identity $Identity -ProhibitSendReceiveQuota $Quota -IssueWarningQuota $warningQuota
        Write-Host "Quota da caixa de correio '$Identity' definida para $Quota (aviso em $warningQuota)." -ForegroundColor Green
    } catch { Write-Error "Erro ao definir quota para '$Identity': $_" }
}

Function Set-S365MailboxForwarding {
    [CmdletBinding()]
    Param (
        [Parameter(Mandatory=$true)][String] $Identity,
        [String] $ForwardingSmtpAddress,
        [Switch] $DeliverToMailboxAndForward = $true,
        [Switch] $RemoveForwarding
    )
    try {
        if ($RemoveForwarding) {
            Set-Mailbox -Identity $Identity -ForwardingSmtpAddress $null -DeliverToMailboxAndForward $false
            Write-Host "Encaminhamento removido para '$Identity'." -ForegroundColor Green
        } else {
            if (-not $ForwardingSmtpAddress) { Write-Error "É necessário fornecer -ForwardingSmtpAddress." ; return }
            Set-Mailbox -Identity $Identity -ForwardingSmtpAddress $ForwardingSmtpAddress -DeliverToMailboxAndForward $DeliverToMailboxAndForward
            Write-Host "Encaminhamento configurado para '$Identity' para '$ForwardingSmtpAddress'." -ForegroundColor Green
        }
    } catch { Write-Error "Erro ao configurar encaminhamento para '$Identity': $_" }
}

Function Get-S365MailboxStatistics {
     [CmdletBinding()]
    Param (
        [Parameter(Mandatory=$true)][String] $Identity
    )
    try {
        Get-MailboxStatistics -Identity $Identity | Select-Object DisplayName, ItemCount, TotalItemSize, LastLogonTime, LastUserActionTime
    } catch { Write-Error "Erro ao buscar estatísticas para '$Identity': $_" }
}

Function New-S365SharedMailbox {
    [CmdletBinding()]
    Param (
        [Parameter(Mandatory=$true)][String] $Name,
        [String] $Alias = ($Name -replace "\s", ""),
        [String] $PrimarySmtpAddress = "$Alias@$(Get-AcceptedDomain | Where-Object { $_.IsDefault -eq $true } | Select-Object -ExpandProperty DomainName)"
    )
    try {
        New-Mailbox -Shared -Name $Name -DisplayName $Name -Alias $Alias -PrimarySmtpAddress $PrimarySmtpAddress
        Write-Host "Caixa compartilhada '$Name' ($PrimarySmtpAddress) criada com sucesso." -ForegroundColor Green
    } catch { Write-Error "Erro ao criar caixa compartilhada '$Name': $_" }
}

Function Set-S365LitigationHold {
     [CmdletBinding()]
    Param (
        [Parameter(Mandatory=$true)][String] $Identity,
        [Switch] $Enable,
        [Switch] $Disable,
        [Int32] $DurationDays # Opcional: Duração em dias
    )
    try {
        if ($Enable) {
            $params = @{ Identity = $Identity; LitigationHoldEnabled = $true }
            if ($DurationDays) { $params.LitigationHoldDuration = $DurationDays }
            Set-Mailbox @params
            Write-Host "Litigation Hold HABILITADO para '$Identity'." -ForegroundColor Green
        } elseif ($Disable) {
            Set-Mailbox -Identity $Identity -LitigationHoldEnabled $false
            Write-Host "Litigation Hold DESABILITADO para '$Identity'." -ForegroundColor Green
        } else {
            Write-Warning "Especifique -Enable ou -Disable."
        }
    } catch { Write-Error "Erro ao configurar Litigation Hold para '$Identity': $_" }
}


#endregion

#region Exchange Online - Permissões

Function Add-S365MailboxPermission {
    [CmdletBinding()]
    Param (
        [Parameter(Mandatory=$true)][String] $Identity,
        [Parameter(Mandatory=$true)][String] $User,
        [Parameter(Mandatory=$true)][ValidateSet("FullAccess", "ExternalAccount", "DeleteItem", "ReadPermission", "ChangePermission", "ChangeOwner")]$AccessRights = "FullAccess",
        [Switch] $AutoMapping = $true
    )
    try {
        Add-MailboxPermission -Identity $Identity -User $User -AccessRights $AccessRights -InheritanceType All -AutoMapping:$AutoMapping
        Write-Host "Permissão '$AccessRights' concedida a '$User' na caixa '$Identity'." -ForegroundColor Green
    } catch { Write-Error "Erro ao adicionar permissão '$AccessRights' a '$User' em '$Identity': $_" }
}

Function Remove-S365MailboxPermission {
    [CmdletBinding(SupportsShouldProcess=$true)]
    Param (
        [Parameter(Mandatory=$true)][String] $Identity,
        [Parameter(Mandatory=$true)][String] $User,
        [Parameter(Mandatory=$true)][ValidateSet("FullAccess", "ExternalAccount", "DeleteItem", "ReadPermission", "ChangePermission", "ChangeOwner")]$AccessRights = "FullAccess"
    )
    try {
        if ($PSCmdlet.ShouldProcess("$User on $Identity", "Remover Permissão de Mailbox ($AccessRights)")) {
            Remove-MailboxPermission -Identity $Identity -User $User -AccessRights $AccessRights -InheritanceType All -Confirm:$false
            Write-Host "Permissão '$AccessRights' removida de '$User' na caixa '$Identity'." -ForegroundColor Green
        }
    } catch { Write-Error "Erro ao remover permissão '$AccessRights' de '$User' em '$Identity': $_" }
}

Function Add-S365RecipientPermission {
    [CmdletBinding()]
    Param (
        [Parameter(Mandatory=$true)][String] $Identity,
        [Parameter(Mandatory=$true)][String] $Trustee, # Usuário que recebe
        [Parameter(Mandatory=$true)][ValidateSet("SendAs", "SendOnBehalf")]$AccessRights = "SendAs"
    )
    try {
        if ($AccessRights -eq "SendAs") {
             Add-RecipientPermission -Identity $Identity -Trustee $Trustee -AccessRights SendAs
             Write-Host "Permissão 'SendAs' concedida a '$Trustee' em '$Identity'." -ForegroundColor Green
        } elseif ($AccessRights -eq "SendOnBehalf") {
             Set-Mailbox -Identity $Identity -GrantSendOnBehalfTo @{Add="$Trustee"}
             Write-Host "Permissão 'SendOnBehalf' concedida a '$Trustee' em '$Identity'." -ForegroundColor Green
        }
    } catch { Write-Error "Erro ao adicionar permissão '$AccessRights' a '$Trustee' em '$Identity': $_" }
}

Function Remove-S365RecipientPermission {
    [CmdletBinding(SupportsShouldProcess=$true)]
    Param (
        [Parameter(Mandatory=$true)][String] $Identity,
        [Parameter(Mandatory=$true)][String] $Trustee,
        [Parameter(Mandatory=$true)][ValidateSet("SendAs", "SendOnBehalf")]$AccessRights = "SendAs"
    )
    try {
       if ($PSCmdlet.ShouldProcess("$Trustee on $Identity", "Remover Permissão de Destinatário ($AccessRights)")) {
            if ($AccessRights -eq "SendAs") {
                Remove-RecipientPermission -Identity $Identity -Trustee $Trustee -AccessRights SendAs -Confirm:$false
                Write-Host "Permissão 'SendAs' removida de '$Trustee' em '$Identity'." -ForegroundColor Green
            } elseif ($AccessRights -eq "SendOnBehalf") {
                Set-Mailbox -Identity $Identity -GrantSendOnBehalfTo @{Remove="$Trustee"}
                Write-Host "Permissão 'SendOnBehalf' removida de '$Trustee' em '$Identity'." -ForegroundColor Green
            }
       }
    } catch { Write-Error "Erro ao remover permissão '$AccessRights' de '$Trustee' em '$Identity': $_" }
}

Function Get-S365MailboxFolderPermission {
     [CmdletBinding()]
    Param (
        [Parameter(Mandatory=$true)][String] $Identity,
        [String] $FolderPath = "\Calendário" # Ex: \Inbox, \Calendário, \Contacts, \Caixa de Entrada
    )
    try {
        # Tenta descobrir o nome correto da pasta Calendário se não for o padrão
        if ($FolderPath -eq "\Calendário" -or $FolderPath -eq "\Calendar") {
            try {
                $calendarPath = Get-MailboxFolderStatistics -Identity $Identity -FolderScope Calendar | Select-Object -First 1 | Select-Object -ExpandProperty FolderPath
                if ($calendarPath) { $FolderPath = $calendarPath }
            } catch { Write-Warning "Não foi possível detectar o nome exato da pasta Calendário, usando '$FolderPath'." }
        }
        Get-MailboxFolderPermission -Identity "$($Identity):$FolderPath"
    } catch { Write-Error "Erro ao buscar permissões da pasta '$FolderPath' para '$Identity': $_" }
}

Function Add-S365MailboxFolderPermission {
     [CmdletBinding()]
    Param (
        [Parameter(Mandatory=$true)][String] $Identity,
        [Parameter(Mandatory=$true)][String] $User,
        [Parameter(Mandatory=$true)][ValidateSet("Owner", "PublishingEditor", "Editor", "PublishingAuthor", "Author", "NoneditingAuthor", "Reviewer", "Contributor", "AvailabilityOnly", "LimitedDetails")]$AccessRights,
        [String] $FolderPath = "\Calendário"
    )
    try {
        Add-MailboxFolderPermission -Identity "$($Identity):$FolderPath" -User $User -AccessRights $AccessRights
        Write-Host "Permissão '$AccessRights' concedida a '$User' na pasta '$FolderPath' de '$Identity'." -ForegroundColor Green
    } catch { Write-Error "Erro ao adicionar permissão de pasta '$AccessRights' a '$User' em '$($Identity):$FolderPath': $_" }
}

Function Remove-S365MailboxFolderPermission {
     [CmdletBinding(SupportsShouldProcess=$true)]
    Param (
        [Parameter(Mandatory=$true)][String] $Identity,
        [Parameter(Mandatory=$true)][String] $User,
        [String] $FolderPath = "\Calendário"
    )
    try {
        # CORRIGIDO AQUI: Usando $($Identity) para ser mais explícito
        if ($PSCmdlet.ShouldProcess("$User on $($Identity):$FolderPath", "Remover Permissão de Pasta")) {
            Remove-MailboxFolderPermission -Identity "$($Identity):$FolderPath" -User $User -Confirm:$false
            Write-Host "Permissões removidas de '$User' na pasta '$FolderPath' de '$Identity'." -ForegroundColor Green
        }
    } catch { Write-Error "Erro ao remover permissão de pasta de '$User' em '$($Identity):$FolderPath': $_" }
}

#endregion

#region Exchange Online - Grupos de Distribuição e Contatos

Function Get-S365DistributionGroup {
    [CmdletBinding()]
    Param (
        [String] $Identity,
        [Switch] $All
    )
    try {
        if ($All) { Get-DistributionGroup -ResultSize Unlimited }
        elseif ($Identity) { Get-DistributionGroup -Identity $Identity }
        else { Get-DistributionGroup }
    } catch { Write-Error "Erro ao buscar grupo(s) de distribuição: $_" }
}

Function Get-S365DistributionGroupMember {
    [CmdletBinding()]
    Param (
        [Parameter(Mandatory=$true)][String] $Identity
    )
    try {
        Get-DistributionGroupMember -Identity $Identity -ResultSize Unlimited
    } catch { Write-Error "Erro ao buscar membros do grupo '$Identity': $_" }
}

Function Add-S365DistributionGroupMember {
    [CmdletBinding()]
    Param (
        [Parameter(Mandatory=$true)][String] $Identity,
        [Parameter(Mandatory=$true)][String] $Member
    )
    try {
        Add-DistributionGroupMember -Identity $Identity -Member $Member
        Write-Host "Membro '$Member' adicionado ao grupo '$Identity'." -ForegroundColor Green
    } catch { Write-Error "Erro ao adicionar membro '$Member' ao grupo '$Identity': $_" }
}

Function Remove-S365DistributionGroupMember {
    [CmdletBinding(SupportsShouldProcess=$true)]
    Param (
        [Parameter(Mandatory=$true)][String] $Identity,
        [Parameter(Mandatory=$true)][String] $Member
    )
    try {
        if ($PSCmdlet.ShouldProcess($Member, "Remover Membro do Grupo de Distribuição $Identity")) {
            Remove-DistributionGroupMember -Identity $Identity -Member $Member -Confirm:$false
            Write-Host "Membro '$Member' removido do grupo '$Identity'." -ForegroundColor Green
        }
    } catch { Write-Error "Erro ao remover membro '$Member' do grupo '$Identity': $_" }
}

Function Get-S365MailContact {
     [CmdletBinding()]
    Param (
        [String] $Identity,
        [Switch] $All
    )
    try {
        if ($All) { Get-MailContact -ResultSize Unlimited }
        elseif ($Identity) { Get-MailContact -Identity $Identity }
        else { Get-MailContact }
    } catch { Write-Error "Erro ao buscar contato(s): $_" }
}

#endregion

#region Exchange Online - Fluxo de Email e Rastreamento

Function Get-S365TransportRule {
    [CmdletBinding()]
    Param ([String] $Identity)
    try {
        if ($Identity) { Get-TransportRule -Identity $Identity }
        else { Get-TransportRule }
    } catch { Write-Error "Erro ao buscar regras de transporte: $_" }
}

Function Set-S365TransportRuleState {
    [CmdletBinding(SupportsShouldProcess=$true)]
    Param (
        [Parameter(Mandatory=$true)][String] $Identity,
        [Parameter(Mandatory=$true)][ValidateSet("Enabled", "Disabled")]$State
    )
    try {
        if ($PSCmdlet.ShouldProcess($Identity, "Alterar Estado da Regra para $State")) {
            if ($State -eq "Enabled") { Enable-TransportRule -Identity $Identity }
            else { Disable-TransportRule -Identity $Identity }
            Write-Host "Regra '$Identity' definida para o estado '$State'." -ForegroundColor Green
        }
    } catch { Write-Error "Erro ao alterar estado da regra '$Identity': $_" }
}

Function Start-S365MessageTrace {
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory=$true)][DateTime] $StartDate,
        [Parameter(Mandatory=$true)][DateTime] $EndDate,
        [String] $SenderAddress,
        [String] $RecipientAddress,
        [String] $Subject,
        [ValidateSet("Pending", "Failed", "Delivered", "Expanded", "Quarantined", "FilteredAsSpam", "GettingStatus")] $Status,
        [String] $MessageId
    )
    Write-Host "Iniciando rastreamento de mensagens (últimos 10 dias)..." -ForegroundColor Yellow
    $params = @{
        StartDate = $StartDate
        EndDate = $EndDate
    }
    if ($SenderAddress) { $params.SenderAddress = $SenderAddress }
    if ($RecipientAddress) { $params.RecipientAddress = $RecipientAddress }
    if ($Subject) { $params.Subject = $Subject }
    if ($PSBoundParameters.ContainsKey('Status')) { $params.Status = $Status }
    if ($MessageId) { $params.MessageId = $MessageId }

    try {
        # Get-MessageTrace é para traces recentes (até 10 dias)
        Get-MessageTrace @params | Select-Object Received, SenderAddress, RecipientAddress, Subject, Status, MessageId, Size, FromIP
    } catch { Write-Error "Erro ao buscar rastreamento de mensagens: $_. Para traces > 10 dias, use Start-HistoricalSearch." }
}

Function Start-S365HistoricalMessageTrace {
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory=$true)][DateTime] $StartDate,
        [Parameter(Mandatory=$true)][DateTime] $EndDate,
        [Parameter(Mandatory=$true)][String] $ReportTitle,
        [Parameter(Mandatory=$true)][String] $NotifyAddress, # Email para receber o relatório
        [String] $SenderAddress,
        [String] $RecipientAddress,
        [String] $Subject,
        [String] $MessageId,
        [ValidateSet("Summary", "Message")]$ReportType = "Summary"
    )
    Write-Host "Iniciando rastreamento histórico de mensagens (pode levar horas)..." -ForegroundColor Yellow
    $params = @{
        StartDate = $StartDate
        EndDate = $EndDate
        ReportTitle = $ReportTitle
        NotifyAddress = $NotifyAddress
        ReportType = $ReportType
    }
    if ($SenderAddress) { $params.SenderAddress = $SenderAddress }
    if ($RecipientAddress) { $params.RecipientAddress = $RecipientAddress }
    if ($Subject) { $params.Subject = $Subject }
    if ($MessageId) { $params.MessageId = $MessageId }

    try {
        $search = Start-HistoricalSearch @params
        Write-Host "Rastreamento histórico iniciado com sucesso. Título: '$ReportTitle'. ID: $($search.JobId). Você será notificado em '$NotifyAddress'." -ForegroundColor Green
        return $search
    } catch { Write-Error "Erro ao iniciar rastreamento histórico: $_" }
}

Function Get-S365HistoricalMessageTrace {
    [CmdletBinding()]
    Param(
        [String] $JobId
    )
    Write-Host "Buscando status de rastreamentos históricos..." -ForegroundColor Yellow
    try {
        if ($JobId) { Get-HistoricalSearch -JobId $JobId }
        else { Get-HistoricalSearch } # Lista todos os searches
    } catch { Write-Error "Erro ao buscar rastreamento histórico: $_" }
}

#endregion

#region Relatórios e Auditoria (Básico)

Function Get-S365SignInActivity {
     [CmdletBinding()]
    Param (
        [Parameter(Mandatory=$true)][String] $UserPrincipalName
    )
    try {
        Get-MgUser -UserId $UserPrincipalName -Property 'SignInActivity' | Select-Object -ExpandProperty 'SignInActivity'
    } catch { Write-Error "Erro ao buscar atividade de login para '$UserPrincipalName': $_" }
}

Function Get-S365LastLogonTime {
    [CmdletBinding()]
    Param (
        [Switch] $All,
        [Int32] $DaysInactive = 90
    )
    try {
        Write-Host "Buscando estatísticas de todas as caixas. Isso pode levar bastante tempo..." -ForegroundColor Yellow
        $mailboxes = Get-Mailbox -ResultSize Unlimited -Filter {RecipientTypeDetails -ne "DiscoveryMailbox"}
        $i = 0
        $total = $mailboxes.Count
        $stats = $mailboxes | ForEach-Object {
            $i++
            Write-Progress -Activity "Buscando Estatísticas" -Status "Processando: $($_.UserPrincipalName) ($i de $total)" -PercentComplete ($i / $total * 100)
            Get-MailboxStatistics -Identity $_.UserPrincipalName -ErrorAction SilentlyContinue
        }

        if ($All) {
             $stats | Select-Object DisplayName, UserPrincipalName, LastLogonTime, TotalItemSize
        } else {
             $threshold = (Get-Date).AddDays(-$DaysInactive)
             Write-Host "Filtrando por usuários inativos há mais de $DaysInactive dias..."
             $stats | Where-Object { $_.LastLogonTime -lt $threshold -or $_.LastLogonTime -eq $null } | Select-Object DisplayName, UserPrincipalName, LastLogonTime, TotalItemSize
        }
    } catch { Write-Error "Erro ao buscar último logon: $_" }
}

Function Search-S365AuditLog {
    [CmdletBinding()]
    Param (
        [Parameter(Mandatory=$true)][DateTime] $StartDate,
        [Parameter(Mandatory=$true)][DateTime] $EndDate,
        [String] $UserIds,
        [String] $Operations, # Ex: "MailboxLogin", "Update user."
        [Int32] $ResultSize = 1000
    )
    Write-Host "Pesquisando Log de Auditoria Unificado (pode demorar)..." -ForegroundColor Yellow
    try {
        Search-UnifiedAuditLog -StartDate $StartDate -EndDate $EndDate -UserIds $UserIds -Operations $Operations -ResultSize $ResultSize | Select-Object CreationDate, UserIds, Operations, AuditData | Sort-Object CreationDate -Descending
    } catch { Write-Error "Erro ao pesquisar o log de auditoria: $_" }
}

#endregion

#region Exemplos de Uso (Descomente para usar)

<#

# --- PASSO 1: Conectar ---
# Connect-Super365Services

# --- Exemplos de Usuários (Graph) ---
# Get-S365User -Identity "alguem@seudominio.com"
# Get-S365User -Filter "startswith(displayName,'Joao')"
# Get-S365User -All | Out-GridView # Ver todos em uma grade
# $senha = ConvertTo-SecureString "SenhaSuperForte123!" -AsPlainText -Force
# New-S365User -UserPrincipalName "novo.usuario@seudominio.com" -DisplayName "Novo Usuario" -MailNickname "novo.usuario" -Password $senha -UsageLocation "BR"
# Set-S365User -UserPrincipalName "novo.usuario@seudominio.com" -JobTitle "Analista" -Department "TI" -ManagerUPN "chefe@seudominio.com"
# Reset-S365UserPassword -UserPrincipalName "novo.usuario@seudominio.com" -NewPassword $senha
# Remove-S365User -UserPrincipalName "velho.usuario@seudominio.com" -WhatIf # Simula a remoção
# Get-S365DeletedUser
# Restore-S365DeletedUser -UserId "<ID_DO_USUARIO_EXCLUIDO>"
# Get-S365MfaStatus -UserPrincipalName "usuario@seudominio.com"

# --- Exemplos de Licenças (Graph) ---
# Get-S365AvailableSkus | Out-GridView
# Get-S365UserLicense -UserPrincipalName "alguem@seudominio.com"
# $e3Sku = (Get-S365AvailableSkus | Where-Object { $_.SkuPartNumber -eq "ENTERPRISEPACK" }).SkuId # Exemplo E3, verifique o seu
# Set-S365UserLicense -UserPrincipalName "alguem@seudominio.com" -AddSkuIds $e3Sku

# --- Exemplos de Grupos (Graph) ---
# Get-S365Group -Filter "startswith(displayName,'Vendas')"
# $grupoTI = New-S365Group -DisplayName "Grupo TI Global" -MailNickname "GrupoTIGlobal" -GroupType "M365" -Description "Grupo para TI" -Visibility Private -OwnersUPN "admin@seudominio.com"
# Add-S365GroupMember -GroupId $grupoTI.Id -UserPrincipalName "membro.ti@seudominio.com"
# Get-S365GroupMember -GroupId $grupoTI.Id
# Remove-S365GroupMember -GroupId $grupoTI.Id -UserPrincipalName "membro.ti@seudominio.com"

# --- Exemplos de Teams (Graph) ---
# $meuTime = New-S365Team -DisplayName "Projeto Alpha" -Description "Time para o Projeto Alpha" -OwnerUPN "admin@seudominio.com"
# Get-S365Team -DisplayName "Projeto Alpha"
# Add-S365TeamMember -TeamId $meuTime.Id -UserPrincipalName "membro1@seudominio.com"
# Get-S365TeamChannel -TeamId $meuTime.Id

# --- Exemplos de Caixas de Correio (Exchange) ---
# Get-S365Mailbox -Identity "caixa@seudominio.com"
# Get-S365Mailbox -Shared -All
# Set-S365MailboxQuota -Identity "caixa@seudominio.com" -Quota "80GB"
# Set-S365MailboxForwarding -Identity "saindo@seudominio.com" -ForwardingSmtpAddress "fica@seudominio.com"
# Get-S365MailboxStatistics -Identity "caixa@seudominio.com"
# New-S365SharedMailbox -Name "Financeiro Compartilhado"
# Set-S365LitigationHold -Identity "importante@seudominio.com" -Enable -DurationDays 3650 # 10 anos

# --- Exemplos de Permissões (Exchange) ---
# Add-S365MailboxPermission -Identity "compartilhada@seudominio.com" -User "usuario@seudominio.com" -AccessRights "FullAccess"
# Add-S365RecipientPermission -Identity "compartilhada@seudominio.com" -Trustee "usuario@seudominio.com" -AccessRights "SendAs"
# Add-S365MailboxFolderPermission -Identity "sala.reuniao@seudominio.com" -User "reservas@seudominio.com" -AccessRights "Editor" -FolderPath "\Calendário"
# Remove-S365MailboxFolderPermission -Identity "sala.reuniao@seudominio.com" -User "reservas@seudominio.com" -FolderPath "\Calendário" -WhatIf

# --- Exemplos de Fluxo de Email (Exchange) ---
# Get-S365TransportRule
# Set-S365TransportRuleState -Identity "Minha Regra de Spam" -State "Disabled" -WhatIf # Simula
# Start-S365MessageTrace -StartDate (Get-Date).AddDays(-1) -EndDate (Get-Date) -RecipientAddress "destino@seudominio.com"
# Start-S365HistoricalMessageTrace -StartDate (Get-Date).AddDays(-30) -EndDate (Get-Date).AddDays(-29) -ReportTitle "Trace_Maio_Externo" -NotifyAddress "admin@seudominio.com" -SenderAddress "alguem@externo.com"

# --- Exemplos de Relatórios ---
# Get-S365SignInActivity -UserPrincipalName "suspeito@seudominio.com"
# Get-S365LastLogonTime -DaysInactive 120 | Export-Csv -Path C:\Reports\InactiveUsers.csv -NoTypeInformation -Encoding UTF8
# $start = (Get-Date).AddDays(-7)
# $end = Get-Date
# Search-S365AuditLog -StartDate $start -EndDate $end -Operations "Remove-MailboxPermission" | Out-GridView

# --- PASSO FINAL: Desconectar ---
# Disconnect-Super365Services

#>

#endregion

Write-Host ""
Write-Host "*****************************************************" -ForegroundColor White -BackgroundColor DarkBlue
Write-Host "*** Módulo SuperAdmin365 (v3.2) Carregado!   ***" -ForegroundColor White -BackgroundColor DarkBlue
Write-Host "*****************************************************" -ForegroundColor White -BackgroundColor DarkBlue
Write-Host ""
Write-Host "Use 'Connect-Super365Services' para iniciar a conexão." -ForegroundColor Cyan
Write-Host "Lembre-se: SALVE este código como .ps1 e execute-o usando '. .\SuperAdmin365.ps1'." -ForegroundColor Yellow
Write-Host "NÃO COLE o script diretamente no console." -ForegroundColor Red
Write-Host "Use 'Disconnect-Super365Services' ao terminar." -ForegroundColor Yellow
