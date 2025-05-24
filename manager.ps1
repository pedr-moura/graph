$Global:ErrorActionPreference = "Stop"

Function Install-RequiredModule {
    Param (
        [String]$ModuleName
    )
    if (-not (Get-Module -ListAvailable -Name $ModuleName)) {
        Write-Host "Módulo '$ModuleName' não encontrado. Tentando instalar via PSGallery..." -ForegroundColor Yellow
        try {
            Install-Module $ModuleName -AllowClobber -Force -Scope CurrentUser -Repository PSGallery -ErrorAction Stop
            Write-Host "Módulo '$ModuleName' instalado para o usuário atual com sucesso." -ForegroundColor Green
        } catch {
            Write-Warning "Falha ao instalar para o usuário atual. Tentando como Administrador..."
            try {
                Install-Module $ModuleName -AllowClobber -Force -Scope AllUsers -Repository PSGallery -ErrorAction Stop
                Write-Host "Módulo '$ModuleName' instalado para todos os usuários com sucesso." -ForegroundColor Green
            } catch {
                Write-Error "FALHA CRÍTICA ao instalar '$ModuleName': $_."
                throw "Instalação de Módulo Essencial Falhou: $ModuleName"
            }
        }
    } else {
        Write-Host "Módulo '$ModuleName' já está disponível." -ForegroundColor Green
    }
    try {
        Import-Module $ModuleName -ErrorAction Stop
    } catch {
         Write-Error "Falha ao importar o módulo '$ModuleName': $_"
         throw "Importação de Módulo Essencial Falhou: $ModuleName"
    }
}

Function Connect-Super365Services {
    [CmdletBinding()]
    Param (
        [Parameter(Mandatory=$false)]
        [String] $UserPrincipalName,
        [Parameter(Mandatory=$false)]
        [String[]] $GraphScopes = @(
            "User.ReadWrite.All", "Group.ReadWrite.All", "Directory.ReadWrite.All",
            "MailboxSettings.ReadWrite", "Mail.ReadWrite", "Mail.Send",
            "Reports.Read.All", "AuditLog.Read.All", "Policy.ReadWrite.All",
            "DeviceManagementManagedDevices.ReadWrite.All", "Team.ReadBasic.All", "TeamMember.ReadWrite.All",
            "TeamSettings.ReadWrite.All", "Channel.ReadWrite.All", "TeamsApp.ReadWrite.All",
            "RoleManagement.ReadWrite.Directory", "Application.ReadWrite.All", "ServiceHealth.Read.All",
            "LicenseAssignment.ReadWrite.All", "Chat.ReadWrite.All", "Sites.ReadWrite.All",
            "UserAuthenticationMethod.ReadWrite.All"
        )
    )

    Write-Host "--- Iniciando Conexão SuperAdmin365 ---" -ForegroundColor Cyan
    Write-Host "[Passo 1/3] Verificando e Instalando Módulos Essenciais..." -ForegroundColor Yellow
    Install-RequiredModule -ModuleName "ExchangeOnlineManagement"
    Install-RequiredModule -ModuleName "Microsoft.Graph"
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
    } catch { Write-Warning "Não foi possível desconectar do Exchange Online." }
    try {
        Disconnect-MgGraph -ErrorAction SilentlyContinue
        Write-Host "- Sessão Microsoft Graph encerrada." -ForegroundColor Gray
    } catch { Write-Warning "Não foi possível desconectar do Microsoft Graph." }
    Write-Host "Desconexão concluída." -ForegroundColor Green
}

Function Get-S365User {
    [CmdletBinding()]
    Param ( [String] $Identity, [Switch] $All, [String] $Filter, [String[]] $Select = @("Id", "DisplayName", "UserPrincipalName", "Mail", "JobTitle", "Department", "AccountEnabled", "CreatedDateTime", "LastPasswordChangeDateTime", "SignInActivity", "UsageLocation", "Manager") )
    try { $params = @{ Select = $Select }; if ($All) { Get-MgUser @params -All } elseif ($Filter) { Get-MgUser @params -Filter $Filter -ConsistencyLevel eventual -CountVariable countVar } elseif ($Identity) { Get-MgUser @params -UserId $Identity } else { Write-Warning "Especifique -Identity, -Filter ou -All." } } catch { Write-Error "Erro ao buscar usuário(s): $_" }
}

Function New-S365User {
    [CmdletBinding()]
    Param ( [Parameter(Mandatory=$true)][String] $UserPrincipalName, [Parameter(Mandatory=$true)][String] $DisplayName, [Parameter(Mandatory=$true)][String] $MailNickname, [Parameter(Mandatory=$true)][System.Security.SecureString] $Password, [Parameter(Mandatory=$true)][String] $UsageLocation, [String] $GivenName, [String] $Surname, [String] $JobTitle, [String] $Department, [Switch] $ForceChangePasswordNextSignIn = $true, [Switch] $AccountEnabled = $true )
    $params = @{ UserPrincipalName = $UserPrincipalName; DisplayName = $DisplayName; MailNickname = $MailNickname; PasswordProfile = @{ ForceChangePasswordNextSignIn = $ForceChangePasswordNextSignIn; Password = $Password }; AccountEnabled = $AccountEnabled; UsageLocation = $UsageLocation }; if ($GivenName) { $params.Add("GivenName", $GivenName) }; if ($Surname) { $params.Add("Surname", $Surname) }; if ($JobTitle) { $params.Add("JobTitle", $JobTitle) }; if ($Department) { $params.Add("Department", $Department) }
    try { $newUser = New-MgUser -BodyParameter $params; Write-Host "Usuário '$UserPrincipalName' (ID: $($newUser.Id)) criado com sucesso." -ForegroundColor Green; return $newUser } catch { Write-Error "Erro ao criar usuário: $_" }
}

Function Set-S365User {
    [CmdletBinding()]
    Param ( [Parameter(Mandatory=$true)][String] $UserPrincipalName, [String] $JobTitle, [String] $Department, [String] $OfficeLocation, [String] $MobilePhone, [String] $ManagerUPN, [String] $UsageLocation, [Switch] $EnableAccount, [Switch] $DisableAccount )
    $params = @{}; if ($JobTitle) { $params.Add("JobTitle", $JobTitle) }; if ($Department) { $params.Add("Department", $Department) }; if ($OfficeLocation) { $params.Add("OfficeLocation", $OfficeLocation) }; if ($MobilePhone) { $params.Add("MobilePhone", $MobilePhone) }; if ($UsageLocation) { $params.Add("UsageLocation", $UsageLocation) }; if ($PSBoundParameters.ContainsKey('EnableAccount')) { $params.Add("AccountEnabled", $true) }; if ($PSBoundParameters.ContainsKey('DisableAccount')) { $params.Add("AccountEnabled", $false) }
    try { if ($params.Count -gt 0) { Update-MgUser -UserId $UserPrincipalName -BodyParameter $params; Write-Host "Propriedades atualizadas para $UserPrincipalName." -ForegroundColor Green }; if ($ManagerUPN) { $manager = Get-S365User -Identity $ManagerUPN; if ($manager) { Set-MgUserManagerByRef -UserId $UserPrincipalName -AdditionalProperties @{"@odata.id" = "https://graph.microsoft.com/v1.0/users/$($manager.Id)"}; Write-Host "Gerente '$ManagerUPN' definido para $UserPrincipalName." -ForegroundColor Green } else { Write-Warning "Gerente '$ManagerUPN' não encontrado." } } } catch { Write-Error "Erro ao atualizar usuário '$UserPrincipalName': $_" }
}

Function Remove-S365User {
    [CmdletBinding(SupportsShouldProcess=$true, ConfirmImpact='High')]
    Param ( [Parameter(Mandatory=$true)][String] $UserPrincipalName )
    try { if ($PSCmdlet.ShouldProcess($UserPrincipalName, "Remover Usuário")) { Remove-MgUser -UserId $UserPrincipalName; Write-Host "Usuário '$UserPrincipalName' movido para a lixeira." -ForegroundColor Green } } catch { Write-Error "Erro ao remover usuário '$UserPrincipalName': $_" }
}

Function Restore-S365DeletedUser {
     [CmdletBinding()]
    Param ( [Parameter(Mandatory=$true)][String] $UserId )
    try { Restore-MgDirectoryDeletedItem -DirectoryObjectId $UserId; Write-Host "Usuário com ID '$UserId' restaurado com sucesso." -ForegroundColor Green } catch { Write-Error "Erro ao restaurar usuário: $_" }
}

Function Get-S365DeletedUser {
     [CmdletBinding()]
    Param ()
    try { Get-MgDirectoryDeletedItem -DirectoryObjectId "microsoft.graph.user" | Select-Object Id, DisplayName, UserPrincipalName, DeletedDateTime } catch { Write-Error "Erro ao buscar usuários excluídos: $_" }
}

Function Reset-S365UserPassword {
    [CmdletBinding()]
    Param ( [Parameter(Mandatory=$true)][String] $UserPrincipalName, [Parameter(Mandatory=$true)][System.Security.SecureString] $NewPassword, [Switch] $ForceChangePasswordNextSignIn = $true )
    $passwordProfile = @{ ForceChangePasswordNextSignIn = $ForceChangePasswordNextSignIn; Password = $NewPassword }
    try { Update-MgUser -UserId $UserPrincipalName -PasswordProfile $passwordProfile; Write-Host "Senha redefinida para '$UserPrincipalName'." -ForegroundColor Green } catch { Write-Error "Erro ao redefinir senha: $_" }
}

Function Get-S365UserManager {
    [CmdletBinding()]
    Param ([Parameter(Mandatory=$true)][String] $UserPrincipalName)
    try { Get-MgUserManager -UserId $UserPrincipalName | Select-Object Id, DisplayName, UserPrincipalName } catch { Write-Error "Erro ao buscar gerente para '$UserPrincipalName': $_" }
}

Function Get-S365UserDirectReports {
    [CmdletBinding()]
    Param ([Parameter(Mandatory=$true)][String] $UserPrincipalName)
    try { Get-MgUserDirectReport -UserId $UserPrincipalName | Select-Object Id, DisplayName, UserPrincipalName } catch { Write-Error "Erro ao buscar subordinados diretos para '$UserPrincipalName': $_" }
}

Function Get-S365MfaStatus {
    [CmdletBinding()]
    Param ([Parameter(Mandatory=$true)][String] $UserPrincipalName)
    try { $methods = Get-MgUserAuthenticationMethod -UserId $UserPrincipalName; if ($methods) { $methods | ForEach-Object { $type = $_.AdditionalProperties.'@odata.type'.Split('.')[-1]; Write-Host "- Tipo: $type" }; if (($methods | Where-Object { $_.AdditionalProperties.'@odata.type' -like "*PhoneAuthenticationMethod*" -or $_.AdditionalProperties.'@odata.type' -like "*MicrosoftAuthenticatorAuthenticationMethod*" })) { Write-Host "=> Status: MFA ATIVO." -ForegroundColor Green } else { Write-Host "=> Status: MFA INATIVO." -ForegroundColor Yellow } } else { Write-Host "=> Status: NENHUM método registrado." -ForegroundColor Red } } catch { Write-Error "Erro ao buscar status MFA: $_" }
}

Function Get-S365AvailableSkus {
    [CmdletBinding()]
    Param ()
    try { Get-MgSubscribedSku | Select-Object SkuId, SkuPartNumber, ConsumedUnits, PrepaidUnits } catch { Write-Error "Erro ao buscar SKUs: $_" }
}

Function Get-S365UserLicense {
    [CmdletBinding()]
    Param ( [Parameter(Mandatory=$true)][String] $UserPrincipalName )
    try { Get-MgUserLicenseDetail -UserId $UserPrincipalName | Select-Object SkuPartNumber, ServicePlans } catch { Write-Error "Erro ao buscar licenças de '$UserPrincipalName': $_" }
}

Function Set-S365UserLicense {
    [CmdletBinding()]
    Param ( [Parameter(Mandatory=$true)][String] $UserPrincipalName, [String[]] $AddSkuIds, [String[]] $RemoveSkuIds )
    try { $user = Get-S365User -Identity $UserPrincipalName -Select UsageLocation; if (-not $user.UsageLocation) { Write-Error "Usuário '$UserPrincipalName' não tem UsageLocation."; return }; $addLicensesObject = @(); if ($AddSkuIds) { $addLicensesObject = $AddSkuIds | ForEach-Object { @{ SkuId = $_ } } }; $params = @{ UserId = $UserPrincipalName; AddLicenses = $addLicensesObject; RemoveLicenses = @($RemoveSkuIds) }; Set-MgUserLicense @params; Write-Host "Licenças atualizadas para '$UserPrincipalName'." -ForegroundColor Green } catch { Write-Error "Erro ao atualizar licenças para '$UserPrincipalName': $_" }
}

Function Get-S365Group {
    [CmdletBinding()]
    Param ( [String] $Identity, [Switch] $All, [String] $Filter, [String[]] $Select = @("Id", "DisplayName", "Mail", "GroupTypes", "SecurityEnabled", "MailEnabled", "Visibility", "Description") )
    try { $params = @{ Select = $Select }; if ($All) { Get-MgGroup @params -All } elseif ($Filter) { Get-MgGroup @params -Filter $Filter -ConsistencyLevel eventual -CountVariable countVar } elseif ($Identity) { Get-MgGroup @params -GroupId $Identity } else { Write-Warning "Especifique -Identity, -Filter ou -All." } } catch { Write-Error "Erro ao buscar grupo(s): $_" }
}

Function Get-S365GroupMember {
    [CmdletBinding()]
    Param ( [Parameter(Mandatory=$true)][String] $GroupId )
    try { Get-MgGroupMember -GroupId $GroupId -All | Select-Object Id, DisplayName, UserPrincipalName, Mail } catch { Write-Error "Erro ao buscar membros do grupo '$GroupId': $_" }
}

Function Add-S365GroupMember {
    [CmdletBinding()]
    Param ( [Parameter(Mandatory=$true)][String] $GroupId, [Parameter(Mandatory=$true)][String] $UserPrincipalName )
    try { $user = Get-S365User -Identity $UserPrincipalName; New-MgGroupMember -GroupId $GroupId -DirectoryObjectId $user.Id; Write-Host "Usuário '$UserPrincipalName' adicionado ao grupo '$GroupId'." -ForegroundColor Green } catch { Write-Error "Erro ao adicionar membro '$UserPrincipalName' ao grupo '$GroupId': $_" }
}

Function Remove-S365GroupMember {
    [CmdletBinding(SupportsShouldProcess=$true)]
    Param ( [Parameter(Mandatory=$true)][String] $GroupId, [Parameter(Mandatory=$true)][String] $UserPrincipalName )
    try { $user = Get-S365User -Identity $UserPrincipalName; if ($PSCmdlet.ShouldProcess("$UserPrincipalName from $GroupId", "Remover Membro")) { Remove-MgGroupMemberByRef -GroupId $GroupId -DirectoryObjectId $user.Id; Write-Host "Usuário '$UserPrincipalName' removido do grupo '$GroupId'." -ForegroundColor Green } } catch { Write-Error "Erro ao remover membro '$UserPrincipalName' do grupo '$GroupId': $_" }
}

Function New-S365Group {
    [CmdletBinding()]
    Param( [Parameter(Mandatory=$true)][String] $DisplayName, [Parameter(Mandatory=$true)][String] $MailNickname, [Parameter(Mandatory=$true)][ValidateSet("M365", "Security")][String] $GroupType, [String] $Description, [String[]] $OwnersUPN, [String[]] $MembersUPN, [Switch] $MailEnabled = ($GroupType -eq "M365"), [Switch] $SecurityEnabled = ($GroupType -eq "Security" -or $GroupType -eq "M365"), [ValidateSet("Public", "Private")][String] $Visibility = "Private" )
    $params = @{ DisplayName = $DisplayName; MailNickname = $MailNickname; MailEnabled = $MailEnabled; SecurityEnabled = $SecurityEnabled; Description = $Description; GroupTypes = @(if ($GroupType -eq "M365") { "Unified" } else { }) }; if ($GroupType -eq "M365") { $params.Visibility = $Visibility }
    try { $newGroup = New-MgGroup -BodyParameter $params; Write-Host "Grupo '$DisplayName' (ID: $($newGroup.Id)) criado." -ForegroundColor Green; if ($OwnersUPN) { $OwnersUPN | ForEach-Object { try { Add-S365GroupOwner -GroupId $newGroup.Id -UserPrincipalName $_ } catch { Write-Warning "Falha ao add '$_' como dono." } } }; if ($MembersUPN) { $MembersUPN | ForEach-Object { try { Add-S365GroupMember -GroupId $newGroup.Id -UserPrincipalName $_ } catch { Write-Warning "Falha ao add '$_' como membro." } } }; return $newGroup } catch { Write-Error "Erro ao criar grupo: $_" }
}

Function Add-S365GroupOwner {
    [CmdletBinding()]
    Param ( [Parameter(Mandatory=$true)][String] $GroupId, [Parameter(Mandatory=$true)][String] $UserPrincipalName )
    try { $user = Get-S365User -Identity $UserPrincipalName; New-MgGroupOwnerByRef -GroupId $GroupId -OdataId "https://graph.microsoft.com/v1.0/users/$($user.Id)"; Write-Host "Usuário '$UserPrincipalName' adicionado como dono do grupo '$GroupId'." -ForegroundColor Green } catch { Write-Error "Erro ao adicionar dono '$UserPrincipalName': $_" }
}

Function Get-S365GroupOwner {
    [CmdletBinding()]
    Param ( [Parameter(Mandatory=$true)][String] $GroupId )
    try { Get-MgGroupOwner -GroupId $GroupId -All | Select-Object Id, DisplayName, UserPrincipalName } catch { Write-Error "Erro ao buscar donos do grupo '$GroupId': $_" }
}

Function Get-S365Team {
    [CmdletBinding()]
    Param ( [String] $DisplayName, [Switch] $All )
    try { if ($DisplayName) { Get-MgTeam -Filter "DisplayName eq '$DisplayName'" } elseif ($All) { Get-MgTeam -All } else { Get-MgTeam } } catch { Write-Error "Erro ao buscar Teams: $_" }
}

Function New-S365Team {
    [CmdletBinding()]
    Param ( [Parameter(Mandatory=$true)][String] $DisplayName, [String] $Description, [String] $OwnerUPN )
    try { $template = "https://graph.microsoft.com/v1.0/teamsTemplates('standard')"; $params = @{ "template@odata.bind" = $template; DisplayName = $DisplayName; Description = $Description }; $newTeam = New-MgTeam -BodyParameter $params; Write-Host "Team '$DisplayName' (ID: $($newTeam.Id)) criado. Aguarde..." -ForegroundColor Yellow; Start-Sleep -Seconds 15; if ($OwnerUPN) { try { Add-S365TeamMember -TeamId $newTeam.Id -UserPrincipalName $OwnerUPN -IsOwner } catch { Write-Warning "Falha ao add '$OwnerUPN' como dono." } } else { Write-Warning "Team criado sem dono." }; Write-Host "Team '$DisplayName' provisionado." -ForegroundColor Green; return $newTeam } catch { Write-Error "Erro ao criar Team: $_" }
}

Function Get-S365TeamChannel {
    [CmdletBinding()]
    Param ([Parameter(Mandatory=$true)][String] $TeamId)
    try { Get-MgTeamChannel -TeamId $TeamId -All | Select-Object Id, DisplayName, WebUrl, MembershipType } catch { Write-Error "Erro ao buscar canais para Team '$TeamId': $_" }
}

Function Add-S365TeamMember {
    [CmdletBinding()]
    Param ( [Parameter(Mandatory=$true)][String] $TeamId, [Parameter(Mandatory=$true)][String] $UserPrincipalName, [Switch] $IsOwner )
    try { $user = Get-S365User -Identity $UserPrincipalName; $params = @{ "@odata.type" = "#microsoft.graph.aadUserConversationMember"; "user@odata.bind" = "https://graph.microsoft.com/v1.0/users('$($user.Id)')"; Roles = @(if ($IsOwner) { "owner" } else { "member" }) }; New-MgTeamMember -TeamId $TeamId -BodyParameter $params; $role = if ($IsOwner) { "Proprietário" } else { "Membro" }; Write-Host "Usuário '$UserPrincipalName' adicionado ao Team '$TeamId' como $role." -ForegroundColor Green } catch { Write-Error "Erro ao adicionar membro '$UserPrincipalName' ao Team '$TeamId': $_" }
}

Function Remove-S365TeamMember {
    [CmdletBinding(SupportsShouldProcess=$true)]
    Param ( [Parameter(Mandatory=$true)][String] $TeamId, [Parameter(Mandatory=$true)][String] $UserPrincipalName )
    try { $user = Get-S365User -Identity $UserPrincipalName; $member = Get-MgTeamMember -TeamId $TeamId -Filter "UserId eq '$($user.Id)'"; if ($member) { if ($PSCmdlet.ShouldProcess("$UserPrincipalName from $TeamId", "Remover Membro")) { Remove-MgTeamMember -TeamId $TeamId -ConversationMemberId $member.Id; Write-Host "Usuário '$UserPrincipalName' removido do Team '$TeamId'." -ForegroundColor Green } } else { Write-Warning "Usuário '$UserPrincipalName' não encontrado no Team '$TeamId'." } } catch { Write-Error "Erro ao remover membro '$UserPrincipalName' do Team '$TeamId': $_" }
}

Function Get-S365Mailbox {
    [CmdletBinding()]
    Param ( [String] $Identity, [Switch] $All, [String] $Filter, [Switch] $Shared, [Switch] $Room, [Switch] $Equipment, [Switch] $Archive )
    try { $params = @{}; if ($Shared) { $params.RecipientTypeDetails = "SharedMailbox" } elseif ($Room) { $params.RecipientTypeDetails = "RoomMailbox" } elseif ($Equipment) { $params.RecipientTypeDetails = "EquipmentMailbox" } elseif ($Archive) { $params.Archive = $true }; if ($All) { Get-Mailbox @params -ResultSize Unlimited } elseif ($Filter) { Get-Mailbox @params -Filter $Filter -ResultSize Unlimited } elseif ($Identity) { Get-Mailbox @params -Identity $Identity } else { Get-Mailbox @params -ResultSize 500 } } catch { Write-Error "Erro ao buscar caixa(s) de correio: $_" }
}

Function Set-S365MailboxQuota {
    [CmdletBinding()]
    Param ( [Parameter(Mandatory=$true)][String] $Identity, [Parameter(Mandatory=$true)][String] $Quota )
    try { $quotaValue = [Microsoft.Exchange.Data.ByteQuantifiedSize]::Parse($Quota); $warningQuota = ([Microsoft.Exchange.Data.ByteQuantifiedSize]($quotaValue.ToBytes() * 0.9)).ToString(); Set-Mailbox -Identity $Identity -ProhibitSendReceiveQuota $Quota -IssueWarningQuota $warningQuota; Write-Host "Quota '$Identity' definida para $Quota." -ForegroundColor Green } catch { Write-Error "Erro ao definir quota para '$Identity': $_" }
}

Function Set-S365MailboxForwarding {
    [CmdletBinding()]
    Param ( [Parameter(Mandatory=$true)][String] $Identity, [String] $ForwardingSmtpAddress, [Switch] $DeliverToMailboxAndForward = $true, [Switch] $RemoveForwarding )
    try { if ($RemoveForwarding) { Set-Mailbox -Identity $Identity -ForwardingSmtpAddress $null -DeliverToMailboxAndForward $false; Write-Host "Encaminhamento removido para '$Identity'." -ForegroundColor Green } else { if (-not $ForwardingSmtpAddress) { Write-Error "-ForwardingSmtpAddress é necessário." ; return }; Set-Mailbox -Identity $Identity -ForwardingSmtpAddress $ForwardingSmtpAddress -DeliverToMailboxAndForward $DeliverToMailboxAndForward; Write-Host "Encaminhamento configurado para '$Identity' -> '$ForwardingSmtpAddress'." -ForegroundColor Green } } catch { Write-Error "Erro ao configurar encaminhamento para '$Identity': $_" }
}

Function Get-S365MailboxStatistics {
    [CmdletBinding()]
    Param ( [Parameter(Mandatory=$true)][String] $Identity )
    try { Get-MailboxStatistics -Identity $Identity | Select-Object DisplayName, ItemCount, TotalItemSize, LastLogonTime, LastUserActionTime } catch { Write-Error "Erro ao buscar estatísticas para '$Identity': $_" }
}

Function New-S365SharedMailbox {
    [CmdletBinding()]
    Param ( [Parameter(Mandatory=$true)][String] $Name, [String] $Alias = ($Name -replace "\s", ""), [String] $PrimarySmtpAddress = "$Alias@$(Get-AcceptedDomain | Where-Object { $_.IsDefault -eq $true } | Select-Object -ExpandProperty DomainName)" )
    try { New-Mailbox -Shared -Name $Name -DisplayName $Name -Alias $Alias -PrimarySmtpAddress $PrimarySmtpAddress; Write-Host "Caixa compartilhada '$Name' ($PrimarySmtpAddress) criada." -ForegroundColor Green } catch { Write-Error "Erro ao criar caixa compartilhada '$Name': $_" }
}

Function Set-S365LitigationHold {
    [CmdletBinding()]
    Param ( [Parameter(Mandatory=$true)][String] $Identity, [Switch] $Enable, [Switch] $Disable, [Int32] $DurationDays )
    try { if ($Enable) { $params = @{ Identity = $Identity; LitigationHoldEnabled = $true }; if ($DurationDays) { $params.LitigationHoldDuration = $DurationDays }; Set-Mailbox @params; Write-Host "Litigation Hold HABILITADO para '$Identity'." -ForegroundColor Green } elseif ($Disable) { Set-Mailbox -Identity $Identity -LitigationHoldEnabled = $false; Write-Host "Litigation Hold DESABILITADO para '$Identity'." -ForegroundColor Green } else { Write-Warning "Use -Enable ou -Disable." } } catch { Write-Error "Erro ao configurar Litigation Hold para '$Identity': $_" }
}

Function Add-S365MailboxPermission {
    [CmdletBinding()]
    Param ( [Parameter(Mandatory=$true)][String] $Identity, [Parameter(Mandatory=$true)][String] $User, [Parameter(Mandatory=$true)][ValidateSet("FullAccess", "ExternalAccount", "DeleteItem", "ReadPermission", "ChangePermission", "ChangeOwner")]$AccessRights = "FullAccess", [Switch] $AutoMapping = $true )
    try { Add-MailboxPermission -Identity $Identity -User $User -AccessRights $AccessRights -InheritanceType All -AutoMapping:$AutoMapping; Write-Host "Permissão '$AccessRights' concedida a '$User' em '$Identity'." -ForegroundColor Green } catch { Write-Error "Erro ao adicionar permissão '$AccessRights' a '$User' em '$Identity': $_" }
}

Function Remove-S365MailboxPermission {
    [CmdletBinding(SupportsShouldProcess=$true)]
    Param ( [Parameter(Mandatory=$true)][String] $Identity, [Parameter(Mandatory=$true)][String] $User, [Parameter(Mandatory=$true)][ValidateSet("FullAccess", "ExternalAccount", "DeleteItem", "ReadPermission", "ChangePermission", "ChangeOwner")]$AccessRights = "FullAccess" )
    try { if ($PSCmdlet.ShouldProcess("$User on $Identity", "Remover Permissão ($AccessRights)")) { Remove-MailboxPermission -Identity $Identity -User $User -AccessRights $AccessRights -InheritanceType All -Confirm:$false; Write-Host "Permissão '$AccessRights' removida de '$User' em '$Identity'." -ForegroundColor Green } } catch { Write-Error "Erro ao remover permissão '$AccessRights' de '$User' em '$Identity': $_" }
}

Function Add-S365RecipientPermission {
    [CmdletBinding()]
    Param ( [Parameter(Mandatory=$true)][String] $Identity, [Parameter(Mandatory=$true)][String] $Trustee, [Parameter(Mandatory=$true)][ValidateSet("SendAs", "SendOnBehalf")]$AccessRights = "SendAs" )
    try { if ($AccessRights -eq "SendAs") { Add-RecipientPermission -Identity $Identity -Trustee $Trustee -AccessRights SendAs; Write-Host "'SendAs' concedida a '$Trustee' em '$Identity'." -ForegroundColor Green } elseif ($AccessRights -eq "SendOnBehalf") { Set-Mailbox -Identity $Identity -GrantSendOnBehalfTo @{Add="$Trustee"}; Write-Host "'SendOnBehalf' concedida a '$Trustee' em '$Identity'." -ForegroundColor Green } } catch { Write-Error "Erro ao adicionar permissão '$AccessRights': $_" }
}

Function Remove-S365RecipientPermission {
    [CmdletBinding(SupportsShouldProcess=$true)]
    Param ( [Parameter(Mandatory=$true)][String] $Identity, [Parameter(Mandatory=$true)][String] $Trustee, [Parameter(Mandatory=$true)][ValidateSet("SendAs", "SendOnBehalf")]$AccessRights = "SendAs" )
    try { if ($PSCmdlet.ShouldProcess("$Trustee on $Identity", "Remover Permissão ($AccessRights)")) { if ($AccessRights -eq "SendAs") { Remove-RecipientPermission -Identity $Identity -Trustee $Trustee -AccessRights SendAs -Confirm:$false; Write-Host "'SendAs' removida de '$Trustee' em '$Identity'." -ForegroundColor Green } elseif ($AccessRights -eq "SendOnBehalf") { Set-Mailbox -Identity $Identity -GrantSendOnBehalfTo @{Remove="$Trustee"}; Write-Host "'SendOnBehalf' removida de '$Trustee' em '$Identity'." -ForegroundColor Green } } } catch { Write-Error "Erro ao remover permissão '$AccessRights': $_" }
}

Function Get-S365MailboxFolderPermission {
     [CmdletBinding()]
    Param ( [Parameter(Mandatory=$true)][String] $Identity, [String] $FolderPath = "\Calendário" )
    try { if ($FolderPath -eq "\Calendário" -or $FolderPath -eq "\Calendar") { try { $calendarPath = Get-MailboxFolderStatistics -Identity $Identity -FolderScope Calendar | Select-Object -First 1 | Select-Object -ExpandProperty FolderPath; if ($calendarPath) { $FolderPath = $calendarPath } } catch { Write-Warning "Usando '$FolderPath'." } }; $IdentityPath = $Identity + ":" + $FolderPath; Get-MailboxFolderPermission -Identity $IdentityPath } catch { Write-Error "Erro ao buscar permissões da pasta '$FolderPath' para '$Identity': $_" }
}

Function Add-S365MailboxFolderPermission {
     [CmdletBinding()]
    Param ( [Parameter(Mandatory=$true)][String] $Identity, [Parameter(Mandatory=$true)][String] $User, [Parameter(Mandatory=$true)][ValidateSet("Owner", "PublishingEditor", "Editor", "PublishingAuthor", "Author", "NoneditingAuthor", "Reviewer", "Contributor", "AvailabilityOnly", "LimitedDetails")]$AccessRights, [String] $FolderPath = "\Calendário" )
    try { $IdentityPath = $Identity + ":" + $FolderPath; Add-MailboxFolderPermission -Identity $IdentityPath -User $User -AccessRights $AccessRights; Write-Host "Permissão '$AccessRights' concedida a '$User' na pasta '$FolderPath' de '$Identity'." -ForegroundColor Green } catch { $IdentityPath = $Identity + ":" + $FolderPath; Write-Error "Erro ao adicionar permissão '$AccessRights' a '$User' em '$IdentityPath': $_" }
}

Function Remove-S365MailboxFolderPermission {
     [CmdletBinding(SupportsShouldProcess=$true)]
    Param ( [Parameter(Mandatory=$true)][String] $Identity, [Parameter(Mandatory=$true)][String] $User, [String] $FolderPath = "\Calendário" )
    try { $IdentityPath = $Identity + ":" + $FolderPath; if ($PSCmdlet.ShouldProcess("$User on $IdentityPath", "Remover Permissão de Pasta")) { Remove-MailboxFolderPermission -Identity $IdentityPath -User $User -Confirm:$false; Write-Host "Permissões removidas de '$User' na pasta '$FolderPath' de '$Identity'." -ForegroundColor Green } } catch { $IdentityPath = $Identity + ":" + $FolderPath; Write-Error "Erro ao remover permissão de '$User' em '$IdentityPath': $_" }
}

Function Get-S365DistributionGroup {
    [CmdletBinding()]
    Param ( [String] $Identity, [Switch] $All )
    try { if ($All) { Get-DistributionGroup -ResultSize Unlimited } elseif ($Identity) { Get-DistributionGroup -Identity $Identity } else { Get-DistributionGroup } } catch { Write-Error "Erro ao buscar grupo(s) de distribuição: $_" }
}

Function Get-S365DistributionGroupMember {
    [CmdletBinding()]
    Param ( [Parameter(Mandatory=$true)][String] $Identity )
    try { Get-DistributionGroupMember -Identity $Identity -ResultSize Unlimited } catch { Write-Error "Erro ao buscar membros do grupo '$Identity': $_" }
}

Function Add-S365DistributionGroupMember {
    [CmdletBinding()]
    Param ( [Parameter(Mandatory=$true)][String] $Identity, [Parameter(Mandatory=$true)][String] $Member )
    try { Add-DistributionGroupMember -Identity $Identity -Member $Member; Write-Host "Membro '$Member' adicionado ao grupo '$Identity'." -ForegroundColor Green } catch { Write-Error "Erro ao adicionar membro '$Member' ao grupo '$Identity': $_" }
}

Function Remove-S365DistributionGroupMember {
    [CmdletBinding(SupportsShouldProcess=$true)]
    Param ( [Parameter(Mandatory=$true)][String] $Identity, [Parameter(Mandatory=$true)][String] $Member )
    try { if ($PSCmdlet.ShouldProcess($Member, "Remover Membro $Identity")) { Remove-DistributionGroupMember -Identity $Identity -Member $Member -Confirm:$false; Write-Host "Membro '$Member' removido do grupo '$Identity'." -ForegroundColor Green } } catch { Write-Error "Erro ao remover membro '$Member': $_" }
}

Function Get-S365MailContact {
     [CmdletBinding()]
    Param ( [String] $Identity, [Switch] $All )
    try { if ($All) { Get-MailContact -ResultSize Unlimited } elseif ($Identity) { Get-MailContact -Identity $Identity } else { Get-MailContact } } catch { Write-Error "Erro ao buscar contato(s): $_" }
}

Function Get-S365TransportRule {
    [CmdletBinding()]
    Param ([String] $Identity)
    try { if ($Identity) { Get-TransportRule -Identity $Identity } else { Get-TransportRule } } catch { Write-Error "Erro ao buscar regras de transporte: $_" }
}

Function Set-S365TransportRuleState {
    [CmdletBinding(SupportsShouldProcess=$true)]
    Param ( [Parameter(Mandatory=$true)][String] $Identity, [Parameter(Mandatory=$true)][ValidateSet("Enabled", "Disabled")]$State )
    try { if ($PSCmdlet.ShouldProcess($Identity, "Alterar Estado para $State")) { if ($State -eq "Enabled") { Enable-TransportRule -Identity $Identity } else { Disable-TransportRule -Identity $Identity }; Write-Host "Regra '$Identity' definida para '$State'." -ForegroundColor Green } } catch { Write-Error "Erro ao alterar estado da regra '$Identity': $_" }
}

Function Start-S365MessageTrace {
    [CmdletBinding()]
    Param( [Parameter(Mandatory=$true)][DateTime] $StartDate, [Parameter(Mandatory=$true)][DateTime] $EndDate, [String] $SenderAddress, [String] $RecipientAddress, [String] $Subject, [ValidateSet("Pending", "Failed", "Delivered", "Expanded", "Quarantined", "FilteredAsSpam", "GettingStatus")] $Status, [String] $MessageId )
    $params = @{ StartDate = $StartDate; EndDate = $EndDate }; if ($SenderAddress) { $params.SenderAddress = $SenderAddress }; if ($RecipientAddress) { $params.RecipientAddress = $RecipientAddress }; if ($Subject) { $params.Subject = $Subject }; if ($PSBoundParameters.ContainsKey('Status')) { $params.Status = $Status }; if ($MessageId) { $params.MessageId = $MessageId }
    try { Get-MessageTrace @params | Select-Object Received, SenderAddress, RecipientAddress, Subject, Status, MessageId, Size, FromIP } catch { Write-Error "Erro ao buscar rastreamento: $_" }
}

Function Start-S365HistoricalMessageTrace {
    [CmdletBinding()]
    Param( [Parameter(Mandatory=$true)][DateTime] $StartDate, [Parameter(Mandatory=$true)][DateTime] $EndDate, [Parameter(Mandatory=$true)][String] $ReportTitle, [Parameter(Mandatory=$true)][String] $NotifyAddress, [String] $SenderAddress, [String] $RecipientAddress, [String] $Subject, [String] $MessageId, [ValidateSet("Summary", "Message")]$ReportType = "Summary" )
    $params = @{ StartDate = $StartDate; EndDate = $EndDate; ReportTitle = $ReportTitle; NotifyAddress = $NotifyAddress; ReportType = $ReportType }; if ($SenderAddress) { $params.SenderAddress = $SenderAddress }; if ($RecipientAddress) { $params.RecipientAddress = $RecipientAddress }; if ($Subject) { $params.Subject = $Subject }; if ($MessageId) { $params.MessageId = $MessageId }
    try { $search = Start-HistoricalSearch @params; Write-Host "Rastreamento histórico iniciado. Título: '$ReportTitle'. ID: $($search.JobId)." -ForegroundColor Green; return $search } catch { Write-Error "Erro ao iniciar rastreamento histórico: $_" }
}

Function Get-S365HistoricalMessageTrace {
    [CmdletBinding()]
    Param( [String] $JobId )
    try { if ($JobId) { Get-HistoricalSearch -JobId $JobId } else { Get-HistoricalSearch } } catch { Write-Error "Erro ao buscar rastreamento histórico: $_" }
}

Function Get-S365SignInActivity {
     [CmdletBinding()]
    Param ( [Parameter(Mandatory=$true)][String] $UserPrincipalName )
    try { Get-MgUser -UserId $UserPrincipalName -Property 'SignInActivity' | Select-Object -ExpandProperty 'SignInActivity' } catch { Write-Error "Erro ao buscar atividade de login para '$UserPrincipalName': $_" }
}

Function Get-S365LastLogonTime {
    [CmdletBinding()]
    Param ( [Switch] $All, [Int32] $DaysInactive = 90 )
    try { Write-Host "Buscando estatísticas (pode demorar)..." -ForegroundColor Yellow; $mailboxes = Get-Mailbox -ResultSize Unlimited -Filter {RecipientTypeDetails -ne "DiscoveryMailbox"}; $i = 0; $total = $mailboxes.Count; $stats = $mailboxes | ForEach-Object { $i++; Write-Progress -Activity "Buscando Estatísticas" -Status "Processando: $($_.UserPrincipalName) ($i de $total)" -PercentComplete ($i / $total * 100); Get-MailboxStatistics -Identity $_.UserPrincipalName -ErrorAction SilentlyContinue }; if ($All) { $stats | Select-Object DisplayName, UserPrincipalName, LastLogonTime, TotalItemSize } else { $threshold = (Get-Date).AddDays(-$DaysInactive); Write-Host "Filtrando por $DaysInactive dias inativos..."; $stats | Where-Object { $_.LastLogonTime -lt $threshold -or $_.LastLogonTime -eq $null } | Select-Object DisplayName, UserPrincipalName, LastLogonTime, TotalItemSize } } catch { Write-Error "Erro ao buscar último logon: $_" }
}

Function Search-S365AuditLog {
    [CmdletBinding()]
    Param ( [Parameter(Mandatory=$true)][DateTime] $StartDate, [Parameter(Mandatory=$true)][DateTime] $EndDate, [String] $UserIds, [String] $Operations, [Int32] $ResultSize = 1000 )
    Write-Host "Pesquisando Log de Auditoria (pode demorar)..." -ForegroundColor Yellow
    try { Search-UnifiedAuditLog -StartDate $StartDate -EndDate $EndDate -UserIds $UserIds -Operations $Operations -ResultSize $ResultSize | Select-Object CreationDate, UserIds, Operations, AuditData | Sort-Object CreationDate -Descending } catch { Write-Error "Erro ao pesquisar o log de auditoria: $_" }
}

Write-Host "Módulo SuperAdmin365 (v3.3-iex) pronto. Use 'Connect-Super365Services'." -ForegroundColor Magenta
