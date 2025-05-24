# =========================================================================
#                    SuperAdmin365 TUI (Terminal User Interface)
#                 Designed for iex (irm 'your_url_here')
#                            Version: 4.0-TUI
# =========================================================================

$Global:ErrorActionPreference = "Stop"

# --- Variáveis Globais de Status ---
$Global:IsExchangeConnected = $false
$Global:IsGraphConnected = $false
$Global:GraphUser = "N/A"

# =========================================================================
#                  CORE FUNCTIONS (Seu Script Original)
# =========================================================================

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

    Write-Host "`n--- Iniciando Conexão SuperAdmin365 ---" -ForegroundColor Cyan
    Write-Host "[Passo 1/3] Verificando e Instalando Módulos Essenciais..." -ForegroundColor Yellow
    try {
        Install-RequiredModule -ModuleName "ExchangeOnlineManagement"
        Install-RequiredModule -ModuleName "Microsoft.Graph"
    } catch {
        Write-Error "Erro Crítico na instalação/importação de módulos. Abortando."
        return
    }

    Write-Host "[Passo 2/3] Conectando ao Exchange Online..." -ForegroundColor Yellow
    try {
        Get-PSSession | Where-Object { $_.ConfigurationName -eq 'Microsoft.Exchange' } | Remove-PSSession -Confirm:$false -ErrorAction SilentlyContinue
        if ($PSBoundParameters.ContainsKey('UserPrincipalName')) {
            Connect-ExchangeOnline -UserPrincipalName $UserPrincipalName -ShowProgress $true
        } else {
            Connect-ExchangeOnline -ShowProgress $true
        }
        Write-Host "=> Conectado ao Exchange Online com sucesso." -ForegroundColor Green
        $Global:IsExchangeConnected = $true
    } catch {
        Write-Error "Falha ao conectar ao Exchange Online: $_"
        $Global:IsExchangeConnected = $false
    }

    Write-Host "[Passo 3/3] Conectando ao Microsoft Graph..." -ForegroundColor Yellow
    try {
        Disconnect-MgGraph -ErrorAction SilentlyContinue
        Connect-MgGraph -Scopes $GraphScopes
        $context = Get-MgContext
        $Global:GraphUser = $context.Account
        Write-Host "=> Conectado ao Microsoft Graph como $($Global:GraphUser)." -ForegroundColor Green
        $Global:IsGraphConnected = $true
    } catch {
        Write-Error "Falha ao conectar ao Microsoft Graph: $_"
        $Global:IsGraphConnected = $false
        $Global:GraphUser = "N/A"
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
        $Global:IsExchangeConnected = $false
    } catch { Write-Warning "Não foi possível desconectar do Exchange Online." }
    try {
        Disconnect-MgGraph -ErrorAction SilentlyContinue
        Write-Host "- Sessão Microsoft Graph encerrada." -ForegroundColor Gray
        $Global:IsGraphConnected = $false
        $Global:GraphUser = "N/A"
    } catch { Write-Warning "Não foi possível desconectar do Microsoft Graph." }
    Write-Host "Desconexão concluída." -ForegroundColor Green
}

# --- COLE TODAS AS SUAS OUTRAS FUNÇÕES S365 (Get-S365User, Set-S365User, etc.) AQUI ---
# ...
# ... (Coloque todas as 50+ funções aqui) ...
# ...
Function Get-S365User { [CmdletBinding()] Param ( [String] $Identity, [Switch] $All, [String] $Filter, [String[]] $Select = @("Id", "DisplayName", "UserPrincipalName", "Mail", "JobTitle", "Department", "AccountEnabled", "CreatedDateTime", "LastPasswordChangeDateTime", "SignInActivity", "UsageLocation", "Manager") ) try { $params = @{ }; if ($Select) { $params.Property = $Select }; if ($All) { Get-MgUser @params -All } elseif ($Filter) { Get-MgUser @params -Filter $Filter -ConsistencyLevel eventual -CountVariable countVar } elseif ($Identity) { Get-MgUser @params -UserId $Identity } else { Write-Warning "Especifique -Identity, -Filter ou -All." } } catch { Write-Error "Erro ao buscar usuário(s): $_" } }
Function New-S365User { [CmdletBinding()] Param ( [Parameter(Mandatory=$true)][String] $UserPrincipalName, [Parameter(Mandatory=$true)][String] $DisplayName, [Parameter(Mandatory=$true)][String] $MailNickname, [Parameter(Mandatory=$true)][System.Security.SecureString] $Password, [Parameter(Mandatory=$true)][String] $UsageLocation, [String] $GivenName, [String] $Surname, [String] $JobTitle, [String] $Department, [Switch] $ForceChangePasswordNextSignIn = $true, [Switch] $AccountEnabled = $true ) $params = @{ UserPrincipalName = $UserPrincipalName; DisplayName = $DisplayName; MailNickname = $MailNickname; PasswordProfile = @{ ForceChangePasswordNextSignIn = $ForceChangePasswordNextSignIn; Password = $Password }; AccountEnabled = $AccountEnabled; UsageLocation = $UsageLocation }; if ($GivenName) { $params.Add("GivenName", $GivenName) }; if ($Surname) { $params.Add("Surname", $Surname) }; if ($JobTitle) { $params.Add("JobTitle", $JobTitle) }; if ($Department) { $params.Add("Department", $Department) }; try { $newUser = New-MgUser -BodyParameter $params; Write-Host "Usuário '$UserPrincipalName' (ID: $($newUser.Id)) criado com sucesso." -ForegroundColor Green; return $newUser } catch { Write-Error "Erro ao criar usuário: $_" } }
Function Set-S365User { [CmdletBinding()] Param ( [Parameter(Mandatory=$true)][String] $UserPrincipalName, [String] $JobTitle, [String] $Department, [String] $OfficeLocation, [String] $MobilePhone, [String] $ManagerUPN, [String] $UsageLocation, [Switch] $EnableAccount, [Switch] $DisableAccount ) $params = @{}; if ($JobTitle) { $params.Add("JobTitle", $JobTitle) }; if ($Department) { $params.Add("Department", $Department) }; if ($OfficeLocation) { $params.Add("OfficeLocation", $OfficeLocation) }; if ($MobilePhone) { $params.Add("MobilePhone", $MobilePhone) }; if ($UsageLocation) { $params.Add("UsageLocation", $UsageLocation) }; if ($PSBoundParameters.ContainsKey('EnableAccount')) { $params.Add("AccountEnabled", $true) }; if ($PSBoundParameters.ContainsKey('DisableAccount')) { $params.Add("AccountEnabled", $false) }; try { if ($params.Count -gt 0) { Update-MgUser -UserId $UserPrincipalName -BodyParameter $params; Write-Host "Propriedades atualizadas para $UserPrincipalName." -ForegroundColor Green }; if ($ManagerUPN) { $manager = Get-S365User -Identity $ManagerUPN; if ($manager) { Set-MgUserManagerByRef -UserId $UserPrincipalName -AdditionalProperties @{"@odata.id" = "https://graph.microsoft.com/v1.0/users/$($manager.Id)"}; Write-Host "Gerente '$ManagerUPN' definido para $UserPrincipalName." -ForegroundColor Green } else { Write-Warning "Gerente '$ManagerUPN' não encontrado." } } } catch { Write-Error "Erro ao atualizar usuário '$UserPrincipalName': $_" } }
Function Remove-S365User { [CmdletBinding(SupportsShouldProcess=$true, ConfirmImpact='High')] Param ( [Parameter(Mandatory=$true)][String] $UserPrincipalName ) try { if ($PSCmdlet.ShouldProcess($UserPrincipalName, "Remover Usuário")) { Remove-MgUser -UserId $UserPrincipalName; Write-Host "Usuário '$UserPrincipalName' movido para a lixeira." -ForegroundColor Green } } catch { Write-Error "Erro ao remover usuário '$UserPrincipalName': $_" } }
Function Restore-S365DeletedUser { [CmdletBinding()] Param ( [Parameter(Mandatory=$true)][String] $UserId ) try { Restore-MgDirectoryDeletedItem -DirectoryObjectId $UserId; Write-Host "Usuário com ID '$UserId' restaurado com sucesso." -ForegroundColor Green } catch { Write-Error "Erro ao restaurar usuário: $_" } }
Function Get-S365DeletedUser { [CmdletBinding()] Param () try { Get-MgDirectoryDeletedItem -DirectoryObjectId "microsoft.graph.user" | Select-Object Id, DisplayName, UserPrincipalName, DeletedDateTime } catch { Write-Error "Erro ao buscar usuários excluídos: $_" } }
Function Reset-S365UserPassword { [CmdletBinding()] Param ( [Parameter(Mandatory=$true)][String] $UserPrincipalName, [Parameter(Mandatory=$true)][System.Security.SecureString] $NewPassword, [Switch] $ForceChangePasswordNextSignIn = $true ) $passwordProfile = @{ ForceChangePasswordNextSignIn = $ForceChangePasswordNextSignIn; Password = $NewPassword }; try { Update-MgUser -UserId $UserPrincipalName -PasswordProfile $passwordProfile; Write-Host "Senha redefinida para '$UserPrincipalName'." -ForegroundColor Green } catch { Write-Error "Erro ao redefinir senha: $_" } }
Function Get-S365UserManager { [CmdletBinding()] Param ([Parameter(Mandatory=$true)][String] $UserPrincipalName) try { Get-MgUserManager -UserId $UserPrincipalName | Select-Object Id, DisplayName, UserPrincipalName } catch { Write-Error "Erro ao buscar gerente para '$UserPrincipalName': $_" } }
Function Get-S365UserDirectReports { [CmdletBinding()] Param ([Parameter(Mandatory=$true)][String] $UserPrincipalName) try { Get-MgUserDirectReport -UserId $UserPrincipalName | Select-Object Id, DisplayName, UserPrincipalName } catch { Write-Error "Erro ao buscar subordinados diretos para '$UserPrincipalName': $_" } }
Function Get-S365MfaStatus { [CmdletBinding()] Param ([Parameter(Mandatory=$true)][String] $UserPrincipalName) try { $methods = Get-MgUserAuthenticationMethod -UserId $UserPrincipalName; if ($methods) { $methods | ForEach-Object { $type = $_.AdditionalProperties.'@odata.type'.Split('.')[-1]; Write-Host "- Tipo: $type" }; if (($methods | Where-Object { $_.AdditionalProperties.'@odata.type' -like "*PhoneAuthenticationMethod*" -or $_.AdditionalProperties.'@odata.type' -like "*MicrosoftAuthenticatorAuthenticationMethod*" })) { Write-Host "=> Status: MFA ATIVO." -ForegroundColor Green } else { Write-Host "=> Status: MFA INATIVO." -ForegroundColor Yellow } } else { Write-Host "=> Status: NENHUM método registrado." -ForegroundColor Red } } catch { Write-Error "Erro ao buscar status MFA: $_" } }
Function Get-S365AvailableSkus { [CmdletBinding()] Param () try { Get-MgSubscribedSku | Select-Object SkuId, SkuPartNumber, ConsumedUnits, PrepaidUnits } catch { Write-Error "Erro ao buscar SKUs: $_" } }
Function Get-S365UserLicense { [CmdletBinding()] Param ( [Parameter(Mandatory=$true)][String] $UserPrincipalName ) try { Get-MgUserLicenseDetail -UserId $UserPrincipalName | Select-Object SkuPartNumber, ServicePlans } catch { Write-Error "Erro ao buscar licenças de '$UserPrincipalName': $_" } }
Function Set-S365UserLicense { [CmdletBinding()] Param ( [Parameter(Mandatory=$true)][String] $UserPrincipalName, [String[]] $AddSkuIds, [String[]] $RemoveSkuIds ) try { $user = Get-S365User -Identity $UserPrincipalName -Select UsageLocation; if (-not $user.UsageLocation) { Write-Error "Usuário '$UserPrincipalName' não tem UsageLocation."; return }; $addLicensesObject = @(); if ($AddSkuIds) { $addLicensesObject = $AddSkuIds | ForEach-Object { @{ SkuId = $_ } } }; $params = @{ UserId = $UserPrincipalName; AddLicenses = $addLicensesObject; RemoveLicenses = @($RemoveSkuIds) }; Set-MgUserLicense @params; Write-Host "Licenças atualizadas para '$UserPrincipalName'." -ForegroundColor Green } catch { Write-Error "Erro ao atualizar licenças para '$UserPrincipalName': $_" } }
Function Get-S365Group { [CmdletBinding()] Param ( [String] $Identity, [Switch] $All, [String] $Filter, [String[]] $Select = @("Id", "DisplayName", "Mail", "GroupTypes", "SecurityEnabled", "MailEnabled", "Visibility", "Description") ) try { $params = @{ }; if ($Select) { $params.Property = $Select }; if ($All) { Get-MgGroup @params -All } elseif ($Filter) { Get-MgGroup @params -Filter $Filter -ConsistencyLevel eventual -CountVariable countVar } elseif ($Identity) { Get-MgGroup @params -GroupId $Identity } else { Write-Warning "Especifique -Identity, -Filter ou -All." } } catch { Write-Error "Erro ao buscar grupo(s): $_" } }
Function Get-S365GroupMember { [CmdletBinding()] Param ( [Parameter(Mandatory=$true)][String] $GroupId ) try { Get-MgGroupMember -GroupId $GroupId -All | Select-Object Id, DisplayName, UserPrincipalName, Mail } catch { Write-Error "Erro ao buscar membros do grupo '$GroupId': $_" } }
Function Add-S365GroupMember { [CmdletBinding()] Param ( [Parameter(Mandatory=$true)][String] $GroupId, [Parameter(Mandatory=$true)][String] $UserPrincipalName ) try { $user = Get-S365User -Identity $UserPrincipalName; New-MgGroupMember -GroupId $GroupId -DirectoryObjectId $user.Id; Write-Host "Usuário '$UserPrincipalName' adicionado ao grupo '$GroupId'." -ForegroundColor Green } catch { Write-Error "Erro ao adicionar membro '$UserPrincipalName' ao grupo '$GroupId': $_" } }
Function Remove-S365GroupMember { [CmdletBinding(SupportsShouldProcess=$true)] Param ( [Parameter(Mandatory=$true)][String] $GroupId, [Parameter(Mandatory=$true)][String] $UserPrincipalName ) try { $user = Get-S365User -Identity $UserPrincipalName; if ($PSCmdlet.ShouldProcess("$UserPrincipalName from $GroupId", "Remover Membro")) { Remove-MgGroupMemberByRef -GroupId $GroupId -DirectoryObjectId $user.Id; Write-Host "Usuário '$UserPrincipalName' removido do grupo '$GroupId'." -ForegroundColor Green } } catch { Write-Error "Erro ao remover membro '$UserPrincipalName' do grupo '$GroupId': $_" } }
Function New-S365Group { [CmdletBinding()] Param( [Parameter(Mandatory=$true)][String] $DisplayName, [Parameter(Mandatory=$true)][String] $MailNickname, [Parameter(Mandatory=$true)][ValidateSet("M365", "Security")][String] $GroupType, [String] $Description, [String[]] $OwnersUPN, [String[]] $MembersUPN, [Switch] $MailEnabled = ($GroupType -eq "M365"), [Switch] $SecurityEnabled = ($GroupType -eq "Security" -or $GroupType -eq "M365"), [ValidateSet("Public", "Private")][String] $Visibility = "Private" ) $params = @{ DisplayName = $DisplayName; MailNickname = $MailNickname; MailEnabled = $MailEnabled; SecurityEnabled = $SecurityEnabled; Description = $Description; GroupTypes = @(if ($GroupType -eq "M365") { "Unified" } else { }) }; if ($GroupType -eq "M365") { $params.Visibility = $Visibility }; try { $newGroup = New-MgGroup -BodyParameter $params; Write-Host "Grupo '$DisplayName' (ID: $($newGroup.Id)) criado." -ForegroundColor Green; if ($OwnersUPN) { $OwnersUPN | ForEach-Object { try { Add-S365GroupOwner -GroupId $newGroup.Id -UserPrincipalName $_ } catch { Write-Warning "Falha ao add '$_' como dono." } } }; if ($MembersUPN) { $MembersUPN | ForEach-Object { try { Add-S365GroupMember -GroupId $newGroup.Id -UserPrincipalName $_ } catch { Write-Warning "Falha ao add '$_' como membro." } } }; return $newGroup } catch { Write-Error "Erro ao criar grupo: $_" } }
Function Add-S365GroupOwner { [CmdletBinding()] Param ( [Parameter(Mandatory=$true)][String] $GroupId, [Parameter(Mandatory=$true)][String] $UserPrincipalName ) try { $user = Get-S365User -Identity $UserPrincipalName; New-MgGroupOwnerByRef -GroupId $GroupId -OdataId "https://graph.microsoft.com/v1.0/users/$($user.Id)"; Write-Host "Usuário '$UserPrincipalName' adicionado como dono do grupo '$GroupId'." -ForegroundColor Green } catch { Write-Error "Erro ao adicionar dono '$UserPrincipalName': $_" } }
Function Get-S365GroupOwner { [CmdletBinding()] Param ( [Parameter(Mandatory=$true)][String] $GroupId ) try { Get-MgGroupOwner -GroupId $GroupId -All | Select-Object Id, DisplayName, UserPrincipalName } catch { Write-Error "Erro ao buscar donos do grupo '$GroupId': $_" } }
Function Get-S365Team { [CmdletBinding()] Param ( [String] $DisplayName, [Switch] $All ) try { if ($DisplayName) { Get-MgTeam -Filter "DisplayName eq '$DisplayName'" } elseif ($All) { Get-MgTeam -All } else { Get-MgTeam } } catch { Write-Error "Erro ao buscar Teams: $_" } }
Function New-S365Team { [CmdletBinding()] Param ( [Parameter(Mandatory=$true)][String] $DisplayName, [String] $Description, [String] $OwnerUPN ) try { $template = "https://graph.microsoft.com/v1.0/teamsTemplates('standard')"; $params = @{ "template@odata.bind" = $template; DisplayName = $DisplayName; Description = $Description }; $newTeam = New-MgTeam -BodyParameter $params; Write-Host "Team '$DisplayName' (ID: $($newTeam.Id)) criado. Aguarde..." -ForegroundColor Yellow; Start-Sleep -Seconds 15; if ($OwnerUPN) { try { Add-S365TeamMember -TeamId $newTeam.Id -UserPrincipalName $OwnerUPN -IsOwner } catch { Write-Warning "Falha ao add '$OwnerUPN' como dono." } } else { Write-Warning "Team criado sem dono." }; Write-Host "Team '$DisplayName' provisionado." -ForegroundColor Green; return $newTeam } catch { Write-Error "Erro ao criar Team: $_" } }
Function Get-S365TeamChannel { [CmdletBinding()] Param ([Parameter(Mandatory=$true)][String] $TeamId) try { Get-MgTeamChannel -TeamId $TeamId -All | Select-Object Id, DisplayName, WebUrl, MembershipType } catch { Write-Error "Erro ao buscar canais para Team '$TeamId': $_" } }
Function Add-S365TeamMember { [CmdletBinding()] Param ( [Parameter(Mandatory=$true)][String] $TeamId, [Parameter(Mandatory=$true)][String] $UserPrincipalName, [Switch] $IsOwner ) try { $user = Get-S365User -Identity $UserPrincipalName; $params = @{ "@odata.type" = "#microsoft.graph.aadUserConversationMember"; "user@odata.bind" = "https://graph.microsoft.com/v1.0/users('$($user.Id)')"; Roles = @(if ($IsOwner) { "owner" } else { "member" }) }; New-MgTeamMember -TeamId $TeamId -BodyParameter $params; $role = if ($IsOwner) { "Proprietário" } else { "Membro" }; Write-Host "Usuário '$UserPrincipalName' adicionado ao Team '$TeamId' como $role." -ForegroundColor Green } catch { Write-Error "Erro ao adicionar membro '$UserPrincipalName' ao Team '$TeamId': $_" } }
Function Remove-S365TeamMember { [CmdletBinding(SupportsShouldProcess=$true)] Param ( [Parameter(Mandatory=$true)][String] $TeamId, [Parameter(Mandatory=$true)][String] $UserPrincipalName ) try { $user = Get-S365User -Identity $UserPrincipalName; $member = Get-MgTeamMember -TeamId $TeamId -Filter "UserId eq '$($user.Id)'"; if ($member) { if ($PSCmdlet.ShouldProcess("$UserPrincipalName from $TeamId", "Remover Membro")) { Remove-MgTeamMember -TeamId $TeamId -ConversationMemberId $member.Id; Write-Host "Usuário '$UserPrincipalName' removido do Team '$TeamId'." -ForegroundColor Green } } else { Write-Warning "Usuário '$UserPrincipalName' não encontrado no Team '$TeamId'." } } catch { Write-Error "Erro ao remover membro '$UserPrincipalName' do Team '$TeamId': $_" } }
Function Get-S365Mailbox { [CmdletBinding()] Param ( [String] $Identity, [Switch] $All, [String] $Filter, [Switch] $Shared, [Switch] $Room, [Switch] $Equipment, [Switch] $Archive ) try { $params = @{}; if ($Shared) { $params.RecipientTypeDetails = "SharedMailbox" } elseif ($Room) { $params.RecipientTypeDetails = "RoomMailbox" } elseif ($Equipment) { $params.RecipientTypeDetails = "EquipmentMailbox" } elseif ($Archive) { $params.Archive = $true }; if ($All) { Get-Mailbox @params -ResultSize Unlimited } elseif ($Filter) { Get-Mailbox @params -Filter $Filter -ResultSize Unlimited } elseif ($Identity) { Get-Mailbox @params -Identity $Identity } else { Get-Mailbox @params -ResultSize 500 } } catch { Write-Error "Erro ao buscar caixa(s) de correio: $_" } }
Function Set-S365MailboxQuota { [CmdletBinding()] Param ( [Parameter(Mandatory=$true)][String] $Identity, [Parameter(Mandatory=$true)][String] $Quota ) try { $quotaValue = [Microsoft.Exchange.Data.ByteQuantifiedSize]::Parse($Quota); $warningQuota = ([Microsoft.Exchange.Data.ByteQuantifiedSize]($quotaValue.ToBytes() * 0.9)).ToString(); Set-Mailbox -Identity $Identity -ProhibitSendReceiveQuota $Quota -IssueWarningQuota $warningQuota; Write-Host "Quota '$Identity' definida para $Quota." -ForegroundColor Green } catch { Write-Error "Erro ao definir quota para '$Identity': $_" } }
Function Set-S365MailboxForwarding { [CmdletBinding()] Param ( [Parameter(Mandatory=$true)][String] $Identity, [String] $ForwardingSmtpAddress, [Switch] $DeliverToMailboxAndForward = $true, [Switch] $RemoveForwarding ) try { if ($RemoveForwarding) { Set-Mailbox -Identity $Identity -ForwardingSmtpAddress $null -DeliverToMailboxAndForward $false; Write-Host "Encaminhamento removido para '$Identity'." -ForegroundColor Green } else { if (-not $ForwardingSmtpAddress) { Write-Error "-ForwardingSmtpAddress é necessário." ; return }; Set-Mailbox -Identity $Identity -ForwardingSmtpAddress $ForwardingSmtpAddress -DeliverToMailboxAndForward $DeliverToMailboxAndForward; Write-Host "Encaminhamento configurado para '$Identity' -> '$ForwardingSmtpAddress'." -ForegroundColor Green } } catch { Write-Error "Erro ao configurar encaminhamento para '$Identity': $_" } }
Function Get-S365MailboxStatistics { [CmdletBinding()] Param ( [Parameter(Mandatory=$true)][String] $Identity ) try { Get-MailboxStatistics -Identity $Identity | Select-Object DisplayName, ItemCount, TotalItemSize, LastLogonTime, LastUserActionTime } catch { Write-Error "Erro ao buscar estatísticas para '$Identity': $_" } }
Function New-S365SharedMailbox { [CmdletBinding()] Param ( [Parameter(Mandatory=$true)][String] $Name, [String] $Alias = ($Name -replace "\s", ""), [String] $PrimarySmtpAddress ) try { if(-not $PrimarySmtpAddress) { $domain = Get-AcceptedDomain | Where-Object { $_.IsDefault -eq $true } | Select-Object -ExpandProperty DomainName; $PrimarySmtpAddress = "$Alias@$domain" }; New-Mailbox -Shared -Name $Name -DisplayName $Name -Alias $Alias -PrimarySmtpAddress $PrimarySmtpAddress; Write-Host "Caixa compartilhada '$Name' ($PrimarySmtpAddress) criada." -ForegroundColor Green } catch { Write-Error "Erro ao criar caixa compartilhada '$Name': $_" } }
Function Set-S365LitigationHold { [CmdletBinding()] Param ( [Parameter(Mandatory=$true)][String] $Identity, [Switch] $Enable, [Switch] $Disable, [Int32] $DurationDays ) try { if ($Enable) { $params = @{ Identity = $Identity; LitigationHoldEnabled = $true }; if ($DurationDays) { $params.LitigationHoldDuration = $DurationDays }; Set-Mailbox @params; Write-Host "Litigation Hold HABILITADO para '$Identity'." -ForegroundColor Green } elseif ($Disable) { Set-Mailbox -Identity $Identity -LitigationHoldEnabled = $false; Write-Host "Litigation Hold DESABILITADO para '$Identity'." -ForegroundColor Green } else { Write-Warning "Use -Enable ou -Disable." } } catch { Write-Error "Erro ao configurar Litigation Hold para '$Identity': $_" } }
Function Add-S365MailboxPermission { [CmdletBinding()] Param ( [Parameter(Mandatory=$true)][String] $Identity, [Parameter(Mandatory=$true)][String] $User, [Parameter(Mandatory=$true)][ValidateSet("FullAccess", "ExternalAccount", "DeleteItem", "ReadPermission", "ChangePermission", "ChangeOwner")]$AccessRights = "FullAccess", [Switch] $AutoMapping = $true ) try { Add-MailboxPermission -Identity $Identity -User $User -AccessRights $AccessRights -InheritanceType All -AutoMapping:$AutoMapping; Write-Host "Permissão '$AccessRights' concedida a '$User' em '$Identity'." -ForegroundColor Green } catch { Write-Error "Erro ao adicionar permissão '$AccessRights' a '$User' em '$Identity': $_" } }
Function Remove-S365MailboxPermission { [CmdletBinding(SupportsShouldProcess=$true)] Param ( [Parameter(Mandatory=$true)][String] $Identity, [Parameter(Mandatory=$true)][String] $User, [Parameter(Mandatory=$true)][ValidateSet("FullAccess", "ExternalAccount", "DeleteItem", "ReadPermission", "ChangePermission", "ChangeOwner")]$AccessRights = "FullAccess" ) try { if ($PSCmdlet.ShouldProcess("$User on $Identity", "Remover Permissão ($AccessRights)")) { Remove-MailboxPermission -Identity $Identity -User $User -AccessRights $AccessRights -InheritanceType All -Confirm:$false; Write-Host "Permissão '$AccessRights' removida de '$User' em '$Identity'." -ForegroundColor Green } } catch { Write-Error "Erro ao remover permissão '$AccessRights' de '$User' em '$Identity': $_" } }
Function Add-S365RecipientPermission { [CmdletBinding()] Param ( [Parameter(Mandatory=$true)][String] $Identity, [Parameter(Mandatory=$true)][String] $Trustee, [Parameter(Mandatory=$true)][ValidateSet("SendAs", "SendOnBehalf")]$AccessRights = "SendAs" ) try { if ($AccessRights -eq "SendAs") { Add-RecipientPermission -Identity $Identity -Trustee $Trustee -AccessRights SendAs; Write-Host "'SendAs' concedida a '$Trustee' em '$Identity'." -ForegroundColor Green } elseif ($AccessRights -eq "SendOnBehalf") { Set-Mailbox -Identity $Identity -GrantSendOnBehalfTo @{Add="$Trustee"}; Write-Host "'SendOnBehalf' concedida a '$Trustee' em '$Identity'." -ForegroundColor Green } } catch { Write-Error "Erro ao adicionar permissão '$AccessRights': $_" } }
Function Remove-S365RecipientPermission { [CmdletBinding(SupportsShouldProcess=$true)] Param ( [Parameter(Mandatory=$true)][String] $Identity, [Parameter(Mandatory=$true)][String] $Trustee, [Parameter(Mandatory=$true)][ValidateSet("SendAs", "SendOnBehalf")]$AccessRights = "SendAs" ) try { if ($PSCmdlet.ShouldProcess("$Trustee on $Identity", "Remover Permissão ($AccessRights)")) { if ($AccessRights -eq "SendAs") { Remove-RecipientPermission -Identity $Identity -Trustee $Trustee -AccessRights SendAs -Confirm:$false; Write-Host "'SendAs' removida de '$Trustee' em '$Identity'." -ForegroundColor Green } elseif ($AccessRights -eq "SendOnBehalf") { Set-Mailbox -Identity $Identity -GrantSendOnBehalfTo @{Remove="$Trustee"}; Write-Host "'SendOnBehalf' removida de '$Trustee' em '$Identity'." -ForegroundColor Green } } } catch { Write-Error "Erro ao remover permissão '$AccessRights': $_" } }
Function Get-S365MailboxFolderPermission { [CmdletBinding()] Param ( [Parameter(Mandatory=$true)][String] $Identity, [String] $FolderPath = "\Calendário" ) try { $calendarPath = $null; try { $calendarPath = (Get-MailboxFolderStatistics -Identity $Identity -FolderScope Calendar | Select-Object -First 1).FolderPath } catch { Write-Warning "Não foi possível detectar a pasta Calendário automaticamente. Usando '$FolderPath'." }; $finalPath = if($calendarPath) { $calendarPath } else { $FolderPath }; $IdentityPath = $Identity + ":" + $finalPath; Get-MailboxFolderPermission -Identity $IdentityPath } catch { Write-Error "Erro ao buscar permissões da pasta '$finalPath' para '$Identity': $_" } }
Function Add-S365MailboxFolderPermission { [CmdletBinding()] Param ( [Parameter(Mandatory=$true)][String] $Identity, [Parameter(Mandatory=$true)][String] $User, [Parameter(Mandatory=$true)][ValidateSet("Owner", "PublishingEditor", "Editor", "PublishingAuthor", "Author", "NoneditingAuthor", "Reviewer", "Contributor", "AvailabilityOnly", "LimitedDetails")]$AccessRights, [String] $FolderPath = "\Calendário" ) try { $IdentityPath = $Identity + ":" + $FolderPath; Add-MailboxFolderPermission -Identity $IdentityPath -User $User -AccessRights $AccessRights; Write-Host "Permissão '$AccessRights' concedida a '$User' na pasta '$FolderPath' de '$Identity'." -ForegroundColor Green } catch { $IdentityPath = $Identity + ":" + $FolderPath; Write-Error "Erro ao adicionar permissão '$AccessRights' a '$User' em '$IdentityPath': $_" } }
Function Remove-S365MailboxFolderPermission { [CmdletBinding(SupportsShouldProcess=$true)] Param ( [Parameter(Mandatory=$true)][String] $Identity, [Parameter(Mandatory=$true)][String] $User, [String] $FolderPath = "\Calendário" ) try { $IdentityPath = $Identity + ":" + $FolderPath; if ($PSCmdlet.ShouldProcess("$User on $IdentityPath", "Remover Permissão de Pasta")) { Remove-MailboxFolderPermission -Identity $IdentityPath -User $User -Confirm:$false; Write-Host "Permissões removidas de '$User' na pasta '$FolderPath' de '$Identity'." -ForegroundColor Green } } catch { $IdentityPath = $Identity + ":" + $FolderPath; Write-Error "Erro ao remover permissão de '$User' em '$IdentityPath': $_" } }
Function Get-S365DistributionGroup { [CmdletBinding()] Param ( [String] $Identity, [Switch] $All ) try { if ($All) { Get-DistributionGroup -ResultSize Unlimited } elseif ($Identity) { Get-DistributionGroup -Identity $Identity } else { Get-DistributionGroup } } catch { Write-Error "Erro ao buscar grupo(s) de distribuição: $_" } }
Function Get-S365DistributionGroupMember { [CmdletBinding()] Param ( [Parameter(Mandatory=$true)][String] $Identity ) try { Get-DistributionGroupMember -Identity $Identity -ResultSize Unlimited } catch { Write-Error "Erro ao buscar membros do grupo '$Identity': $_" } }
Function Add-S365DistributionGroupMember { [CmdletBinding()] Param ( [Parameter(Mandatory=$true)][String] $Identity, [Parameter(Mandatory=$true)][String] $Member ) try { Add-DistributionGroupMember -Identity $Identity -Member $Member; Write-Host "Membro '$Member' adicionado ao grupo '$Identity'." -ForegroundColor Green } catch { Write-Error "Erro ao adicionar membro '$Member' ao grupo '$Identity': $_" } }
Function Remove-S365DistributionGroupMember { [CmdletBinding(SupportsShouldProcess=$true)] Param ( [Parameter(Mandatory=$true)][String] $Identity, [Parameter(Mandatory=$true)][String] $Member ) try { if ($PSCmdlet.ShouldProcess($Member, "Remover Membro $Identity")) { Remove-DistributionGroupMember -Identity $Identity -Member $Member -Confirm:$false; Write-Host "Membro '$Member' removido do grupo '$Identity'." -ForegroundColor Green } } catch { Write-Error "Erro ao remover membro '$Member': $_" } }
Function Get-S365MailContact { [CmdletBinding()] Param ( [String] $Identity, [Switch] $All ) try { if ($All) { Get-MailContact -ResultSize Unlimited } elseif ($Identity) { Get-MailContact -Identity $Identity } else { Get-MailContact } } catch { Write-Error "Erro ao buscar contato(s): $_" } }
Function Get-S365TransportRule { [CmdletBinding()] Param ([String] $Identity) try { if ($Identity) { Get-TransportRule -Identity $Identity } else { Get-TransportRule } } catch { Write-Error "Erro ao buscar regras de transporte: $_" } }
Function Set-S365TransportRuleState { [CmdletBinding(SupportsShouldProcess=$true)] Param ( [Parameter(Mandatory=$true)][String] $Identity, [Parameter(Mandatory=$true)][ValidateSet("Enabled", "Disabled")]$State ) try { if ($PSCmdlet.ShouldProcess($Identity, "Alterar Estado para $State")) { if ($State -eq "Enabled") { Enable-TransportRule -Identity $Identity } else { Disable-TransportRule -Identity $Identity }; Write-Host "Regra '$Identity' definida para '$State'." -ForegroundColor Green } } catch { Write-Error "Erro ao alterar estado da regra '$Identity': $_" } }
Function Start-S365MessageTrace { [CmdletBinding()] Param( [Parameter(Mandatory=$true)][DateTime] $StartDate, [Parameter(Mandatory=$true)][DateTime] $EndDate, [String] $SenderAddress, [String] $RecipientAddress, [String] $Subject, [ValidateSet("Pending", "Failed", "Delivered", "Expanded", "Quarantined", "FilteredAsSpam", "GettingStatus")] $Status, [String] $MessageId ) $params = @{ StartDate = $StartDate; EndDate = $EndDate }; if ($SenderAddress) { $params.SenderAddress = $SenderAddress }; if ($RecipientAddress) { $params.RecipientAddress = $RecipientAddress }; if ($Subject) { $params.Subject = $Subject }; if ($PSBoundParameters.ContainsKey('Status')) { $params.Status = $Status }; if ($MessageId) { $params.MessageId = $MessageId }; try { Get-MessageTrace @params | Select-Object Received, SenderAddress, RecipientAddress, Subject, Status, MessageId, Size, FromIP } catch { Write-Error "Erro ao buscar rastreamento: $_" } }
Function Start-S365HistoricalMessageTrace { [CmdletBinding()] Param( [Parameter(Mandatory=$true)][DateTime] $StartDate, [Parameter(Mandatory=$true)][DateTime] $EndDate, [Parameter(Mandatory=$true)][String] $ReportTitle, [Parameter(Mandatory=$true)][String] $NotifyAddress, [String] $SenderAddress, [String] $RecipientAddress, [String] $Subject, [String] $MessageId, [ValidateSet("Summary", "Message")]$ReportType = "Summary" ) $params = @{ StartDate = $StartDate; EndDate = $EndDate; ReportTitle = $ReportTitle; NotifyAddress = $NotifyAddress; ReportType = $ReportType }; if ($SenderAddress) { $params.SenderAddress = $SenderAddress }; if ($RecipientAddress) { $params.RecipientAddress = $RecipientAddress }; if ($Subject) { $params.Subject = $Subject }; if ($MessageId) { $params.MessageId = $MessageId }; try { $search = Start-HistoricalSearch @params; Write-Host "Rastreamento histórico iniciado. Título: '$ReportTitle'. ID: $($search.JobId)." -ForegroundColor Green; return $search } catch { Write-Error "Erro ao iniciar rastreamento histórico: $_" } }
Function Get-S365HistoricalMessageTrace { [CmdletBinding()] Param( [String] $JobId ) try { if ($JobId) { Get-HistoricalSearch -JobId $JobId } else { Get-HistoricalSearch } } catch { Write-Error "Erro ao buscar rastreamento histórico: $_" } }
Function Get-S365SignInActivity { [CmdletBinding()] Param ( [Parameter(Mandatory=$true)][String] $UserPrincipalName ) try { Get-MgUser -UserId $UserPrincipalName -Property 'SignInActivity' | Select-Object -ExpandProperty 'SignInActivity' } catch { Write-Error "Erro ao buscar atividade de login para '$UserPrincipalName': $_" } }
Function Get-S365LastLogonTime { [CmdletBinding()] Param ( [Switch] $All, [Int32] $DaysInactive = 90 ) try { Write-Host "Buscando estatísticas (pode demorar)..." -ForegroundColor Yellow; $mailboxes = Get-Mailbox -ResultSize Unlimited -Filter {RecipientTypeDetails -ne "DiscoveryMailbox"}; $i = 0; $total = $mailboxes.Count; $stats = $mailboxes | ForEach-Object { $i++; Write-Progress -Activity "Buscando Estatísticas" -Status "Processando: $($_.UserPrincipalName) ($i de $total)" -PercentComplete ($i / $total * 100); Get-MailboxStatistics -Identity $_.UserPrincipalName -ErrorAction SilentlyContinue }; if ($All) { $stats | Select-Object DisplayName, UserPrincipalName, LastLogonTime, TotalItemSize } else { $threshold = (Get-Date).AddDays(-$DaysInactive); Write-Host "Filtrando por $DaysInactive dias inativos..."; $stats | Where-Object { $_.LastLogonTime -lt $threshold -or $_.LastLogonTime -eq $null } | Select-Object DisplayName, UserPrincipalName, LastLogonTime, TotalItemSize } } catch { Write-Error "Erro ao buscar último logon: $_" } }
Function Search-S365AuditLog { [CmdletBinding()] Param ( [Parameter(Mandatory=$true)][DateTime] $StartDate, [Parameter(Mandatory=$true)][DateTime] $EndDate, [String] $UserIds, [String] $Operations, [Int32] $ResultSize = 1000 ) Write-Host "Pesquisando Log de Auditoria (pode demorar)..." -ForegroundColor Yellow; try { Search-UnifiedAuditLog -StartDate $StartDate -EndDate $EndDate -UserIds $UserIds -Operations $Operations -ResultSize $ResultSize | Select-Object CreationDate, UserIds, Operations, AuditData | Sort-Object CreationDate -Descending } catch { Write-Error "Erro ao pesquisar o log de auditoria: $_" } }


# =========================================================================
#                    TUI HELPER FUNCTIONS
# =========================================================================

Function Get-InteractiveParams {
    Param (
        [System.Management.Automation.FunctionInfo]$FunctionInfo
    )

    $params = @{}
    $functionName = $FunctionInfo.Name
    $parameters = $FunctionInfo.Parameters.Values | Where-Object {
        $_.Name -ne "Confirm" -and $_.Name -ne "WhatIf" -and $_.Name -ne "Verbose" -and $_.Name -ne "Debug" -and $_.Name -ne "ErrorAction" -and $_.Name -ne "WarningAction" -and $_.Name -ne "InformationAction" -and $_.Name -ne "ErrorVariable" -and $_.Name -ne "WarningVariable" -and $_.Name -ne "InformationVariable" -and $_.Name -ne "OutVariable" -and $_.Name -ne "OutBuffer" -and $_.Name -ne "PipelineVariable"
    }

    Write-Host "`n--- Forneça os parâmetros para '$functionName' ---" -ForegroundColor Yellow
    Write-Host "Pressione Enter para pular opcionais. Digite '$$' para cancelar."

    foreach ($param in $parameters) {
        $isMandatory = $param.Attributes | Where-Object { $_.TypeId.Name -eq 'ParameterAttribute' -and $_.Mandatory }
        $paramType = $param.ParameterType

        $prompt = "  [$($param.Name)]"
        if ($isMandatory) { $prompt += " (Obrigatório)" } else { $prompt += " (Opcional)" }
        $prompt += " <$($paramType.Name)>: "

        while ($true) {
            $value = $null
            if ($paramType.Name -eq 'SwitchParameter') {
                $response = Read-Host "$prompt (S/N)"
                if ($response -eq '$$') { return $null }
                if ($response -match '^[SsYy1]') { $params.Add($param.Name, $true) }
                break # Sai do while
            } elseif ($paramType.Name -eq 'SecureString') {
                $value = Read-Host "$prompt" -AsSecureString
                if ($value -eq '$$') { return $null }
            } else {
                $value = Read-Host "$prompt"
                if ($value -eq '$$') { return $null }
            }

            if ($value -ne $null -and $value.Length -gt 0) {
                try {
                    if ($paramType.Name -eq 'String[]') {
                        $params.Add($param.Name, ($value -split ',' | ForEach-Object { $_.Trim() }))
                    } elseif ($paramType.Name -eq 'DateTime') {
                        $params.Add($param.Name, ([DateTime]$value))
                    } elseif ($paramType.Name -eq 'Int32') {
                        $params.Add($param.Name, ([Int32]$value))
                    } elseif ($paramType.Name -eq 'SecureString') {
                        $params.Add($param.Name, $value)
                    } else {
                        $params.Add($param.Name, $value)
                    }
                    break # Sai do while, valor válido
                } catch {
                    Write-Warning "Entrada inválida. $_.Exception.Message"
                    # Continua no loop para pedir novamente
                }
            } elseif ($isMandatory) {
                Write-Warning "Parâmetro obrigatório '$($param.Name)' não pode ser vazio."
                # Continua no loop para pedir novamente
            } else {
                break # Sai do while, opcional vazio
            }
        }
    }
    return $params
}

Function Invoke-FunctionWithParams {
    Param(
        [string]$FunctionName
    )
    try {
        $funcInfo = Get-Command $FunctionName -ErrorAction Stop
        $parameters = Get-InteractiveParams -FunctionInfo $funcInfo

        if ($parameters -ne $null) {
            Write-Host "`nExecutando $FunctionName..." -ForegroundColor Cyan
            & $FunctionName @parameters | Out-Host # Usar Out-Host para garantir a exibição no console
            Write-Host "`nExecução de $FunctionName concluída." -ForegroundColor Green
        } else {
             Write-Host "`nOperação '$FunctionName' cancelada." -ForegroundColor Yellow
        }
    } catch {
        Write-Error "Erro ao preparar ou executar $FunctionName: $_"
    }
    Read-Host "`nPressione Enter para continuar..."
}


# =========================================================================
#                    TUI DISPLAY FUNCTIONS
# =========================================================================
Function Show-Header {
    $exchStatus = if ($Global:IsExchangeConnected) { "(Conectado)" } else { "(Desconectado)" }
    $graphStatus = if ($Global:IsGraphConnected) { "(Conectado como $($Global:GraphUser))" } else { "(Desconectado)" }
    Write-Host "====================== SuperAdmin365 TUI ======================" -ForegroundColor Cyan
    Write-Host " Exchange: " -NoNewline; Write-Host $exchStatus -ForegroundColor (if ($Global:IsExchangeConnected) { "Green" } else { "Red" }) -NoNewline
    Write-Host " | Graph: " -NoNewline; Write-Host $graphStatus -ForegroundColor (if ($Global:IsGraphConnected) { "Green" } else { "Red" })
    Write-Host "---------------------------------------------------------------"
}

Function Show-MainMenu {
    Show-Header
    Write-Host "   [1] Conexão" -ForegroundColor White
    Write-Host "   [2] Usuários (Graph)" -ForegroundColor White
    Write-Host "   [3] Grupos & Teams (Graph)" -ForegroundColor White
    Write-Host "   [4] Exchange Online" -ForegroundColor White
    Write-Host "   [5] Licenças & Auditoria" -ForegroundColor White
    Write-Host "---------------------------------------------------------------"
    Write-Host "   [Q] Sair" -ForegroundColor Yellow
    Write-Host "==============================================================="
}

Function Show-ConnectionMenu {
    Show-Header
    Write-Host "=========== Menu de Conexão ===========" -ForegroundColor Green
    Write-Host "  [1] Conectar Serviços"
    Write-Host "  [2] Desconectar Serviços"
    Write-Host "---------------------------------------"
    Write-Host "  [B] Voltar" -ForegroundColor Yellow
    Write-Host "======================================="
}

Function Show-UserMenu {
    Show-Header
    Write-Host "========= Menu de Usuários (Graph) =========" -ForegroundColor Green
    Write-Host "  [1] Buscar Usuário"
    Write-Host "  [2] Criar Usuário"
    Write-Host "  [3] Modificar Usuário"
    Write-Host "  [4] Remover Usuário"
    Write-Host "  [5] Restaurar Usuário Excluído"
    Write-Host "  [6] Listar Usuários Excluídos"
    Write-Host "  [7] Resetar Senha"
    Write-Host "  [8] Ver Gerente"
    Write-Host "  [9] Ver Subordinados"
    Write-Host "  [10] Ver Status MFA"
    Write-Host "  [11] Ver Atividade de Login"
    Write-Host "------------------------------------------"
    Write-Host "  [B] Voltar" -ForegroundColor Yellow
    Write-Host "=========================================="
}

Function Show-GroupMenu {
    Show-Header
    Write-Host "======= Menu de Grupos & Teams (Graph) =======" -ForegroundColor Green
    Write-Host " --- Grupos ---"
    Write-Host "  [1] Buscar Grupo"
    Write-Host "  [2] Criar Grupo (M365/Segurança)"
    Write-Host "  [3] Listar Membros de Grupo"
    Write-Host "  [4] Adicionar Membro a Grupo"
    Write-Host "  [5] Remover Membro de Grupo"
    Write-Host "  [6] Listar Donos de Grupo"
    Write-Host "  [7] Adicionar Dono a Grupo"
    Write-Host " --- Teams ---"
    Write-Host "  [8] Buscar Team"
    Write-Host "  [9] Criar Team"
    Write-Host "  [10] Listar Canais do Team"
    Write-Host "  [11] Adicionar Membro ao Team"
    Write-Host "  [12] Remover Membro do Team"
    Write-Host "---------------------------------------------"
    Write-Host "  [B] Voltar" -ForegroundColor Yellow
    Write-Host "============================================="
}

Function Show-ExchangeMenu {
    Show-Header
    Write-Host "========= Menu Exchange Online =========" -ForegroundColor Green
    Write-Host " --- Mailboxes ---"
    Write-Host "  [1] Buscar Mailbox (User/Shared/Room/Eqp)"
    Write-Host "  [2] Ver Estatísticas de Mailbox"
    Write-Host "  [3] Definir Quota de Mailbox"
    Write-Host "  [4] Configurar Encaminhamento"
    Write-Host "  [5] Criar Mailbox Compartilhada"
    Write-Host "  [6] Configurar Litigation Hold"
    Write-Host " --- Permissões ---"
    Write-Host "  [7] Adicionar Permissão de Mailbox (FullAccess)"
    Write-Host "  [8] Remover Permissão de Mailbox"
    Write-Host "  [9] Adicionar Permissão de Destinatário (SendAs/Behalf)"
    Write-Host "  [10] Remover Permissão de Destinatário"
    Write-Host "  [11] Ver Permissões de Pasta (Ex: Calendário)"
    Write-Host "  [12] Adicionar Permissão de Pasta"
    Write-Host "  [13] Remover Permissão de Pasta"
    Write-Host " --- Grupos & Contatos (ExO) ---"
    Write-Host "  [14] Buscar Grupo de Distribuição"
    Write-Host "  [15] Listar Membros de Grupo de Distribuição"
    Write-Host "  [16] Adicionar Membro a Grupo de Distribuição"
    Write-Host "  [17] Remover Membro de Grupo de Distribuição"
    Write-Host "  [18] Buscar Contato de Email"
    Write-Host " --- Fluxo de Email ---"
    Write-Host "  [19] Buscar Regra de Transporte"
    Write-Host "  [20] Ativar/Desativar Regra de Transporte"
    Write-Host "  [21] Rastrear Mensagem (Recente)"
    Write-Host "  [22] Iniciar Rastreamento Histórico"
    Write-Host "  [23] Ver Rastreamento Histórico"
    Write-Host "-----------------------------------------"
    Write-Host "  [B] Voltar" -ForegroundColor Yellow
    Write-Host "========================================="
}

Function Show-LicenseAuditMenu {
    Show-Header
    Write-Host "====== Menu Licenças & Auditoria ======" -ForegroundColor Green
    Write-Host "  [1] Listar SKUs Disponíveis"
    Write-Host "  [2] Ver Licenças de Usuário"
    Write-Host "  [3] Atribuir/Remover Licença"
    Write-Host "  [4] Ver Último Logon (Mailbox)"
    Write-Host "  [5] Pesquisar Log de Auditoria"
    Write-Host "---------------------------------------"
    Write-Host "  [B] Voltar" -ForegroundColor Yellow
    Write-Host "======================================="
}

# =========================================================================
#                    TUI MAIN LOOP
# =========================================================================

Function Start-SuperAdminTUI {
    $menuStack = New-Object System.Collections.Stack
    $currentMenu = "Main"

    while ($true) {
        Clear-Host
        switch ($currentMenu) {
            "Main" { Show-MainMenu }
            "Connection" { Show-ConnectionMenu }
            "User" { Show-UserMenu }
            "Group" { Show-GroupMenu }
            "Exchange" { Show-ExchangeMenu }
            "LicenseAudit" { Show-LicenseAuditMenu }
        }

        $selection = Read-Host "Sua escolha"

        # --- NAVEGAÇÃO GERAL (Voltar / Sair) ---
        if ($selection -eq 'b' -or $selection -eq 'B') {
            if ($menuStack.Count -gt 0) {
                $currentMenu = $menuStack.Pop()
            } else {
                $currentMenu = "Main" # Garante que volte ao principal se a pilha estiver vazia
            }
            continue
        }
        if ($selection -eq 'q' -or $selection -eq 'Q') {
            if ($currentMenu -eq "Main") {
                Write-Host "Saindo..." -ForegroundColor Yellow
                Disconnect-Super365Services -ErrorAction SilentlyContinue
                break # Sai do loop while
            } else {
                 Write-Host "Use 'B' para voltar ou 'Q' no menu principal para sair." -ForegroundColor Yellow
                 Start-Sleep 2
            }
             continue
        }

        # --- LÓGICA DO MENU ATUAL ---
        try {
            switch ($currentMenu) {
                "Main" {
                    switch ($selection) {
                        "1" { $menuStack.Push($currentMenu); $currentMenu = "Connection" }
                        "2" { $menuStack.Push($currentMenu); $currentMenu = "User" }
                        "3" { $menuStack.Push($currentMenu); $currentMenu = "Group" }
                        "4" { $menuStack.Push($currentMenu); $currentMenu = "Exchange" }
                        "5" { $menuStack.Push($currentMenu); $currentMenu = "LicenseAudit" }
                        default { Write-Warning "Seleção inválida." }
                    }
                }
                "Connection" {
                    switch ($selection) {
                        "1" { Invoke-FunctionWithParams -FunctionName "Connect-Super365Services" }
                        "2" { Invoke-FunctionWithParams -FunctionName "Disconnect-Super365Services" }
                        default { Write-Warning "Seleção inválida." }
                    }
                }
                "User" {
                    switch ($selection) {
                        "1" { Invoke-FunctionWithParams -FunctionName "Get-S365User" }
                        "2" { Invoke-FunctionWithParams -FunctionName "New-S365User" }
                        "3" { Invoke-FunctionWithParams -FunctionName "Set-S365User" }
                        "4" { Invoke-FunctionWithParams -FunctionName "Remove-S365User" }
                        "5" { Invoke-FunctionWithParams -FunctionName "Restore-S365DeletedUser" }
                        "6" { Invoke-FunctionWithParams -FunctionName "Get-S365DeletedUser" }
                        "7" { Invoke-FunctionWithParams -FunctionName "Reset-S365UserPassword" }
                        "8" { Invoke-FunctionWithParams -FunctionName "Get-S365UserManager" }
                        "9" { Invoke-FunctionWithParams -FunctionName "Get-S365UserDirectReports" }
                        "10" { Invoke-FunctionWithParams -FunctionName "Get-S365MfaStatus" }
                        "11" { Invoke-FunctionWithParams -FunctionName "Get-S365SignInActivity" }
                        default { Write-Warning "Seleção inválida." }
                    }
                }
                "Group" {
                    switch ($selection) {
                        "1" { Invoke-FunctionWithParams -FunctionName "Get-S365Group" }
                        "2" { Invoke-FunctionWithParams -FunctionName "New-S365Group" }
                        "3" { Invoke-FunctionWithParams -FunctionName "Get-S365GroupMember" }
                        "4" { Invoke-FunctionWithParams -FunctionName "Add-S365GroupMember" }
                        "5" { Invoke-FunctionWithParams -FunctionName "Remove-S365GroupMember" }
                        "6" { Invoke-FunctionWithParams -FunctionName "Get-S365GroupOwner" }
                        "7" { Invoke-FunctionWithParams -FunctionName "Add-S365GroupOwner" }
                        "8" { Invoke-FunctionWithParams -FunctionName "Get-S365Team" }
                        "9" { Invoke-FunctionWithParams -FunctionName "New-S365Team" }
                        "10" { Invoke-FunctionWithParams -FunctionName "Get-S365TeamChannel" }
                        "11" { Invoke-FunctionWithParams -FunctionName "Add-S365TeamMember" }
                        "12" { Invoke-FunctionWithParams -FunctionName "Remove-S365TeamMember" }
                        default { Write-Warning "Seleção inválida." }
                    }
                }
                "Exchange" {
                    switch ($selection) {
                        "1" { Invoke-FunctionWithParams -FunctionName "Get-S365Mailbox" }
                        "2" { Invoke-FunctionWithParams -FunctionName "Get-S365MailboxStatistics" }
                        "3" { Invoke-FunctionWithParams -FunctionName "Set-S365MailboxQuota" }
                        "4" { Invoke-FunctionWithParams -FunctionName "Set-S365MailboxForwarding" }
                        "5" { Invoke-FunctionWithParams -FunctionName "New-S365SharedMailbox" }
                        "6" { Invoke-FunctionWithParams -FunctionName "Set-S365LitigationHold" }
                        "7" { Invoke-FunctionWithParams -FunctionName "Add-S365MailboxPermission" }
                        "8" { Invoke-FunctionWithParams -FunctionName "Remove-S365MailboxPermission" }
                        "9" { Invoke-FunctionWithParams -FunctionName "Add-S365RecipientPermission" }
                        "10" { Invoke-FunctionWithParams -FunctionName "Remove-S365RecipientPermission" }
                        "11" { Invoke-FunctionWithParams -FunctionName "Get-S365MailboxFolderPermission" }
                        "12" { Invoke-FunctionWithParams -FunctionName "Add-S365MailboxFolderPermission" }
                        "13" { Invoke-FunctionWithParams -FunctionName "Remove-S365MailboxFolderPermission" }
                        "14" { Invoke-FunctionWithParams -FunctionName "Get-S365DistributionGroup" }
                        "15" { Invoke-FunctionWithParams -FunctionName "Get-S365DistributionGroupMember" }
                        "16" { Invoke-FunctionWithParams -FunctionName "Add-S365DistributionGroupMember" }
                        "17" { Invoke-FunctionWithParams -FunctionName "Remove-S365DistributionGroupMember" }
                        "18" { Invoke-FunctionWithParams -FunctionName "Get-S365MailContact" }
                        "19" { Invoke-FunctionWithParams -FunctionName "Get-S365TransportRule" }
                        "20" { Invoke-FunctionWithParams -FunctionName "Set-S365TransportRuleState" }
                        "21" { Invoke-FunctionWithParams -FunctionName "Start-S365MessageTrace" }
                        "22" { Invoke-FunctionWithParams -FunctionName "Start-S365HistoricalMessageTrace" }
                        "23" { Invoke-FunctionWithParams -FunctionName "Get-S365HistoricalMessageTrace" }
                        default { Write-Warning "Seleção inválida." }
                    }
                }
                "LicenseAudit" {
                     switch ($selection) {
                        "1" { Invoke-FunctionWithParams -FunctionName "Get-S365AvailableSkus" }
                        "2" { Invoke-FunctionWithParams -FunctionName "Get-S365UserLicense" }
                        "3" { Invoke-FunctionWithParams -FunctionName "Set-S365UserLicense" }
                        "4" { Invoke-FunctionWithParams -FunctionName "Get-S365LastLogonTime" }
                        "5" { Invoke-FunctionWithParams -FunctionName "Search-S365AuditLog" }
                         default { Write-Warning "Seleção inválida." }
                     }
                }
            }
        } catch {
             Write-Error "Ocorreu um erro inesperado no menu: $_"
             Read-Host "`nPressione Enter para continuar..."
        }

        # Pausa para ver o resultado da ação (exceto se for apenas navegação)
        if ($selection -notmatch '^[bBqQ]$' -and ($currentMenu -ne "Main" -or $selection -notin @("1","2","3","4","5"))) {
            # A pausa já está dentro de Invoke-FunctionWithParams
            # Mas se a seleção for inválida, damos uma pausa aqui.
            if($LASTEXITCODE -ne 0 -and -not $?) { Start-Sleep 1 }
        }

    } # Fim do While
}

# =========================================================================
#                    START TUI
# =========================================================================
Write-Host "`nCarregando SuperAdmin365 TUI..." -ForegroundColor Magenta
Start-SuperAdminTUI
