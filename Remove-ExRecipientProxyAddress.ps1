[CmdletBinding(DefaultParameterSetName = 'byDomain')]
param (
    [Parameter(Mandatory)]
    [string[]]
    $Identity,

    [Parameter(ParameterSetName = 'byDomain', Mandatory)]
    [string[]]
    $Domain,

    [Parameter(ParameterSetName = 'byProxyAddress', Mandatory)]
    [string[]]
    $ProxyAddress
)

# $InformationPreference = 'Continue'

#Region Function
function Test-IsValidFqdn {
    param (
        [string]$Fqdn
    )

    $FqdnPattern = '^(?=.{1,253}$)(([a-zA-Z0-9]|[a-zA-Z0-9][a-zA-Z0-9\-]*[a-zA-Z0-9])\.)*([A-Za-z0-9]|[A-Za-z0-9][A-Za-z0-9\-]*[A-Za-z0-9])$'

    if ($Fqdn -match $FqdnPattern) {
        return $true
    }
    else {
        return $false
    }
}
#EndRegion Function

$suported_recipient_types = @(
    'EquipmentMailbox',
    'MailContact',
    'MailNonUniversalGroup',
    'MailUniversalDistributionGroup',
    'MailUniversalSecurityGroup',
    'DynamicDistributionGroup',
    'MailUser',
    'RemoteEquipmentMailbox',
    'RemoteRoomMailbox',
    'RemoteSharedMailbox',
    'RemoteTeamMailbox',
    'RemoteUserMailbox',
    'RoomMailbox',
    'SchedulingMailbox',
    'SharedMailbox',
    'UserMailbox'
)

$remote_mailbox_types = $suported_recipient_types | Where-Object { $_ -like "Remote*" }
$mailbox_types = $suported_recipient_types | Where-Object { $_ -like "*Mailbox" -and $_ -notin $remote_mailbox_types }
$distribution_group_types = $suported_recipient_types | Where-Object { $_ -like "*Group" -and $_ -ne 'DynamicDistributionGroup' }

foreach ($id in @($Identity)) {
    # Make sure the recipient exists
    try {
        $recipient_object = Get-Recipient -Identity $id -ErrorAction Stop
        Write-Information "[$($recipient_object.DisplayName)]: Recipient found!"
    }
    catch {
        Write-Information $_.Exception.Message
    }

    # Check if recipient type is supported
    if ($recipient_object.RecipientTypeDetails -notin $suported_recipient_types) {
        Write-Information "[$($recipient_object.PrimarySmtpAddress.Address)]: This recipient type [$($recipient_object.RecipientTypeDetails)] is not supported."
        continue
    }

    # Check if proxyaddress domain match is present
    if ($PSCmdlet.ParameterSetName -eq 'byDomain') {
        $proxy_address_to_remove = [system.collections.arraylist]@()
        $recipient_object.EmailAddresses | Where-Object {
            $emailDomain = $_.AddressString.Split('@')[-1]
            $Domain -contains $emailDomain
        } | ForEach-Object {
            $_.ProxyAddressString
        } | ForEach-Object { $null = $proxy_address_to_remove.Add($_) }

        # Skip if there are zero matches
        if ($proxy_address_to_remove.Count -eq 0) {
            Write-Information "[$($recipient_object.DisplayName)]: There are ($($proxy_address_to_remove.Count)) proxy address domain match to remove."
            continue
        }
        Write-Information "[$($recipient_object.DisplayName)]: There are ($($proxy_address_to_remove.Count)) proxy address domain match to remove = $($proxy_address_to_remove -join " ,")"
    }

    # Check if proxyaddress email address match is present
    if ($PSCmdlet.ParameterSetName -eq 'byProxyAddress') {
        $proxy_address_to_remove = [system.collections.arraylist]@()
        $recipient_object.EmailAddresses | Where-Object {
            $emailAddress = $_.AddressString
            $ProxyAddress -contains $emailAddress

        } | ForEach-Object {
            $_.ProxyAddressString
        } | ForEach-Object { $null = $proxy_address_to_remove.Add($_) }

        # Skip if there are zero matches
        if ($proxy_address_to_remove.Count -eq 0) {
            Write-Information "[$($recipient_object.DisplayName)]: There are ($($proxy_address_to_remove.Count)) proxy address domain match to remove."
            continue
        }
        Write-Information "[$($recipient_object.DisplayName)]: There are ($($proxy_address_to_remove.Count)) proxy address domain match to remove = $($proxy_address_to_remove -join " ,")"
    }

    $params = @{
        Identity       = $recipient_object.Identity
        EmailAddresses = @{remove = $proxy_address_to_remove }
    }
    if ($recipient_object.EmailAddressPolicyEnabled) {
        $params.Add('EmailAddressPolicyEnabled', $false)
    }

    # If the recipient type is mailbox (not remote)
    if ($recipient_object.RecipientTypeDetails -in $mailbox_types) {
        try {
            Set-Mailbox @params -ErrorAction Stop
            Write-Information "[$($recipient_object.DisplayName)]: Removed proxy addresses OK."
        }
        catch {
            Write-Information $_.Exception.Message
        }

    }

    # If the recipient type is remote mailbox
    if ($recipient_object.RecipientTypeDetails -in $remote_mailbox_types) {
        try {
            Set-RemoteMailbox @params -ErrorAction Stop
            Write-Information "[$($recipient_object.DisplayName)]: Removed proxy addresses OK."
        }
        catch {
            Write-Information $_.Exception.Message
        }
    }

    # If the recipient type is a distribution group (DG, SG, DDL)
    if ($recipient_object.RecipientTypeDetails -in $distribution_group_types) {
        try {
            Set-DistributionGroup @params -ErrorAction Stop
            Write-Information "[$($recipient_object.DisplayName)]: Removed proxy addresses OK."
        }
        catch {
            Write-Information $_.Exception.Message
        }
    }

    # If the recipient type is MailUser
    if ($recipient_object.RecipientTypeDetails -eq 'MailUser') {
        try {
            Set-MailUser @params -ErrorAction Stop
            Write-Information "[$($recipient_object.DisplayName)]: Removed proxy addresses OK."
        }
        catch {
            Write-Information $_.Exception.Message
        }
    }

    # If the recipient type is MailContact
    if ($recipient_object.RecipientTypeDetails -eq 'MailContact') {
        try {
            Set-MailContact @params -ErrorAction Stop
            Write-Information "[$($recipient_object.DisplayName)]: Removed proxy addresses OK."
        }
        catch {
            Write-Information $_.Exception.Message
        }
    }

    # If the recipient type is Dynamic
    if ($recipient_object.RecipientTypeDetails -eq 'DynamicDistributionGroup') {
        try {
            Set-DynamicDistributionGroup @params -ErrorAction Stop
            Write-Information "[$($recipient_object.DisplayName)]: Removed proxy addresses OK."
        }
        catch {
            Write-Information $_.Exception.Message
        }
    }
}

