[CmdletBinding(DefaultParameterSetName = 'byDomain')]
param (
    [Parameter(Mandatory)]
    [string[]]$Identity,

    [Parameter(ParameterSetName = 'byDomain', Mandatory)]
    [string[]]$Domain,

    [Parameter(ParameterSetName = 'byProxyAddress', Mandatory)]
    [string[]]$ProxyAddress,

    [string]$TargetDomainController,
    [string]$OutputCsv,
    [switch]$ReturnResult
)

function Test-IsValidFqdn {
    param ([string]$Fqdn)
    return $Fqdn -match '^(?=.{1,253}$)(([a-zA-Z0-9]|[a-zA-Z0-9][a-zA-Z0-9\-]*[a-zA-Z0-9])\.)*([A-Za-z0-9]|[A-Za-z0-9][A-Za-z0-9\-]*[A-Za-z0-9])$'
}

$now = [datetime]::Now

if (!$OutputCsv) {
    $OutputCsv = "$($PSScriptRoot)\remove_proxy_result-$($now.ToString('yyyy-MM-ddTHH-mm-ss')).csv"
}

$supported_recipient_types = @(Get-Content "$PSScriptRoot\supported_recipient_types.list")
$remote_mailbox_types = $supported_recipient_types | Where-Object { $_ -like "Remote*" }
$mailbox_types = $supported_recipient_types | Where-Object { $_ -like "*Mailbox" -and $_ -notin $remote_mailbox_types }
$distribution_group_types = $supported_recipient_types | Where-Object { $_ -like "*Group" -and $_ -ne 'DynamicDistributionGroup' }

# Results will be collected here
$results = @()

:outer foreach ($id in $Identity) {
    try {
        $recipient = Get-Recipient -Identity $id -ErrorAction Stop
        Write-Information "[$($recipient.DisplayName)]: Recipient found!"
    }
    catch {
        Write-Information "[UNKNOWN]: NOT OK - $($_.Exception.Message)"
        continue
    }

    if ($recipient.RecipientTypeDetails -notin $supported_recipient_types) {
        Write-Information "[$($recipient.PrimarySmtpAddress.Address)]: Unsupported recipient type [$($recipient.RecipientTypeDetails)]"
        continue
    }

    $proxyToRemove = @()
    # $matchCriteria = ""

    switch ($PSCmdlet.ParameterSetName) {
        'byDomain' {
            $proxyToRemove = $recipient.EmailAddresses |
            Where-Object {
                $emailDomain = $_.AddressString.Split('@')[-1]
                $Domain -contains $emailDomain -and $_.AddressString -ne $recipient.PrimarySmtpAddress.Address
            } |
            Select-Object -ExpandProperty ProxyAddressString

            if (-not $proxyToRemove) {
                Write-Information "[$($recipient.DisplayName)]: No proxy addresses matched for domain removal."
                continue outer
            }

            Write-Information "[$($recipient.DisplayName)]: Proxy addresses to remove (domain match): $($proxyToRemove -join ', ')"
        }

        'byProxyAddress' {
            $proxyToRemove = $recipient.EmailAddresses |
            Where-Object { $ProxyAddress -contains $_.AddressString } |
            Select-Object -ExpandProperty ProxyAddressString

            if (-not $proxyToRemove) {
                Write-Information "[$($recipient.DisplayName)]: No proxy addresses matched for email address removal."
                continue outer
            }

            Write-Information "[$($recipient.DisplayName)]: Proxy addresses to remove (exact match): $($proxyToRemove -join ', ')"
        }
    }

    $params = @{
        Identity       = $recipient.Identity
        EmailAddresses = @{ remove = $proxyToRemove }
    }

    if ($recipient.EmailAddressPolicyEnabled) {
        $params.EmailAddressPolicyEnabled = $false
    }

    if ($TargetDomainController) {
        $params.DomainController = $TargetDomainController
    }

    $setCommand = switch ($recipient.RecipientTypeDetails) {
        { $_ -in $mailbox_types } { 'Set-Mailbox' }
        { $_ -in $remote_mailbox_types } { 'Set-RemoteMailbox' }
        { $_ -in $distribution_group_types } { 'Set-DistributionGroup' }
        'MailUser' { 'Set-MailUser' }
        'MailContact' { 'Set-MailContact' }
        'DynamicDistributionGroup' { 'Set-DynamicDistributionGroup' }
        default { $null }
    }

    if ($setCommand) {
        try {
            & $setCommand @params -ErrorAction Stop
            Write-Information "[$($recipient.DisplayName)]: Removed proxy addresses OK."

            # Re-fetch updated recipient to get remaining proxy addresses
            $updatedRecipient = Get-Recipient -Identity $recipient.Identity -ErrorAction Stop

            $remainingProxies = $updatedRecipient.EmailAddresses |
            Where-Object { $_.AddressString -ne $updatedRecipient.PrimarySmtpAddress.Address } |
            Select-Object -ExpandProperty ProxyAddressString

            # Add success entry to results
            $results += [PSCustomObject]@{
                DisplayName           = $updatedRecipient.DisplayName
                PrimarySMTPAddress    = $updatedRecipient.PrimarySmtpAddress.Address
                RemovedProxyAddress   = ($proxyToRemove -join ', ')
                RemainingProxyAddress = ($remainingProxies -join ', ')
            }
        }
        catch {
            Write-Information "[$($recipient.DisplayName)]: NOT OK - $($_.Exception.Message)"
        }
    }

    else {
        Write-Information "[$($recipient.DisplayName)]: No matching Set-* command for recipient type [$($recipient.RecipientTypeDetails)]"
    }
}

# Export results to CSV if path provided
# if ($OutputCsv) {
if ($results) {
    try {
        $results | Export-Csv -Path $OutputCsv -NoTypeInformation -Encoding UTF8
        Write-Information "Output written to: $OutputCsv"
    }
    catch {
        Write-Information "ERROR writing to CSV: $($_.Exception.Message)"
    }
    if ($ReturnResult) {
        $results
    }
}



