[CmdletBinding(DefaultParameterSetName = 'byDomain')]
param (
    [Parameter(Mandatory)]
    $Identity,

    [Parameter(ParameterSetName = 'byDomain', Mandatory)]
    [string[]]$DomainName,

    [Parameter(ParameterSetName = 'byProxyAddress', Mandatory)]
    [string[]]$ProxyAddress,

    [string]$TargetDomainController,
    [string]$OutputCsv,
    [switch]$ReturnResult,
    [switch]$Quiet
)



. "$PSScriptRoot\generic_functions.ps1"

if (!$Quiet) {
    $InformationPreference = 'Continue'
}
else {
    $InformationPreference = 'SilentlyContinue'
}

if ($PSCmdlet.ParameterSetName -eq 'byDomain') {
    $acceptedDomain = (Get-AcceptedDomain).DomainName
    $tempDomainCollection = [string[]]@()

    foreach ($domain in $DomainName) {
        if ($domain -notin $acceptedDomain) {
            SayInfo "The domain [$($domain)] is not an accepted domain in this organization and will be ignored."
        }
        else {
            $tempDomainCollection += $domain
        }
    }

    if (!$tempDomainCollection) {
        SayInfo "No valid domain to be removed. Script terminated."
        return $null
    }
    else {
        $DomainName = $tempDomainCollection
        SayInfo "Proxy address domain to match = $($DomainName -join ", ")"
    }
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
    SayInfo $recipient.DisplayName
    try {
        $recipient = Get-Recipient -Identity $id -ErrorAction Stop
        SayInfo "  -> Recipient found!"
    }
    catch {
        SayInfo "  -> NOT OK - $($_.Exception.Message)"
        continue
    }

    if ($recipient.RecipientTypeDetails -notin $supported_recipient_types) {
        SayInfo "  -> Unsupported recipient type [$($recipient.RecipientTypeDetails)]"
        continue
    }

    $proxyToRemove = @()

    switch ($PSCmdlet.ParameterSetName) {
        'byDomain' {
            $proxyToRemove = $recipient.EmailAddresses |
            Where-Object {
                $emailDomain = $_.AddressString.Split('@')[-1]
                $DomainName -contains $emailDomain -and $_.AddressString -ne $recipient.PrimarySmtpAddress.Address
            } |
            Select-Object -ExpandProperty ProxyAddressString

            if (-not $proxyToRemove) {
                SayInfo "  -> No proxy addresses matched for domain removal."
                continue outer
            }

            SayInfo "  -> Proxy addresses to remove (domain match): $($proxyToRemove -join ', ')"
        }

        'byProxyAddress' {
            $proxyToRemove = $recipient.EmailAddresses |
            Where-Object { $ProxyAddress -contains $_.AddressString } |
            Select-Object -ExpandProperty ProxyAddressString

            if (-not $proxyToRemove) {
                SayInfo "  -> No proxy addresses matched for email address removal."
                continue outer
            }

            SayInfo "  -> Proxy addresses to remove (exact match): $($proxyToRemove -join ', ')"
        }
    }

    $params = @{
        Identity       = $recipient.GUID.ToString()
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
            SayInfo "  -> Removed proxy addresses OK."

            # Re-fetch updated recipient to get remaining proxy addresses
            $updatedRecipient = Get-Recipient -Identity $recipient.GUID.ToString() -ErrorAction Stop

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
            SayInfo " -> NOT OK - $($_.Exception.Message)"
        }
    }

    else {
        SayInfo " -> No matching Set-* command for recipient type [$($recipient.RecipientTypeDetails)]"
    }
}

# Export results to CSV if path provided
# if ($OutputCsv) {
if ($results) {
    try {
        $results | Export-Csv -Path $OutputCsv -NoTypeInformation -Encoding UTF8
        SayInfo "Output written to: $OutputCsv"
    }
    catch {
        SayInfo "ERROR writing to CSV: $($_.Exception.Message)"
    }
    if ($ReturnResult) {
        $results
    }
}
else {
    SayInfo "No results."
}



