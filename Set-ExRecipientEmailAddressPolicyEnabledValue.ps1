[CmdletBinding(DefaultParameterSetName = 'Id')]
param (
    [Parameter(Mandatory, ParameterSetName = 'Id')]
    $Identity,

    [Parameter(Mandatory)]
    [boolean]
    $Value,

    [Parameter()]
    [string]
    $TargetDomainController,

    [Parameter(ParameterSetName = 'All', Mandatory)]
    [switch]
    $All,

    [switch]$Quiet
)

. "$PSScriptRoot\generic_functions.ps1"

$reverseValue = -not $Value

$supported_recipient_types = @(Get-Content "$($PSScriptRoot)\supported_recipient_types.list")

$remote_mailbox_types = $supported_recipient_types | Where-Object { $_ -like "Remote*" }
$mailbox_types = $supported_recipient_types | Where-Object { $_ -like "*Mailbox" -and $_ -notin $remote_mailbox_types }
$distribution_group_types = $supported_recipient_types | Where-Object { $_ -like "*Group" -and $_ -ne 'DynamicDistributionGroup' }

if ($PSCmdlet.ParameterSetName -eq 'All') {
    $Identity = @(Get-Recipient -ResultSize Unlimited -Filter "EmailAddressPolicyEnabled -eq '$($reverseValue.ToString())'")
}

foreach ($id in @($Identity)) {
    # Make sure the recipient exists
    try {
        if ($id.psobject.typenames[0] -like "*Microsoft.Exchange.Data.Directory.Management.ReducedRecipient") {
            $recipient_object = $id
        }
        else {
            $recipient_object = Get-Recipient -Identity $id -ErrorAction Stop
            Write-Information "[$($recipient_object.DisplayName)]: Recipient found!"
        }
    }
    catch {
        Write-Information $_.Exception.Message
    }

    # Check if recipient type is supported
    if ($recipient_object.RecipientTypeDetails -notin $supported_recipient_types) {
        Write-Information "[$($recipient_object.PrimarySmtpAddress.Address)]: This recipient type [$($recipient_object.RecipientTypeDetails)] is not supported."
        continue
    }

    $params = @{
        Identity                  = $recipient_object.Identity
        EmailAddressPolicyEnabled = $Value
    }

    if ($TargetDomainController) {
        $params.Add('DomainController', $TargetDomainController)
    }

    # If the recipient type is mailbox (not remote)
    if ($recipient_object.RecipientTypeDetails -in $mailbox_types) {
        try {
            Set-Mailbox @params -ErrorAction Stop
            Write-Information "[$($recipient_object.DisplayName)]: Set email address policy enabled - $($Value)."
        }
        catch {
            Write-Information $_.Exception.Message
        }

    }

    # If the recipient type is remote mailbox
    if ($recipient_object.RecipientTypeDetails -in $remote_mailbox_types) {
        try {
            Set-RemoteMailbox @params -ErrorAction Stop
            Write-Information "[$($recipient_object.DisplayName)]: Set email address policy enabled - $($Value)."
        }
        catch {
            Write-Information $_.Exception.Message
        }
    }

    # If the recipient type is a distribution group (DG, SG, DDL)
    if ($recipient_object.RecipientTypeDetails -in $distribution_group_types) {
        try {
            Set-DistributionGroup @params -ErrorAction Stop
            Write-Information "[$($recipient_object.DisplayName)]: Set email address policy enabled - $($Value)."
        }
        catch {
            Write-Information $_.Exception.Message
        }
    }

    # If the recipient type is MailUser
    if ($recipient_object.RecipientTypeDetails -eq 'MailUser') {
        try {
            Set-MailUser @params -ErrorAction Stop
            Write-Information "[$($recipient_object.DisplayName)]: Set email address policy enabled - $($Value)."
        }
        catch {
            Write-Information $_.Exception.Message
        }
    }

    # If the recipient type is MailContact
    if ($recipient_object.RecipientTypeDetails -eq 'MailContact') {
        try {
            Set-MailContact @params -ErrorAction Stop
            Write-Information "[$($recipient_object.DisplayName)]: Set email address policy enabled - $($Value)."
        }
        catch {
            Write-Information $_.Exception.Message
        }
    }

    # If the recipient type is Dynamic
    if ($recipient_object.RecipientTypeDetails -eq 'DynamicDistributionGroup') {
        try {
            Set-DynamicDistributionGroup @params -ErrorAction Stop
            Write-Information "[$($recipient_object.DisplayName)]: Set email address policy enabled - $($Value)."
        }
        catch {
            Write-Information $_.Exception.Message
        }
    }
}