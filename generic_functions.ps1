function Test-IsValidFqdn {
    param ([string]$Fqdn)
    return $Fqdn -match '^(?=.{1,253}$)(([a-zA-Z0-9]|[a-zA-Z0-9][a-zA-Z0-9\-]*[a-zA-Z0-9])\.)*([A-Za-z0-9]|[A-Za-z0-9][A-Za-z0-9\-]*[A-Za-z0-9])$'
}

function SayInfo {
    param (
        [string]$Text
    )

    Write-Information "[$(Get-Date -Format G)]: $($Text)"
}