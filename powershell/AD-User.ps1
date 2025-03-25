

function Get-CurrentUserOU {
    # Get the current user's distinguished name
    $currentUserDN = [System.DirectoryServices.AccountManagement.UserPrincipal]::Current.DistinguishedName
    # Extract the OU from the distinguished name
    $ou = $currentUserDN -replace '^CN=.*?,(OU=.*)$','$1'
    # Output the OU
    Write-Output $ou
}

$currentUserOU  = Get-CurrentUserOU

$DefaultLDAPPath = ("LDAP://$currentUserOU")

function Get-UsersInOU {
    param (
        [string]$LDAPPath = $DefaultLDAPPath
    )
    $searcher = New-Object DirectoryServices.DirectorySearcher([ADSI]$LDAPPath)
    $searcher.PageSize = 10000
    $searcher.Filter = "(objectClass=user)"
    $searcher.FindAll() | ForEach-Object {
        $entry = $_.GetDirectoryEntry()
        [PSCustomObject]@{ 
            Name = $entry.Name
            Path = $entry.Path
            Mail = $entry.mail
        }
    }
}

$users = Get-UsersInOU
$users