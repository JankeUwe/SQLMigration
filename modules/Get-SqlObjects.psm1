# =============================================================================
# Modul: Get-SqlObjects.psm1
# Zweck: SQL-Objekte vom Quell- und Ziel-Server auslesen
#        Datenbanken, Logins, User/Rollen, Linked Server,
#        SQL Agent Jobs, Credentials, Proxies
# =============================================================================

function Get-SqlDatabases {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]$Server,
        [switch]$ExcludeSystem
    )
    try {
        $params = @{ SqlInstance = $Server }
        if ($ExcludeSystem) { $params['ExcludeSystem'] = $true }
        $dbs = Get-DbaDatabase @params -ErrorAction Stop |
               Select-Object Name, Status, RecoveryModel, SizeMB,
                             Compatibility, Owner, CreateDate, IsAccessible,
                             @{N='SizeGB';E={[math]::Round($_.SizeMB/1024,2)}}
        Write-MigrationLog -Level 'INFO' -Category 'GET-DB' `
            -Message "Datenbanken gelesen: $($dbs.Count)" -Detail $Server.Name
        return $dbs
    }
    catch {
        Write-MigrationLog -Level 'ERROR' -Category 'GET-DB' `
            -Message "Fehler beim Lesen der Datenbanken" -Detail $_.Exception.Message
        return @()
    }
}

function Get-SqlLogins {
    [CmdletBinding()]
    param([Parameter(Mandatory)]$Server)
    try {
        $logins = Get-DbaLogin -SqlInstance $Server -ErrorAction Stop |
                  Select-Object Name, LoginType, IsDisabled, IsLocked,
                                HasAccess, MustChangePassword, CreateDate,
                                @{N='IsSystem';E={$_.Name -like '##*' -or $_.Name -eq 'sa' -or $_.Name -like 'NT *' -or $_.Name -like 'BUILTIN*'}}
        Write-MigrationLog -Level 'INFO' -Category 'GET-LOGIN' `
            -Message "Logins gelesen: $($logins.Count)" -Detail $Server.Name
        return $logins
    }
    catch {
        Write-MigrationLog -Level 'ERROR' -Category 'GET-LOGIN' `
            -Message "Fehler beim Lesen der Logins" -Detail $_.Exception.Message
        return @()
    }
}

function Get-SqlDbUsers {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]$Server,
        [string[]]$Databases
    )
    $result = [System.Collections.Generic.List[PSObject]]::new()
    try {
        $dbList = if ($Databases) {
            Get-DbaDatabase -SqlInstance $Server -Database $Databases -ErrorAction Stop
        } else {
            Get-DbaDatabase -SqlInstance $Server -ExcludeSystem -ErrorAction Stop
        }

        foreach ($db in $dbList) {
            $users = Get-DbaDbUser -SqlInstance $Server -Database $db.Name -ErrorAction SilentlyContinue |
                     Select-Object @{N='Database';E={$db.Name}},
                                   Name, LoginType, Login, AuthenticationType,
                                   IsSystemObject, CreateDate
            if ($users) { $result.AddRange([System.Management.Automation.PSObject[]]$users) }
        }
        Write-MigrationLog -Level 'INFO' -Category 'GET-USER' `
            -Message "DB-User gelesen: $($result.Count)" -Detail $Server.Name
    }
    catch {
        Write-MigrationLog -Level 'ERROR' -Category 'GET-USER' `
            -Message "Fehler beim Lesen der DB-User" -Detail $_.Exception.Message
    }
    return $result
}

function Get-SqlLinkedServers {
    [CmdletBinding()]
    param([Parameter(Mandatory)]$Server)
    try {
        $ls = Get-DbaLinkedServer -SqlInstance $Server -ErrorAction Stop |
              Select-Object Name, ProductName, ProviderName, DataSource,
                            IsLinkedServerLocal, RemoteUser, Publisher
        Write-MigrationLog -Level 'INFO' -Category 'GET-LS' `
            -Message "Linked Server gelesen: $($ls.Count)" -Detail $Server.Name
        return $ls
    }
    catch {
        Write-MigrationLog -Level 'ERROR' -Category 'GET-LS' `
            -Message "Fehler beim Lesen der Linked Server" -Detail $_.Exception.Message
        return @()
    }
}

function Get-SqlAgentJobs {
    [CmdletBinding()]
    param([Parameter(Mandatory)]$Server)
    try {
        $jobs = Get-DbaAgentJob -SqlInstance $Server -ErrorAction Stop |
                Select-Object Name, Category, OwnerLoginName, IsEnabled,
                              LastRunDate, LastRunOutcome, NextRunDate,
                              HasSchedule, Description
        Write-MigrationLog -Level 'INFO' -Category 'GET-JOB' `
            -Message "Agent Jobs gelesen: $($jobs.Count)" -Detail $Server.Name
        return $jobs
    }
    catch {
        Write-MigrationLog -Level 'ERROR' -Category 'GET-JOB' `
            -Message "Fehler beim Lesen der Agent Jobs" -Detail $_.Exception.Message
        return @()
    }
}

function Get-SqlCredentials {
    [CmdletBinding()]
    param([Parameter(Mandatory)]$Server)
    try {
        $creds = Get-DbaCredential -SqlInstance $Server -ErrorAction Stop |
                 Select-Object Name, Identity, CreateDate, ModifyDate
        Write-MigrationLog -Level 'INFO' -Category 'GET-CRED' `
            -Message "Credentials gelesen: $($creds.Count)" -Detail $Server.Name
        return $creds
    }
    catch {
        Write-MigrationLog -Level 'ERROR' -Category 'GET-CRED' `
            -Message "Fehler beim Lesen der Credentials" -Detail $_.Exception.Message
        return @()
    }
}

function Get-SqlProxies {
    [CmdletBinding()]
    param([Parameter(Mandatory)]$Server)
    try {
        $proxies = Get-DbaAgentProxy -SqlInstance $Server -ErrorAction Stop |
                   Select-Object Name, CredentialName, IsEnabled, Description
        Write-MigrationLog -Level 'INFO' -Category 'GET-PROXY' `
            -Message "Proxies gelesen: $($proxies.Count)" -Detail $Server.Name
        return $proxies
    }
    catch {
        Write-MigrationLog -Level 'ERROR' -Category 'GET-PROXY' `
            -Message "Fehler beim Lesen der Proxies" -Detail $_.Exception.Message
        return @()
    }
}

Export-ModuleMember -Function Get-SqlDatabases, Get-SqlLogins, Get-SqlDbUsers,
                               Get-SqlLinkedServers, Get-SqlAgentJobs,
                               Get-SqlCredentials, Get-SqlProxies
