# =============================================================================
# Modul: Connect-SqlServer.psm1
# Zweck: Verbindungsaufbau zu SQL Server (Windows-Auth + SQL-Login)
#        Kompatibel mit PS 5.1 und PS 7.x / dbaTools 1.x und 2.x
# =============================================================================

function New-SqlConnection {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$ServerInstance,
        [ValidateSet('Windows','SqlLogin')]
        [string]$AuthMode = 'Windows',
        [string]$SqlUser,
        [System.Security.SecureString]$SqlPassword,
        [int]$ConnectTimeout = 30,
        [switch]$TrustServerCertificate
    )

    Write-MigrationLog -Level 'STEP' -Category 'CONNECT' -Message "Verbinde mit $ServerInstance (Auth: $AuthMode)"

    try {
        $params = @{
            SqlInstance              = $ServerInstance
            ConnectTimeout           = $ConnectTimeout
            TrustServerCertificate   = $TrustServerCertificate.IsPresent
        }

        if ($AuthMode -eq 'SqlLogin') {
            if (-not $SqlUser -or -not $SqlPassword) {
                throw "SQL-Login erfordert SqlUser und SqlPassword"
            }
            $cred = New-Object System.Management.Automation.PSCredential($SqlUser, $SqlPassword)
            $params['SqlCredential'] = $cred
        }

        $server = Connect-DbaInstance @params -ErrorAction Stop

        $version = $server.VersionString
        $edition = $server.Edition
        Write-MigrationLog -Level 'SUCCESS' -Category 'CONNECT' `
            -Message "Verbindung hergestellt: $ServerInstance" `
            -Detail "Version: $version | Edition: $edition"

        return $server
    }
    catch {
        Write-MigrationLog -Level 'ERROR' -Category 'CONNECT' `
            -Message "Verbindung fehlgeschlagen: $ServerInstance" `
            -Detail $_.Exception.Message
        throw
    }
}

function Test-SqlConnection {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$ServerInstance,
        [ValidateSet('Windows','SqlLogin')]
        [string]$AuthMode = 'Windows',
        [string]$SqlUser,
        [System.Security.SecureString]$SqlPassword,
        [int]$ConnectTimeout = 10,
        [switch]$TrustServerCertificate
    )

    try {
        $params = @{
            SqlInstance            = $ServerInstance
            ConnectTimeout         = $ConnectTimeout
            TrustServerCertificate = $TrustServerCertificate.IsPresent
        }
        if ($AuthMode -eq 'SqlLogin') {
            $cred = New-Object System.Management.Automation.PSCredential($SqlUser, $SqlPassword)
            $params['SqlCredential'] = $cred
        }
        $result = Test-DbaConnection @params -ErrorAction Stop
        return $result.ConnectSuccess
    }
    catch {
        return $false
    }
}

Export-ModuleMember -Function New-SqlConnection, Test-SqlConnection
