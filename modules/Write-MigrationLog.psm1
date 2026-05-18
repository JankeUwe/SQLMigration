# =============================================================================
# Modul: Write-MigrationLog.psm1
# Zweck: Einheitliches Logging fuer die SQL-Migration (Text + CSV)
# =============================================================================

$script:LogFile  = $null
$script:CsvFile  = $null
$script:LogMutex = New-Object System.Threading.Mutex($false, 'SQLMigrationLogMutex')

function Initialize-MigrationLog {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$LogDirectory,
        [string]$Prefix = 'SQL-Migration'
    )

    if (-not (Test-Path $LogDirectory)) {
        New-Item -ItemType Directory -Path $LogDirectory -Force | Out-Null
    }

    $stamp = Get-Date -Format 'yyyyMMdd_HHmmss'
    $script:LogFile = Join-Path $LogDirectory "$Prefix`_$stamp.log"
    $script:CsvFile = Join-Path $LogDirectory "$Prefix`_$stamp.csv"

    # CSV-Header
    "Timestamp;Level;Category;Message;Detail" | Out-File -FilePath $script:CsvFile -Encoding UTF8

    Write-MigrationLog -Level 'INFO' -Category 'INIT' -Message 'Logdatei initialisiert' -Detail $script:LogFile
    return $script:LogFile
}

function Write-MigrationLog {
    [CmdletBinding()]
    param(
        [ValidateSet('INFO','WARN','ERROR','SUCCESS','DEBUG','STEP')]
        [string]$Level = 'INFO',
        [string]$Category = 'GENERAL',
        [Parameter(Mandatory)][string]$Message,
        [string]$Detail = ''
    )

    $ts     = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
    $line   = "$ts  [$($Level.PadRight(7))]  [$($Category.PadRight(15))]  $Message"
    if ($Detail) { $line += "  >>  $Detail" }

    $csvLine = '"' + $ts + '";"' + $Level + '";"' + $Category + '";"' +
               $Message.Replace('"','""') + '";"' + $Detail.Replace('"','""') + '"'

    # Thread-sicheres Schreiben
    $null = $script:LogMutex.WaitOne(5000)
    try {
        if ($script:LogFile) {
            Add-Content -Path $script:LogFile -Value $line  -Encoding UTF8
            Add-Content -Path $script:CsvFile -Value $csvLine -Encoding UTF8
        }
    }
    finally {
        $script:LogMutex.ReleaseMutex()
    }

    # Farbige Konsolenausgabe
    $color = switch ($Level) {
        'ERROR'   { 'Red' }
        'WARN'    { 'Yellow' }
        'SUCCESS' { 'Green' }
        'STEP'    { 'Cyan' }
        'DEBUG'   { 'DarkGray' }
        default   { 'Gray' }
    }
    Write-Host $line -ForegroundColor $color
}

function Get-MigrationLogPath { return $script:LogFile }
function Get-MigrationCsvPath { return $script:CsvFile }

Export-ModuleMember -Function Initialize-MigrationLog, Write-MigrationLog,
                               Get-MigrationLogPath, Get-MigrationCsvPath
