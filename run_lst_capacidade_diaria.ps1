Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

$baseDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$scriptPath = Join-Path $baseDir "LST_Capacidade_Diaria_extrator_ecargo_infos.pyw"
$logDir = Join-Path $baseDir "logs"
$logFile = Join-Path $logDir ("lst_capacidade_diaria_" + (Get-Date -Format "yyyy-MM-dd") + ".log")
$mutexName = "Global\LST_Capacidade_Diaria_Ecargo"

function Write-Log {
    param([string]$Message)

    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    Add-Content -LiteralPath $logFile -Value "[$timestamp] $Message"
}

if (-not (Test-Path -LiteralPath $scriptPath)) {
    throw "Arquivo principal nao encontrado: $scriptPath"
}

New-Item -ItemType Directory -Path $logDir -Force | Out-Null

$pythonCommand = Get-Command py -ErrorAction SilentlyContinue
if (-not $pythonCommand) {
    $pythonCommand = Get-Command python -ErrorAction SilentlyContinue
}

if (-not $pythonCommand) {
    throw "Python nao encontrado no PATH."
}

$pythonExe = $pythonCommand.Source
$mutex = [System.Threading.Mutex]::new($false, $mutexName)
$hasHandle = $false

try {
    $hasHandle = $mutex.WaitOne(0, $false)

    if (-not $hasHandle) {
        Write-Log "Execucao ignorada porque ja existe outra instancia em andamento."
        exit 0
    }

    Write-Log "Execucao iniciada."
    Push-Location $baseDir
    try {
        & $pythonExe $scriptPath *>> $logFile
        $exitCode = $LASTEXITCODE
    }
    finally {
        Pop-Location
    }

    if ($exitCode -ne 0) {
        throw "Python retornou codigo $exitCode."
    }

    Write-Log "Execucao finalizada com sucesso."
}
catch {
    Write-Log ("Falha: " + $_.Exception.Message)
    throw
}
finally {
    if ($hasHandle) {
        [void]$mutex.ReleaseMutex()
    }

    $mutex.Dispose()
}
