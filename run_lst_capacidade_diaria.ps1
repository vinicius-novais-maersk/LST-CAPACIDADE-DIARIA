Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

$baseDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$scriptPath = Join-Path $baseDir "LST_Capacidade_Diaria_extrator_ecargo_infos.pyw"
$logDir = Join-Path $baseDir "logs"
$logFile = Join-Path $logDir ("lst_capacidade_diaria_" + (Get-Date -Format "yyyy-MM-dd") + ".log")
$mutexName = "Global\LST_Capacidade_Diaria_Ecargo"
$executionTimeout = New-TimeSpan -Minutes 25

function Write-Log {
    param([string]$Message)

    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    "[$timestamp] $Message" | Out-File -LiteralPath $logFile -Append -Encoding utf8
}

function Write-ProcessOutput {
    param([string]$Text)

    if ([string]::IsNullOrEmpty($Text)) {
        return
    }

    $reader = [System.IO.StringReader]::new($Text)
    try {
        while ($true) {
            $line = $reader.ReadLine()
            if ($null -eq $line) {
                break
            }

            if ($line.Length -gt 0) {
                $line | Out-File -LiteralPath $logFile -Append -Encoding utf8
            }
        }
    }
    finally {
        $reader.Dispose()
    }
}

function Stop-ProcessTree {
    param([int]$ProcessId)

    $childProcessIds = @(
        Get-CimInstance Win32_Process -Filter "ParentProcessId = $ProcessId" -ErrorAction SilentlyContinue |
            Select-Object -ExpandProperty ProcessId
    )

    foreach ($childProcessId in $childProcessIds) {
        Stop-ProcessTree -ProcessId $childProcessId
    }

    Stop-Process -Id $ProcessId -Force -ErrorAction SilentlyContinue
}

if (-not (Test-Path -LiteralPath $scriptPath)) {
    throw "Arquivo principal nao encontrado: $scriptPath"
}

New-Item -ItemType Directory -Path $logDir -Force | Out-Null

$pythonCommand = Get-Command python -ErrorAction SilentlyContinue
if (-not $pythonCommand) {
    $pythonCommand = Get-Command py -ErrorAction SilentlyContinue
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
    $startInfo = [System.Diagnostics.ProcessStartInfo]::new()
    $startInfo.FileName = $pythonExe
    $startInfo.Arguments = "-X utf8 -u `"$scriptPath`""
    $startInfo.WorkingDirectory = $baseDir
    $startInfo.UseShellExecute = $false
    $startInfo.CreateNoWindow = $true
    $startInfo.RedirectStandardOutput = $true
    $startInfo.RedirectStandardError = $true
    $startInfo.EnvironmentVariables["PYTHONIOENCODING"] = "utf-8"

    if ($startInfo.PSObject.Properties.Name -contains "StandardOutputEncoding") {
        $startInfo.StandardOutputEncoding = [System.Text.Encoding]::UTF8
        $startInfo.StandardErrorEncoding = [System.Text.Encoding]::UTF8
    }

    $process = [System.Diagnostics.Process]::new()
    $process.StartInfo = $startInfo

    try {
        [void]$process.Start()
        $stdoutTask = $process.StandardOutput.ReadToEndAsync()
        $stderrTask = $process.StandardError.ReadToEndAsync()
        $timedOut = -not $process.WaitForExit([int]$executionTimeout.TotalMilliseconds)

        if ($timedOut) {
            Write-Log(
                "Falha: Python excedeu o limite de $([int]$executionTimeout.TotalMinutes) minutos. " +
                "Encerrando a arvore de processos."
            )
            Stop-ProcessTree -ProcessId $process.Id
        }

        $process.WaitForExit()
        $stdout = $stdoutTask.GetAwaiter().GetResult()
        $stderr = $stderrTask.GetAwaiter().GetResult()
        $exitCode = $process.ExitCode

        if ($timedOut) {
            throw "Python excedeu o limite de $([int]$executionTimeout.TotalMinutes) minutos."
        }
    }
    finally {
        $process.Dispose()
    }

    Write-ProcessOutput -Text $stdout
    Write-ProcessOutput -Text $stderr

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
