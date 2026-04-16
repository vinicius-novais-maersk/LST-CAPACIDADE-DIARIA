Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

$taskName = "LST_Capacidade_Diaria_30min"
$baseDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$runnerPath = Join-Path $baseDir "run_lst_capacidade_diaria.ps1"

if (-not (Test-Path -LiteralPath $runnerPath)) {
    throw "Runner nao encontrado: $runnerPath"
}

$now = Get-Date
$nextRun = Get-Date -Hour $now.Hour -Minute 0 -Second 0

if ($now.Minute -lt 30) {
    $nextRun = $nextRun.AddMinutes(30)
}
else {
    $nextRun = $nextRun.AddHours(1)
}

if ($nextRun -le $now) {
    $nextRun = $nextRun.AddMinutes(30)
}

$action = New-ScheduledTaskAction `
    -Execute "powershell.exe" `
    -Argument "-NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -File `"$runnerPath`""

$trigger = New-ScheduledTaskTrigger `
    -Once `
    -At $nextRun `
    -RepetitionInterval (New-TimeSpan -Minutes 30) `
    -RepetitionDuration (New-TimeSpan -Days 3650)

$settings = New-ScheduledTaskSettingsSet `
    -StartWhenAvailable `
    -AllowStartIfOnBatteries `
    -DontStopIfGoingOnBatteries `
    -MultipleInstances IgnoreNew `
    -ExecutionTimeLimit (New-TimeSpan -Hours 2)

$userId = [System.Security.Principal.WindowsIdentity]::GetCurrent().Name
$principal = New-ScheduledTaskPrincipal `
    -UserId $userId `
    -LogonType Interactive `
    -RunLevel Limited

Register-ScheduledTask `
    -TaskName $taskName `
    -Action $action `
    -Trigger $trigger `
    -Settings $settings `
    -Principal $principal `
    -Force | Out-Null

Write-Host "Tarefa registrada com sucesso."
Write-Host "Nome: $taskName"
Write-Host "Primeira execucao: $($nextRun.ToString('yyyy-MM-dd HH:mm:ss'))"
Write-Host "Usuario: $userId"
