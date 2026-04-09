Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

$taskName = "LST_Capacidade_Diaria_30min"

if (Get-ScheduledTask -TaskName $taskName -ErrorAction SilentlyContinue) {
    Unregister-ScheduledTask -TaskName $taskName -Confirm:$false
    Write-Host "Tarefa removida: $taskName"
}
else {
    Write-Host "Tarefa nao encontrada: $taskName"
}
