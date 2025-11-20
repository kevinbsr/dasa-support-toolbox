<#
.SYNOPSIS
    DASA SUPPORT TOOLBOX v1.0 (Stable CLI)
.DESCRIPTION
    Versao compatibilidade (ASCII) para evitar erros de caracteres.
#>

$currentPrincipal = New-Object Security.Principal.WindowsPrincipal([Security.Principal.WindowsIdentity]::GetCurrent())
if (-not $currentPrincipal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    Write-Warning "ERRO: Execute como ADMINISTRADOR!"
    Start-Sleep 3
    Exit
}

$RepoURL   = "https://raw.githubusercontent.com/kevinbsr/dasa-support-toolbox/main"
$AssetsURL = "$RepoURL/assets"
$TempDir   = "$env:TEMP\DasaToolbox"

if (Test-Path $TempDir) { Remove-Item $TempDir -Recurse -Force -ErrorAction SilentlyContinue }
New-Item -ItemType Directory -Force -Path $TempDir | Out-Null

function Get-RemoteCode {
    param($ModuleName)
    Write-Host ">> Carregando: $ModuleName..." -ForegroundColor DarkGray
    try {
        return Invoke-RestMethod -Uri "$RepoURL/modules/$ModuleName?v=$(Get-Random)" -ErrorAction Stop
    } catch {
        Write-Host "X Erro ao baixar $ModuleName." -ForegroundColor Red
        return $null
    }
}

 $UtilsCode = Get-RemoteCode "utils.ps1"
if ($UtilsCode) { 
    Invoke-Expression $UtilsCode 
} else { 
    Write-Warning "Falha de conexao. Verifique a internet."
    Pause; Exit 
}

do {
    Clear-Host
    try { Show-Header } catch { Write-Host "--- DASA TOOLBOX ---" }
    
    Write-Host " [1] Manutencao de Impressoras"
    Write-Host " [2] Instalar Drivers"
    Write-Host " [3] Instalar Anydesk"
    Write-Host " [4] Limpar Spooler"
    Write-Host " [Q] Sair"
    Write-Host ""

    $main = Read-Host " Opcao"

    if ($main -eq '1') { 
        $Code = Get-RemoteCode "maintenance.ps1"
        if ($Code) { Invoke-Expression $Code }
    }
    elseif ($main -eq '2') { 
        $Code = Get-RemoteCode "drivers.ps1"
        if ($Code) { Invoke-Expression $Code }
    }
    elseif ($main -eq '3') {
        Write-Host ">> Baixando Anydesk..." -ForegroundColor Cyan
        try {
            Invoke-WebRequest "https://download.anydesk.com/AnyDesk.exe" -OutFile "$env:USERPROFILE\Desktop\AnyDesk.exe"
            Write-Host "OK: Salvo na Area de Trabalho!" -ForegroundColor Green
        } catch { Write-Host "Erro no download." }
        Start-Sleep 2
    }
    elseif ($main -eq '4') {
        Write-Host ">> Parando Spooler..." -ForegroundColor Yellow
        Stop-Service spooler -Force
        Remove-Item "$env:systemroot\System32\spool\PRINTERS\*" -Force -ErrorAction SilentlyContinue
        Start-Service spooler
        Write-Host "OK: Spooler Reiniciado!" -ForegroundColor Green
        Start-Sleep 2
    }

} while ($main -ne 'q')