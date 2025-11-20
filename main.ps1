<#
.SYNOPSIS
    DASA SUPPORT TOOLBOX - Entry Point
.DESCRIPTION
    Versão com correção de Encoding (UTF-8) para exibir emojis e acentos corretamente.
#>

[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
$OutputEncoding = [System.Text.Encoding]::UTF8

$currentPrincipal = New-Object Security.Principal.WindowsPrincipal([Security.Principal.WindowsIdentity]::GetCurrent())
if (-not $currentPrincipal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    Write-Warning "⚠️  ERRO: Execute como ADMINISTRADOR!"
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
    Write-Host "⌛ Baixando módulo: $ModuleName..." -ForegroundColor DarkGray
    try {
        $WebClient = New-Object System.Net.WebClient
        $WebClient.Encoding = [System.Text.Encoding]::UTF8
        $Url = "$RepoURL/modules/$ModuleName?v=$(Get-Random)"
        return $WebClient.DownloadString($Url)
    } catch {
        Write-Host "❌ Erro ao baixar $ModuleName. Verifique a internet." -ForegroundColor Red
        return $null
    }
}

$UtilsCode = Get-RemoteCode "utils.ps1"
if ($UtilsCode) {
    Invoke-Expression $UtilsCode
} else {
    Write-Warning "Falha crítica ao carregar o núcleo. Encerrando."
    Pause
    Exit
}

do {
    try { Show-Header } catch { Clear-Host; Write-Host "--- DASA TOOLBOX ---" }
    
    Write-Host " [1] 🔧 Manutenção de Impressoras"
    Write-Host " [2] 📂 Instalar Drivers"
    Write-Host " [3] 🖥️  Instalar Anydesk"
    Write-Host " [4] 🧹 Limpar Spooler"
    Write-Host " [Q] Sair"
    Write-Host ""

    $main = Read-Host " Opção"

    if ($main -eq '1') { 
        $Code = Get-RemoteCode "maintenance.ps1"
        if ($Code) { Invoke-Expression $Code }
    }
    elseif ($main -eq '2') { 
        $Code = Get-RemoteCode "drivers.ps1"
        if ($Code) { Invoke-Expression $Code }
    }
    elseif ($main -eq '3') {
        Write-Host "⬇️  Baixando Anydesk..." -ForegroundColor Cyan
        try {
            Invoke-WebRequest "https://download.anydesk.com/AnyDesk.exe" -OutFile "$env:USERPROFILE\Desktop\AnyDesk.exe"
            Write-Host "✅ Salvo na Área de Trabalho!" -ForegroundColor Green
        } catch {
            Write-Host "❌ Erro no download." -ForegroundColor Red
        }
        Start-Sleep 2
    }
    elseif ($main -eq '4') {
        Write-Host "🛑 Parando Spooler..." -ForegroundColor Yellow
        Stop-Service spooler -Force
        Remove-Item "$env:systemroot\System32\spool\PRINTERS\*" -Force -ErrorAction SilentlyContinue
        Start-Service spooler
        Write-Host "✅ Spooler Limpo e Reiniciado!" -ForegroundColor Green
        Start-Sleep 2
    }

} while ($main -ne 'q')