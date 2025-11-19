<#
.SYNOPSIS
    DASA SUPPORT TOOLBOX - Entry Point
#>

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

function Load-Module {
    param($ModuleName)
    Write-Host "⌛ Carregando módulo: $ModuleName..." -ForegroundColor DarkGray
    try {
        $Code = Invoke-RestMethod -Uri "$RepoURL/modules/$ModuleName"
        Invoke-Expression $Code
    } catch {
        Write-Host "❌ Erro ao carregar módulo $ModuleName. Verifique a internet." -ForegroundColor Red
        Pause
    }
}

Load-Module "utils.ps1"

do {
    Show-Header
    Write-Host " [1] 🔧 Manutenção de Impressoras"
    Write-Host " [2] 📂 Instalar Drivers"
    Write-Host " [3] 🖥️  Instalar Anydesk"
    Write-Host " [4] 🧹 Limpar Spooler"
    Write-Host " [Q] Sair"
    Write-Host ""

    $main = Read-Host " Opção"

    if ($main -eq '1') { 
        Load-Module "maintenance.ps1" 
    }
    elseif ($main -eq '2') { 
       Load-Module "drivers.ps1" 
    }
    elseif ($main -eq '3') {
        Write-Host "Baixando Anydesk..."
        Invoke-WebRequest "https://download.anydesk.com/AnyDesk.exe" -OutFile "$env:USERPROFILE\Desktop\AnyDesk.exe"
        Write-Host "✅ Salvo na Área de Trabalho!" -ForegroundColor Green
        Start-Sleep 2
    }
    elseif ($main -eq '4') {
        Write-Host "Reiniciando Spooler..."
        Stop-Service spooler -Force
        Remove-Item "$env:systemroot\System32\spool\PRINTERS\*" -Force -ErrorAction SilentlyContinue
        Start-Service spooler
        Write-Host "✅ Pronto!" -ForegroundColor Green
        Start-Sleep 2
    }

} while ($main -ne 'q')