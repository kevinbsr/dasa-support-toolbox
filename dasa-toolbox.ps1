<#
.SYNOPSIS
    DASA SUPPORT TOOLBOX v1.0 - Stable Monolith (ASCII Safe)
.DESCRIPTION
    Author: Kevin Benevides
#>

$ErrorActionPreference = "Stop"
Set-StrictMode -Version 2.0

$currentPrincipal = New-Object Security.Principal.WindowsPrincipal([Security.Principal.WindowsIdentity]::GetCurrent())
if (-not $currentPrincipal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    Write-Warning "ERRO: Execute como ADMINISTRADOR!"
    Start-Sleep 3
    Exit
}

$BaseURL = "https://raw.githubusercontent.com/kevinbsr/dasa-support-toolbox/main/assets"
$TempDir = "$env:TEMP\DasaToolbox"

if (!(Test-Path $TempDir)) { New-Item -ItemType Directory -Force -Path $TempDir | Out-Null }

Import-Module BitsTransfer -ErrorAction SilentlyContinue

$signature = "
using System;
using System.IO;
using System.Runtime.InteropServices;

public class RawPrinterHelper {
    [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Ansi)]
    public class DOCINFOA {
        [MarshalAs(UnmanagedType.LPStr)] public string pDocName;
        [MarshalAs(UnmanagedType.LPStr)] public string pOutputFile;
        [MarshalAs(UnmanagedType.LPStr)] public string pDataType;
    }
    [DllImport(`"winspool.Drv`", EntryPoint = `"OpenPrinterA`", SetLastError = true, CharSet = CharSet.Ansi, ExactSpelling = true, CallingConvention = CallingConvention.StdCall)]
    public static extern bool OpenPrinter([MarshalAs(UnmanagedType.LPStr)] string szPrinter, out IntPtr hPrinter, IntPtr pd);

    [DllImport(`"winspool.Drv`", EntryPoint = `"ClosePrinter`", SetLastError = true, ExactSpelling = true, CallingConvention = CallingConvention.StdCall)]
    public static extern bool ClosePrinter(IntPtr hPrinter);

    [DllImport(`"winspool.Drv`", EntryPoint = `"StartDocPrinterA`", SetLastError = true, CharSet = CharSet.Ansi, ExactSpelling = true, CallingConvention = CallingConvention.StdCall)]
    public static extern bool StartDocPrinter(IntPtr hPrinter, Int32 level, [In, MarshalAs(UnmanagedType.LPStruct)] DOCINFOA di);

    [DllImport(`"winspool.Drv`", EntryPoint = `"EndDocPrinter`", SetLastError = true, ExactSpelling = true, CallingConvention = CallingConvention.StdCall)]
    public static extern bool EndDocPrinter(IntPtr hPrinter);

    [DllImport(`"winspool.Drv`", EntryPoint = `"StartPagePrinter`", SetLastError = true, ExactSpelling = true, CallingConvention = CallingConvention.StdCall)]
    public static extern bool StartPagePrinter(IntPtr hPrinter);

    [DllImport(`"winspool.Drv`", EntryPoint = `"EndPagePrinter`", SetLastError = true, ExactSpelling = true, CallingConvention = CallingConvention.StdCall)]
    public static extern bool EndPagePrinter(IntPtr hPrinter);

    [DllImport(`"winspool.Drv`", EntryPoint = `"WritePrinter`", SetLastError = true, ExactSpelling = true, CallingConvention = CallingConvention.StdCall)]
    public static extern bool WritePrinter(IntPtr hPrinter, IntPtr pBytes, Int32 dwCount, out Int32 dwWritten);

    public static bool SendStringToPrinter(string szPrinterName, string szString) {
        IntPtr pBytes;
        Int32 dwCount;
        dwCount = szString.Length;
        pBytes = Marshal.StringToCoTaskMemAnsi(szString);
        SendBytesToPrinter(szPrinterName, pBytes, dwCount);
        Marshal.FreeCoTaskMem(pBytes);
        return true;
    }
    public static bool SendBytesToPrinter(string szPrinterName, IntPtr pBytes, Int32 dwCount) {
        Int32 dwWritten = 0;
        IntPtr hPrinter = new IntPtr(0);
        DOCINFOA di = new DOCINFOA();
        bool bSuccess = false;
        di.pDocName = `"DASA_CMD`";
        di.pDataType = `"RAW`";
        if (OpenPrinter(szPrinterName.Normalize(), out hPrinter, IntPtr.Zero)) {
            if (StartDocPrinterA(hPrinter, 1, di)) {
                if (StartPagePrinter(hPrinter)) {
                    bSuccess = WritePrinter(hPrinter, pBytes, dwCount, out dwWritten);
                    EndPagePrinter(hPrinter);
                }
                EndDocPrinter(hPrinter);
            }
            ClosePrinter(hPrinter);
        }
        return bSuccess;
    }
}
"

try { Add-Type -TypeDefinition $signature } catch {}

function Show-Header {
    Clear-Host
    Write-Host "========================================================" -ForegroundColor Blue
    Write-Host "              DASA SUPPORT TOOLBOX v1.0                 " -ForegroundColor White
    Write-Host "         Dev: Kevin Benevides (Compass UOL)             " -ForegroundColor Gray
    Write-Host "========================================================" -ForegroundColor Blue
    Write-Host ""
}

function Download-Bits {
    param ($Url, $Destino, $Descricao)
    try {
        Start-BitsTransfer -Source $Url -Destination $Destino -DisplayName $Descricao -ErrorAction Stop
        return $true
    } catch {
        Write-Warning "Falha no BITS. Tentando metodo legado..."
        try {
            [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
            (New-Object System.Net.WebClient).DownloadFile($Url, $Destino)
            return $true
        } catch {
            Write-Host "[ERRO] Falha total no download: $_" -ForegroundColor Red
            return $false
        }
    }
}

function Baixar-Arquivo {
    param (
        [string]$Nome, 
        [string]$ArquivoZip, 
        [string]$ExecutavelInterno
    )
    
    $ZipPath = "$TempDir\$ArquivoZip"
    $FolderName = $ArquivoZip.Replace(".zip", "")
    $ExtractPath = "$TempDir\$FolderName"
    
    if (Test-Path $ExtractPath) {
        if ($ExecutavelInterno -ne "") {
            $ExePath = "$ExtractPath\$ExecutavelInterno"
            if (Test-Path $ExePath) { return $ExtractPath }
        } elseif ((Get-ChildItem $ExtractPath).Count -gt 0) {
            return $ExtractPath
        }
    }

    if (Test-Path $ZipPath) {
        Write-Host ">> Extraindo $Nome (Cache)..." -ForegroundColor Cyan
        try {
            Expand-Archive -Path $ZipPath -DestinationPath $ExtractPath -Force -ErrorAction Stop
            return $ExtractPath
        } catch {
            Remove-Item $ZipPath -Force
        }
    }

    Write-Host ">> Iniciando Download: $Nome..." -ForegroundColor Cyan
    $Success = Download-Bits "$BaseURL/$ArquivoZip" $ZipPath "Baixando $Nome"
    
    if (-not $Success) { return $null }

    try {
        Write-Host "   Extraindo..." -ForegroundColor Cyan
        Expand-Archive -Path $ZipPath -DestinationPath $ExtractPath -Force
    } catch {
        Write-Host "[ERRO] Falha na extracao." -ForegroundColor Red
        return $null
    }

    return $ExtractPath
}

function Instalar-AnyDesk {
    Write-Host ">> Instalacao do AnyDesk..." -ForegroundColor Yellow
    
    $InstallPath = "$env:ProgramFiles(x86)\AnyDesk"
    $TargetExe = "$InstallPath\AnyDesk.exe"

    if (Test-Path $TargetExe) {
        Write-Host "   [INFO] O AnyDesk ja esta instalado." -ForegroundColor Green
        Start-Sleep 2
        return
    }

    $Installer = "$TempDir\AnyDesk_Setup.exe"
    
    $Success = Download-Bits "https://download.anydesk.com/AnyDesk.exe" $Installer "Baixando AnyDesk"
    if (-not $Success) { return }

    Write-Host ">> Instalando (Aguarde)..." -ForegroundColor Cyan
    $Args = "--install `"$InstallPath`" --silent --create-shortcuts --create-desktop-icon"
    
    try {
        $proc = Start-Process -FilePath $Installer -ArgumentList $Args -Wait -PassThru
        Start-Sleep 5 

        if (Test-Path $TargetExe) {
            Write-Host "[OK] AnyDesk instalado com sucesso!" -ForegroundColor Green
        } else {
            Write-Host "[ERRO] Falha na instalacao." -ForegroundColor Red
        }
    } catch {
        Write-Host "[ERRO] Falha ao executar instalador: $_" -ForegroundColor Red
    }
    
    Remove-Item $Installer -Force -ErrorAction SilentlyContinue
    Pause
}

function Injetar-Drivers-Inf {
    param ([string]$PastaRaiz, [string]$NomeAmigavel)

    Write-Host ">> Iniciando Injecao Universal de Drivers ($NomeAmigavel)..." -ForegroundColor Yellow
    Write-Host "   Procurando arquivos .INF..." -ForegroundColor Cyan
    
    $InfFiles = Get-ChildItem -Path $PastaRaiz -Recurse -Filter "*.inf"

    if ($InfFiles) {
        $Count = 0
        foreach ($inf in $InfFiles) {
            $content = Get-Content $inf.FullName -Head 50
            if ($content -match "Signature" -and ($content -match "Version" -or $content -match "Class")) {
                Write-Host "   [+] Injetando: $($inf.Name)" -ForegroundColor Cyan
                $proc = Start-Process -FilePath "pnputil.exe" -ArgumentList "/add-driver `"$($inf.FullName)`" /install" -Wait -PassThru -NoNewWindow
                if ($proc.ExitCode -eq 0) { $Count++ }
            }
        }
        if ($Count -gt 0) {
            Write-Host "[OK] Drivers Carregados. Conecte a impressora USB para instalacao automatica." -ForegroundColor Green
        } else {
            Write-Warning "Nenhum driver novo foi aceito (sistema ja atualizado)."
        }
    } else {
        Write-Host "[ERRO] Nenhum arquivo .INF encontrado." -ForegroundColor Red
    }
    Pause
}

function Remover-Impressora-Individual {
    Write-Host ">> SELECIONE A IMPRESSORA PARA REMOVER:" -ForegroundColor Yellow
    $NomeImpressora = Listar-Impressoras-Selecao
    
    if (-not $NomeImpressora) { return }

    Write-Host ""
    Write-Host ">> TEM CERTEZA que deseja remover: '$NomeImpressora'?" -ForegroundColor Red
    $Confirm = Read-Host "   Digite 'S' para confirmar"
    if ($Confirm -ne "s" -and $Confirm -ne "S") { return }

    Write-Host ">> Removendo Fila de Impressao..." -ForegroundColor Cyan
    try {
        Remove-Printer -Name $NomeImpressora -ErrorAction Stop
        Write-Host "   [OK] Fila removida." -ForegroundColor Green
    } catch {
        Write-Host "   [ERRO] Falha ao remover fila: $_" -ForegroundColor Red
        Pause
        return
    }

    Write-Host ">> Tentando limpar driver associado (Best Effort)..." -ForegroundColor Cyan
    Restart-Service spooler -Force

    $Driver = Get-PrinterDriver | Where-Object { $_.Name -match $NomeImpressora } | Select-Object -First 1
    
    if ($Driver) {
        try {
            Remove-PrinterDriver -Name $Driver.Name -ErrorAction Stop
            Write-Host "   [OK] Driver removido." -ForegroundColor Green
        } catch {
            Write-Host "   [INFO] O driver nao pode ser removido (Pode estar em uso por outra impressora)." -ForegroundColor Gray
        }
    } else {
        Write-Host "   [INFO] Driver especifico nao encontrado ou nome difere." -ForegroundColor Gray
    }

    Write-Host ""
    Write-Host "[OK] Processo concluido." -ForegroundColor Green
    Pause
}

function Remover-Plugin-AOL {
    Write-Host ">> REMOVENDO PLUGIN AOL..." -ForegroundColor Red
    
    Write-Host "   [1] Encerrando processos do plugin..." -ForegroundColor Cyan
    Stop-Process -Name "AOL2.0*" -Force -ErrorAction SilentlyContinue
    Stop-Process -Name "javaw" -Force -ErrorAction SilentlyContinue
    
    $Paths = @(
        "$env:ProgramFiles(x86)\AOL2.0 Plugin de Impressão",
        "$env:ProgramFiles\AOL2.0 Plugin de Impressão",
        "$env:ProgramData\DasaPlugin"
    )

    $Found = $false
    foreach ($Path in $Paths) {
        if (Test-Path $Path) {
            $Uninstaller = "$Path\unins000.exe"
            
            if (Test-Path $Uninstaller) {
                Write-Host "   [2] Executando desinstalador oficial..." -ForegroundColor Yellow
                $proc = Start-Process -FilePath $Uninstaller -ArgumentList "/VERYSILENT /SUPPRESSMSGBOXES" -Wait -PassThru
                
                if ($proc.ExitCode -eq 0) {
                    Write-Host "       [OK] Desinstalado com sucesso." -ForegroundColor Green
                } else {
                    Write-Host "       [ERRO] Falha no desinstalador. Tentando forcar remocao..." -ForegroundColor Red
                    Remove-Item $Path -Recurse -Force -ErrorAction SilentlyContinue
                }
            } else {
                Write-Host "   [2] Desinstalador nao encontrado. Removendo arquivos na forca..." -ForegroundColor Yellow
                Remove-Item $Path -Recurse -Force -ErrorAction SilentlyContinue
            }
            $Found = $true
        }
    }

    if (-not $Found) {
        Write-Host "   [INFO] Pasta do Plugin nao encontrada (ja removido?)." -ForegroundColor Gray
    }

    Write-Host "   [3] Limpando atalhos..." -ForegroundColor Cyan
    $Shortcuts = @(
        "$env:USERPROFILE\Desktop\AOL2.0 Plugin de Impressão.lnk",
        "$env:ProgramData\Microsoft\Windows\Start Menu\Programs\StartUp\Plugin AOL DASA.lnk",
        "$env:ProgramData\Microsoft\Windows\Start Menu\Programs\StartUp\AOL2.0 Plugin de Impressão.lnk"
    )
    foreach ($lnk in $Shortcuts) {
        if (Test-Path $lnk) { Remove-Item $lnk -Force }
    }

    Write-Host ""
    Write-Host "[OK] Remocao finalizada." -ForegroundColor Green
    Pause
}

function Executar-Ferramenta {
    param ($Dir, $NomeFerramenta)
    
    if (-not $Dir) { return }
    
    $Exe = Get-ChildItem -Path $Dir -Filter "*.exe" -Recurse | Select-Object -First 1
    
    if ($Exe) {
        Write-Host ">> Abrindo $NomeFerramenta..." -ForegroundColor Cyan
        Start-Process -FilePath $Exe.FullName
        Write-Host "[OK] Ferramenta iniciada." -ForegroundColor Green
    } else {
        Write-Host "[ERRO] Nenhum executavel (.exe) encontrado no pacote." -ForegroundColor Red
    }
    Pause
}

function Enviar-Comando {
    param($Printer, $Cmd, $Desc)
    Write-Host ">> Aplicando: $Desc..." -ForegroundColor Cyan
    try {
        [RawPrinterHelper]::SendStringToPrinter($Printer, $Cmd) | Out-Null
        if (-not $?) { Write-Host "[ERRO] Falha ao enviar comando (Impressora Offline?)" -ForegroundColor Red }
        else { Write-Host "[OK] Comando enviado!" -ForegroundColor Green }
    } catch {
        Write-Host "[ERRO] Falha de comunicacao: $_" -ForegroundColor Red
    }
}

function Listar-Impressoras-Selecao {
    $printers = @(Get-Printer | Where-Object { $_.Name -notmatch "PDF|XPS|OneNote|Fax" })
    
    if ($printers.Count -eq 0) { 
        Write-Warning "Nenhuma impressora encontrada."
        return $null 
    }

    $i = 1
    foreach ($p in $printers) { Write-Host "   [$i] $($p.Name)"; $i++ }
    Write-Host ""
    $sel = Read-Host " Digite o numero"
    if ($sel -match "^\d+$" -and $sel -le $printers.Count -and $sel -gt 0) { return $printers[$sel - 1].Name }
    return $null
}

function Menu-Manutencao-Selecao {
    Write-Host ">> SELECIONE A IMPRESSORA:" -ForegroundColor Yellow
    $Nome = Listar-Impressoras-Selecao
    return $Nome
}

$EPL_Config_Str     = "`nN`nOD`nq400`nQ200,24`nS2`nD10`n"
$EPL_Reset_Str      = "`nN`n"
$EPL_Calibrate_Str  = "`nxa`n"
$EPL_Test_Str       = "
N
ZB
q400
D15
A10,0,0,1,1,1,N,`"SOCIAL_NAME_TEST`"
A10,12,0,1,1,1,N,`"DN:01/04/2025 COLETA: 06/05 09:28`"
A10,24,0,1,1,1,N,`"174714707525 - 999999999999`"
B60,45,0,2,3,6,90,N,`"999999999999`"
A10,145,0,1,1,1,N,`"TSH`"
A10,157,0,1,1,1,N,`"SORO-BASAL`"
A10,169,0,1,1,1,N,`"VOLUME MINIMO: 1.5 ml`"
A300,155,0,2,1,1,R,`"ALP.R.1`"
A380,45,1,2,1,1,R,`"`"
A10,135,3,1,1,1,N,`"ID: 0000`"
A22,142,3,1,1,1,N,`"999999999`"
P1
"

$ZPL_Config_Str     = "^XA^MTD^PW406^LL203^JUS^XZ"
$ZPL_Calibrate_Str  = "~JC"
$ZPL_Reset_Str      = "^XA^JZA^XZ"
$ZPL_Test_Str       = "^XA^FO50,50^A0N,50,50^FDTESTE DASA - ZPL^FS^XZ"

function Menu-Instalacao {
    do {
        Show-Header
        Write-Host ">> INSTALACAO DE DRIVERS (MODO UNIVERSAL)" -ForegroundColor Yellow
        Write-Host " [1] Zebra (Todas)"
        Write-Host " [2] Elgin (Todas)"
        Write-Host " [3] Honeywell (Todas)"
        Write-Host " [4] Plugin AOL (Download + Instalacao Assistida)"
        Write-Host " [B] Voltar"
        Write-Host ""

        $opt = Read-Host " Escolha"
        if ($opt -eq "b") { break }

        switch ($opt) {
            "1" { 
                $Dir = Baixar-Arquivo "Zebra" "Zebra.zip" ""
                if ($Dir) { Injetar-Drivers-Inf $Dir "Zebra" }
            }
            "2" {
                $Dir = Baixar-Arquivo "Elgin" "Elgin.zip" ""
                if ($Dir) { Injetar-Drivers-Inf $Dir "Elgin" }
            }
            "3" {
                $Dir = Baixar-Arquivo "Honeywell" "Honeywell.zip" ""
                if ($Dir) { Injetar-Drivers-Inf $Dir "Honeywell" }
            }
            "4" {
                $PluginUrl = "https://storageapoiob2bprd.blob.core.windows.net/printer-plugin/AOL2_0PlugindeImpress%C3%A3o.exe"
                $PluginLocal = "$TempDir\Plugin_AOL_Setup.exe"

                $Success = Download-Bits $PluginUrl $PluginLocal "Baixando Plugin AOL"
                
                if ($Success) {
                    Write-Host ">> Abrindo instalador..." -ForegroundColor Cyan
                    Write-Host "   Por favor, realize a instalacao na janela que se abriu." -ForegroundColor White
                    
                    # Nao usamos mais -Wait para nao travar se o usuario demorar
                    Start-Process -FilePath $PluginLocal
                    
                    Write-Host ""
                    Read-Host ">> Quando terminar a instalacao, pressione ENTER para continuar..."
                    
                    Remove-Item $PluginLocal -Force
                }
            }
        }
    } while ($true)
}

function Menu-Utilitarios {
    do {
        Show-Header
        Write-Host ">> FERRAMENTAS DO FABRICANTE (CONFIGURACAO AVANCADA)" -ForegroundColor Yellow
        Write-Host "   (Ethernet, Densidade Fina, Sensores)"
        Write-Host ""
        Write-Host " [1] Elgin Utility"
        Write-Host " [2] Honeywell Tools (PrintSet)"
        Write-Host " [B] Voltar"
        Write-Host ""

        $opt = Read-Host " Escolha"
        if ($opt -eq "b") { break }

        switch ($opt) {
            "1" {
                $Dir = Baixar-Arquivo "Elgin Utility" "ElginUtility.zip" ""
                Executar-Ferramenta $Dir "Elgin Utility"
            }
            "2" {
                $Dir = Baixar-Arquivo "Honeywell Tools" "HoneywellUtility.zip" ""
                Executar-Ferramenta $Dir "Honeywell Config Tool"
            }
        }
    } while ($true)
}

function Menu-Desinstalacao {
    do {
        Show-Header
        Write-Host ">> DESINSTALACAO" -ForegroundColor Yellow
        Write-Host "   ATENCAO: Remove dispositivos ou software."
        Write-Host ""
        Write-Host " [1] Remover Impressora (Selecao Individual)"
        Write-Host " [2] Remover Plugin AOL"
        Write-Host " [B] Voltar"
        Write-Host ""

        $opt = Read-Host " Escolha"
        if ($opt -eq "b") { break }

        switch ($opt) {
            "1" { Remover-Impressora-Individual }
            "2" { Remover-Plugin-AOL }
        }
    } while ($true)
}

function Menu-Manutencao {
    $Printer = Menu-Manutencao-Selecao
    if (!$Printer) { Start-Sleep 2; return }

    $Protocolo = "EPL"
    if ($Printer -match "Honeywell" -or $Printer -match "PC42") { $Protocolo = "ZPL" }

    do {
        Show-Header
        Write-Host ">> MANUTENCAO: $Printer" -ForegroundColor Cyan
        Write-Host "   Protocolo: $Protocolo" -ForegroundColor Gray
        Write-Host ""
        Write-Host " [1] Configurar DASA (5x2.5cm + Termico)" -ForegroundColor Green
        Write-Host " [2] Calibrar (Auto-Sense)" -ForegroundColor Yellow
        Write-Host " [3] Resetar (Factory Default)" -ForegroundColor Red
        Write-Host " [4] Teste de Impressao (Modelo DASA)" -ForegroundColor Cyan
        Write-Host " [5] Alternar Protocolo (EPL/ZPL)"
        Write-Host " [B] Voltar"
        Write-Host ""

        $opt = Read-Host " Escolha"
        if ($opt -eq "b") { break }

        if ($opt -eq "5") {
            if ($Protocolo -eq "EPL") { $Protocolo = "ZPL" } else { $Protocolo = "EPL" }
            continue
        }

        if ($Protocolo -eq "EPL") {
            switch ($opt) {
                "1" { Enviar-Comando $Printer $EPL_Config_Str "Config DASA (EPL)" }
                "2" { Enviar-Comando $Printer $EPL_Calibrate_Str "Calibracao (EPL)" }
                "3" { Enviar-Comando $Printer $EPL_Reset_Str "Reset Buffer (EPL)" }
                "4" { Enviar-Comando $Printer $EPL_Test_Str "Teste DASA (EPL)" }
            }
        } else {
            switch ($opt) {
                "1" { Enviar-Comando $Printer $ZPL_Config_Str "Config DASA (ZPL)" }
                "2" { Enviar-Comando $Printer $ZPL_Calibrate_Str "Calibracao (ZPL)" }
                "3" { Enviar-Comando $Printer $ZPL_Reset_Str "Factory Reset (ZPL)" }
                "4" { Enviar-Comando $Printer $ZPL_Test_Str "Teste DASA (ZPL)" }
            }
        }
        Pause
    } while ($true)
}

do {
    Show-Header
    Write-Host " [1] Manutencao de Impressoras (Reset/Calibrar)"
    Write-Host " [2] Instalar Drivers (Universal)"
    Write-Host " [3] Instalar Anydesk"
    Write-Host " [4] Limpar Spooler"
    Write-Host " [5] Desinstalar (Impressoras/Plugins)"
    Write-Host " [6] Ferramentas do Fabricante (Elgin/Honeywell)"
    Write-Host " [Q] Sair"
    Write-Host ""

    $main = Read-Host " Opcao"

    switch ($main) {
        "1" { Menu-Manutencao }
        "2" { Menu-Instalacao }
        "3" { Instalar-AnyDesk }
        "4" {
            Write-Host ">> Parando Spooler..." -ForegroundColor Yellow
            Stop-Service spooler -Force -ErrorAction SilentlyContinue
            Remove-Item "$env:systemroot\System32\spool\PRINTERS\*" -Force -ErrorAction SilentlyContinue
            Start-Service spooler
            Write-Host "[OK] Spooler Limpo!" -ForegroundColor Green
            Start-Sleep 2
        }
        "5" { Menu-Desinstalacao }
        "6" { Menu-Utilitarios }
    }

} while ($main -ne "q")
