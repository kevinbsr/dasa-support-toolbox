<#
.SYNOPSIS
    DASA SUPPORT TOOLBOX v10.1 (Zebra Submenu and Installation Fix)
.DESCRIPTION
    Versao que implementa submenu de selecao de modelo/protocolo para impressoras Zebra.
#>

$currentPrincipal = New-Object Security.Principal.WindowsPrincipal([Security.Principal.WindowsIdentity]::GetCurrent())
if (-not $currentPrincipal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    Write-Warning "ERRO: Execute como ADMINISTRADOR!"
    Start-Sleep 3
    Exit
}

$BaseURL = "https://raw.githubusercontent.com/kevinbsr/dasa-support-toolbox/main/assets"
$TempDir = "$env:TEMP\DasaToolbox"
if (Test-Path $TempDir) { Remove-Item $TempDir -Recurse -Force -ErrorAction SilentlyContinue }
New-Item -ItemType Directory -Force -Path $TempDir | Out-Null

$signature = @'
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
    [DllImport("winspool.Drv", EntryPoint = "OpenPrinterA", SetLastError = true, CharSet = CharSet.Ansi, ExactSpelling = true, CallingConvention = CallingConvention.StdCall)]
    public static extern bool OpenPrinter([MarshalAs(UnmanagedType.LPStr)] string szPrinter, out IntPtr hPrinter, IntPtr pd);

    [DllImport("winspool.Drv", EntryPoint = "ClosePrinter", SetLastError = true, ExactSpelling = true, CallingConvention = CallingConvention.StdCall)]
    public static extern bool ClosePrinter(IntPtr hPrinter);

    [DllImport("winspool.Drv", EntryPoint = "StartDocPrinterA", SetLastError = true, CharSet = CharSet.Ansi, ExactSpelling = true, CallingConvention = CallingConvention.StdCall)]
    public static extern bool StartDocPrinter(IntPtr hPrinter, Int32 level, [In, MarshalAs(UnmanagedType.LPStruct)] DOCINFOA di);

    [DllImport("winspool.Drv", EntryPoint = "EndDocPrinter", SetLastError = true, ExactSpelling = true, CallingConvention = CallingConvention.StdCall)]
    public static extern bool EndDocPrinter(IntPtr hPrinter);

    [DllImport("winspool.Drv", EntryPoint = "StartPagePrinter", SetLastError = true, ExactSpelling = true, CallingConvention = CallingConvention.StdCall)]
    public static extern bool StartPagePrinter(IntPtr hPrinter);

    [DllImport("winspool.Drv", EntryPoint = "EndPagePrinter", SetLastError = true, ExactSpelling = true, CallingConvention = CallingConvention.StdCall)]
    public static extern bool EndPagePrinter(IntPtr hPrinter);

    [DllImport("winspool.Drv", EntryPoint = "WritePrinter", SetLastError = true, ExactSpelling = true, CallingConvention = CallingConvention.StdCall)]
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
        di.pDocName = "DASA_CMD";
        di.pDataType = "RAW";
        if (OpenPrinter(szPrinterName.Normalize(), out hPrinter, IntPtr.Zero)) {
            if (StartDocPrinter(hPrinter, 1, di)) {
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
'@

try { Add-Type -TypeDefinition $signature } catch {}

function Show-Header {
    Clear-Host
    Write-Host "========================================================" -ForegroundColor Blue
    Write-Host "         DASA SUPPORT TOOLBOX v1.0                      " -ForegroundColor White
    Write-Host "         Dev: Kevin Benevides (Compass UOL)             " -ForegroundColor Gray
    Write-Host "========================================================" -ForegroundColor Blue
    Write-Host ""
}

function Baixar-Arquivo {
    param ($Nome, $ArquivoZip, $ExecutavelInterno)
    $ZipPath = "$TempDir\$ArquivoZip"
    $FolderName = $ArquivoZip.Replace(".zip", "")
    $ExtractPath = "$TempDir\$FolderName"
    
    Write-Host ">> Baixando $Nome..." -ForegroundColor Cyan
    try {
        Invoke-WebRequest -Uri "$BaseURL/$ArquivoZip" -OutFile $ZipPath -ErrorAction Stop
        Expand-Archive -Path $ZipPath -DestinationPath $ExtractPath -Force
    } catch {
        Write-Host "X Erro no download/extracao. Verifique a conexao ou nome do ZIP no GitHub." -ForegroundColor Red
        return $null
    }

    $ExePath = "$ExtractPath\$ExecutavelInterno"
    if (!(Test-Path $ExePath)) { $ExePath = "$ExtractPath\$FolderName\$ExecutavelInterno" }

    if (Test-Path $ExePath) { return $ExePath }
    Write-Host "X Executavel nao encontrado: $ExecutavelInterno" -ForegroundColor Red
    return $null
}

function Enviar-Comando {
    param($Printer, $Cmd, $Desc)
    Write-Host ">> Aplicando: $Desc..." -ForegroundColor Cyan
    try {
        [RawPrinterHelper]::SendStringToPrinter($Printer, $Cmd) | Out-Null
        if (-not $?) { Write-Host "X Falha ao enviar comando (Impressora Offline?)" -ForegroundColor Red }
        else { Write-Host "[OK] Comando enviado!" -ForegroundColor Green }
    } catch {
        Write-Host "X Erro de comunicacao: $_" -ForegroundColor Red
    }
}

function Listar-Impressoras {
    Write-Host ">> SELECIONE A IMPRESSORA:" -ForegroundColor Yellow
    $printers = Get-Printer | Where-Object { $_.Name -notmatch "PDF|XPS|OneNote|Fax" }
    if ($printers.Count -eq 0) { return $null }
    
    $i = 1
    foreach ($p in $printers) { Write-Host "   [$i] $($p.Name)"; $i++ }
    Write-Host ""
    $sel = Read-Host " Digite o numero"
    if ($sel -match '^\d+$' -and $sel -le $printers.Count -and $sel -gt 0) { return $printers[$sel - 1].Name }
    return $null
}

$EPL_Config = "`nN`nOD`nq400`nQ200,24`nS2`nD10`n"
$EPL_Test = "`nN`nA50,50,0,4,1,1,N,""TESTE DASA - EPL""`nP1`n"
$ZPL_Config = "^XA^MTD^PW406^LL203^JUS^XZ"
$ZPL_Test = "^XA^FO50,50^A0N,50,50^FDTESTE DASA - ZPL^FS^XZ"

function Menu-Zebra-Instalacao {
    $ZebraExe = Baixar-Arquivo "Zebra" "Zebra.zip" "PrnInst.exe"
    if (!$ZebraExe) { return }

    do {
        Show-Header
        Write-Host ">> SELECAO DE MODELO ZEBRA (PRE-INSTALACAO)" -ForegroundColor Yellow
        Write-Host " [1] GC420t (EPL) - Padrao DASA"
        Write-Host " [2] GC420t (ZPL) - Emulacao"
        Write-Host " [3] ZD220 (EPL)"
        Write-Host " [4] ZD220 (ZPL)"
        Write-Host " [5] ZD230 (EPL)"
        Write-Host " [6] TLP2844 (EPL)"
        Write-Host " [B] Voltar"
        Write-Host ""

        $opt = Read-Host " Escolha"
        if ($opt -eq 'b') { break }

        $Modelos = @{
            '1' = "ZDesigner GC420t (EPL)"
            '2' = "ZDesigner GC420t (ZPL)"
            '3' = "ZDesigner ZD220-203dpi EPL"
            '4' = "ZDesigner ZD220-203dpi ZPL"
            '5' = "ZDesigner ZD230-203dpi EPL"
            '6' = "ZDesigner TLP 2844"
        }
        
        $ModeloSelecionado = $Modelos[$opt]
        
        if ($ModeloSelecionado) {
            $Args = "/PREINSTALL `"$ModeloSelecionado`""
            Write-Host ">> Instalando drivers para $ModeloSelecionado..." -ForegroundColor Cyan
            
            # Executa o PrnInst.exe
            Start-Process -FilePath $ZebraExe -ArgumentList $Args -Wait -NoNewWindow
            
            Write-Host "✅ Drivers pre-instalados. Conecte a impressora na USB." -ForegroundColor Green
        }
        else {
             Write-Host "X Opcao invalida." -ForegroundColor Red
        }

        Pause
    } while ($true)
}

function Menu-Instalacao {
    do {
        Show-Header
        Write-Host "📂 INSTALACAO DE DRIVERS" -ForegroundColor Yellow
        Write-Host " [1] Zebra (GC420t / ZD220 / TLP2844)"
        Write-Host " [2] Elgin L42 Pro"
        Write-Host " [3] Honeywell PC42t"
        Write-Host " [4] Plugin AOL"
        Write-Host " [B] Voltar"
        Write-Host ""

        $opt = Read-Host " Escolha"
        if ($opt -eq 'b') { break }

        if ($opt -eq '1') {
            Menu-Zebra-Instalacao 
        }
        elseif ($opt -eq '2') {
            $Exe = Baixar-Arquivo "Elgin" "Elgin.zip" "DriverWizard.exe"
            if ($Exe) { 
                $argsElgin = 'install /product:"Elgin L42 Pro" /quiet'
                Start-Process $Exe -ArgumentList $argsElgin -Wait 
                Write-Host "✅ Instalacao concluida!" -ForegroundColor Green
            }
        }
        elseif ($opt -eq '3') {
            $Exe = Baixar-Arquivo "Honeywell" "Honeywell.zip" "001 - QuickInstaller.exe"
            if ($Exe) { 
                Start-Process $Exe -ArgumentList "/VERYSILENT" -Wait
                Write-Host "✅ Instalacao concluida!" -ForegroundColor Green
            }
        }
        elseif ($opt -eq '4') {
            $Exe = Baixar-Arquivo "Plugin" "Plugin.zip" "Plugin_AOL.exe"
            if ($Exe) { Start-Process $Exe -ArgumentList "/S" -Wait }
        }
        if ($opt -ne 'b') { Pause }
    } while ($true)
}

function Menu-Manutencao {
    $Printer = Listar-Impressoras
    if (!$Printer) { Start-Sleep 2; return }

    $Protocolo = "EPL"
    if ($Printer -match "Honeywell" -or $Printer -match "PC42") { $Protocolo = "ZPL" }

    $EPL_Config = "
    N
    OD
    q400
    Q200,24
    S2
    D10
    "
    $EPL_Test = "
    N
    A50,50,0,4,1,1,N,""TESTE DASA - EPL""
    P1
    "
    $ZPL_Config = "^XA^MTD^PW406^LL203^JUS^XZ"
    $ZPL_Test = "^XA^FO50,50^A0N,50,50^FDTESTE DASA - ZPL^FS^XZ"

    do {
        Show-Header
        Write-Host ">> MANUTENCAO: $Printer" -ForegroundColor Cyan
        Write-Host "   Protocolo: $Protocolo" -ForegroundColor Gray
        Write-Host ""
        Write-Host " [1] Configurar DASA (5x2.5cm + Termico)" -ForegroundColor Green
        Write-Host " [2] Calibrar (Auto-Sense)" -ForegroundColor Yellow
        Write-Host " [3] Resetar (Factory Default)" -ForegroundColor Red
        Write-Host " [4] Teste de Impressao"
        Write-Host " [5] Alternar Protocolo (EPL/ZPL)"
        Write-Host " [B] Voltar"
        Write-Host ""

        $opt = Read-Host " Escolha"
        if ($opt -eq 'b') { break }
        
        if ($opt -eq '5') { 
            if ($Protocolo -eq "EPL") { $Protocolo = "ZPL" } else { $Protocolo = "EPL" }
            continue 
        }

        if ($Protocolo -eq "EPL") {
            if ($opt -eq '1') { Enviar-Comando $Printer $EPL_Config "Config DASA (EPL)" }
            if ($opt -eq '2') { Enviar-Comando $Printer "`nxa`n" "Calibracao (EPL)" }
            if ($opt -eq '3') { Enviar-Comando $Printer "`nN`n" "Reset Buffer" }
            if ($opt -eq '4') { Enviar-Comando $Printer $EPL_Test "Teste" }
        } else {
            if ($opt -eq '1') { Enviar-Comando $Printer $ZPL_Config "Config DASA (ZPL)" }
            if ($opt -eq '2') { Enviar-Comando $Printer "~JC" "Calibracao (ZPL)" }
            if ($opt -eq '3') { Enviar-Comando $Printer "^XA^JZA^XZ" "Factory Reset" }
            if ($opt -eq '4') { Enviar-Comando $Printer $ZPL_Test "Teste" }
        }
        Pause
    } while ($true)
}

do {
    Show-Header
    Write-Host " [1] Manutencao de Impressoras"
    Write-Host " [2] Instalar Drivers"
    Write-Host " [3] Instalar Anydesk"
    Write-Host " [4] Limpar Spooler"
    Write-Host " [Q] Sair"
    Write-Host ""

    $main = Read-Host " Opcao"

    if ($main -eq '1') { Menu-Manutencao }
    elseif ($main -eq '2') { Menu-Instalacao }
    elseif ($main -eq '3') {
        Write-Host ">> Baixando Anydesk..." -ForegroundColor Cyan
        try {
            [Net.ServicePointManager]::SecurityProtocol=[Net.SecurityProtocolType]::Tls12
            $web = New-Object System.Net.WebClient
            $web.DownloadFile("https://download.anydesk.com/AnyDesk.exe", "$env:USERPROFILE\Desktop\AnyDesk.exe")
            Write-Host "[OK] Salvo na Area de Trabalho!" -ForegroundColor Green
        } catch { Write-Host "X Erro no download." -ForegroundColor Red }
        Start-Sleep 2
    }
    elseif ($main -eq '4') {
        Write-Host ">> Reiniciando Spooler..." -ForegroundColor Yellow
        Stop-Service spooler -Force -ErrorAction SilentlyContinue
        Remove-Item "$env:systemroot\System32\spool\PRINTERS\*" -Force -ErrorAction SilentlyContinue
        Start-Service spooler
        Write-Host "[OK] Spooler Limpo!" -ForegroundColor Green
        Start-Sleep 2
    }
} while ($main -ne 'q')