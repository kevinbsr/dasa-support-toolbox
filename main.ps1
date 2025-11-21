<#
.SYNOPSIS
    DASA SUPPORT TOOLBOX v1.0 (Standalone Edition)
.DESCRIPTION
    Versao de arquivo unico para contornar bloqueios de rede corporativa.
    Compatibilidade total ASCII (sem emojis).
#>

$TempDir = "C:\DasaToolbox\Temp"
$BaseURL = "https://github.com/kevinbsr/dasa-support-toolbox/raw/main/assets"

if (!(Test-Path $TempDir)) { New-Item -ItemType Directory -Force -Path $TempDir | Out-Null }

# CLASSE C# PARA COMUNICACAO DIRETA COM IMPRESSORA (RAW PRINTING)
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
    Write-Host "         DASA SUPPORT TOOLBOX - AOL 2.0                 " -ForegroundColor White
    Write-Host "         Dev: Kevin Benevides (Compass UOL)             " -ForegroundColor Gray
    Write-Host "========================================================" -ForegroundColor Blue
    Write-Host ""
}

function Baixar-Arquivo {
    param ($Nome, $ArquivoZip, $ExecutavelInterno)
    $ZipPath = "$TempDir\$ArquivoZip"
    $FolderName = $ArquivoZip.Replace(".zip", "")
    $ExtractPath = "$TempDir\$FolderName"
    
    [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

    Write-Host ">> Baixando $Nome..." -ForegroundColor Cyan
    try {
        $WebClient = New-Object System.Net.WebClient
        $WebClient.DownloadFile("$BaseURL/$ArquivoZip", $ZipPath)
        
        Write-Host ">> Extraindo..." -ForegroundColor DarkGray
        Expand-Archive -Path $ZipPath -DestinationPath $ExtractPath -Force
    } catch {
        Write-Host "X Erro no download/extracao. Verifique a internet ou proxy." -ForegroundColor Red
        return $null
    }

    $ExeRaiz = "$ExtractPath\$ExecutavelInterno"
    $ExeSub = "$ExtractPath\$FolderName\$ExecutavelInterno"

    if (Test-Path $ExeRaiz) { return $ExeRaiz }
    if (Test-Path $ExeSub) { return $ExeSub }
    
    Write-Host "X Executavel nao encontrado: $ExecutavelInterno" -ForegroundColor Red
    return $null
}

function Enviar-Comando {
    param($Printer, $Cmd, $Desc)
    Write-Host ">> Aplicando: $Desc..." -ForegroundColor Cyan
    try {
        [RawPrinterHelper]::SendStringToPrinter($Printer, $Cmd) | Out-Null
        Write-Host "[OK] Comando enviado!" -ForegroundColor Green
    } catch {
        Write-Host "X Falha de comunicacao com a impressora." -ForegroundColor Red
    }
}

function Listar-Impressoras {
    Write-Host ">> SELECIONE A IMPRESSORA:" -ForegroundColor Yellow
    $printers = Get-Printer | Where-Object { $_.Name -notmatch "PDF|XPS|OneNote|Fax" }
    if ($printers.Count -eq 0) { Write-Host "X Nenhuma impressora encontrada." -ForegroundColor Red; return $null }
    
    $i = 1
    foreach ($p in $printers) { Write-Host "   [$i] $($p.Name)"; $i++ }
    Write-Host ""
    $sel = Read-Host " Digite o numero"
    if ($sel -match '^\d+$' -and $sel -le $printers.Count -and $sel -gt 0) { return $printers[$sel - 1].Name }
    return $null
}

function Menu-Instalacao {
    do {
        Show-Header
        Write-Host ">> INSTALACAO DE DRIVERS" -ForegroundColor Yellow
        Write-Host " [1] Zebra (GC420t / ZD220)"
        Write-Host " [2] Elgin L42 Pro"
        Write-Host " [3] Honeywell PC42t"
        Write-Host " [4] Plugin AOL"
        Write-Host " [B] Voltar"
        Write-Host ""

        $opt = Read-Host " Escolha"
        if ($opt -eq 'b') { break }

        if ($opt -eq '1') {
            $Exe = Baixar-Arquivo "Zebra" "Zebra.zip" "PrnInst.exe"
            if ($Exe) { Start-Process $Exe -Wait }
        }
        elseif ($opt -eq '2') {
            $Exe = Baixar-Arquivo "Elgin" "Elgin.zip" "DriverWizard.exe"
            if ($Exe) { 
                $args = 'install /product:"Elgin L42 Pro" /quiet'
                Start-Process $Exe -ArgumentList $args -Wait 
            }
        }
        elseif ($opt -eq '3') {
            $Exe = Baixar-Arquivo "Honeywell" "Honeywell.zip" "001 - QuickInstaller.exe"
            if ($Exe) { Start-Process $Exe -ArgumentList "/VERYSILENT" -Wait }
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
            $web = New-Object System.Net.WebClient
            $web.DownloadFile("https://download.anydesk.com/AnyDesk.exe", "$env:USERPROFILE\Desktop\AnyDesk.exe")
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