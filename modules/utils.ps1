# CLASSE C# (RAW PRINTER)
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
    Write-Host "         🚀 DASA SUPPORT TOOLBOX v11.0 (MODULAR)        " -ForegroundColor White
    Write-Host "         Dev: Kevin Benevides (Compass UOL)             " -ForegroundColor Gray
    Write-Host "========================================================" -ForegroundColor Blue
    Write-Host ""
}

function Baixar-Arquivo {
    param ($Nome, $ArquivoZip, $ExecutavelInterno)
    $ZipPath = "$TempDir\$ArquivoZip"
    $ExtractPath = "$TempDir\$($ArquivoZip.Replace('.zip',''))"
    
    Write-Host "⬇️  Baixando $Nome..." -ForegroundColor Cyan
    try {
        Invoke-WebRequest -Uri "$AssetsURL/$ArquivoZip" -OutFile $ZipPath -ErrorAction Stop
        Expand-Archive -Path $ZipPath -DestinationPath $ExtractPath -Force
    } catch {
        Write-Host "❌ Erro no download/extração. Verifique conexão." -ForegroundColor Red
        return $null
    }

    $ExePath = "$ExtractPath\$ExecutavelInterno"
    if (!(Test-Path $ExePath)) { $ExePath = "$ExtractPath\$($ArquivoZip.Replace('.zip',''))\$ExecutavelInterno" }

    if (Test-Path $ExePath) { return $ExePath }
    Write-Host "❌ Executável não encontrado: $ExecutavelInterno" -ForegroundColor Red
    return $null
}

function Enviar-Comando {
    param($Printer, $Cmd, $Desc)
    Write-Host "⚙️  Aplicando: $Desc..." -ForegroundColor Cyan
    try {
        [RawPrinterHelper]::SendStringToPrinter($Printer, $Cmd) | Out-Null
        Write-Host "✅ Comando enviado!" -ForegroundColor Green
    } catch {
        Write-Host "❌ Falha de comunicação." -ForegroundColor Red
    }
}

function Listar-Impressoras {
    Write-Host "🖨️  SELECIONE A IMPRESSORA:" -ForegroundColor Yellow
    $printers = Get-Printer | Where-Object { $_.Name -notmatch "PDF|XPS|OneNote|Fax" }
    if ($printers.Count -eq 0) { return $null }
    
    $i = 1
    foreach ($p in $printers) { Write-Host "   [$i] $($p.Name)"; $i++ }
    Write-Host ""
    $sel = Read-Host " Digite o número"
    if ($sel -match '^\d+$' -and $sel -le $printers.Count -and $sel -gt 0) { return $printers[$sel - 1].Name }
    return $null
}