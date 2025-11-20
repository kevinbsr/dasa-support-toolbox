# --- CONFIGURACOES (EPL/ZPL) ---

# EPL: N=Clear, OD=Thermal Direct, q400=Width 50mm, Q200,24=Height 25mm
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

function Menu-Manutencao {
    $Printer = Listar-Impressoras
    if (!$Printer) { Write-Host "X Nenhuma impressora encontrada."; Start-Sleep 2; return }

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

Menu-Manutencao