function Menu-Instalacao {
    do {
        Show-Header
        Write-Host "📂 INSTALAÇÃO DE DRIVERS" -ForegroundColor Yellow
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
                $argsElgin = 'install /product:"Elgin L42 Pro" /quiet'
                Start-Process $Exe -ArgumentList $argsElgin -Wait 
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

Menu-Instalacao