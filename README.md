# 🚀 DASA Support Toolbox (Alvaro Apoio)

Ferramenta de automação em linha de comando (CLI) para a equipe de suporte técnico.
Automatiza a instalação de drivers de impressoras térmicas, plugins web e configurações específicas (ZPL/EPL) para o sistema Alvaro Apoio.

![PowerShell](https://img.shields.io/badge/PowerShell-5.1%2B-blue?logo=powershell)
![Status](https://img.shields.io/badge/Status-Stable-green)

## ✨ Funcionalidades

* **Instalação Automática de Drivers:**
    * 🦓 Zebra (GC420t, ZD220, TLP2844)
    * 🏷️ Elgin (L42 Pro)
    * 🐝 Honeywell (PC42t)
* **Manutenção e Configuração:**
    * Calibração automática (Auto-Sense).
    * Configuração forçada de tamanho (5x2.5cm) via RAW Printing (USB).
    * Alternância de protocolos (EPL/ZPL).
    * Limpeza de Spooler de Impressão.
* **Ferramentas:**
    * Instalação silenciosa do Anydesk.
    * Instalação do Plugin de Impressão AOL.

## ⚡ Como Usar (Quick Start)

Para executar a ferramenta, abra o **PowerShell como Administrador** e cole o comando abaixo:

```powershell
$D="C:\DasaToolbox"; $F="$D\dasa-toolbox.ps1"; if(!(Test-Path $D)){New-Item -ItemType Directory -Path $D -Force}; Invoke-WebRequest -Uri 'https://raw.githubusercontent.com/kevinbsr/dasa-support-toolbox/main/dasa-toolbox.ps1' -OutFile $F -UseBasicParsing -ErrorAction Stop; Unblock-File $F -ErrorAction SilentlyContinue; & $F
```

**Nota**: Este comando baixa e executa a versão mais recente diretamente da memória, sem precisar salvar arquivos no computador do cliente.

## 🛠️ Arquitetura

O projeto é modular para facilitar a manutenção:

* `main.ps1`: Orquestrador principal. Verifica permissões e baixa os módulos.
* `modules/utils.ps1`: Contém a classe C# **RawPrinterHelper** para comunicação direta com impressoras via `winspool.drv`.
* `modules/drivers.ps1`: Lógica de download e instalação silenciosa de drivers.
* `modules/maintenance.ps1`: Comandos ZPL/EPL para configuração física das impressoras.

## 📦 Dependências

Os drivers e instaladores são baixados sob demanda da pasta `/assets` deste repositório.

* **Zebra:** Utiliza `PrnInst.exe` (Driver oficial Zebra).
* **Elgin:** Utiliza `DriverWizard.exe` (Seagull Scientific).
* **Honeywell:** Utiliza `QuickInstaller.exe`.

## 🤝 Contribuição

1. Clone o repositório.
2. Crie uma branch para sua feature:
    ```bash
    git checkout -b feature/nova-impressora
    ```
3. Faça o Commit:
    ```bash
    git commit -m 'Add: Suporte a Argox'
    ```
4. Faça o Push:
    ```bash
    git push origin feature/nova-impressora
    ```
5. Abra um Pull Request.

---
Desenvolvido por **Kevin Benevides** | Compass UOL