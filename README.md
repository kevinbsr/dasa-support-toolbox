# 🚀 DASA Support Toolbox (Alvaro Apoio)

Ferramenta de automação em linha de comando (CLI) para a equipe de suporte técnico.
Automatiza a instalação de drivers de impressoras térmicas, plugins web e configurações específicas (ZPL/EPL) para o sistema Alvaro Apoio.

![PowerShell](https://img.shields.io/badge/PowerShell-5.1%2B-blue?logo=powershell)
![Status](https://img.shields.io/badge/Status-Stable-green)
![Architecture](https://img.shields.io/badge/Architecture-Monolithic_Stand_Alone-orange)

## ✨ Funcionalidades
* **Drivers Universais:** Instala Zebra, Elgin e Honeywell automaticamente (via injeção de .INF).

* **Manutenção:** Calibração, Reset e Teste de Impressão (Layout DASA Real).

* **Plugin AOL:** Download direto do servidor DASA e instalação assistida.

* **Ferramentas**: Instalação do AnyDesk e Utilitários de Fabricante.

* **Limpeza:** Desinstalação cirúrgica de impressoras e limpeza de Spooler.

## ⚡ Como Usar (Quick Start)

1. Acesse a máquina do cliente via **AnyDesk**.
2. Abra o **PowerShell como Administrador**.
3. Copie e cole o comando abaixo:

```powershell
Set-ExecutionPolicy Bypass -Scope Process -Force; [System.Net.ServicePointManager]::SecurityProtocol = [System.Net.ServicePointManager]::SecurityProtocol -bor 3072; iex ((New-Object System.Net.WebClient).DownloadString('https://raw.githubusercontent.com/kevinbsr/dasa-support-toolbox/main/dasa-toolbox.ps1'))
```
**Nota**: Este comando baixa e executa a versão mais recente diretamente da memória, sem precisar salvar arquivos no computador do cliente. Se houver um erro de `404 Not Found`, o arquivo pode estar com outro nome. Verifique o repositório no navegador.

## 🛠️ Arquitetura

-  `dasa-toolbox.ps1`: **Arquivo Principal** que contém todo o código C# (para comunicação RAW com impressoras), funções de utilidade, menus e lógica de drivers.

- O design foi consolidado em um único arquivo para garantir a **estabilidade** e **confiabilidade** no ambiente corporativo, principalmente para:

1.  **Bypass de Firewall:** Evita que a segurança da rede bloqueie o script por tentar fazer múltiplos downloads de código em tempo de execução.
2.  **Garantia de Execução:** Depois do download inicial, a ferramenta é totalmente funcional mesmo sem conexão, pois toda a inteligência está embutida.
3.  **Core Técnico:** A classe C# `RawPrinterHelper` para comandos ZPL/EPL está embutida, tornando-o um artefato único e poderoso para o suporte.

## 📦 Dependências

* **Download:** Os drivers e instaladores são baixados por HTTPS da pasta `/assets` deste repositório.
* **Drivers:** Utiliza instaladores oficiais (**ZDesigner/Seagull**) com argumentos de instalação silenciosa (`/S`, `/VERYSILENT`).

## 🤝 Contribuição

1. Clone o repositório.
2. Crie uma branch para sua feature:
    ```bash
    git checkout -b feature/nova-impressora
    ```
3. Faça o Commit:
    ```bash
    git commit -m 'feat: adicionar suporte a [Nova Impressora]'
    ```
4. Faça o Push e abra um Pull Request.

---
Desenvolvido por **Kevin Benevides** | Compass **UOL**
