# EMIVE Smart Alarme — Gerenciador de Cartões SD

Aplicativo desktop para Windows desenvolvido em **Visual Basic .NET (Windows Forms / .NET 8)** para preparação, formatação, diagnóstico, recuperação e clonagem bit-a-bit de cartões SD utilizados em **câmeras EZVIZ** (Filial Brasília).

---

## Funcionalidades

- **Detecção automática de unidades removíveis e cartões SD**
  - Identifica unidades removíveis montadas (`DriveInfo`) e dispositivos USB/removíveis via WMI (`Win32_DiskDrive` e `Win32_LogicalDisk`).
  - Reconhece inclusive cartões sem partição ou em formato RAW, listando-os como `Disco X [NÃO FORMATADO]`.
  - Exibe formato atual do sistema de arquivos, capacidade total, espaço utilizado e espaço livre.

- **Formatação com múltiplos sistemas de arquivos**
  - **FAT32** (Padrão)
  - **exFAT** (Recomendado para cartões de 64 GB ou maiores)
  - **NTFS**
  - **EZVIZ (Formato Otimizado para Câmeras)**: formata em `exFAT` e gera a estrutura esperada pela câmera.
  - Suporte a **Formatação Rápida** (com timeout de segurança de 3 minutos) e definição personalizada do **rótulo do volume** (padrão: `EMIVE_CAM`).

- **Preparação de estrutura para câmeras EZVIZ**
  - Criação do arquivo `config.ini` na raiz com metadados da formatação (`[EZVIZ]`, `Version=2.0`, `ClusterSize=64KB`).
  - Pré-alocação de 236 arquivos de vídeo na raiz (`hiv00000.mp4` a `hiv00235.mp4`, de 256 MB cada) e arquivos de índice (`index00.bin`, `index01.bin`, `logMainFile.bin`, `logCurFile.bin`).

- **Formatação Física (Low-Level) em 5 fases**
  - Voltada para recuperação de cartões com erro de leitura, sem tabela de partição ou setores defeituosos:
    1. Limpeza completa de setores via `diskpart` (`clean all`).
    2. Recriação de partição primária ativa e atribuição de letra de unidade via `diskpart`.
    3. Formatação completa (sem modo rápido) para verificação de setores.
    4. Teste de integridade de gravação e leitura (`test_integrity.tmp` de 1 MB com bytes aleatórios).
    5. Otimização de acesso (`fsutil behavior set disablelastaccess 1`) e criação dos diretórios `EZVIZ\Record`, `EZVIZ\Picture` e `EZVIZ\Log`.

- **Verificação e Recuperação de erros**
  - **Verificar Erros (`[V]`)**: executa análise do sistema de arquivos via `chkdsk <unidade>:`.
  - **Recuperar (`[R]`)**: executa reparo automático de erros e setores lógicos via `chkdsk <unidade>: /F`.

- **Clonagem e Restauração bit-a-bit**
  - **Clonar Cartão (`[C]`)**: lê diretamente o disco físico (`\\.\PhysicalDriveN`) e gera uma imagem fiel (`.img`) em `C:\EMIVE_Clones\CartaoEZVIZ_Clone.img`.
  - **Restaurar Clone (`[X]`)**: grava a imagem de clone bit-a-bit de volta para o cartão SD de destino.
  - Rotina de verificação por **hash SHA-1** no código para validar a integridade entre o cartão original e o arquivo de clone.

- **Interface com log em tempo real e cancelamento**
  - Barra de progresso contínua e console de logs coloridos com horário de cada etapa.
  - Botão **Cancelar** para interromper operações em andamento de forma controlada.

---

## Requisitos

- **Sistema Operacional:** Windows 10 / Windows 11 (utiliza Windows Forms, WMI e utilitários nativos como `format`, `chkdsk`, `diskpart`, `wmic` e `fsutil`).
- **Runtime / SDK:** [.NET 8.0 SDK / Desktop Runtime (`net8.0-windows`)](https://dotnet.microsoft.com/) ou Visual Studio 2022 (v17.14+).
- **Privilégios:** **Execução obrigatória como Administrador** (verificado na inicialização pelo `Module1.vb` para permitir operações de disco e acesso físico ao dispositivo).

---

## Estrutura do Projeto

```text
Emive-Format/
├── Emive Format.sln                  # Solução do Visual Studio
└── Emive Format/
    ├── Emive Format.vbproj           # Projeto .NET 8 Windows Forms (VB.NET)
    ├── App.config                    # Configurações e metadados da aplicação
    ├── Module1.vb                    # Ponto de entrada e verificação de Administrador
    ├── ApplicationEvents.vb          # Eventos do ciclo de vida da aplicação
    ├── FormPrincipal.vb              # Interface principal e lógica de disco/EZVIZ/clone
    ├── FormPrincipal.Designer.vb     # Definição parcial do formulário principal
    ├── FormSobre.vb                  # Janela "Sobre" com informações do projeto e autor
    ├── FormSobre.Designer.vb         # Designer da janela "Sobre"
    ├── FormErro.vb                   # Diálogo customizado para exibição de erros
    └── My Project/                   # Configurações padrão do projeto VB.NET
```

---

## Como Executar ou Compilar

1. Clone o repositório:
   ```bash
   git clone https://github.com/mamedio1/Emive-Format.git
   ```
2. Abra o arquivo `Emive Format.sln` no **Visual Studio 2022** (com carga de trabalho *Desenvolvimento para desktop com .NET*) **executando o Visual Studio como Administrador** — ou compile pela linha de comando:
   ```bash
   dotnet build "Emive Format.sln" -c Release
   ```
3. Execute o aplicativo gerado com privilégios de **Administrador**.

---

## Aviso Importante

> **Atenção:** Operações de formatação, formatação física (`diskpart clean all`) e restauração de clone **apagam permanentemente todos os dados** da unidade selecionada. Certifique-se de selecionar corretamente o cartão SD antes de iniciar qualquer ação e não remova o cartão durante o processo.

---

## Autor

**Robson Mamedio Araujo**  
Supervisor Técnico de Brasília — EMIVE Smart Alarme (2026)
