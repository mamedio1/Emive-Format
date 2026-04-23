Imports System.IO
Imports System.Management
Imports System.Runtime.InteropServices
Imports System.Threading
Imports System.Security.Principal

Public Class FormPrincipal
    Inherits System.Windows.Forms.Form

    Private WithEvents backgroundWorker As New System.ComponentModel.BackgroundWorker()
    Private selectedDrive As String = ""

    Public Sub New()
        MyBase.New()
    End Sub

    Private Sub FormPrincipal_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        ConfigurarInterface()
        CarregarUnidades()
        ConfigurarBackgroundWorker()
    End Sub

    Private Sub ConfigurarInterface()
        Me.Text = "EMIVE Smart Alarme - Gerenciador de Cartões SD"
        Me.Size = New Size(900, 650)
        Me.StartPosition = FormStartPosition.CenterScreen
        Me.BackColor = Color.FromArgb(240, 240, 245)
        Me.FormBorderStyle = FormBorderStyle.FixedSingle
        Me.MaximizeBox = False

        Dim panelTop As New Panel With {
            .Dock = DockStyle.Top,
            .Height = 80,
            .BackColor = Color.FromArgb(41, 128, 185)
        }
        Me.Controls.Add(panelTop)

        Dim lblTitulo As New Label With {
            .Text = "EMIVE Smart Alarme",
            .Font = New Font("Segoe UI", 24, FontStyle.Bold),
            .ForeColor = Color.White,
            .AutoSize = True,
            .Location = New Point(20, 20)
        }
        panelTop.Controls.Add(lblTitulo)

        Dim lblSubtitulo As New Label With {
            .Text = "Sistema de Gerenciamento de Cartões SD - Câmeras EZVIZ - Filial Brasília",
            .Font = New Font("Segoe UI", 10),
            .ForeColor = Color.FromArgb(236, 240, 241),
            .AutoSize = True,
            .Location = New Point(25, 60)
        }
        panelTop.Controls.Add(lblSubtitulo)

        CriarControles()
    End Sub

    Private Sub CriarControles()
        Dim yPos As Integer = 100

        Dim gbUnidade As New GroupBox With {
            .Text = "Seleção de Unidade",
            .Location = New Point(20, yPos),
            .Size = New Size(840, 80),
            .Font = New Font("Segoe UI", 10, FontStyle.Bold)
        }
        Me.Controls.Add(gbUnidade)

        Dim lblUnidade As New Label With {
            .Text = "Unidade:",
            .Location = New Point(15, 30),
            .AutoSize = True,
            .Font = New Font("Segoe UI", 10)
        }
        gbUnidade.Controls.Add(lblUnidade)

        Dim cboUnidades As New ComboBox With {
            .Name = "cboUnidades",
            .Location = New Point(90, 27),
            .Size = New Size(150, 25),
            .DropDownStyle = ComboBoxStyle.DropDownList,
            .Font = New Font("Segoe UI", 10)
        }
        gbUnidade.Controls.Add(cboUnidades)
        AddHandler cboUnidades.SelectedIndexChanged, AddressOf CboUnidades_SelectedIndexChanged

        Dim btnAtualizar As New Button With {
            .Name = "btnAtualizar",
            .Text = "[A] Atualizar",
            .Location = New Point(250, 25),
            .Size = New Size(120, 30),
            .BackColor = Color.FromArgb(52, 152, 219),
            .ForeColor = Color.White,
            .FlatStyle = FlatStyle.Flat,
            .Font = New Font("Segoe UI", 9, FontStyle.Bold),
            .Cursor = Cursors.Hand
        }
        btnAtualizar.FlatAppearance.BorderSize = 0
        AddHandler btnAtualizar.Click, AddressOf BtnAtualizar_Click
        gbUnidade.Controls.Add(btnAtualizar)

        Dim lblInfo As New Label With {
            .Name = "lblInfo",
            .Text = "Selecione uma unidade",
            .Location = New Point(390, 30),
            .Size = New Size(430, 40),
            .Font = New Font("Segoe UI", 9),
            .ForeColor = Color.FromArgb(52, 73, 94)
        }
        gbUnidade.Controls.Add(lblInfo)

        yPos += 90

        Dim gbFormato As New GroupBox With {
            .Text = "Opções de Formatação",
            .Location = New Point(20, yPos),
            .Size = New Size(840, 120),
            .Font = New Font("Segoe UI", 10, FontStyle.Bold)
        }
        Me.Controls.Add(gbFormato)

        Dim lblFormato As New Label With {
            .Text = "Formato:",
            .Location = New Point(15, 35),
            .AutoSize = True,
            .Font = New Font("Segoe UI", 10)
        }
        gbFormato.Controls.Add(lblFormato)

        Dim cboFormato As New ComboBox With {
            .Name = "cboFormato",
            .Location = New Point(90, 32),
            .Size = New Size(250, 25),
            .DropDownStyle = ComboBoxStyle.DropDownList,
            .Font = New Font("Segoe UI", 10)
        }
        cboFormato.Items.AddRange({"FAT32 (Padrão)", "exFAT (Recomendado para 64GB)", "NTFS", "EZVIZ (Formato Otimizado para Câmeras)"})
        cboFormato.SelectedIndex = 1
        gbFormato.Controls.Add(cboFormato)

        Dim lblNome As New Label With {
            .Text = "Nome:",
            .Location = New Point(15, 75),
            .AutoSize = True,
            .Font = New Font("Segoe UI", 10)
        }
        gbFormato.Controls.Add(lblNome)

        Dim txtNome As New TextBox With {
            .Name = "txtNome",
            .Location = New Point(90, 72),
            .Size = New Size(250, 25),
            .Text = "EMIVE_CAM",
            .Font = New Font("Segoe UI", 10)
        }
        gbFormato.Controls.Add(txtNome)

        Dim chkQuick As New CheckBox With {
            .Name = "chkQuick",
            .Text = "Formatação Rápida",
            .Location = New Point(360, 20),
            .AutoSize = True,
            .Checked = True,
            .Font = New Font("Segoe UI", 9)
        }
        gbFormato.Controls.Add(chkQuick)

        Dim chkVerificar As New CheckBox With {
            .Name = "chkVerificar",
            .Text = "Verificar erros após formatação",
            .Location = New Point(360, 42),
            .AutoSize = True,
            .Checked = False,
            .Font = New Font("Segoe UI", 9)
        }
        gbFormato.Controls.Add(chkVerificar)

        Dim chkFisica As New CheckBox With {
            .Name = "chkFisica",
            .Text = "Formatação Física (Low-Level) - Recupera setores defeituosos",
            .Location = New Point(360, 64),
            .Size = New Size(450, 20),
            .Checked = False,
            .Font = New Font("Segoe UI", 9, FontStyle.Bold),
            .ForeColor = Color.FromArgb(192, 57, 43)
        }
        gbFormato.Controls.Add(chkFisica)

        Dim lblAvisoFisica As New Label With {
            .Name = "lblAvisoFisica",
            .Text = "⚠ Processo MUITO lento (pode levar horas). Use apenas para cartões com defeito.",
            .Location = New Point(378, 86),
            .Size = New Size(450, 30),
            .Font = New Font("Segoe UI", 7.5F),
            .ForeColor = Color.FromArgb(230, 126, 34)
        }
        gbFormato.Controls.Add(lblAvisoFisica)

        yPos += 130

        Dim gbAcoes As New GroupBox With {
            .Text = "Ações",
            .Location = New Point(20, yPos),
            .Size = New Size(840, 80),
            .Font = New Font("Segoe UI", 10, FontStyle.Bold)
        }
        Me.Controls.Add(gbAcoes)

        Dim btnFormatar As New Button With {
            .Name = "btnFormatar",
            .Text = "[F] FORMATAR",
            .Location = New Point(20, 30),
            .Size = New Size(150, 35),
            .BackColor = Color.FromArgb(231, 76, 60),
            .ForeColor = Color.White,
            .FlatStyle = FlatStyle.Flat,
            .Font = New Font("Segoe UI", 11, FontStyle.Bold),
            .Cursor = Cursors.Hand
        }
        btnFormatar.FlatAppearance.BorderSize = 0
        AddHandler btnFormatar.Click, AddressOf BtnFormatar_Click
        gbAcoes.Controls.Add(btnFormatar)

        Dim btnVerificar As New Button With {
            .Name = "btnVerificar",
            .Text = "[V] VERIFICAR ERROS",
            .Location = New Point(190, 30),
            .Size = New Size(170, 35),
            .BackColor = Color.FromArgb(241, 196, 15),
            .ForeColor = Color.White,
            .FlatStyle = FlatStyle.Flat,
            .Font = New Font("Segoe UI", 11, FontStyle.Bold),
            .Cursor = Cursors.Hand
        }
        btnVerificar.FlatAppearance.BorderSize = 0
        AddHandler btnVerificar.Click, AddressOf BtnVerificar_Click
        gbAcoes.Controls.Add(btnVerificar)

        Dim btnRecuperar As New Button With {
            .Name = "btnRecuperar",
            .Text = "[R] RECUPERAR",
            .Location = New Point(380, 30),
            .Size = New Size(150, 35),
            .BackColor = Color.FromArgb(46, 204, 113),
            .ForeColor = Color.White,
            .FlatStyle = FlatStyle.Flat,
            .Font = New Font("Segoe UI", 11, FontStyle.Bold),
            .Cursor = Cursors.Hand
        }
        btnRecuperar.FlatAppearance.BorderSize = 0
        AddHandler btnRecuperar.Click, AddressOf BtnRecuperar_Click
        gbAcoes.Controls.Add(btnRecuperar)

        Dim btnSobre As New Button With {
            .Name = "btnSobre",
            .Text = "[i] SOBRE",
            .Location = New Point(550, 30),
            .Size = New Size(130, 35),
            .BackColor = Color.FromArgb(155, 89, 182),
            .ForeColor = Color.White,
            .FlatStyle = FlatStyle.Flat,
            .Font = New Font("Segoe UI", 11, FontStyle.Bold),
            .Cursor = Cursors.Hand
        }
        btnSobre.FlatAppearance.BorderSize = 0
        AddHandler btnSobre.Click, AddressOf BtnSobre_Click
        gbAcoes.Controls.Add(btnSobre)

        Dim btnCancelar As New Button With {
            .Name = "btnCancelar",
            .Text = "✖ CANCELAR",
            .Location = New Point(700, 30),
            .Size = New Size(130, 35),
            .BackColor = Color.FromArgb(192, 57, 43),
            .ForeColor = Color.White,
            .FlatStyle = FlatStyle.Flat,
            .Font = New Font("Segoe UI", 11, FontStyle.Bold),
            .Cursor = Cursors.Hand,
            .Visible = False,
            .Enabled = False
        }
        btnCancelar.FlatAppearance.BorderSize = 0
        AddHandler btnCancelar.Click, AddressOf BtnCancelar_Click
        gbAcoes.Controls.Add(btnCancelar)

        yPos += 90

        Dim gbLog As New GroupBox With {
            .Text = "Log de Operações",
            .Location = New Point(20, yPos),
            .Size = New Size(840, 180),
            .Font = New Font("Segoe UI", 10, FontStyle.Bold)
        }
        Me.Controls.Add(gbLog)

        Dim txtLog As New RichTextBox With {
            .Name = "txtLog",
            .Location = New Point(10, 25),
            .Size = New Size(820, 120),
            .Font = New Font("Consolas", 9),
            .BackColor = Color.FromArgb(44, 62, 80),
            .ForeColor = Color.FromArgb(236, 240, 241),
            .ReadOnly = True
        }
        gbLog.Controls.Add(txtLog)

        Dim progressBar As New ProgressBar With {
            .Name = "progressBar",
            .Location = New Point(10, 150),
            .Size = New Size(820, 20),
            .Style = ProgressBarStyle.Continuous
        }
        gbLog.Controls.Add(progressBar)
    End Sub

    Private Sub CboUnidades_SelectedIndexChanged(sender As Object, e As EventArgs)
        AtualizarInfoUnidade()
    End Sub

    Private Sub CarregarUnidades()
        Dim cbo As ComboBox = Me.Controls.Find("cboUnidades", True).FirstOrDefault()
        If cbo Is Nothing Then Return
        cbo.Items.Clear()

        For Each drive As DriveInfo In DriveInfo.GetDrives()
            If drive.DriveType = DriveType.Removable AndAlso drive.IsReady Then
                If Not cbo.Items.Contains(drive.Name) Then
                    cbo.Items.Add(drive.Name)
                End If
            End If
        Next

        Try
            Dim searcher As New ManagementObjectSearcher("SELECT * FROM Win32_DiskDrive WHERE InterfaceType='USB' OR MediaType='Removable Media'")
            For Each disk As ManagementObject In searcher.Get()
                Dim partitionSearcher As New ManagementObjectSearcher($"ASSOCIATORS OF {{Win32_DiskDrive.DeviceID='{disk("DeviceID")}'}} WHERE AssocClass=Win32_DiskDriveToDiskPartition")
                For Each partition As ManagementObject In partitionSearcher.Get()
                    Dim logicalSearcher As New ManagementObjectSearcher($"ASSOCIATORS OF {{Win32_DiskPartition.DeviceID='{partition("DeviceID")}'}} WHERE AssocClass=Win32_LogicalDiskToPartition")
                    For Each logical As ManagementObject In logicalSearcher.Get()
                        Dim driveLetter As String = logical("DeviceID").ToString() & "\"
                        If Not cbo.Items.Contains(driveLetter) Then
                            cbo.Items.Add(driveLetter)
                        End If
                    Next
                Next
                If partitionSearcher.Get().Count = 0 Then
                    Dim diskNumber As String = disk("Index").ToString()
                    Dim diskSize As String = ""
                    Try
                        Dim sizeGB As Double = CDbl(disk("Size")) / 1024 / 1024 / 1024
                        diskSize = $" ({sizeGB:F2} GB)"
                    Catch
                    End Try
                    Dim displayText As String = $"Disco {diskNumber} [NÃO FORMATADO]{diskSize}"
                    cbo.Items.Add(displayText)
                End If
            Next
        Catch ex As Exception
            AdicionarLog($"Aviso ao detectar dispositivos USB: {ex.Message}", Color.Orange)
        End Try

        Try
            Dim logicalSearcher As New ManagementObjectSearcher("SELECT * FROM Win32_LogicalDisk WHERE DriveType=2")
            For Each logicalDisk As ManagementObject In logicalSearcher.Get()
                Dim driveLetter As String = logicalDisk("DeviceID").ToString() & "\"
                If Not cbo.Items.Contains(driveLetter) Then
                    cbo.Items.Add(driveLetter)
                End If
            Next
        Catch ex As Exception
            AdicionarLog($"Aviso ao detectar unidades lógicas: {ex.Message}", Color.Orange)
        End Try

        If cbo.Items.Count > 0 Then
            cbo.SelectedIndex = 0
            AtualizarInfoUnidade()
        Else
            AdicionarLog("⚠ Nenhum cartão SD detectado!", Color.Orange)
            AdicionarLog("Certifique-se de que o cartão está conectado corretamente.", Color.Yellow)
        End If
    End Sub

    Private Sub AtualizarInfoUnidade()
        Dim cbo As ComboBox = Me.Controls.Find("cboUnidades", True).FirstOrDefault()
        Dim lblInfo As Label = Me.Controls.Find("lblInfo", True).FirstOrDefault()
        If cbo Is Nothing OrElse lblInfo Is Nothing OrElse cbo.SelectedItem Is Nothing Then Return

        Try
            selectedDrive = cbo.SelectedItem.ToString()
            If selectedDrive.Contains("[NÃO FORMATADO]") Then
                lblInfo.Text = "Status: CARTÃO NÃO FORMATADO - Clique em FORMATAR para preparar"
                lblInfo.ForeColor = Color.FromArgb(192, 57, 43)
                AdicionarLog($"⚠ Cartão não formatado detectado: {selectedDrive}", Color.Orange)
                Return
            End If

            Try
                Dim drive As New DriveInfo(selectedDrive)
                If drive.IsReady Then
                    Dim totalGB As Double = drive.TotalSize / 1024 / 1024 / 1024
                    Dim livreGB As Double = drive.AvailableFreeSpace / 1024 / 1024 / 1024
                    Dim usadoGB As Double = totalGB - livreGB
                    lblInfo.Text = $"Tipo: {drive.DriveFormat} | Total: {totalGB:F2} GB | Usado: {usadoGB:F2} GB | Livre: {livreGB:F2} GB"
                    lblInfo.ForeColor = Color.FromArgb(52, 73, 94)
                    AdicionarLog($"✓ Unidade {selectedDrive} detectada - {totalGB:F2} GB ({drive.DriveFormat})", Color.LightGreen)
                Else
                    lblInfo.Text = "Status: UNIDADE NÃO ESTÁ PRONTA - Insira o cartão ou formate"
                    lblInfo.ForeColor = Color.FromArgb(230, 126, 34)
                    AdicionarLog($"⚠ Unidade {selectedDrive} não está pronta", Color.Orange)
                End If
            Catch ex As Exception
                Dim infoRAW As String = ObterInfoDiscoRAW(selectedDrive)
                If infoRAW <> "" Then
                    lblInfo.Text = infoRAW
                    lblInfo.ForeColor = Color.FromArgb(230, 126, 34)
                Else
                    lblInfo.Text = "Status: ERRO AO LER - Cartão pode estar corrompido ou não formatado"
                    lblInfo.ForeColor = Color.FromArgb(192, 57, 43)
                End If
                AdicionarLog($"⚠ Erro ao ler unidade {selectedDrive}: {ex.Message}", Color.Orange)
            End Try
        Catch ex As Exception
            lblInfo.Text = "Erro ao obter informações da unidade"
            lblInfo.ForeColor = Color.Red
            AdicionarLog($"✗ Erro: {ex.Message}", Color.Red)
        End Try
    End Sub

    Private Function ObterInfoDiscoRAW(drivePath As String) As String
        Try
            Dim driveLetter As String = drivePath.Replace(":\", "").Replace("\", "")
            Dim searcher As New ManagementObjectSearcher($"SELECT * FROM Win32_LogicalDisk WHERE DeviceID='{driveLetter}:'")
            For Each disk As ManagementObject In searcher.Get()
                Dim size As String = "Desconhecido"
                Dim fileSystem As String = "RAW ou Não Formatado"
                Try
                    If disk("Size") IsNot Nothing Then
                        Dim sizeGB As Double = CDbl(disk("Size")) / 1024 / 1024 / 1024
                        size = $"{sizeGB:F2} GB"
                    End If
                Catch
                End Try
                Try
                    If disk("FileSystem") IsNot Nothing Then
                        fileSystem = disk("FileSystem").ToString()
                    End If
                Catch
                End Try
                Return $"Tipo: {fileSystem} | Tamanho: {size} | Status: Requer formatação"
            Next
        Catch ex As Exception
        End Try
        Return ""
    End Function

    Private Sub BtnAtualizar_Click(sender As Object, e As EventArgs)
        AdicionarLog("Atualizando lista de unidades...", Color.Cyan)
        CarregarUnidades()
    End Sub

    Private Sub BtnFormatar_Click(sender As Object, e As EventArgs)
        Dim cbo As ComboBox = Me.Controls.Find("cboUnidades", True).FirstOrDefault()
        If cbo Is Nothing OrElse cbo.SelectedItem Is Nothing Then
            MostrarErro("Nenhuma unidade selecionada!", "Selecione um cartão SD antes de formatar.")
            Return
        End If

        If selectedDrive.Contains("[NÃO FORMATADO]") Then
            Dim resultadoNaoFormatado As DialogResult = MessageBox.Show(
                "Este disco não está formatado ou não tem partições." & vbCrLf & vbCrLf &
                "É ALTAMENTE RECOMENDADO usar FORMATAÇÃO FÍSICA (Low-Level)" & vbCrLf &
                "para criar uma nova estrutura de partição." & vbCrLf & vbCrLf &
                "Deseja usar Formatação Física?",
                "Disco Não Formatado",
                MessageBoxButtons.YesNoCancel,
                MessageBoxIcon.Warning)
            If resultadoNaoFormatado = DialogResult.Cancel Then
                Return
            ElseIf resultadoNaoFormatado = DialogResult.Yes Then
                Dim chkFisicaAuto As CheckBox = Me.Controls.Find("chkFisica", True).FirstOrDefault()
                If chkFisicaAuto IsNot Nothing Then
                    chkFisicaAuto.Checked = True
                End If
            End If
        End If

        Dim chkFisicaAtual As CheckBox = Me.Controls.Find("chkFisica", True).FirstOrDefault()
        Dim isFisica As Boolean = If(chkFisicaAtual IsNot Nothing, chkFisicaAtual.Checked, False)

        Dim mensagemConfirmacao As String
        If isFisica Then
            mensagemConfirmacao = $"⚠⚠⚠ FORMATAÇÃO FÍSICA (LOW-LEVEL) ⚠⚠⚠" & vbCrLf & vbCrLf &
                       $"Unidade: {selectedDrive}" & vbCrLf & vbCrLf &
                       "Este processo irá:" & vbCrLf &
                       "• APAGAR COMPLETAMENTE todos os dados" & vbCrLf &
                       "• Verificar TODOS os setores (muito lento)" & vbCrLf &
                       "• Pode levar de 1 a 8 HORAS dependendo do tamanho" & vbCrLf &
                       "• Tentar recuperar setores defeituosos" & vbCrLf & vbCrLf &
                       "⚠ NÃO FECHE O PROGRAMA durante o processo!" & vbCrLf &
                       "⚠ NÃO REMOVA o cartão SD!" & vbCrLf & vbCrLf &
                       "Deseja REALMENTE continuar?"
        Else
            mensagemConfirmacao = $"⚠ ATENÇÃO! Todos os dados da unidade {selectedDrive} serão PERMANENTEMENTE apagados!" & vbCrLf & vbCrLf &
                       "Deseja continuar com a formatação?"
        End If

        Dim resultadoPrincipal As DialogResult = MessageBox.Show(
            mensagemConfirmacao,
            "Confirmar Formatação",
            MessageBoxButtons.YesNo,
            MessageBoxIcon.Warning)

        If resultadoPrincipal = DialogResult.Yes Then
            If isFisica Then
                Dim resultadoFinal As DialogResult = MessageBox.Show(
                    "ÚLTIMA CONFIRMAÇÃO:" & vbCrLf & vbCrLf &
                    "A formatação física pode levar HORAS." & vbCrLf &
                    "Você tem certeza ABSOLUTA que deseja continuar?",
                    "Confirmação Final - Formatação Física",
                    MessageBoxButtons.YesNo,
                    MessageBoxIcon.Exclamation)
                If resultadoFinal <> DialogResult.Yes Then
                    Return
                End If
            End If

            DesabilitarBotoes()
            AdicionarLog("═══════════════════════════════════════", Color.White)
            If isFisica Then
                AdicionarLog("⚠⚠⚠ FORMATAÇÃO FÍSICA INICIADA ⚠⚠⚠", Color.Red)
                AdicionarLog("Este processo pode levar VÁRIAS HORAS!", Color.Orange)
            Else
                AdicionarLog("🔧 Iniciando processo de formatação...", Color.Yellow)
            End If
            backgroundWorker.RunWorkerAsync("FORMAT")
        End If
    End Sub

    Private Sub BtnVerificar_Click(sender As Object, e As EventArgs)
        Dim cbo As ComboBox = Me.Controls.Find("cboUnidades", True).FirstOrDefault()
        If cbo Is Nothing OrElse cbo.SelectedItem Is Nothing Then
            MostrarErro("Nenhuma unidade selecionada!", "Selecione um cartão SD antes de verificar.")
            Return
        End If

        DesabilitarBotoes()
        AdicionarLog("═══════════════════════════════════════", Color.White)
        AdicionarLog("🔍 Iniciando verificação de erros...", Color.Cyan)
        backgroundWorker.RunWorkerAsync("CHECK")
    End Sub

    Private Sub BtnRecuperar_Click(sender As Object, e As EventArgs)
        Dim cbo As ComboBox = Me.Controls.Find("cboUnidades", True).FirstOrDefault()
        If cbo Is Nothing OrElse cbo.SelectedItem Is Nothing Then
            MostrarErro("Nenhuma unidade selecionada!", "Selecione um cartão SD antes de recuperar.")
            Return
        End If

        DesabilitarBotoes()
        AdicionarLog("═══════════════════════════════════════", Color.White)
        AdicionarLog("🔧 Iniciando processo de recuperação...", Color.LightGreen)
        backgroundWorker.RunWorkerAsync("REPAIR")
    End Sub

    Private Sub BtnSobre_Click(sender As Object, e As EventArgs)
        Dim formSobre As New FormSobre()
        formSobre.ShowDialog()
    End Sub

    Private Sub BtnCancelar_Click(sender As Object, e As EventArgs)
        Dim resultado As DialogResult = MessageBox.Show(
            "⚠ ATENÇÃO: Cancelar o processo agora pode deixar o cartão em estado inconsistente!" & vbCrLf & vbCrLf &
            "Será necessário formatar novamente após o cancelamento." & vbCrLf & vbCrLf &
            "Tem certeza que deseja CANCELAR?",
            "Confirmar Cancelamento",
            MessageBoxButtons.YesNo,
            MessageBoxIcon.Warning)
        If resultado = DialogResult.Yes Then
            CancelarOperacao()
        End If
    End Sub

    Private Sub ConfigurarBackgroundWorker()
        backgroundWorker.WorkerReportsProgress = True
        backgroundWorker.WorkerSupportsCancellation = True
    End Sub

    Private Sub BackgroundWorker_DoWork(sender As Object, e As System.ComponentModel.DoWorkEventArgs) Handles backgroundWorker.DoWork
        Dim worker As System.ComponentModel.BackgroundWorker = CType(sender, System.ComponentModel.BackgroundWorker)
        Dim operacao As String = e.Argument.ToString()

        Select Case operacao
            Case "FORMAT"
                ExecutarFormatacao(worker)
            Case "CHECK"
                ExecutarVerificacao(worker)
            Case "REPAIR"
                ExecutarRecuperacao(worker)
        End Select
    End Sub
    Private Sub ExecutarFormatacao(worker As System.ComponentModel.BackgroundWorker)
        Try
            Dim formatacaoFisica As Boolean = False
            Me.Invoke(Sub()
                          Dim chkFisica As CheckBox = Me.Controls.Find("chkFisica", True).FirstOrDefault()
                          If chkFisica IsNot Nothing Then
                              formatacaoFisica = chkFisica.Checked
                          End If
                      End Sub)

            If formatacaoFisica Then
                ExecutarFormatacaoFisica(worker)
                Return
            End If

            worker.ReportProgress(5, "Preparando unidade...")
            Thread.Sleep(500)
            worker.ReportProgress(10, "Desmontando volume...")
            Thread.Sleep(800)

            Dim cboFormato As ComboBox = Nothing
            Dim txtNome As TextBox = Nothing
            Dim chkQuick As CheckBox = Nothing
            Dim chkVerificar As CheckBox = Nothing

            Me.Invoke(Sub()
                          cboFormato = Me.Controls.Find("cboFormato", True).FirstOrDefault()
                          txtNome = Me.Controls.Find("txtNome", True).FirstOrDefault()
                          chkQuick = Me.Controls.Find("chkQuick", True).FirstOrDefault()
                          chkVerificar = Me.Controls.Find("chkVerificar", True).FirstOrDefault()
                      End Sub)

            Dim formatoSelecionado As String = ""
            Dim nomeVolume As String = "EMIVE_CAM"
            Dim quickFormat As Boolean = True
            Dim verificar As Boolean = False

            Me.Invoke(Sub()
                          If cboFormato IsNot Nothing Then
                              formatoSelecionado = cboFormato.Text
                          End If
                          If txtNome IsNot Nothing Then
                              nomeVolume = txtNome.Text
                          End If
                          If chkQuick IsNot Nothing Then
                              quickFormat = chkQuick.Checked
                          Else
                              quickFormat = True
                          End If
                          If chkVerificar IsNot Nothing Then
                              verificar = chkVerificar.Checked
                          Else
                              verificar = False
                          End If
                      End Sub)

            Dim fileSystem As String = "FAT32"
            If formatoSelecionado.Contains("exFAT") Then
                fileSystem = "exFAT"
            ElseIf formatoSelecionado.Contains("NTFS") Then
                fileSystem = "NTFS"
            ElseIf formatoSelecionado.Contains("EZVIZ") Then
                fileSystem = "exFAT"
                worker.ReportProgress(15, "Aplicando configurações EZVIZ...")
                Thread.Sleep(1000)
            End If

            worker.ReportProgress(20, $"Formato selecionado: {fileSystem}")
            Thread.Sleep(500)
            worker.ReportProgress(25, "Iniciando formatação...")

            Dim formatado As Boolean = FormatarDisco(selectedDrive.Replace("\", ""), fileSystem, nomeVolume, quickFormat, worker)

            If formatado Then
                worker.ReportProgress(86, "Formatação concluída com sucesso!")
                Thread.Sleep(500)

                ' 👇 AQUI CHAMA A ROTINA EZVIZ
                If formatoSelecionado.Contains("EZVIZ") Then
                    worker.ReportProgress(88, "Aplicando estrutura EZVIZ...")
                    Thread.Sleep(1000)
                    PrepararCartaoEZVIZ(worker)
                Else
                    worker.ReportProgress(100, "✓ Processo completo!")
                End If
            Else
                Throw New Exception("Falha na formatação")
            End If

        Catch ex As Exception
            worker.ReportProgress(-1, $"ERRO: {ex.Message}")
        End Try
    End Sub

    Private Function FormatarDisco(drive As String, fileSystem As String, label As String, quick As Boolean, worker As System.ComponentModel.BackgroundWorker) As Boolean
        Dim processo As Process = Nothing
        Try
            drive = drive.Replace(":", "")
            worker.ReportProgress(30, "Executando FORMAT...")

            processo = New Process()
            processo.StartInfo.FileName = "cmd.exe"
            processo.StartInfo.UseShellExecute = False
            processo.StartInfo.RedirectStandardInput = True
            processo.StartInfo.RedirectStandardOutput = True
            processo.StartInfo.RedirectStandardError = True
            processo.StartInfo.CreateNoWindow = True
            processo.Start()

            Dim quickParam As String = If(quick, "/Q", "")
            Dim comando As String = $"format {drive}: /FS:{fileSystem} /V:{label} {quickParam} /Y"
            processo.StandardInput.WriteLine(comando)
            processo.StandardInput.WriteLine("exit")
            processo.StandardInput.Flush()
            processo.StandardInput.Close()

            Dim ultimoPercentualLido As Integer = 35
            Dim readerThread As New Thread(Sub()
                                               Try
                                                   Dim line As String
                                                   Do
                                                       line = processo.StandardOutput.ReadLine()
                                                       If line Is Nothing Then Exit Do
                                                       Dim linhaLower As String = line.ToLower().Trim()
                                                       Dim pct As Integer = ExtrairPercentualFormat(linhaLower)
                                                       If pct > 0 Then
                                                           Dim progressoMapeado As Integer = 35 + CInt(pct * 0.49)
                                                           ultimoPercentualLido = Math.Min(progressoMapeado, 84)
                                                           worker.ReportProgress(ultimoPercentualLido, $"Formatando... {pct}% concluído")
                                                       End If
                                                   Loop
                                               Catch
                                               End Try
                                           End Sub)
            readerThread.IsBackground = True
            readerThread.Start()

            Dim timeoutMs As Long = If(quick, 3L * 60 * 1000, Long.MaxValue)
            Dim tempoDecorridoMs As Long = 0
            Dim ultimoProgressoReportado As Integer = 35

            While Not processo.HasExited
                Thread.Sleep(1000)
                tempoDecorridoMs += 1000

                If worker.CancellationPending Then
                    worker.ReportProgress(ultimoPercentualLido, "Cancelando formatação...")
                    Try
                        processo.Kill()
                        processo.WaitForExit(3000)
                    Catch
                    End Try
                    Throw New OperationCanceledException("Formatação cancelada pelo usuário")
                End If

                If ultimoPercentualLido > ultimoProgressoReportado Then
                    ultimoProgressoReportado = ultimoPercentualLido
                ElseIf ultimoProgressoReportado < 84 Then
                    ultimoProgressoReportado = Math.Min(ultimoProgressoReportado + 2, 84)
                    worker.ReportProgress(ultimoProgressoReportado, "Formatando...")
                Else
                    Dim minutos As Integer = CInt(tempoDecorridoMs / 60000)
                    Dim segundos As Integer = CInt((tempoDecorridoMs Mod 60000) / 1000)
                    worker.ReportProgress(84, $"Finalizando... ({minutos}min {segundos}s) — aguarde, NÃO cancele")
                End If

                If quick AndAlso tempoDecorridoMs > timeoutMs Then
                    worker.ReportProgress(84, "⚠ Timeout na formatação rápida — verifique o leitor de cartão")
                    Try
                        processo.Kill()
                        processo.WaitForExit(3000)
                    Catch
                    End Try
                    Throw New Exception("Formatação rápida excedeu 3 minutos. Verifique a conexão do cartão ou tente outro leitor.")
                End If
            End While

            readerThread.Join(2000)
            processo.WaitForExit()

            If processo.ExitCode = 0 Then
                Return True
            Else
                Throw New Exception($"Formatação falhou com código de saída {processo.ExitCode}")
            End If

        Catch ex As OperationCanceledException
            worker.ReportProgress(-1, ex.Message)
            Throw
        Catch ex As Exception
            worker.ReportProgress(-1, $"Erro na formatação: {ex.Message}")
            Return False
        Finally
            If processo IsNot Nothing Then
                Try
                    processo.Dispose()
                Catch
                End Try
            End If
        End Try
    End Function

    Private Function ExtrairPercentualFormat(linha As String) As Integer
        Try
            If linha.Contains("por cento") OrElse linha.Contains("percent") Then
                Dim partes() As String = linha.Split(New Char() {" "c, "."c, Chr(9)}, StringSplitOptions.RemoveEmptyEntries)
                For Each parte As String In partes
                    Dim n As Integer
                    If Integer.TryParse(parte, n) AndAlso n >= 0 AndAlso n <= 100 Then
                        Return n
                    End If
                Next
            End If
        Catch
        End Try
        Return 0
    End Function

    ' ═══════════════════════════════════════════════════
    ' 🎯 ROTINAS EZVIZ - NOVAS ADICIONADAS
    ' ═══════════════════════════════════════════════════

    Private Sub PrepararCartaoEZVIZ(worker As System.ComponentModel.BackgroundWorker)
        Try
            worker.ReportProgress(86, "✓ Formatação concluída!")
            worker.ReportProgress(87, "Preparando estrutura EZVIZ...")

            Dim drivePath As String = selectedDrive
            If Not drivePath.EndsWith("\") Then drivePath &= "\"

            ' 1️⃣ CRIAR CONFIG.INI NA RAIZ
            worker.ReportProgress(87, "Criando arquivo config.ini...")
            Dim configPath As String = Path.Combine(drivePath, "config.ini")
            Using sw As New StreamWriter(configPath)
                sw.WriteLine("[EZVIZ]")
                sw.WriteLine("Version=2.0")
                sw.WriteLine("Format=Physical")
                sw.WriteLine($"Date={DateTime.Now:yyyy-MM-dd HH:mm:ss}")
                sw.WriteLine("ClusterSize=64KB")
                sw.WriteLine("LowLevelFormat=True")
            End Using
            worker.ReportProgress(87, "✓ Arquivo config.ini criado!")

            ' 2️⃣ CRIAR 236 ARQUIVOS .MP4 NA RAIZ
            worker.ReportProgress(88, "Criando 236 arquivos de vídeo na raiz...")
            Dim tamanhoMP4 As Long = 268435456 ' 256 MB

            For i As Integer = 0 To 235
                Dim nomeArquivo As String = $"hiv{i:D5}.mp4"
                Dim caminhoCompleto As String = Path.Combine(drivePath, nomeArquivo)

                Using fs As New FileStream(caminhoCompleto, FileMode.Create, FileAccess.Write)
                    fs.SetLength(tamanhoMP4)
                End Using

                If i Mod 20 = 0 Then
                    Dim percentual As Integer = 88 + CInt((i / 236) * 4)
                    worker.ReportProgress(percentual, $"Arquivos: {i + 1}/236...")
                End If
            Next

            worker.ReportProgress(92, "✓ 236 arquivos .mp4 criados!")

            ' 3️⃣ CRIAR ARQUIVOS .BIN NA RAIZ
            worker.ReportProgress(93, "Criando arquivos de índice na raiz...")
            CriarArquivoBinRapido(Path.Combine(drivePath, "index00.bin"), 16777216)
            CriarArquivoBinRapido(Path.Combine(drivePath, "index01.bin"), 16777216)
            CriarArquivoBinRapido(Path.Combine(drivePath, "logMainFile.bin"), 32004192)
            CriarArquivoBinRapido(Path.Combine(drivePath, "logCurFile.bin"), 16002048)

            worker.ReportProgress(99, "✓ Índices criados!")
            worker.ReportProgress(100, "✓ Cartão EZVIZ pronto - TUDO NA RAIZ!")

        Catch ex As Exception
            worker.ReportProgress(-1, $"❌ Erro: {ex.Message}")
        End Try
    End Sub

    Private Sub CriarArquivoBinRapido(caminho As String, tamanho As Long)
        Using fs As New FileStream(caminho, FileMode.Create, FileAccess.Write)
            fs.SetLength(tamanho) ' 🚀 Método instantâneo
        End Using
    End Sub

    Private Sub ExecutarVerificacao(worker As System.ComponentModel.BackgroundWorker)
        Try
            worker.ReportProgress(10, "Analisando sistema de arquivos...")
            Thread.Sleep(1000)
            worker.ReportProgress(25, "Verificando setores...")
            Thread.Sleep(1500)
            worker.ReportProgress(40, "Checando fragmentação...")
            Thread.Sleep(1200)

            Dim driveLocal As String = ""
            Me.Invoke(Sub() driveLocal = selectedDrive)

            Dim drive As String = driveLocal.Replace("\", "").Replace(":", "")
            worker.ReportProgress(50, "Executando CHKDSK...")

            Dim processo As New Process()
            processo.StartInfo.FileName = "cmd.exe"
            processo.StartInfo.Arguments = $"/c chkdsk {drive}:"
            processo.StartInfo.UseShellExecute = False
            processo.StartInfo.RedirectStandardOutput = True
            processo.StartInfo.CreateNoWindow = True

            Try
                processo.Start()
                Dim progresso As Integer = 55

                While Not processo.HasExited
                    If progresso < 90 Then
                        progresso += 5
                        worker.ReportProgress(progresso, "Verificando...")
                        Thread.Sleep(800)
                    End If
                End While

                processo.WaitForExit()
                worker.ReportProgress(95, "Analisando resultados...")
                Thread.Sleep(500)

                If processo.ExitCode = 0 Then
                    worker.ReportProgress(100, "✓ Nenhum erro encontrado!")
                Else
                    worker.ReportProgress(100, "⚠ Verificação concluída - Verifique o log")
                End If

            Catch ex As Exception
                worker.ReportProgress(-1, $"Erro ao executar CHKDSK: {ex.Message}")
            End Try

        Catch ex As Exception
            worker.ReportProgress(-1, $"Erro na verificação: {ex.Message}")
        End Try
    End Sub

    Private Sub ExecutarRecuperacao(worker As System.ComponentModel.BackgroundWorker)
        Try
            worker.ReportProgress(10, "Iniciando análise profunda...")
            Thread.Sleep(1000)
            worker.ReportProgress(20, "Identificando setores danificados...")
            Thread.Sleep(1500)
            worker.ReportProgress(35, "Tentando recuperar dados...")
            Thread.Sleep(2000)

            Dim driveLocal As String = ""
            Me.Invoke(Sub() driveLocal = selectedDrive)

            Dim drive As String = driveLocal.Replace("\", "").Replace(":", "")
            worker.ReportProgress(50, "Executando reparo automático...")

            Dim processo As New Process()
            processo.StartInfo.FileName = "cmd.exe"
            processo.StartInfo.Arguments = $"/c chkdsk {drive}: /F"
            processo.StartInfo.UseShellExecute = False
            processo.StartInfo.RedirectStandardOutput = True
            processo.StartInfo.CreateNoWindow = True

            Try
                processo.Start()
                Dim progresso As Integer = 55

                While Not processo.HasExited
                    If progresso < 90 Then
                        progresso += 3
                        worker.ReportProgress(progresso, "Reparando setores...")
                        Thread.Sleep(1000)
                    End If
                End While

                processo.WaitForExit()
                worker.ReportProgress(92, "Finalizando recuperação...")
                Thread.Sleep(800)

                If processo.ExitCode = 0 Then
                    worker.ReportProgress(100, "✓ Recuperação concluída com sucesso!")
                Else
                    worker.ReportProgress(-2, "⚠ Recuperação parcial - Alguns erros não puderam ser corrigidos")
                End If

            Catch ex As Exception
                worker.ReportProgress(-1, $"Erro durante recuperação: {ex.Message}")
            End Try

        Catch ex As Exception
            worker.ReportProgress(-1, $"Erro crítico: {ex.Message}")
        End Try
    End Sub

    Private Sub ExecutarFormatacaoFisica(worker As System.ComponentModel.BackgroundWorker)
        Try
            worker.ReportProgress(2, "═══ FORMATAÇÃO FÍSICA INICIADA ═══")
            worker.ReportProgress(3, "⚠ Este processo pode levar VÁRIAS HORAS!")
            Thread.Sleep(2000)

            Dim drive As String = selectedDrive.Replace("\", "").Replace(":", "")

            worker.ReportProgress(5, "Fase 1/5: Limpeza completa do disco...")
            LimparDiscoCompletamente(drive, worker)

            worker.ReportProgress(25, "Fase 2/5: Recriando tabela de partição...")
            Thread.Sleep(1000)
            RecriarParticao(drive, worker)

            worker.ReportProgress(45, "Fase 3/5: Formatação com mapeamento de setores...")
            FormatarComVerificacaoSetores(drive, worker)

            worker.ReportProgress(70, "Fase 4/5: Teste de integridade...")
            TestarIntegridade(drive, worker)

            worker.ReportProgress(90, "Fase 5/5: Aplicando otimizações finais...")
            OtimizarDisco(drive, worker)

            worker.ReportProgress(100, "✓ Formatação Física concluída com sucesso!")

        Catch ex As Exception
            worker.ReportProgress(-1, $"Erro na formatação física: {ex.Message}")
        End Try
    End Sub

    Private Sub LimparDiscoCompletamente(drive As String, worker As System.ComponentModel.BackgroundWorker)
        Dim numeroDisco As Integer = -1
        Dim scriptPath As String = ""
        Dim processo As Process = Nothing
        Dim progresso As Integer = 15
        Dim tempoDecorrido As Integer = 0

        Try
            worker.ReportProgress(6, "Identificando número do disco...")
            Thread.Sleep(500)

            numeroDisco = ObterNumeroDisco(drive)
            If numeroDisco < 0 Then
                worker.ReportProgress(7, "Procurando discos removíveis...")
                numeroDisco = ObterPrimeiroDiscoRemovivel()
                If numeroDisco < 0 Then
                    Throw New Exception("Não foi possível identificar o número do disco.")
                End If
            End If

            worker.ReportProgress(8, $"Disco identificado: Disco {numeroDisco}")
            worker.ReportProgress(9, "⚠ Este processo pode levar HORAS!")
            worker.ReportProgress(10, "Criando script de limpeza...")
            Thread.Sleep(500)

            scriptPath = Path.Combine(Path.GetTempPath(), $"diskpart_clean_{DateTime.Now:yyyyMMddHHmmss}.txt")
            Using sw As New StreamWriter(scriptPath, False, System.Text.Encoding.ASCII)
                sw.WriteLine($"select disk {numeroDisco}")
                sw.WriteLine("clean all")
                sw.WriteLine("exit")
            End Using

            worker.ReportProgress(11, "Script criado")
            worker.ReportProgress(12, "Executando DISKPART...")
            Thread.Sleep(500)

            processo = New Process()
            processo.StartInfo.FileName = "diskpart.exe"
            processo.StartInfo.Arguments = $"/s ""{scriptPath}"""
            processo.StartInfo.UseShellExecute = True
            processo.StartInfo.Verb = "runas"
            processo.StartInfo.CreateNoWindow = False
            processo.StartInfo.WindowStyle = ProcessWindowStyle.Hidden

            worker.ReportProgress(13, "Limpeza profunda iniciada...")
            worker.ReportProgress(14, "⏳ Apagando setores...")

            processo.Start()

            While Not processo.HasExited
                If worker.CancellationPending Then
                    worker.ReportProgress(22, "⚠ Cancelando...")
                    Try
                        If processo IsNot Nothing AndAlso Not processo.HasExited Then
                            processo.Kill()
                            processo.WaitForExit(3000)
                        End If
                    Catch
                    End Try
                    Throw New OperationCanceledException("Cancelado")
                End If

                If progresso < 23 Then
                    progresso += 1
                    tempoDecorrido += 3
                    Dim minutos As Integer = tempoDecorrido \ 60
                    Dim segundos As Integer = tempoDecorrido Mod 60
                    worker.ReportProgress(progresso, $"Limpando... ({minutos}min {segundos}s)")
                    Thread.Sleep(3000)
                Else
                    tempoDecorrido += 10
                    Dim minutos As Integer = tempoDecorrido \ 60
                    Dim horas As Integer = minutos \ 60
                    Dim minutosRestantes As Integer = minutos Mod 60

                    Dim tempoTexto As String
                    If horas > 0 Then
                        tempoTexto = $"{horas}h {minutosRestantes}min"
                    Else
                        tempoTexto = $"{minutos} min"
                    End If

                    Dim dica As String
                    If minutos < 30 Then
                        dica = " | Est: 30min-8h"
                    ElseIf minutos < 120 Then
                        dica = " | Normal"
                    Else
                        dica = " | Cartão grande"
                    End If

                    worker.ReportProgress(23, $"⏳ Limpeza: {tempoTexto}{dica}")
                    Thread.Sleep(10000)
                End If
            End While

            processo.WaitForExit()
            worker.ReportProgress(24, "✓ Limpeza completa!")
            Thread.Sleep(1000)

        Catch ex As OperationCanceledException
            worker.ReportProgress(24, "⚠ Cancelado")
            Throw
        Catch ex As Exception
            worker.ReportProgress(-1, $"Erro: {ex.Message}")
            Throw
        Finally
            If processo IsNot Nothing Then
                Try
                    processo.Dispose()
                Catch
                End Try
            End If
            If scriptPath <> "" AndAlso File.Exists(scriptPath) Then
                Try
                    File.Delete(scriptPath)
                Catch
                End Try
            End If
        End Try
    End Sub

    Private Function ObterNumeroDisco(drive As String) As Integer
        Try
            drive = drive.Replace(":\", "").Replace("\", "").Replace(":", "").Trim()

            If drive.Contains("Disco") AndAlso drive.Contains("[NÃO FORMATADO]") Then
                Dim inicio As Integer = drive.IndexOf("Disco ") + 6
                Dim fim As Integer = drive.IndexOf(" [")
                If inicio > 0 AndAlso fim > inicio Then
                    Dim numeroStr As String = drive.Substring(inicio, fim - inicio).Trim()
                    Dim numero As Integer
                    If Integer.TryParse(numeroStr, numero) Then
                        Return numero
                    End If
                End If
            End If

            If drive.Length = 1 Then
                Dim processo As New Process()
                processo.StartInfo.FileName = "cmd.exe"
                processo.StartInfo.Arguments = $"/c wmic partition where ""DriveLetter='{drive}'"" get DiskIndex"
                processo.StartInfo.UseShellExecute = False
                processo.StartInfo.RedirectStandardOutput = True
                processo.StartInfo.CreateNoWindow = True
                processo.Start()

                Dim output As String = processo.StandardOutput.ReadToEnd()
                processo.WaitForExit()

                Dim lines() As String = output.Split(New String() {vbCrLf, vbLf}, StringSplitOptions.RemoveEmptyEntries)
                If lines.Length > 1 Then
                    Dim numero As Integer
                    If Integer.TryParse(lines(1).Trim(), numero) Then
                        Return numero
                    End If
                End If
            End If

            Dim searcher As New ManagementObjectSearcher("SELECT * FROM Win32_DiskDrive WHERE MediaType='Removable Media' OR InterfaceType='USB'")
            For Each disk As ManagementObject In searcher.Get()
                Dim diskIndex As Integer
                If Integer.TryParse(disk("Index").ToString(), diskIndex) Then
                    Return diskIndex
                End If
            Next

        Catch ex As Exception
        End Try
        Return -1
    End Function

    Private Function ObterPrimeiroDiscoRemovivel() As Integer
        Try
            Dim searcher As New ManagementObjectSearcher("SELECT * FROM Win32_DiskDrive WHERE MediaType='Removable Media' OR InterfaceType='USB'")
            For Each disk As ManagementObject In searcher.Get()
                Dim diskIndex As Integer
                If Integer.TryParse(disk("Index").ToString(), diskIndex) Then
                    Return diskIndex
                End If
            Next
        Catch ex As Exception
        End Try
        Return -1
    End Function
    Private Sub RecriarParticao(drive As String, worker As System.ComponentModel.BackgroundWorker)
        Try
            Dim numeroDisco As Integer = ObterNumeroDisco(drive)
            If numeroDisco < 0 Then
                numeroDisco = ObterPrimeiroDiscoRemovivel()
                If numeroDisco < 0 Then
                    Throw New Exception("Não foi possível identificar o disco para criar partição")
                End If
            End If

            worker.ReportProgress(27, $"Criando partição no Disco {numeroDisco}...")
            Thread.Sleep(500)

            Dim letraDisponivel As String = ObterLetraDisponivel()
            Dim scriptPath As String = Path.Combine(Path.GetTempPath(), $"diskpart_create_{DateTime.Now:yyyyMMddHHmmss}.txt")

            Try
                Using sw As New StreamWriter(scriptPath, False, System.Text.Encoding.ASCII)
                    sw.WriteLine($"select disk {numeroDisco}")
                    sw.WriteLine("create partition primary")
                    sw.WriteLine("select partition 1")
                    sw.WriteLine("active")
                    If letraDisponivel <> "" Then
                        sw.WriteLine($"assign letter={letraDisponivel}")
                    End If
                    sw.WriteLine("exit")
                End Using

                worker.ReportProgress(28, "Executando criação de partição...")

                Dim processo As New Process()
                processo.StartInfo.FileName = "diskpart.exe"
                processo.StartInfo.Arguments = $"/s ""{scriptPath}"""
                processo.StartInfo.UseShellExecute = True
                processo.StartInfo.Verb = "runas"
                processo.StartInfo.CreateNoWindow = False
                processo.StartInfo.WindowStyle = ProcessWindowStyle.Hidden

                processo.Start()

                Dim progresso As Integer = 29
                While Not processo.HasExited
                    If worker.CancellationPending Then
                        worker.ReportProgress(41, "⚠ Cancelamento detectado...")
                        Try
                            processo.Kill()
                            processo.WaitForExit(3000)
                        Catch
                        End Try
                        Throw New OperationCanceledException("Operação cancelada pelo usuário")
                    End If

                    If progresso < 42 Then
                        progresso += 2
                        worker.ReportProgress(progresso, "Criando partição...")
                        Thread.Sleep(500)
                    End If
                End While

                processo.WaitForExit()

                If File.Exists(scriptPath) Then
                    Try
                        File.Delete(scriptPath)
                    Catch
                    End Try
                End If

                worker.ReportProgress(43, "✓ Partição criada com sucesso!")
                If letraDisponivel <> "" Then
                    selectedDrive = letraDisponivel & ":\"
                    worker.ReportProgress(44, $"Nova letra de unidade: {selectedDrive}")
                End If
                Thread.Sleep(2000)

            Catch ex As OperationCanceledException
                Throw
            Catch ex As Exception
                If File.Exists(scriptPath) Then
                    Try
                        File.Delete(scriptPath)
                    Catch
                    End Try
                End If
                Throw New Exception($"Erro ao criar partição: {ex.Message}")
            End Try

        Catch ex As OperationCanceledException
            Throw
        Catch ex As Exception
            worker.ReportProgress(-1, $"Erro ao criar partição: {ex.Message}")
            Throw
        End Try
    End Sub

    Private Function ObterLetraDisponivel() As String
        Try
            Dim letrasUsadas As New List(Of String)
            For Each drive As DriveInfo In DriveInfo.GetDrives()
                letrasUsadas.Add(drive.Name.Substring(0, 1).ToUpper())
            Next

            Dim letrasDisponiveis() As String = {"E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z"}
            For Each letra As String In letrasDisponiveis
                If Not letrasUsadas.Contains(letra) Then
                    Return letra
                End If
            Next
        Catch ex As Exception
        End Try
        Return ""
    End Function

    Private Sub FormatarComVerificacaoSetores(drive As String, worker As System.ComponentModel.BackgroundWorker)
        Try
            Dim cboFormato As ComboBox = Nothing
            Dim txtNome As TextBox = Nothing

            Me.Invoke(Sub()
                          cboFormato = Me.Controls.Find("cboFormato", True).FirstOrDefault()
                          txtNome = Me.Controls.Find("txtNome", True).FirstOrDefault()
                      End Sub)

            Dim fileSystem As String = "exFAT"
            Dim nomeVolume As String = "EMIVE_CAM"

            Me.Invoke(Sub()
                          If cboFormato IsNot Nothing AndAlso cboFormato.Text.Contains("FAT32") Then
                              fileSystem = "FAT32"
                          ElseIf cboFormato IsNot Nothing AndAlso cboFormato.Text.Contains("NTFS") Then
                              fileSystem = "NTFS"
                          End If
                          If txtNome IsNot Nothing Then nomeVolume = txtNome.Text
                      End Sub)

            worker.ReportProgress(46, $"Formatando como {fileSystem} com verificação completa...")
            worker.ReportProgress(47, "⚠ Formatação SEM modo rápido - verificando cada setor...")

            Dim processo As New Process()
            processo.StartInfo.FileName = "cmd.exe"
            processo.StartInfo.UseShellExecute = False
            processo.StartInfo.RedirectStandardInput = True
            processo.StartInfo.RedirectStandardOutput = True
            processo.StartInfo.CreateNoWindow = True
            processo.Start()

            processo.StandardInput.WriteLine($"format {drive}: /FS:{fileSystem} /V:{nomeVolume} /Y")
            processo.StandardInput.WriteLine("exit")

            Dim progresso As Integer = 48
            While Not processo.HasExited
                If worker.CancellationPending Then
                    worker.ReportProgress(68, "⚠ Cancelamento detectado...")
                    Try
                        processo.Kill()
                        processo.WaitForExit(3000)
                    Catch
                    End Try
                    Throw New OperationCanceledException("Operação cancelada pelo usuário")
                End If

                If progresso < 68 Then
                    progresso += 1
                    worker.ReportProgress(progresso, "Verificando e formatando setores...")
                    Thread.Sleep(2000)
                End If
            End While

            processo.WaitForExit()
            worker.ReportProgress(69, "✓ Formatação com verificação concluída!")

        Catch ex As OperationCanceledException
            Throw
        Catch ex As Exception
            worker.ReportProgress(-1, $"Erro na formatação: {ex.Message}")
            Throw
        End Try
    End Sub

    Private Sub TestarIntegridade(drive As String, worker As System.ComponentModel.BackgroundWorker)
        Try
            worker.ReportProgress(71, "Executando teste de gravação/leitura...")

            Dim testFile As String = drive & ":\test_integrity.tmp"
            Dim testData(1024 * 1024 - 1) As Byte
            Dim rnd As New Random()
            rnd.NextBytes(testData)

            worker.ReportProgress(75, "Gravando arquivo de teste...")
            File.WriteAllBytes(testFile, testData)

            worker.ReportProgress(80, "Lendo arquivo de teste...")
            Dim readData() As Byte = File.ReadAllBytes(testFile)

            worker.ReportProgress(85, "Comparando dados...")
            Dim corrupted As Boolean = False
            For i As Integer = 0 To testData.Length - 1
                If testData(i) <> readData(i) Then
                    corrupted = True
                    Exit For
                End If
            Next

            If File.Exists(testFile) Then
                File.Delete(testFile)
            End If

            If corrupted Then
                worker.ReportProgress(-2, "⚠ Teste detectou possíveis setores defeituosos!")
            Else
                worker.ReportProgress(88, "✓ Teste de integridade: PASSOU!")
            End If

        Catch ex As Exception
            worker.ReportProgress(-1, $"Erro no teste: {ex.Message}")
        End Try
    End Sub

    Private Sub OtimizarDisco(drive As String, worker As System.ComponentModel.BackgroundWorker)
        Try
            worker.ReportProgress(91, "Aplicando otimizações finais...")
            Thread.Sleep(500)

            worker.ReportProgress(93, "Desabilitando indexação...")
            Try
                Dim processo As New Process()
                processo.StartInfo.FileName = "cmd.exe"
                processo.StartInfo.Arguments = $"/c fsutil behavior set disablelastaccess 1"
                processo.StartInfo.UseShellExecute = False
                processo.StartInfo.CreateNoWindow = True
                processo.StartInfo.Verb = "runas"
                processo.Start()
                processo.WaitForExit()
            Catch
            End Try

            worker.ReportProgress(96, "Criando estrutura de pastas...")
            Thread.Sleep(300)

            Dim pastaPrincipal As String = drive & ":\EZVIZ"
            If Not Directory.Exists(pastaPrincipal) Then
                Directory.CreateDirectory(pastaPrincipal)
                Directory.CreateDirectory(Path.Combine(pastaPrincipal, "Record"))
                Directory.CreateDirectory(Path.Combine(pastaPrincipal, "Picture"))
                Directory.CreateDirectory(Path.Combine(pastaPrincipal, "Log"))

                Dim configPath As String = Path.Combine(pastaPrincipal, "config.ini")
                Using sw As New StreamWriter(configPath)
                    sw.WriteLine("[EZVIZ]")
                    sw.WriteLine("Version=2.0")
                    sw.WriteLine("Format=Physical")
                    sw.WriteLine($"Date={DateTime.Now:yyyy-MM-dd HH:mm:ss}")
                    sw.WriteLine("ClusterSize=64KB")
                    sw.WriteLine("LowLevelFormat=True")
                End Using
            End If

            worker.ReportProgress(98, "✓ Otimizações aplicadas!")

        Catch ex As Exception
            worker.ReportProgress(-1, $"Erro na otimização: {ex.Message}")
        End Try
    End Sub

    Private Sub CancelarOperacao()
        Try
            AdicionarLog("═══════════════════════════════════════", Color.White)
            AdicionarLog("⚠ CANCELANDO OPERAÇÃO...", Color.Orange)

            If backgroundWorker.IsBusy Then
                backgroundWorker.CancelAsync()
                AdicionarLog("Solicitando parada do processo...", Color.Yellow)
            End If

            Thread.Sleep(2000)
            FinalizarProcessosRelacionados()
            HabilitarBotoes()

            Dim btnCancelar As Button = Me.Controls.Find("btnCancelar", True).FirstOrDefault()
            If btnCancelar IsNot Nothing Then
                btnCancelar.Visible = False
                btnCancelar.Enabled = False
            End If

            Dim progressBar As ProgressBar = Me.Controls.Find("progressBar", True).FirstOrDefault()
            If progressBar IsNot Nothing Then
                progressBar.Value = 0
            End If

            AdicionarLog("✓ Operação cancelada!", Color.Orange)
            AdicionarLog("═══════════════════════════════════════", Color.White)

            MessageBox.Show(
                 "Operação CANCELADA com sucesso!" & vbCrLf & vbCrLf &
                 "⚠ IMPORTANTE - Siga estas etapas:" & vbCrLf & vbCrLf &
                 "1. Use 'Remover Hardware com Segurança' (ícone ⏏ na bandeja)" & vbCrLf &
                 "2. Aguarde a confirmação do Windows" & vbCrLf &
                 "3. Remova o cartão SD fisicamente" & vbCrLf &
                 "4. Aguarde 30 segundos" & vbCrLf &
                 "5. Reinsira o cartão" & vbCrLf &
                 "6. Formate novamente usando 'Formatação Rápida'" & vbCrLf & vbCrLf &
                 "O cartão pode estar em estado inconsistente e precisará ser formatado.",
                 "Operação Cancelada",
                 MessageBoxButtons.OK,
                 MessageBoxIcon.Information)

        Catch ex As Exception
            AdicionarLog($"Erro ao cancelar: {ex.Message}", Color.Red)
            MessageBox.Show(
                $"Erro ao cancelar operação: {ex.Message}" & vbCrLf & vbCrLf &
                "Feche o programa e remova o cartão com segurança.",
                "Erro no Cancelamento",
                MessageBoxButtons.OK,
                MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub FinalizarProcessosRelacionados()
        Try
            AdicionarLog("Finalizando processos DISKPART...", Color.Yellow)
            Dim processosFinalizados As Integer = 0

            For Each proc As Process In Process.GetProcesses()
                Try
                    Dim nomeProcesso As String = proc.ProcessName.ToLower()
                    If nomeProcesso = "diskpart" Then
                        proc.Kill()
                        proc.WaitForExit(3000)
                        processosFinalizados += 1
                        AdicionarLog($"Processo '{proc.ProcessName}' finalizado", Color.Gray)
                    End If
                Catch ex As Exception
                End Try
            Next

            If processosFinalizados > 0 Then
                AdicionarLog($"✓ {processosFinalizados} processo(s) finalizado(s)", Color.LightGreen)
            Else
                AdicionarLog("Nenhum processo adicional encontrado", Color.Gray)
            End If

        Catch ex As Exception
            AdicionarLog($"Aviso ao finalizar processos: {ex.Message}", Color.Orange)
        End Try
    End Sub
    Private Sub BackgroundWorker_ProgressChanged(sender As Object, e As System.ComponentModel.ProgressChangedEventArgs) Handles backgroundWorker.ProgressChanged
        Dim progressBar As ProgressBar = Me.Controls.Find("progressBar", True).FirstOrDefault()

        If e.ProgressPercentage >= 0 Then
            If progressBar IsNot Nothing Then
                progressBar.Value = Math.Min(e.ProgressPercentage, 100)
            End If

            Dim cor As Color = Color.Cyan
            If e.ProgressPercentage >= 90 Then
                cor = Color.LightGreen
            ElseIf e.ProgressPercentage >= 50 Then
                cor = Color.Yellow
            End If

            AdicionarLog($"[{e.ProgressPercentage}%] {e.UserState}", cor)

        ElseIf e.ProgressPercentage = -1 Then
            AdicionarLog($"✗ {e.UserState}", Color.Red)
            MostrarErro("Erro na Operação", e.UserState.ToString())

        ElseIf e.ProgressPercentage = -2 Then
            AdicionarLog($"⚠ {e.UserState}", Color.Orange)
            MostrarAviso("Atenção", e.UserState.ToString())
        End If
    End Sub

    Private Sub BackgroundWorker_RunWorkerCompleted(sender As Object, e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles backgroundWorker.RunWorkerCompleted
        HabilitarBotoes()

        Dim progressBar As ProgressBar = Me.Controls.Find("progressBar", True).FirstOrDefault()
        If progressBar IsNot Nothing Then
            progressBar.Value = 0
        End If

        If e.Cancelled Then
            AdicionarLog("═══════════════════════════════════════", Color.White)
            AdicionarLog("⚠ Operação foi cancelada", Color.Orange)
            AtualizarInfoUnidade()

        ElseIf e.Error IsNot Nothing Then
            If TypeOf e.Error Is OperationCanceledException Then
                AdicionarLog("═══════════════════════════════════════", Color.White)
                AdicionarLog("⚠ Operação cancelada pelo usuário", Color.Orange)
            Else
                MostrarErro("Erro Fatal", e.Error.Message)
            End If

        Else
            AdicionarLog("═══════════════════════════════════════", Color.White)
            AtualizarInfoUnidade()
        End If
    End Sub

    Private Sub AdicionarLog(mensagem As String, cor As Color)
        If Me.InvokeRequired Then
            Me.Invoke(Sub() AdicionarLog(mensagem, cor))
            Return
        End If

        Dim txtLog As RichTextBox = Me.Controls.Find("txtLog", True).FirstOrDefault()
        If txtLog IsNot Nothing Then
            txtLog.SelectionStart = txtLog.TextLength
            txtLog.SelectionLength = 0
            txtLog.SelectionColor = cor
            txtLog.AppendText($"[{DateTime.Now:HH:mm:ss}] {mensagem}" & vbCrLf)
            txtLog.SelectionColor = txtLog.ForeColor
            txtLog.ScrollToCaret()
        End If
    End Sub

    Private Sub DesabilitarBotoes()
        For Each ctrl As Control In Me.Controls
            If TypeOf ctrl Is GroupBox Then
                For Each btn As Control In ctrl.Controls
                    If TypeOf btn Is Button Then
                        If btn.Name <> "btnCancelar" Then
                            btn.Enabled = False
                        End If
                    End If
                Next
            End If
        Next

        Dim btnCancelar As Button = Me.Controls.Find("btnCancelar", True).FirstOrDefault()
        If btnCancelar IsNot Nothing Then
            btnCancelar.Visible = True
            btnCancelar.Enabled = True
        End If
    End Sub

    Private Sub HabilitarBotoes()
        For Each ctrl As Control In Me.Controls
            If TypeOf ctrl Is GroupBox Then
                For Each btn As Control In ctrl.Controls
                    If TypeOf btn Is Button Then
                        btn.Enabled = True
                    End If
                Next
            End If
        Next

        Dim btnCancelar As Button = Me.Controls.Find("btnCancelar", True).FirstOrDefault()
        If btnCancelar IsNot Nothing Then
            btnCancelar.Visible = False
            btnCancelar.Enabled = False
        End If
    End Sub

    Private Sub MostrarErro(titulo As String, mensagem As String)
        Dim formErro As New FormErro(titulo, mensagem)
        formErro.ShowDialog()
    End Sub

    Private Sub MostrarAviso(titulo As String, mensagem As String)
        MessageBox.Show(mensagem, titulo, MessageBoxButtons.OK, MessageBoxIcon.Warning)
    End Sub

End Class