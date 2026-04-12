Imports System.Drawing
Imports System.Windows.Forms

Public Class FormErro
    Inherits System.Windows.Forms.Form

    Private tituloErro As String
    Private mensagemErro As String

    Public Sub New(titulo As String, mensagem As String)
        MyBase.New()
        tituloErro = titulo
        mensagemErro = mensagem
    End Sub

    Private Sub FormErro_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Me.Text = "EMIVE Smart Alarme - Erro"
        Me.Size = New Size(600, 400)
        Me.StartPosition = FormStartPosition.CenterScreen
        Me.FormBorderStyle = FormBorderStyle.FixedDialog
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.BackColor = Color.FromArgb(44, 62, 80)

        ' Painel superior com ícone de erro
        Dim panelTop As New Panel With {
            .Dock = DockStyle.Top,
            .Height = 120,
            .BackColor = Color.FromArgb(192, 57, 43)
        }
        Me.Controls.Add(panelTop)

        ' Ícone de erro
        Dim lblIcone As New Label With {
            .Text = "⚠",
            .Font = New Font("Segoe UI", 48, FontStyle.Bold),
            .ForeColor = Color.White,
            .AutoSize = True,
            .Location = New Point(250, 25)
        }
        panelTop.Controls.Add(lblIcone)

        ' Título do erro
        Dim lblTitulo As New Label With {
            .Text = tituloErro,
            .Font = New Font("Segoe UI", 16, FontStyle.Bold),
            .ForeColor = Color.White,
            .Location = New Point(20, 140),
            .Size = New Size(560, 35),
            .TextAlign = ContentAlignment.MiddleCenter
        }
        Me.Controls.Add(lblTitulo)

        ' Painel de mensagem
        Dim panelMsg As New Panel With {
            .Location = New Point(20, 185),
            .Size = New Size(560, 120),
            .BackColor = Color.FromArgb(52, 73, 94),
            .BorderStyle = BorderStyle.FixedSingle
        }
        Me.Controls.Add(panelMsg)

        ' Mensagem de erro
        Dim lblMensagem As New Label With {
            .Text = mensagemErro,
            .Font = New Font("Segoe UI", 11),
            .ForeColor = Color.FromArgb(236, 240, 241),
            .Location = New Point(15, 15),
            .Size = New Size(530, 90),
            .TextAlign = ContentAlignment.MiddleCenter
        }
        panelMsg.Controls.Add(lblMensagem)

        ' Botão OK
        Dim btnOK As New Button With {
            .Text = "OK",
            .Location = New Point(230, 320),
            .Size = New Size(140, 40),
            .BackColor = Color.FromArgb(41, 128, 185),
            .ForeColor = Color.White,
            .FlatStyle = FlatStyle.Flat,
            .Font = New Font("Segoe UI", 12, FontStyle.Bold),
            .Cursor = Cursors.Hand,
            .DialogResult = DialogResult.OK
        }
        btnOK.FlatAppearance.BorderSize = 0
        Me.Controls.Add(btnOK)
        Me.AcceptButton = btnOK
    End Sub
End Class