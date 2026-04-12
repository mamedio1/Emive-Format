Imports System.Drawing
Imports System.Windows.Forms

Public Class FormSobre
    Inherits System.Windows.Forms.Form

    Public Sub New()
        MyBase.New()
    End Sub

    Private Sub FormSobre_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Me.Text = "Sobre - EMIVE Smart Alarme"
        Me.Size = New Size(550, 500)
        Me.StartPosition = FormStartPosition.CenterScreen
        Me.FormBorderStyle = FormBorderStyle.FixedDialog
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.BackColor = Color.FromArgb(236, 240, 241)

        ' Header
        Dim panelHeader As New Panel With {
            .Dock = DockStyle.Top,
            .Height = 100,
            .BackColor = Color.FromArgb(41, 128, 185)
        }
        Me.Controls.Add(panelHeader)

        Dim lblTitulo As New Label With {
            .Text = "EMIVE Smart Alarme",
            .Font = New Font("Segoe UI", 24, FontStyle.Bold),
            .ForeColor = Color.White,
            .AutoSize = True,
            .Location = New Point(20, 30)
        }
        panelHeader.Controls.Add(lblTitulo)

        Dim lblVersao As New Label With {
            .Text = "Versão 1.0.0 - 2026",
            .Font = New Font("Segoe UI", 10),
            .ForeColor = Color.FromArgb(236, 240, 241),
            .AutoSize = True,
            .Location = New Point(25, 70)
        }
        panelHeader.Controls.Add(lblVersao)

        ' Informações
        Dim yPos As Integer = 120

        Dim lblDescricao As New Label With {
            .Text = "Sistema Profissional de Gerenciamento de Cartões SD" & vbCrLf &
                    "para Câmeras EZVIZ",
            .Font = New Font("Segoe UI", 11, FontStyle.Bold),
            .ForeColor = Color.FromArgb(44, 62, 80),
            .Location = New Point(30, yPos),
            .Size = New Size(490, 50),
            .TextAlign = ContentAlignment.TopCenter
        }
        Me.Controls.Add(lblDescricao)

        yPos += 70

        ' Box de funcionalidades
        Dim panelFunc As New Panel With {
            .Location = New Point(30, yPos),
            .Size = New Size(490, 140),
            .BackColor = Color.White,
            .BorderStyle = BorderStyle.FixedSingle
        }
        Me.Controls.Add(panelFunc)

        Dim lblFuncTitulo As New Label With {
            .Text = "Funcionalidades:",
            .Font = New Font("Segoe UI", 11, FontStyle.Bold),
            .ForeColor = Color.FromArgb(41, 128, 185),
            .Location = New Point(15, 10),
            .AutoSize = True
        }
        panelFunc.Controls.Add(lblFuncTitulo)

        Dim lblFuncLista As New Label With {
            .Text = "✓ Formatação em múltiplos formatos (FAT32, exFAT, NTFS)" & vbCrLf &
                    "✓ Formato otimizado especial para câmeras EZVIZ" & vbCrLf &
                    "✓ Verificação e detecção de erros" & vbCrLf &
                    "✓ Recuperação automática de setores danificados" & vbCrLf &
                    "✓ Interface intuitiva e profissional" & vbCrLf &
                    "✓ Log detalhado de todas as operações",
            .Font = New Font("Segoe UI", 9),
            .ForeColor = Color.FromArgb(52, 73, 94),
            .Location = New Point(15, 40),
            .Size = New Size(460, 95)
        }
        panelFunc.Controls.Add(lblFuncLista)

        yPos += 160

        ' Painel do autor
        Dim panelAutor As New Panel With {
            .Location = New Point(30, yPos),
            .Size = New Size(490, 100),
            .BackColor = Color.FromArgb(52, 152, 219),
            .BorderStyle = BorderStyle.FixedSingle
        }
        Me.Controls.Add(panelAutor)

        Dim lblAutorTitulo As New Label With {
            .Text = "Desenvolvido por:",
            .Font = New Font("Segoe UI", 11, FontStyle.Bold),
            .ForeColor = Color.White,
            .Location = New Point(15, 15),
            .AutoSize = True
        }
        panelAutor.Controls.Add(lblAutorTitulo)

        Dim lblAutor As New Label With {
            .Text = "Robson Mamedio Araujo",
            .Font = New Font("Segoe UI", 14, FontStyle.Bold),
            .ForeColor = Color.White,
            .Location = New Point(15, 40),
            .AutoSize = True
        }
        panelAutor.Controls.Add(lblAutor)

        Dim lblCargo As New Label With {
            .Text = "Supervisor Técnico de Brasília",
            .Font = New Font("Segoe UI", 10),
            .ForeColor = Color.FromArgb(236, 240, 241),
            .Location = New Point(15, 68),
            .AutoSize = True
        }
        panelAutor.Controls.Add(lblCargo)

        yPos += 110

        ' Botão Fechar
        Dim btnFechar As New Button With {
            .Text = "Fechar",
            .Location = New Point(200, yPos + 10),
            .Size = New Size(150, 35),
            .BackColor = Color.FromArgb(149, 165, 166),
            .ForeColor = Color.White,
            .FlatStyle = FlatStyle.Flat,
            .Font = New Font("Segoe UI", 11, FontStyle.Bold),
            .Cursor = Cursors.Hand,
            .DialogResult = DialogResult.OK
        }
        btnFechar.FlatAppearance.BorderSize = 0
        Me.Controls.Add(btnFechar)
        Me.AcceptButton = btnFechar
    End Sub
End Class