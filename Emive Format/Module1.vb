Imports System.Windows.Forms
Imports System.Security.Principal

Module Module1
    <STAThread>
    Sub Main()
        ' Verificar se está executando como administrador
        If Not IsAdministrator() Then
            MessageBox.Show(
                "Este aplicativo requer privilégios de administrador para formatar e verificar unidades." & vbCrLf & vbCrLf &
                "Por favor, execute como Administrador.",
                "Privilégios Insuficientes",
                MessageBoxButtons.OK,
                MessageBoxIcon.Warning)
            Application.Exit()
            Return
        End If

        Application.EnableVisualStyles()
        Application.SetCompatibleTextRenderingDefault(False)
        Application.Run(New FormPrincipal())
    End Sub

    Private Function IsAdministrator() As Boolean
        Try
            Dim identity = WindowsIdentity.GetCurrent()
            Dim principal = New WindowsPrincipal(identity)
            Return principal.IsInRole(WindowsBuiltInRole.Administrator)
        Catch ex As Exception
            Return False
        End Try
    End Function
End Module