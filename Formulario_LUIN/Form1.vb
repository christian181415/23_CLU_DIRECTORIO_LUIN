Public Class Form1
    Dim NuevoProceso As ProcessApp = New ProcessApp()

    '------------------------------------------------------------------------------------------------- LOAD INICIAL
    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Try
            NuevoProceso.CircleSize(Me, P_Logo, Lb_Hora.Text, Lb_Fecha.Text, Timer1.Interval)
            Timer1.Start()

            Dim r As Rectangle = My.Computer.Screen.WorkingArea
            Dim Largo = r.Width - Width
            Dim Alto = r.Height - Height
            Location = New Point(r.Width - Width - 10, 10)

            'Dim Empresa = "INEXISTENTE"
            'Dim Empresa = "FGI"
            'Dim Empresa = "MEDICINE"
            'Dim Empresa = "MOLDAPTA"
            'Dim Empresa = "PLASTIKISIMO"
            'Dim Empresa = "PLASTIKA"
            Dim Empresa = System.Environment.GetEnvironmentVariable("COMPUTERNAME")
            NuevoProceso.CambiarColor(Empresa, Footer1, Footer2, Footer3, Footer4, Footer5, Footer6, Footer7, btnLogo,
                                      FondoEmpresa, btnBackSearch, btnBackList, btnBackResult, btnBackTicket,
                                      btnBackResultAviso, btnBackAviso, btnCrearTicket, btnGuardarTicket, NotifyIcon, Me, btnMostrarAvisos)
        Catch ex As Exception
            Me.NotifyIcon.Visible = False
            Me.NotifyIcon.Visible = True
            Me.NotifyIcon.BalloonTipTitle = "Problemas en el software"
            Me.NotifyIcon.BalloonTipText = "Contacte con sistemas para resolver su problema."
            Me.NotifyIcon.ShowBalloonTip(2000)
            Me.Enabled = False
            btnLogo.Enabled = False
            btnMostrarAvisos.Enabled = False
            btnCrearTicket.Enabled = False
        End Try
    End Sub

    Private Sub MenuCerrar_Click(sender As Object, e As EventArgs) Handles MenuCerrar.Click
        Application.Exit()
    End Sub






    '-------------------------------------------------------------------------------------------------- P_LOGO
    'Click para ir a P_BUSQUEDA
    Private Sub btnLogo_Click(sender As Object, e As EventArgs) Handles btnLogo.Click
        P_Logo.Location = New Point(319, 0)
        P_Busqueda.Location = New Point(0, 0)

        NuevoProceso.SesionSistema(Timer2)
    End Sub


    'Click para ir a P_AVISOS
    Private Sub btnMostrarAvisos_Click(sender As Object, e As EventArgs) Handles btnMostrarAvisos.Click
        P_Logo.Location = New Point(319, 0)
        P_ResultadoAviso.Location = New Point(0, 0)

        NuevoProceso.SesionSistema(Timer2)
    End Sub
    'Click para ir a P_TICKET
    Private Sub btnCrearTicket_Click(sender As Object, e As EventArgs) Handles btnCrearTicket.Click
        P_Logo.Location = New Point(319, 0)
        P_Ticket.Location = New Point(0, 0)

        NuevoProceso.SesionSistema(Timer2)
    End Sub
    'Timer control de hora
    Private Sub Timer1_Tick(sender As Object, e As EventArgs) Handles Timer1.Tick
        Lb_Hora.Text = String.Format("{0:hh:mm tt}", DateTime.Now)
    End Sub
    '------------------------------------------------------------------------------------------------------------------------------------------------TIMER SESION USUARIO
    Private Sub Timer2_Tick(sender As Object, e As EventArgs) Handles Timer2.Tick
        NuevoProceso.TimerAction(Timer2, P_Logo, P_Busqueda, P_ResultadoPerfil, P_Perfil, P_Ticket, P_ResultadoAviso, P_Aviso, Me.NotifyIcon)
    End Sub



    '--------------------------------------------------------------------------------------------------- P_BUSQUEDA
    'Click Regresar a P_LOGO
    Private Sub btnBackSearch_Click(sender As Object, e As EventArgs) Handles btnBackSearch.Click
        P_Busqueda.Location = New Point(634, 0)
        P_Logo.Location = New Point(0, 0)
        txtBuscar.Text = ""
    End Sub
    'Click para ir a P_RESULTADOS
    Private Sub btnSearch_Click(sender As Object, e As EventArgs) Handles btnSearch.Click
        P_Busqueda.Location = New Point(634, 0)
        P_ResultadoPerfil.Location = New Point(0, 0)
        txtBuscar.Text = ""
    End Sub



    '--------------------------------------------------------------------------------------------------- P_RESULTADO_PERFIL
    'Click Regresar a P_BUSQUEDA
    Private Sub btnBackList_Click_1(sender As Object, e As EventArgs) Handles btnBackList.Click
        P_ResultadoPerfil.Location = New Point(947, 0)
        P_Busqueda.Location = New Point(0, 0)
    End Sub
    'Click para ir a P_PERFIL
    Private Sub lbResultado_SelectedIndexChanged(sender As Object, e As EventArgs) Handles lbResultado.SelectedIndexChanged
        If lbResultado.SelectedIndex = 0 Then
            P_ResultadoPerfil.Location = New Point(947, 0)
            P_Perfil.Location = New Point(0, 0)
        ElseIf lbResultado.SelectedIndex = 1 Then
            P_ResultadoPerfil.Location = New Point(947, 0)
            P_Ticket.Location = New Point(0, 0)
        ElseIf lbResultado.SelectedIndex = 2 Then
            P_ResultadoPerfil.Location = New Point(947, 0)
            P_ResultadoAviso.Location = New Point(0, 0)
        ElseIf lbResultado.SelectedIndex = 3 Then
            P_ResultadoPerfil.Location = New Point(947, 0)
            P_Aviso.Location = New Point(0, 0)
        End If
    End Sub



    '----------------------------------------------------------------------------------------------------- P_PERFIL
    'Click para regresar a P_BUSQUEDA
    Private Sub btnBackResult_Click_1(sender As Object, e As EventArgs) Handles btnBackResult.Click
        P_Perfil.Location = New Point(0, 364)
        P_ResultadoPerfil.Location = New Point(0, 0)
    End Sub
    Private Sub btnCorreo_Click(sender As Object, e As EventArgs) Handles btnCorreo.Click
        'System.Diagnostics.Process.Start("Outlook.exe")
        'System.Diagnostics.Process.Start("http://corporativoluin.mx/")
        'System.Diagnostics.Process.Start("mailto:" & "christian181415@gmail.com")
        Dim RutaOutlook As New ProcessStartInfo
        RutaOutlook.FileName = "C:\Program Files (x86)\Microsoft Office\Office15\OUTLOOK.exe"
        Process.Start(RutaOutlook)
    End Sub



    '---------------------------------------------------------------------------------------------------- P_TICKET
    Private Sub btnBackTicket_Click(sender As Object, e As EventArgs) Handles btnBackTicket.Click
        P_Ticket.Location = New Point(319, 364)
        P_Logo.Location = New Point(0, 0)
        txtTitulo.Text = ""
        txtProblema.Text = ""
    End Sub



    '---------------------------------------------------------------------------------------------------- P_RESULTADO_AVISO
    'Click para regresar a P_LOGO
    Private Sub btnBackResultAviso_Click(sender As Object, e As EventArgs) Handles btnBackResultAviso.Click
        P_ResultadoAviso.Location = New Point(634, 364)
        P_Logo.Location = New Point(0, 0)
    End Sub
    Private Sub lbResultadoAviso_SelectedIndexChanged(sender As Object, e As EventArgs) Handles lbResultadoAviso.SelectedIndexChanged
        If lbResultadoAviso.SelectedIndex = 0 Then
            P_ResultadoAviso.Location = New Point(634, 364)
            P_Aviso.Location = New Point(0, 0)
        End If
    End Sub



    '---------------------------------------------------------------------------------------------------- P_AVISO
    Private Sub btnBackAviso_Click(sender As Object, e As EventArgs) Handles btnBackAviso.Click
        P_Aviso.Location = New Point(947, 364)
        P_ResultadoAviso.Location = New Point(0, 0)
    End Sub
End Class
