Imports System.Reflection.Emit
Imports System.Runtime.InteropServices
Imports System.Threading
Imports System.Timers

Public Class ProcessApp
    Private Declare Function GetLastInputInfo Lib "user32" (ByRef plii As UltimoClick) As Boolean
    'FUNCION PARA CAMBIAR A LOS COLORES DE LAS EMPRESAS
    Public Function CambiarColor(Empresa As String, P_Logo As Panel, P_Busqueda As Panel, P_ResultadoPerfil As Panel, P_Perfil As Panel, P_Ticket As Panel,
                                 P_ResultadoAviso As Panel, P_Aviso As Panel, Btn_Logo As PictureBox, Fondo_Empresa As PictureBox, Btn_BackSearch As Button,
                                 Btn_BackList As Button, Btn_BackResult As Button, Btn_BackTicket As Button, Btn_BackResultAviso As Button, Btn_BackAviso As Button,
                                 Btn_CrearTicket As PictureBox, Btn_GuardarTicket As Button, NotifyIcon As NotifyIcon, BloqueoForm As Object, Btn_MostrarAvisos As PictureBox)

        '------------------------------------------------------- COLORES EMPRESARIALES
        Dim CORP = "#8A2A2B"
        Dim OverBackCORP = "#FF0003" 'FB0101
        Dim MEDICINE = "#16A6D9"
        Dim OverBackMDN = "#00BCFF"  '00BCFF
        Dim PLASTIKISIMO = "#00577F"
        Dim OverBackPSKI = "#00AFFF"  '0087C4
        Dim PLASTIKA = "#001871"
        Dim OverBackPSK = "#0036FF"  '0036FF
        Select Case Empresa
            Case "SSIT02"
                '------------------------------------------------CORP
                P_Logo.BackColor = ColorTranslator.FromHtml(CORP)
                P_Busqueda.BackColor = ColorTranslator.FromHtml(CORP)
                P_ResultadoPerfil.BackColor = ColorTranslator.FromHtml(CORP)
                P_Perfil.BackColor = ColorTranslator.FromHtml(CORP)
                P_Ticket.BackColor = ColorTranslator.FromHtml(CORP)
                P_ResultadoAviso.BackColor = ColorTranslator.FromHtml(CORP)
                P_Aviso.BackColor = ColorTranslator.FromHtml(CORP)
                Btn_CrearTicket.BackColor = ColorTranslator.FromHtml(CORP)
                Btn_GuardarTicket.BackColor = ColorTranslator.FromHtml(CORP)
                Btn_BackSearch.FlatAppearance.MouseOverBackColor = ColorTranslator.FromHtml(OverBackCORP)
                Btn_BackList.FlatAppearance.MouseOverBackColor = ColorTranslator.FromHtml(OverBackCORP)
                Btn_BackResult.FlatAppearance.MouseOverBackColor = ColorTranslator.FromHtml(OverBackCORP)
                Btn_BackTicket.FlatAppearance.MouseOverBackColor = ColorTranslator.FromHtml(OverBackCORP)
                Btn_BackResultAviso.FlatAppearance.MouseOverBackColor = ColorTranslator.FromHtml(OverBackCORP)
                Btn_BackAviso.FlatAppearance.MouseOverBackColor = ColorTranslator.FromHtml(OverBackCORP)
                Btn_GuardarTicket.FlatAppearance.MouseOverBackColor = ColorTranslator.FromHtml(OverBackCORP)
                Btn_Logo.BackgroundImage = Image.FromFile(System.IO.Directory.GetCurrentDirectory + "\Resources\Empresas\Logos\Logo_CORP.jpg")
                Fondo_Empresa.BackgroundImage = Image.FromFile(System.IO.Directory.GetCurrentDirectory + "\Resources\Empresas\Iconos\Logo_CORP_Hoja.png")

            Case "FGI"
                '-------------------------------------------------FGI

                P_Logo.BackColor = ColorTranslator.FromHtml(MEDICINE)
                P_Busqueda.BackColor = ColorTranslator.FromHtml(MEDICINE)
                P_ResultadoPerfil.BackColor = ColorTranslator.FromHtml(MEDICINE)
                P_Perfil.BackColor = ColorTranslator.FromHtml(MEDICINE)
                P_Ticket.BackColor = ColorTranslator.FromHtml(MEDICINE)
                P_ResultadoAviso.BackColor = ColorTranslator.FromHtml(MEDICINE)
                P_Aviso.BackColor = ColorTranslator.FromHtml(MEDICINE)
                Btn_CrearTicket.BackColor = ColorTranslator.FromHtml(MEDICINE)
                Btn_GuardarTicket.BackColor = ColorTranslator.FromHtml(MEDICINE)
                Btn_BackSearch.FlatAppearance.MouseOverBackColor = ColorTranslator.FromHtml(OverBackMDN)
                Btn_BackList.FlatAppearance.MouseOverBackColor = ColorTranslator.FromHtml(OverBackMDN)
                Btn_BackResult.FlatAppearance.MouseOverBackColor = ColorTranslator.FromHtml(OverBackMDN)
                Btn_BackTicket.FlatAppearance.MouseOverBackColor = ColorTranslator.FromHtml(OverBackMDN)
                Btn_BackResultAviso.FlatAppearance.MouseOverBackColor = ColorTranslator.FromHtml(OverBackMDN)
                Btn_BackAviso.FlatAppearance.MouseOverBackColor = ColorTranslator.FromHtml(OverBackMDN)
                Btn_GuardarTicket.FlatAppearance.MouseOverBackColor = ColorTranslator.FromHtml(OverBackMDN)
                Btn_Logo.BackgroundImage = Image.FromFile(System.IO.Directory.GetCurrentDirectory + "\Resources\Empresas\Logos\Logo_FGI.jpg")
                Fondo_Empresa.BackgroundImage = Image.FromFile(System.IO.Directory.GetCurrentDirectory + "\Resources\Empresas\Iconos\Logo_FGI.jpg")

            Case "MEDICINE"
                '--------------------------------------------MEDICINE
                P_Logo.BackColor = ColorTranslator.FromHtml(MEDICINE)
                P_Busqueda.BackColor = ColorTranslator.FromHtml(MEDICINE)
                P_ResultadoPerfil.BackColor = ColorTranslator.FromHtml(MEDICINE)
                P_Perfil.BackColor = ColorTranslator.FromHtml(MEDICINE)
                P_Ticket.BackColor = ColorTranslator.FromHtml(MEDICINE)
                P_ResultadoAviso.BackColor = ColorTranslator.FromHtml(MEDICINE)
                P_Aviso.BackColor = ColorTranslator.FromHtml(MEDICINE)
                Btn_CrearTicket.BackColor = ColorTranslator.FromHtml(MEDICINE)
                Btn_GuardarTicket.BackColor = ColorTranslator.FromHtml(MEDICINE)
                Btn_BackSearch.FlatAppearance.MouseOverBackColor = ColorTranslator.FromHtml(OverBackMDN)
                Btn_BackList.FlatAppearance.MouseOverBackColor = ColorTranslator.FromHtml(OverBackMDN)
                Btn_BackResult.FlatAppearance.MouseOverBackColor = ColorTranslator.FromHtml(OverBackMDN)
                Btn_BackTicket.FlatAppearance.MouseOverBackColor = ColorTranslator.FromHtml(OverBackMDN)
                Btn_BackResultAviso.FlatAppearance.MouseOverBackColor = ColorTranslator.FromHtml(OverBackMDN)
                Btn_BackAviso.FlatAppearance.MouseOverBackColor = ColorTranslator.FromHtml(OverBackMDN)
                Btn_GuardarTicket.FlatAppearance.MouseOverBackColor = ColorTranslator.FromHtml(OverBackMDN)
                Btn_Logo.BackgroundImage = Image.FromFile(System.IO.Directory.GetCurrentDirectory + "\Resources\Empresas\Logos\Logo_Medicine.jpg")
                Fondo_Empresa.BackgroundImage = Image.FromFile(System.IO.Directory.GetCurrentDirectory + "\Resources\Empresas\Iconos\Logo_Medicine.jpg")

            Case "MOLDAPTA"
                '--------------------------------------------MOLDAPTA
                P_Logo.BackColor = ColorTranslator.FromHtml(PLASTIKA)
                P_Busqueda.BackColor = ColorTranslator.FromHtml(PLASTIKA)
                P_ResultadoPerfil.BackColor = ColorTranslator.FromHtml(PLASTIKA)
                P_Perfil.BackColor = ColorTranslator.FromHtml(PLASTIKA)
                P_Ticket.BackColor = ColorTranslator.FromHtml(PLASTIKA)
                P_ResultadoAviso.BackColor = ColorTranslator.FromHtml(PLASTIKA)
                P_Aviso.BackColor = ColorTranslator.FromHtml(PLASTIKA)
                Btn_CrearTicket.BackColor = ColorTranslator.FromHtml(PLASTIKA)
                Btn_GuardarTicket.BackColor = ColorTranslator.FromHtml(PLASTIKA)
                Btn_BackSearch.FlatAppearance.MouseOverBackColor = ColorTranslator.FromHtml(OverBackPSK)
                Btn_BackList.FlatAppearance.MouseOverBackColor = ColorTranslator.FromHtml(OverBackPSK)
                Btn_BackResult.FlatAppearance.MouseOverBackColor = ColorTranslator.FromHtml(OverBackPSK)
                Btn_BackTicket.FlatAppearance.MouseOverBackColor = ColorTranslator.FromHtml(OverBackPSK)
                Btn_BackResultAviso.FlatAppearance.MouseOverBackColor = ColorTranslator.FromHtml(OverBackPSK)
                Btn_BackAviso.FlatAppearance.MouseOverBackColor = ColorTranslator.FromHtml(OverBackPSK)
                Btn_GuardarTicket.FlatAppearance.MouseOverBackColor = ColorTranslator.FromHtml(OverBackPSK)
                Btn_Logo.BackgroundImage = Image.FromFile(System.IO.Directory.GetCurrentDirectory + "\Resources\Empresas\Logos\Logo_Moldapta.jpg")
                Fondo_Empresa.BackgroundImage = Image.FromFile(System.IO.Directory.GetCurrentDirectory + "\Resources\Empresas\Iconos\Logo_Moldapta.jpg")

            Case "PLASTIKISIMO"
                '----------------------------------------PLASTIKISIMO
                P_Logo.BackColor = ColorTranslator.FromHtml(PLASTIKISIMO)
                P_Busqueda.BackColor = ColorTranslator.FromHtml(PLASTIKISIMO)
                P_ResultadoPerfil.BackColor = ColorTranslator.FromHtml(PLASTIKISIMO)
                P_Perfil.BackColor = ColorTranslator.FromHtml(PLASTIKISIMO)
                P_Ticket.BackColor = ColorTranslator.FromHtml(PLASTIKISIMO)
                P_ResultadoAviso.BackColor = ColorTranslator.FromHtml(PLASTIKISIMO)
                P_Aviso.BackColor = ColorTranslator.FromHtml(PLASTIKISIMO)
                Btn_CrearTicket.BackColor = ColorTranslator.FromHtml(PLASTIKISIMO)
                Btn_GuardarTicket.BackColor = ColorTranslator.FromHtml(PLASTIKISIMO)
                Btn_BackSearch.FlatAppearance.MouseOverBackColor = ColorTranslator.FromHtml(OverBackPSKI)
                Btn_BackList.FlatAppearance.MouseOverBackColor = ColorTranslator.FromHtml(OverBackPSKI)
                Btn_BackResult.FlatAppearance.MouseOverBackColor = ColorTranslator.FromHtml(OverBackPSKI)
                Btn_BackTicket.FlatAppearance.MouseOverBackColor = ColorTranslator.FromHtml(OverBackPSKI)
                Btn_BackResultAviso.FlatAppearance.MouseOverBackColor = ColorTranslator.FromHtml(OverBackPSKI)
                Btn_BackAviso.FlatAppearance.MouseOverBackColor = ColorTranslator.FromHtml(OverBackPSKI)
                Btn_GuardarTicket.FlatAppearance.MouseOverBackColor = ColorTranslator.FromHtml(OverBackPSKI)
                Btn_Logo.BackgroundImage = Image.FromFile(System.IO.Directory.GetCurrentDirectory + "\Resources\Empresas\Logos\Logo_Plastikisimo.jpg")
                Fondo_Empresa.BackgroundImage = Image.FromFile(System.IO.Directory.GetCurrentDirectory + "\Resources\Empresas\Iconos\Logo_Plastikisimo.jpg")

            Case "PLASTIKA"
                '--------------------------------------------PLASTIKA
                P_Logo.BackColor = ColorTranslator.FromHtml(PLASTIKA)
                P_Busqueda.BackColor = ColorTranslator.FromHtml(PLASTIKA)
                P_ResultadoPerfil.BackColor = ColorTranslator.FromHtml(PLASTIKA)
                P_Perfil.BackColor = ColorTranslator.FromHtml(PLASTIKA)
                P_Ticket.BackColor = ColorTranslator.FromHtml(PLASTIKA)
                P_ResultadoAviso.BackColor = ColorTranslator.FromHtml(PLASTIKA)
                P_Aviso.BackColor = ColorTranslator.FromHtml(PLASTIKA)
                Btn_CrearTicket.BackColor = ColorTranslator.FromHtml(PLASTIKA)
                Btn_GuardarTicket.BackColor = ColorTranslator.FromHtml(PLASTIKA)
                Btn_BackSearch.FlatAppearance.MouseOverBackColor = ColorTranslator.FromHtml(OverBackPSK)
                Btn_BackList.FlatAppearance.MouseOverBackColor = ColorTranslator.FromHtml(OverBackPSK)
                Btn_BackResult.FlatAppearance.MouseOverBackColor = ColorTranslator.FromHtml(OverBackPSK)
                Btn_BackTicket.FlatAppearance.MouseOverBackColor = ColorTranslator.FromHtml(OverBackPSK)
                Btn_BackResultAviso.FlatAppearance.MouseOverBackColor = ColorTranslator.FromHtml(OverBackPSK)
                Btn_BackAviso.FlatAppearance.MouseOverBackColor = ColorTranslator.FromHtml(OverBackPSK)
                Btn_GuardarTicket.FlatAppearance.MouseOverBackColor = ColorTranslator.FromHtml(OverBackPSK)
                Btn_Logo.BackgroundImage = Image.FromFile(System.IO.Directory.GetCurrentDirectory + "\Resources\Empresas\Logos\Logo_PSK.png")
                Fondo_Empresa.BackgroundImage = Image.FromFile(System.IO.Directory.GetCurrentDirectory + "\Resources\Empresas\Iconos\Logo_PSK.png")

            Case Else
                'MessageBox.Show("Tu equipo no esta registrado por la empresa" + Chr(10) + "Comunicate con sistemas", "LUIN | ERROR")
                NotifyIcon.Visible = False
                NotifyIcon.Visible = True
                NotifyIcon.BalloonTipTitle = "Equipo no registrado"
                NotifyIcon.BalloonTipText = "Comunicate con sistemas para resolver tu problema."
                NotifyIcon.ShowBalloonTip(2000)
                BloqueoForm.Enabled = False
                Btn_Logo.Enabled = False
                Btn_MostrarAvisos.Enabled = False
                Btn_CrearTicket.Enabled = False
                P_Logo.BackColor = ColorTranslator.FromHtml(CORP)
        End Select
    End Function

    'TAMAÑO DE CIRCULO DEL SISTEMA, POSICION Y FECHA
    Public Function CircleSize(Medida As Object, Logo As Panel, Hora As String, Fecha As String, Tiempo As Integer)
        'Prepara formulario 
        Medida.Size = New Size(300, 300)
        Logo.Location = New Point(0, 0)
        'P_Perfil.Location = New Point(0, 0)
        Hora = String.Format("{0:hh:mm tt}", DateTime.Now)
        Fecha = DateTime.Now.ToLongDateString()
        Tiempo = 1000  ' Un segundo

        'La propiedad GraphicsPath representa 
        'una serie de lineas y curvas conectadas.
        Dim miPath As New System.Drawing.Drawing2D.GraphicsPath

        'Esta linea agrega un elipse al grafico
        'usando las propiedades de ancho y alto del From
        miPath.AddEllipse(0, 0, Medida.Width, Medida.Height)

        'Agregamos el path a una nueva propiedad de Region
        Dim miRegion As New Region(miPath)

        'Por ultimo asignamos nuestra region a la del Formulario
        Medida.Region = miRegion
    End Function




    '--------------------------------------------------------------------------- MONITOREAR EL MOVIMIENTO DEL SISTEMA Y ENVIO DE NOTIFICACION
    Structure UltimoClick
        Public cbSize As Integer
        Public dwTime As Integer
    End Structure

    Dim Click As New UltimoClick()
    Public Function SesionSistema(Tiempo As Object)
        Click.cbSize = Marshal.SizeOf(Click)
        Tiempo.Interval = 1000
        Tiempo.Start()
    End Function
    Public Function TimerAction(Tiempo As Object, P_Logo As Panel, P_Buscar As Panel, P_ResultP As Panel, P_Perfil As Panel, P_Ticket As Panel, P_ResultA As Panel, P_Aviso As Panel, NotifyIcon As NotifyIcon)
        GetLastInputInfo(Click)

        Dim Total As Integer = Environment.TickCount
        Dim Ultimo As Integer = Click.dwTime
        Dim Intervalo As Integer = (Total - Ultimo) / 1000

        Dim Segundos As Integer = 30
        If Intervalo >= Segundos Then
            Tiempo.Stop()
            'MsgBox(Convert.ToString(Segundos) + " SEGUNDOS DE INACTIVIDAD")
            NotifyIcon.Visible = False
            NotifyIcon.Visible = True
            NotifyIcon.BalloonTipTitle = "Inactividad en el sistema"
            NotifyIcon.BalloonTipText = "Será devuelto a la pantalla principal."
            NotifyIcon.ShowBalloonTip(2000)
            P_Logo.Location = New Point(0, 0)
            P_Buscar.Location = New Point(634, 0)
            P_ResultP.Location = New Point(947, 0)
            P_Perfil.Location = New Point(0, 364)
            P_Ticket.Location = New Point(319, 364)
            P_ResultA.Location = New Point(634, 364)
            P_Aviso.Location = New Point(947, 364)
        End If
    End Function
End Class
