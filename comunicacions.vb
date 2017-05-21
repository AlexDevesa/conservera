Imports System
Imports System.Net.Sockets
Imports System.Net
Imports System.Data
Imports MySql.Data.MySqlClient


Module comunicacions
#Region "Comunicación Modbus"

    Public falloMb As Boolean = False
    Public falloBBDD As Boolean = False
    Public estadoSocket As Boolean = False
    Public ERRO As Boolean = False

    Public conexion As New Socket(AddressFamily.InterNetwork, SocketType.Stream, ProtocolType.Tcp)
    Public conexion2 As New Socket(AddressFamily.InterNetwork, SocketType.Stream, ProtocolType.Tcp)
    Public Function comprobaConexion() As Boolean
        If conexion.Connected Then
            estadoSocket = True
        Else
            estadoSocket = False
            hmi.reloj.Enabled = False
            com.Abort()
            hmi.Button1.Enabled = True
        End If
        Return estadoSocket
    End Function

    Public Sub conectaMb()
        Try
            Dim sServidor As New IPEndPoint(IPAddress.Parse("192.168.1.150"), 502)
            conexion.Connect(sServidor)
            conexion2.Connect(sServidor)
            falloMb = False
        Catch ex As Exception
            If falloMb = False Then
                falloMb = True
                MsgBox("ERRO CONECTANDO AO AUTÓMATA")
            End If
        End Try
    End Sub
   

    Public Sub desconectaMb()
        conexion.Dispose()
    End Sub

    Public Sub mbler(ByVal primero As UShort, ByVal cuantos As UShort, ByVal valores() As UShort)
        Dim smd(11) As Byte
        smd(0) = 0
        smd(1) = 0
        smd(2) = 0
        smd(3) = 0
        smd(4) = 0
        smd(5) = 6
        smd(6) = 1
        smd(7) = Hex(3)
        smd(8) = CByte(Val(primero) \ 256)
        smd(9) = CByte(Val(primero) Mod 256)
        smd(10) = CByte(Val(cuantos) \ 256)
        smd(11) = CByte(Val(cuantos) Mod 256)

        Try
            falloMb = False
            conexion.Send(smd)
            Dim respuesta(9 + cuantos * 2) As Byte
            conexion.Receive(respuesta)
            If (respuesta(7) & Hex(80)) = Hex(80) Then
                Throw New Exception("Código de error enviado por el esclavo: 0x" + respuesta(7).ToString("X2"))
            End If
            For i As Integer = 0 To cuantos - 1
                valores(i) = CUShort(respuesta(9 + i * 2) * 256) + CUShort(respuesta(9 + i * 2 + 1))
            Next
        Catch
            falloMb = True
        Finally

        End Try

        
    End Sub
    Public Sub mbcomunica()
        Dim mmmb(15) As Byte
        mmmb(0) = 0
        mmmb(1) = 0
        mmmb(2) = 0
        mmmb(3) = 0
        mmmb(4) = 0
        mmmb(5) = 6
        mmmb(6) = 1
        mmmb(7) = Hex(6)
        mmmb(8) = 0
        mmmb(9) = 15
        mmmb(10) = 0
        mmmb(11) = 1
       
        Try
            conexion2.Send(mmmb)
            Dim respuestaa(12) As Byte
            Dim nRecibidos As Integer = conexion2.Receive(respuestaa)
          
        Catch ex As Exception

        End Try
    End Sub

    Public Sub mbescribe0(ByVal direccion As UShort)
        Dim mmb(15) As Byte
        mmb(0) = 0
        mmb(1) = 0
        mmb(2) = 0
        mmb(3) = 0
        mmb(4) = 0
        mmb(5) = 6
        mmb(6) = 1
        mmb(7) = Hex(6)
        mmb(8) = 0
        mmb(9) = direccion
        mmb(10) = 0
        mmb(11) = 0

        Try
            conexion.Send(mmb)
            Dim respuestaa(12) As Byte
            Dim nRecibidos As Integer = conexion.Receive(respuestaa)

        Catch ex As Exception

        End Try
    End Sub

    Public Sub mbescribir(ByVal primero As UShort, ByVal cuantos As UShort, ByVal valores() As UShort)
        Dim asd(13 + cuantos * 2) As Byte
        asd(0) = 0
        asd(1) = 0
        asd(2) = 0
        asd(3) = 0
        asd(4) = 0
        asd(5) = CByte(7 + cuantos * 2)
        asd(6) = CByte(1)
        asd(7) = CByte(16) '16
        asd(8) = CByte(primero \ 256)
        asd(9) = CByte(primero Mod 256)
        asd(10) = CByte(cuantos \ 256)
        asd(11) = CByte(cuantos Mod 256)
        asd(12) = CByte(cuantos * 2)
        For i As Integer = 0 To cuantos - 1
            asd(13 + i * 2) = CByte(valores(i) \ 256)
            asd(13 + i * 2 + 1) = CByte(valores(i) Mod 256)
        Next
        'asd(13) = CByte(1000 Mod 256)
        'asd(14) = CByte(1000 / 256)
        Try
            falloMb = False
            conexion.Send(asd)
            Dim respuesta(12) As Byte
            Dim nRecibidos As Integer = conexion.Receive(respuesta)
            If (respuesta(7) & Hex(80)) = Hex(80) Then
                Throw New Exception("Código de error enviado pol0 esclavo: 0x" + respuesta(7).ToString("X2"))
            End If
            If (respuesta(8) <> asd(8)) Or (respuesta(9) <> asd(9)) Then
                Throw New Exception("Dirección do primeiro rexistro incorrecto en resposta")
            End If
            If (respuesta(10) <> asd(10)) Or (respuesta(11) <> asd(11)) Then
                Throw New Exception("Número de rexistros incorrecto en resposta")
            End If
        Catch
            falloMb = True
        Finally

        End Try


    End Sub
#End Region

#Region "BBDD"
    Public conecta As MySqlConnection = New MySqlConnection
    Public Sub conectaBBDD()
        Try
            conecta = New MySqlConnection()

            conecta.ConnectionString = "server=127.0.0.1;" &
                                  "user id=root;" &
                                "password=;" &
                                 "port=3306;" &
                                "database=bd_conservas;"

            'conecta.ConnectionString = "server=" + Mb_BBDD.TextBox3.Text + ";" &
            '                     "user id=root;" &
            '                   "password=;" &
            '                  "port=3306;" &
            '               "database=bd_conservas;"


            conecta.Open()
            falloBBDD = False
        Catch
            If falloBBDD = False Then
                falloBBDD = True
                MsgBox("ERRO CONECTANDO A BASE DE DATOS")
            End If
        End Try
    End Sub
    Public Sub desconectaBBDD()
        conecta.Close()
        conecta.Dispose()
    End Sub
#End Region
End Module
