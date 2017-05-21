Imports System.Data
Imports MySql.Data.MySqlClient
Imports MySql.Data
Imports System.Threading

Module globales
    Public findelote As Boolean
    Public novoRexistro As Boolean
    Public loteActivo As UInteger
    Public loteActual As UInteger
    Public lectura(15) As UShort
    Public lectura2(10) As UShort
    Public respuesta(255) As Integer
    Public consulta As New MySqlDataAdapter
    Public inserta As New MySqlCommand
    Public lock As New Object
    Public bloquea As New Object
    Public com, fail As Thread

    Public Function buscaLote() As UInteger
        Dim lote As New DataTable
        Try
            conectaBBDD()

            Dim fc As New MySqlDataAdapter("SELECT NumLote FROM lote WHERE estado=1", comunicacions.conecta)
            fc.SelectCommand.CommandType = CommandType.Text
            fc.Fill(lote)
        Catch

        End Try
        Try
            loteActivo = lote.Rows(0).Item("NumLote").ToString.Trim
            Dim text As String = ("UPDATE `lote` SET `Estado`=1 WHERE `NumLote`='" + loteActivo.ToString.Trim + "'")
            inserta.Connection = comunicacions.conecta
            inserta.CommandText = text
            inserta.ExecuteNonQuery()
        Catch
            Try
                Dim dc As New MySqlDataAdapter("SELECT NumLote FROM lote WHERE estado=0 and cola=1 order by prioridad desc limit 1", comunicacions.conecta)
                dc.SelectCommand.CommandType = CommandType.Text
                dc.Fill(lote)
            Catch
            End Try
            Try
                loteActivo = lote.Rows(0).Item("NumLote").ToString.Trim
                Dim text As String = ("UPDATE `lote` SET `Estado`=1 WHERE `NumLote`='" + loteActivo.ToString.Trim + "'")
                inserta.Connection = comunicacions.conecta
                inserta.CommandText = text
                inserta.ExecuteNonQuery()
            Catch
                'SetText("NON HAI LOTES EN COLA")
                'TextBox1.Text = "NON HAI LOTES EN COLA"
                hmi.TextBox1.Text = "NON HAI LOTES EN COLA"

            End Try
        Finally
            desconectaBBDD()
        End Try
        lote.Clear()
        Return loteActivo
    End Function

    Public Sub Rexistro()
        Dim tarxeta, tempo, seg, min, hh As UShort
        Dim loteCaixa As UInteger
        Dim pesoEntrada, pesoSalida As Double
        Dim fallo As Boolean = False
        'SyncLock lock
        mbler(1, 8, lectura)
        'End SyncLock
        tempo = lectura(1)

        Dim ss As UInt32                                'pasase hora actual a segundos e restaselle o tempo
        Dim fch As Date = Now
        Dim horaF, horaI, fecha, PE, PS, tempoS As String
        Try
            seg = Format(fch, "ss")
            min = Format(fch, "mm")
            hh = Format(fch, "HH")
            ss = (hh * 3600 + min * 60 + seg) - tempo
            hh = ss \ 3600
            min = ss Mod 3600
            seg = min Mod 60
            min = min \ 60
            fallo = False
        Catch
            fallo = True
        End Try

        horaI = CDate(String.Format("{00}:{01}:{02}", hh, min, seg))
        horaF = CDate(Format(fch, "HH:mm:ss"))
        fecha = Format(fch, "yyyy-MM-dd")
        pesoEntrada = (CInt(lectura(2) * 65536 + lectura(3))) / 1000    'SE E NECESARIO, PASAR A Kg Resoucion de 2 cifras despois de .
        pesoSalida = (CInt(lectura(4) * 65536 + lectura(5))) / 1000
        loteCaixa = CInt(lectura(6) * 65536 + lectura(7))
        tarxeta = lectura(0)
        PE = pesoEntrada.ToString("0.000", System.Globalization.CultureInfo.InvariantCulture)
        PS = pesoSalida.ToString("0.000", System.Globalization.CultureInfo.InvariantCulture)
        tempoS = tempo.ToString


        If fallo = False Then
            Dim p1 As Boolean
            SyncLock bloquea
                Try
                    conectaBBDD()
                    Dim datos As New DataTable

                    Dim da As New MySqlDataAdapter("select IdOperario from tarjeta where Numero='" + tarxeta.ToString.Trim + "'", comunicacions.conecta)
                    da.SelectCommand.CommandType = CommandType.Text
                    da.Fill(datos)
                    Dim idemp As String = datos.Rows(0).Item("IdOperario")

                    Dim dbs As New MySqlDataAdapter("SELECT IdLote FROM lote WHERE NumLote='" + loteCaixa.ToString.Trim + "'", conecta)
                    dbs.SelectCommand.CommandType = CommandType.Text
                    dbs.Fill(datos)
                    Dim IdLot As String = datos.Rows(1).Item("IdLote")

                    Dim text As String = ("INSERT INTO cajas (Fecha, IdOperario, IdLote, PesoInicial, PesoFinal, HoraInicial, HoraFinal, Tiempo) values ('" + fecha + "','" + idemp + "','" + IdLot.Trim + "','" + PE.Trim + "','" + PS.Trim + "','" + horaI + "','" + horaF + "','" + tempoS + "')")
                    inserta.Connection = comunicacions.conecta
                    inserta.CommandText = text
                    inserta.ExecuteNonQuery()
                    datos.Clear()
                    p1 = False
                Catch
                    If p1 = False Then
                        MsgBox("Erro ao gardar BBDD")
                        p1 = True
                    End If

                Finally
                    desconectaBBDD()
                End Try
            End SyncLock
        End If
        'lectura(0) = 0
        'SyncLock lock
        'mbescribir(13, 1, lectura)
        'End SyncLock
        mbescribe0(13)
        novoRexistro = False
    End Sub



    Public Sub finalizaLote()
        Dim primeraLote As Boolean
        SyncLock bloquea
            Try
                conectaBBDD()
                Dim text As String = ("UPDATE `lote` SET `Estado`=2,`Cola`=0,`Prioridad`=0 WHERE `NumLote`='" + loteActivo.ToString.Trim + "'")
                inserta.Connection = comunicacions.conecta
                inserta.CommandText = text
                inserta.ExecuteNonQuery()
                loteActivo = 0
                lectura2(0) = 0
                lectura2(1) = 0
                'SyncLock lock
                mbescribir(9, 2, lectura2) 'ESCRIBE 0 EN LOTEACTIVO

                mbescribe0(14)
                'lectura2(0) = 0
                'mbescribir(14, 1, lectura2) 'CONFIRMACION FIN DE LOTE
                'End SyncLock
                primeraLote = False
            Catch
                If (primeraLote = False) Then
                    MsgBox("Fallo finalizando lote")
                    primeraLote = True
                End If
            End Try
        End SyncLock
        findelote = False
    End Sub
End Module
