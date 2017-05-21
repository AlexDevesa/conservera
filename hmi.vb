Imports System
Imports System.Net.Sockets
Imports System.Net
Imports System.Data
Imports MySql.Data.MySqlClient
Imports MySql.Data
Imports System.Threading

Public Class hmi
#Region "variables"
    Dim finalAPP, correcto As Boolean
#End Region


    Private Sub Mb_BBDD_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Me.CenterToScreen()
        conectaMb()
        reloj.Enabled = True
        com = New Thread(AddressOf Me.comunica)
        loteActual = 0
        com.Start()
        finalAPP = False
    End Sub

    Private Sub Timer1_Tick(sender As Object, e As EventArgs) Handles reloj.Tick

        comprobaConexion()
        If estadoSocket = True Then
            Dim loteAct As ULong
            'SyncLock lock
            mbler(9, 2, lectura2)
            'End SyncLock
            loteAct = lectura2(0) * 65536 + lectura2(1)

            If loteAct < 1 Then
                SyncLock bloquea
                    buscaLote()
                End SyncLock
                lectura2(0) = loteActivo \ 65536
                lectura2(1) = loteActivo Mod 65536
                'SyncLock lock
                mbescribir(9, 2, lectura2)
                'End SyncLock
            Else
                TextBox1.Text = loteAct.ToString
                loteActivo = loteAct
            End If

            '__________________________________________--------------------_______________________________
            'SyncLock lock
            mbler(13, 2, lectura) 'lee words ESTADOS
            'End SyncLock
            If (lectura(0) > 0) Then  'NOVO REXISTRO
                novoRexistro = True
            End If
            If (lectura(1) > 0) Then  'FIN DE LOTE
                findelote = True
            End If

            If novoRexistro Then
                Rexistro()
            End If

            If findelote = True Then
                finalizaLote()
            End If
        End If
        '____________________________________________---------------------------______________________________

    End Sub

    Sub comunica()
        Do Until finalAPP = True
            mbcomunica()
            Thread.Sleep(2333)
        Loop
    End Sub


    Private Sub Mb_BBDD_FormClosed(sender As Object, e As FormClosedEventArgs) Handles MyBase.FormClosed
        com.Abort()
    End Sub

    Private Sub Mb_BBDD_FormClosing(sender As Object, e As FormClosingEventArgs) Handles MyBase.FormClosing
        finalAPP = True
    End Sub


    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Application.ExitThread()
        Application.Exit()
    End Sub
End Class

