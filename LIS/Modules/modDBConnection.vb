Imports System.IO
Imports System.IO.Ports
Imports System.Threading
Imports System.Threading.Thread
Imports MySql.Data.MySqlClient
Imports Microsoft.Reporting.WinForms
Imports System.Data.OleDb

Module modDBConnection

#Region "MySQL Connection Settings"
    Public Sub Connect()
        Try
            conn.ConnectionString = MyConnectionString
            conn.Open()
        Catch ex As MySqlException
            MessageBox.Show(ex.Message, "System Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            conn.Close()
            Exit Sub
        End Try
    End Sub

    Public Sub Disconnect()
        If conn.State = ConnectionState.Open Then
            conn.Close()
        End If
    End Sub

    Public Sub Connect1()
        Try
            conn1.ConnectionString = MyConnectionString
            conn1.Open()
        Catch ex As MySqlException
            MessageBox.Show(ex.Message, "System Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            conn1.Close()
            Exit Sub
        End Try
    End Sub

    Public Sub Disconnect1()
        If conn1.State = ConnectionState.Open Then
            conn1.Close()
        End If
    End Sub

    Public Sub Connect2()
        Try
            conn2.ConnectionString = MyConnectionString
            conn2.Open()
        Catch ex As MySqlException
            MessageBox.Show(ex.Message, "System Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            conn2.Close()
            Exit Sub
        End Try
    End Sub

    Public Sub Disconnect2()
        If conn2.State = ConnectionState.Open Then
            conn2.Close()
        End If
    End Sub

    Public Sub Connect3()
        Try
            conn3.ConnectionString = MyConnectionString
            conn3.Open()
        Catch ex As MySqlException
            MessageBox.Show(ex.Message, "System Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            conn3.Close()
            Exit Sub
        End Try
    End Sub

    Public Sub Disconnect3()
        If conn3.State = ConnectionState.Open Then
            conn3.Close()
        End If
    End Sub

    Public Sub ConnectSQL()
        Try
            connSQL.ConnectionString = MyConnectionStringSQL
            connSQL.Open()
        Catch ex As MySqlException
            MessageBox.Show(ex.Message, "System Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            connSQL.Close()
            Exit Sub
        End Try
    End Sub

    Public Sub DisconnectSQL()
        If connSQL.State = ConnectionState.Open Then
            connSQL.Close()
        End If
    End Sub
#End Region


End Module
