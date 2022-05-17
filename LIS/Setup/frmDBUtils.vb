Public Class frmDBUtils

    Public ID As String

    Private Sub btnBack_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnClose.Click
        Me.Close()
        Me.Dispose()
    End Sub

    Private Sub btnSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSave.Click
        If Me.txtServerName.Text = "" Or Me.txtRoot.Text = "" Or txtDBName.Text = "" Then
            MessageBox.Show("Please fill all required information.", "Empty String", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            Exit Sub
        End If

        My.Settings.Server = txtServerName.Text
        My.Settings.UID = txtRoot.Text
        My.Settings.PWD = txtPassword.Text
        My.Settings.DBName = txtDBName.Text
        My.Settings.MyConnectionString = "SERVER = " & txtServerName.Text &
            ";PORT = " & txtPort.Text &
            ";DATABASE =" & txtDBName.Text &
            ";UID = " & txtRoot.Text &
            ";PWD = " & txtPassword.Text & ";"
        My.Settings.Save()
        My.Settings.Reload()
        MessageBox.Show("Database settings saved. Please Restart the Application", "Saved. Please Restart the Application", MessageBoxButtons.OK, MessageBoxIcon.Information)
        Application.Restart()
    End Sub

    Private Sub frmDBUtils_FormClosed(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
        Me.Dispose()
    End Sub

    Private Sub frmDBUtils_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Me.txtServerName.Text = My.Settings.Server
        Me.txtDBName.Text = My.Settings.DBName
        Me.txtRoot.Text = My.Settings.UID
        Me.txtPassword.Text = My.Settings.PWD
        Me.txtPort.Text = My.Settings.TCPPort
    End Sub

    'Private Sub txtDBName_Click(sender As Object, e As EventArgs) Handles txtDBName.Click
    '    Try
    '        conn.ConnectionString = "SERVER = " & txtServerName.Text &
    '        ";DATABASE =" & txtDBName.Text &
    '        ";UID = " & txtRoot.Text &
    '        ";PWD = " & txtPassword.Text & ";"
    '        conn.Open()
    '        rs.Connection = conn
    '        rs.CommandText = "SELECT * FROM sys.DATABASES"
    '        reader = rs.ExecuteReader
    '        While reader.Read()
    '            txtDBName.Properties.Items.Add(reader(0).ToString)
    '        End While
    '        conn.Close()
    '    Catch ex As Exception
    '        MessageBox.Show(ex.Message)
    '        Exit Sub
    '    End Try
    'End Sub
End Class