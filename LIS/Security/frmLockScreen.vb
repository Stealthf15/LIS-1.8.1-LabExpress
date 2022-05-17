Public Class frmLockScreen

    Public ID As String

    Private Sub btnSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSave.Click
        Try
            Connect()
            rs.Connection = conn
            rs.CommandText = "SELECT * FROM `user_account` WHERE `Username` = '" & Me.txtUsername.Text & "' AND `password` = '" & txtPassword.Text & "' and `user_id` LIKE '" & CurrID & "'"
            reader = rs.ExecuteReader
            reader.Read()
            If (reader.HasRows = True) Then
                If (reader("username") = Me.txtUsername.Text) Then
                    CurrID = reader(0).ToString
                    CurrUser = reader(1).ToString
                    CurrPost = reader(7).ToString
                    CurrDept = reader(6).ToString
                    CurrType = reader(4).ToString
                    CurrEmail = reader(5).ToString

                    frmMain.WindowState = FormWindowState.Maximized
                    frmMain.Show()
                    Me.Hide()
                    txtPassword.Text = ""
                    txtUsername.Text = ""
                Else
                    MessageBox.Show("Invalid Username or Password.", "Authentication Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    Disconnect()
                    Exit Sub
                End If
            Else
                MessageBox.Show("Invalid Username or Password.", "Authentication Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Disconnect()
                Exit Sub
            End If
            Disconnect()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "System Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Exit Sub
        End Try
    End Sub

End Class