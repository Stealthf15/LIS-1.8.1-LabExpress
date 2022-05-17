Public Class frmUserAE

    Public ID As String
    Private medtechName As String
    Dim Verificator As Integer = 0

    Private Sub btnBack_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnClose.Click
        Me.Close()
    End Sub

    Private Sub btnSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSave.Click
        If Me.txtFName.Text = "" Or Me.txtLName.Text = "" Then
            MessageBox.Show("Please fill all required information.", "Empty String", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            Exit Sub
        End If

        rs.Parameters.Clear()
        rs.Parameters.AddWithValue("@id", ID)
        rs.Parameters.AddWithValue("@fname", txtFName.Text)
        rs.Parameters.AddWithValue("@mname", txtMName.Text)
        rs.Parameters.AddWithValue("@lname", txtLName.Text)
        rs.Parameters.AddWithValue("@license", txtLicense.Text)
        rs.Parameters.AddWithValue("@designation", txtDesign.Text)

        If Me.btnSave.Text = "&Save" Then
            SaveRecord("INSERT INTO `users` (`fname`, `mname`, `lname`, `license`, `designation`) VALUES (@fname, @mname, @lname, @license, @designation)")
            Me.Close()
            frmMedTech.LoadRecords()
        Else
            Try
                medtechName = txtFName.Text + " " + txtMName.Text + " " + txtLName.Text + ", " + txtDesign.Text
                UpdateRecord("UPDATE `users` SET `fname` = @fname, `mname` = @mname, `lname` = @lname, `license` = @license, `designation` = @designation WHERE `id` = @id")
                UpdateRecordwthoutMSG("UPDATE `user_account` SET `name` = '" & medtechName & "' WHERE `email` = @id")
                Me.Close()
                frmMedTech.LoadRecords()
            Catch ex As Exception
                MessageBox.Show(ex.ToString, "Error", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            End Try
        End If
    End Sub

    Private Sub frmMedTechAE_FormClosed(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
        Me.Dispose()
    End Sub

End Class