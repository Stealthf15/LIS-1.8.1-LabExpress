Public Class frmDoctorAE

    Public ID As String

    Private Sub btnBack_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnClose.Click
        Me.Close()
    End Sub

    Private Sub btnSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSave.Click
        If Me.txtFName.Text = "" Then
            MessageBox.Show("Please fill all required information.", "Empty String", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            Exit Sub
        End If

        rs.Parameters.Clear()
        rs.Parameters.AddWithValue("@id", ID)
        rs.Parameters.AddWithValue("@name", txtFName.Text)

        If Me.btnSave.Text = "&Save" Then
            SaveRecord("INSERT INTO `requestor` (`name`) VALUES (@name)")
            frmDoctor.LoadRecords()
            txtFName.Text = ""
        Else
            UpdateRecord("UPDATE `requestor` SET `name` = @name WHERE `id` = @id")
            frmDoctor.LoadRecords()
            Me.Close()
        End If
    End Sub

    Private Sub frmDoctorAE_FormClosed(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
        Me.Dispose()
    End Sub

End Class