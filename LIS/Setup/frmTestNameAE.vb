Public Class frmTestNameAE

    Public ID As String

    Private Sub btnBack_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnClose.Click
        Me.Close()
    End Sub

    Private Sub btnSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSave.Click
        If Me.txtTestName.Text = "" Then
            MessageBox.Show("Please fill all required information.", "Empty String", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            Exit Sub
        End If
        rs.Parameters.Clear()
        rs.Parameters.AddWithValue("@id", ID)
        rs.Parameters.AddWithValue("@channel", txtChannel.Text)
        rs.Parameters.AddWithValue("@name", txtTestName.Text)

        If Me.btnSave.Text = "&Save" Then
            SaveRecord("INSERT INTO `test_name` (`channel`, `test_name`) VALUES (@channel, @name)")
            Me.Close()
            frmTestName.LoadRecords()
        Else
            UpdateRecord("UPDATE `test_name` SET `channel` = @channel, `test_name` = @name WHERE `channel` = @id")
            Me.Close()
            frmTestName.LoadRecords()
        End If
    End Sub

    Private Sub frmTestAE_FormClosed(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
        Me.Dispose()
    End Sub
End Class