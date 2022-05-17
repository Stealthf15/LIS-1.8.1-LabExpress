Public Class frmDepartmentAE

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
        rs.Parameters.AddWithValue("@name", txtTestName.Text)

        If Me.btnSave.Text = "&Save" Then
            SaveRecord("INSERT INTO `department` (`dept_name`) VALUES (@name)")
            Me.Close()
            frmDeparment.LoadRecords()
        Else
            UpdateRecord("UPDATE `department` SET `dept_name` = @name WHERE `id` = @id")
            Me.Close()
            frmDeparment.LoadRecords()
        End If
    End Sub
End Class