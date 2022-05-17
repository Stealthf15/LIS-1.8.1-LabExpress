Public Class frmCancelOrder

    Public MainID As String

    Private Sub btnSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSave.Click
        rs.Parameters.Clear()
        rs.Parameters.AddWithValue("@mainID", MainID)
        rs.Parameters.AddWithValue("@comment", txtComment.Text)

        SaveRecordwthoutMSG("INSERT INTO `lab_comment` (`comment`, `sample_id`) VALUES (@comment, @mainID)")

        UpdateRecordwthoutMSG("UPDATE `tmpWorklist` SET `status` = 'Canceled' WHERE main_id LIKE @mainID")

        Me.Close()
        Me.Dispose()
        frmChemWorklist.LoadRecords()

    End Sub

    Private Sub frmChangeUser_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Me.txtComment.Focus()
    End Sub
End Class