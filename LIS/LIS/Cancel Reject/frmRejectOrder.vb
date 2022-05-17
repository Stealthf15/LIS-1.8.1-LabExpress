Public Class frmRejectOrder

    Public MainID As String
    Public ID As String
    Public sID As String
    Public pID As String
    Public pName As String
    Public pTest As String
    Public pSection As String
    Public pSubSection As String

    Private Sub btnSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSave.Click
        rs.Parameters.Clear()
        rs.Parameters.AddWithValue("@ID", ID)
        Connect()
        rs.Connection = conn
        rs.CommandType = CommandType.Text
        rs.CommandText = "UPDATE `tmpworklist` SET `status` = 'Rejected' WHERE `id` = @ID"
        rs.ExecuteNonQuery()
        Disconnect()

        'Log activity
        SpecimenActivity("z_logs_specimen", sID, pID, pName, CurrUser, "Specimen Rejected", txtComment.Text, pTest, pSection, pSubSection)

        Me.Close()

        MessageBox.Show("Patient order successfully rejected.", "Reject Sample", MessageBoxButtons.OK, MessageBoxIcon.Information)
    End Sub

    Private Sub frmRejectOrder_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Me.txtComment.Focus()
    End Sub

    Private Sub frmRejectOrder_FormClosed(sender As Object, e As FormClosedEventArgs) Handles Me.FormClosed
        Me.Dispose()
    End Sub
End Class