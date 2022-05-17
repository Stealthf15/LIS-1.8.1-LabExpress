Public Class frmShiftAE
    Public _ID As String

    Private Sub btnSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSave.Click
        If Me.txtShiftName.Text = "" Or Me.dtStartTime.Text = "" Then
            MessageBox.Show("Please Fill up all the information above.", "Field is Empty", MessageBoxButtons.OK, MessageBoxIcon.Asterisk)
        End If
        rs.Parameters.Clear()
        rs.Parameters.AddWithValue("@ID", _ID)
        rs.Parameters.AddWithValue("@shift", txtShiftName.Text)
        rs.Parameters.AddWithValue("@starttime", dtStartTime.Value)
        rs.Parameters.AddWithValue("@endtime", dtEndTime.Value)

        If Me.btnSave.Text = "&Save" Then
            SaveRecord("INSERT INTO `shift` (`startTime`, `endTime`, `shift`) VALUES (@starttime, @endtime, @shift)")
        Else
            UpdateRecord("UPDATE `shift` SET `shift` = @shift, `startTime` = @startTime, `endTime` = @endTime WHERE `id` = @ID")
        End If
        frmShift.loadRecords()
        txtShiftName.Text = ""
        dtStartTime.Text = ""
        Me.Close()
    End Sub

    Private Sub frmShiftAE_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        dtStartTime.ShowUpDown = False
    End Sub

    Private Sub btnClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnClose.Click
        Me.Close()
    End Sub
End Class