Public Class frmSpecimenPAE

    Public id As String

    Private Sub btnBack_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnClose.Click
        Me.Close()
    End Sub

    Private Sub btnSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSave.Click
        If Me.txtPackageName.Text = "" Or Me.cboTestName.Text = "" Then
            MessageBox.Show("Please fill all required information.", "Empty String", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            Exit Sub
        End If

        rs.Parameters.Clear()
        rs.Parameters.AddWithValue("@ID", id)
        rs.Parameters.AddWithValue("@Package", txtPackageName.Text)
        rs.Parameters.AddWithValue("@HISCode", txtHISCode.Text)
        rs.Parameters.AddWithValue("@Section", cboTestName.Text)
        rs.Parameters.AddWithValue("@Status", "1")

        If Me.btnSave.Text = "&Save" Then
            SaveRecord("INSERT INTO `packages` (`packagename`, `packagecode`, `section`, `subsection`, `display`) VALUES (@Package, @HISCode, @Section, @HISCode, @Status)")
            frmSpecimen.LoadRecords()
            Me.Close()
        Else
            UpdateRecord("UPDATE `packages` SET `packagename` = @Package, `packagecode` = @HISCode, `section` = @Section, `subsection` = @HISCode WHERE `id` = @ID")
            frmSpecimen.LoadRecords()
            Me.Close()
        End If
    End Sub

    Private Sub txtChannel_KeyPress(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyPressEventArgs)
        If InStr("1234567890.", e.KeyChar) = 0 And Not Chr(AscW(e.KeyChar)) = vbBack Then
            e.KeyChar = ChrW(0)
        End If
    End Sub

    Private Sub frmTestTypeAE_FormClosed(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
        Me.Dispose()
    End Sub

End Class