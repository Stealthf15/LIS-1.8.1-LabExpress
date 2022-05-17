Public Class frmSettings

    Public name, id As String

    Private Sub btnBack_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnClose.Click
        Me.Close()
        Me.Dispose()
    End Sub

    Private Sub btnViewReference_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnViewReference.Click
        frmRangeList.ShowDialog()
    End Sub

    Private Sub SimpleButton1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles SimpleButton1.Click
        Me.Close()
        Me.Dispose()
    End Sub
End Class