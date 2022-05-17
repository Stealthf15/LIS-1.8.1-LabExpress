Public Class frmTestName

    Public Sub LoadRecords()
        Try
            LoadRecordsOnLV(lvList, "test_name", 1, "test_name")
        Catch ex As Exception
            MessageBox.Show(ex.Message, "System Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub frmMachineList_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        LoadRecords()
    End Sub

    Private Sub SimpleButton2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnClose.ItemClick
        Me.Close()
        Me.Dispose()
    End Sub

    Private Sub txtSearch_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtSearch.TextChanged
        Try
            SearchRecordsOnLV(lvList, "test_name", 1, "test_name", txtSearch.Text, "test_name")
        Catch ex As Exception
            MessageBox.Show(ex.Message, "System Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub btnDelete_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnDelete.ItemClick
        Try
            If lvList.SelectedItems.Count > 0 Then
                DeleteRecord("test_name", "channel", lvList.FocusedItem.SubItems(0).Text)
                LoadRecords()
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message, "System Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub btnEdit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnEdit.ItemClick
        If lvList.SelectedItems.Count > 0 Then
            frmTestNameAE.ID = lvList.FocusedItem.SubItems(0).Text
            frmTestNameAE.txtChannel.Text = lvList.FocusedItem.SubItems(0).Text
            frmTestNameAE.txtTestName.Text = lvList.FocusedItem.SubItems(1).Text
            frmTestNameAE.btnSave.Text = "&Update"
            frmTestNameAE.ShowDialog()
        End If
    End Sub

    Private Sub btnRefresh_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnRefresh.ItemClick
        LoadRecords()
    End Sub

    Private Sub btnAdd_ItemClick(ByVal sender As Object, ByVal e As DevExpress.XtraBars.ItemClickEventArgs) Handles btnAdd.ItemClick
        frmTestNameAE.btnSave.Text = "&Save"
        frmTestNameAE.ShowDialog()
    End Sub
End Class