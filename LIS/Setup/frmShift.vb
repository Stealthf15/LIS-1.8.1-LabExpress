Public Class frmShift

    Public Sub loadRecords()
        Try
            LoadRecordsOnLV(lvList, "shift", 3, "shift")
        Catch ex As Exception
            MessageBox.Show(ex.Message, "System Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub
    Private Sub frmShift_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        loadRecords()
    End Sub

    Private Sub btnAdd_ItemClick(ByVal sender As System.Object, ByVal e As DevExpress.XtraBars.ItemClickEventArgs) Handles btnAdd.ItemClick
        frmShiftAE.btnSave.Text = "&Save"
        frmShiftAE.ShowDialog()
    End Sub

    Private Sub btnDelete_ItemClick(ByVal sender As System.Object, ByVal e As DevExpress.XtraBars.ItemClickEventArgs) Handles btnDelete.ItemClick
        Try
            DeleteRecord("shift", "id", lvList.FocusedItem.SubItems(0).Text)
            loadRecords()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "System Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub btnRefresh_ItemClick(ByVal sender As System.Object, ByVal e As DevExpress.XtraBars.ItemClickEventArgs) Handles btnRefresh.ItemClick
        loadRecords()
    End Sub

    Private Sub btnEdit_ItemClick(ByVal sender As System.Object, ByVal e As DevExpress.XtraBars.ItemClickEventArgs) Handles btnEdit.ItemClick
        If lvList.SelectedItems.Count > 0 Then
            frmShiftAE._ID = lvList.FocusedItem.SubItems(0).Text
            frmShiftAE.txtShiftName.Text = lvList.FocusedItem.SubItems(1).Text
            frmShiftAE.dtStartTime.Text = lvList.FocusedItem.SubItems(2).Text
            frmShiftAE.dtEndTime.Text = lvList.FocusedItem.SubItems(3).Text
            frmShiftAE.btnSave.Text = "&Update"
            frmShiftAE.ShowDialog()
        End If
    End Sub

    Private Sub btnClose_ItemClick(ByVal sender As System.Object, ByVal e As DevExpress.XtraBars.ItemClickEventArgs) Handles btnClose.ItemClick
        Me.Close()
        Me.Dispose()
    End Sub
End Class