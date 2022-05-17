Public Class frmUnit

    Public Sub LoadRecords()
        Try
            LoadRecordsOnLV(lvList, "units", 1, "unit")
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
            SearchRecordsOnLV(lvList, "units", 1, "unit", txtSearch.Text, "id")
        Catch ex As Exception
            MessageBox.Show(ex.Message, "System Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub btnDelete_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnDelete.ItemClick
        Try
            If lvList.SelectedItems.Count > 0 Then
                DeleteRecord("units", "id", lvList.FocusedItem.SubItems(0).Text)
                LoadRecords()
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message, "System Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub btnEdit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnEdit.ItemClick
        If lvList.SelectedItems.Count > 0 Then
            frmUnitAE.ID = lvList.FocusedItem.SubItems(0).Text
            frmUnitAE.txtChannel.Text = lvList.FocusedItem.SubItems(0).Text
            frmUnitAE.txtTestName.Text = lvList.FocusedItem.SubItems(1).Text
            frmUnitAE.btnSave.Text = "&Update"
            frmUnitAE.ShowDialog()
        End If
    End Sub

    Private Sub btnRefresh_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnRefresh.ItemClick
        LoadRecords()
    End Sub

    Private Sub btnAdd_ItemClick(ByVal sender As Object, ByVal e As DevExpress.XtraBars.ItemClickEventArgs) Handles btnAdd.ItemClick
        frmUnitAE.btnSave.Text = "&Save"
        frmUnitAE.ShowDialog()
    End Sub
End Class