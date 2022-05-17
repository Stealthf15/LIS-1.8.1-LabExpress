Public Class frmPatient

    Public Sub LoadRecords()
        Try
            LoadRecordsOnLV(lvList, "`patient_info`", 8, "patient_id")
        Catch ex As Exception
            MessageBox.Show(ex.Message, "System Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub frmPatient_FormClosed(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
        Me.Dispose()
    End Sub

    Private Sub frmNewOrder_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        LoadRecords()
    End Sub

    Private Sub btnBack_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnClose.ItemClick
        Me.Close()
    End Sub

    Private Sub SimpleButton2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnEdit.ItemClick
        If Me.lvList.SelectedItems.Count > 0 Then
            frmPatientAE.ID = lvList.FocusedItem.SubItems(0).Text
            frmPatientAE.txtPatientID.Text = lvList.FocusedItem.SubItems(0).Text
            frmPatientAE.txtFName.Text = lvList.FocusedItem.SubItems(1).Text
            frmPatientAE.cboSex.Text = lvList.FocusedItem.SubItems(2).Text
            frmPatientAE.dtBDate.Text = lvList.FocusedItem.SubItems(3).Text
            frmPatientAE.txtAge.Text = lvList.FocusedItem.SubItems(4).Text
            frmPatientAE.cboAgeType.Text = lvList.FocusedItem.SubItems(5).Text
            frmPatientAE.txtAddress.Text = lvList.FocusedItem.SubItems(6).Text
            frmPatientAE.txtContact.Text = lvList.FocusedItem.SubItems(7).Text
            frmPatientAE.btnSave.Text = "&Update"
            frmPatientAE.ShowDialog()
        End If
    End Sub

    Private Sub btnDelete_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnDelete.ItemClick
        If Me.lvList.SelectedItems.Count > 0 Then
            DeleteRecord("patient_info", "patient_id", lvList.FocusedItem.SubItems(0).Text)
            LoadRecords()
        End If
    End Sub

    Private Sub btnRefresh_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnRefresh.ItemClick
        LoadRecords()
    End Sub

    Private Sub txtSearch_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtSearch.TextChanged
        SearchRecordsOnLV(lvList, "`patient_info`", 8, "name", Me.txtSearch.Text, "patient_id")
    End Sub

    Private Sub bntAdd_ItemClick(ByVal sender As Object, ByVal e As DevExpress.XtraBars.ItemClickEventArgs) Handles bntAdd.ItemClick
        frmPatientAE.btnSave.Text = "&Save"
        frmPatientAE.ShowDialog()
    End Sub

    Private Sub frm_Activated(sender As Object, e As EventArgs) Handles Me.Activated
        'MainFOrm.aceFecal.Appearance.Normal.BackColor = Color.FromArgb(6, 31, 71)
        MainFOrm.btnPatient.Appearance.Normal.BackColor = Color.FromArgb(16, 110, 190)
        MainFOrm.btnPatient.Appearance.Normal.ForeColor = Color.FromArgb(255, 255, 255)
    End Sub

    Private Sub frm_FormClosed(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
        MainFOrm.lblTitle.Text = ""
        MainFOrm.btnPatient.Appearance.Normal.BackColor = Color.FromArgb(240, 240, 240)
        MainFOrm.btnPatient.Appearance.Normal.ForeColor = Color.FromArgb(27, 41, 62)
        Me.Dispose()
    End Sub

End Class