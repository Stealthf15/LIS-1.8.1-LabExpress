Public Class frmControlSetting

    Public mainID As String = ""

    Public Sub LoadRecords()
        Try
            LoadRecordsOnLVSQL(lvList, "SELECT `id`, `ctrl_id`, `test_type`, `type`, `lot_no`, `control_name`, `instrument`, `entry_date` FROM `control_setting`", 7)
        Catch ex As Exception
            MessageBox.Show(ex.Message, "System Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub frmNewOrder_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        LoadRecords()
    End Sub

    Private Sub btnRefresh_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnRefresh.ItemClick
        LoadRecords()
    End Sub

    Private Sub btnClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnClose.ItemClick
        Me.Close()
        Me.Dispose()
    End Sub

    Private Sub txtSearch_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtSearch.TextChanged
        Try
            Connect()
            rs.Connection = conn
            rs.CommandText = "SELECT `id`, `ctrl_id`, `test_type`, `type`, `lot_no`, `control_name`, `instrument`, `entry_date` FROM `control_setting`"
            reader = rs.ExecuteReader
            While reader.Read
                iItem = New ListViewItem(reader(0).ToString, 0)
                iItem.Checked = True
                For x As Integer = 1 To 7 Step 1
                    iItem.SubItems.Add(reader(x).ToString())
                Next
                lvList.Items.Add(iItem)
            End While
            Disconnect()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "System Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub btnDelete_ItemClick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnDelete.ItemClick
        If lvList.SelectedItems.Count > 0 Then
            DeleteRecord("`control_setting`", "id", Me.lvList.FocusedItem.SubItems(0).Text)
            LoadRecords()
        End If
    End Sub

    Private Sub btnAdd_ItemClick(ByVal sender As Object, ByVal e As DevExpress.XtraBars.ItemClickEventArgs) Handles btnAdd.ItemClick
        frmControlSettingAE.btnSave.Text = "&Save"
        frmControlSettingAE.ShowDialog()
    End Sub

    Private Sub btnTargetAdd_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnTargetAdd.Click
        If lvList.SelectedItems.Count > 0 Then
            frmControlTargetAE.txtControlID.Text = lvList.FocusedItem.SubItems(1).Text
            frmControlTargetAE.txtTestType.Text = lvList.FocusedItem.SubItems(2).Text
            frmControlTargetAE.cboInstrument.Text = lvList.FocusedItem.SubItems(6).Text
            frmControlTargetAE.btnSave.Text = "&Save"
            frmControlTargetAE.ShowDialog()
        End If
    End Sub


    Private Sub lvList_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles lvList.SelectedIndexChanged
        rs.Parameters.Clear()
        rs.Parameters.AddWithValue("@ControlID", lvList.FocusedItem.SubItems(1).Text)
        rs.Parameters.AddWithValue("@Instrument", lvList.FocusedItem.SubItems(6).Text)

        LoadRecordsOnLVSQL(lvTarget, "SELECT `id`, `test_name`, `test_code`, `target`, `sd`, `ll`, `ul`, `plus_one`, `minus_one`, `plus_three`, `minus_three` FROM `control_target` WHERE `control_id` LIKE @ControlID AND `instrument` LIKE @Instrument", 10)
    End Sub

    Private Sub btnTargetEdit_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnTargetEdit.Click
        If lvTarget.SelectedItems.Count > 0 Then
            frmControlTargetAE.txtControlID.Text = lvList.FocusedItem.SubItems(1).Text
            frmControlTargetAE.txtTestType.Text = lvList.FocusedItem.SubItems(2).Text
            frmControlTargetAE.cboInstrument.Text = lvList.FocusedItem.SubItems(6).Text
            frmControlTargetAE.ID = lvTarget.FocusedItem.SubItems(0).Text
            frmControlTargetAE.cboSpecimen.Text = lvTarget.FocusedItem.SubItems(1).Text
            frmControlTargetAE.txtTestCode.Text = lvTarget.FocusedItem.SubItems(2).Text
            frmControlTargetAE.txtTarget.Text = lvTarget.FocusedItem.SubItems(3).Text
            frmControlTargetAE.txtTolerance.Text = lvTarget.FocusedItem.SubItems(4).Text
            frmControlTargetAE.txtLSD.Text = lvTarget.FocusedItem.SubItems(5).Text
            frmControlTargetAE.txtUSD.Text = lvTarget.FocusedItem.SubItems(6).Text
            frmControlTargetAE.txtP1.Text = lvTarget.FocusedItem.SubItems(7).Text
            frmControlTargetAE.txtN1.Text = lvTarget.FocusedItem.SubItems(8).Text
            frmControlTargetAE.txtP3.Text = lvTarget.FocusedItem.SubItems(9).Text
            frmControlTargetAE.txtN3.Text = lvTarget.FocusedItem.SubItems(10).Text
            frmControlTargetAE.btnSave.Text = "&Update"
            frmControlTargetAE.ShowDialog()
        End If
    End Sub


End Class