Public Class frmRangeList

    Public name, id As String

    Public Sub LoadRecords()
        Try
            rs.Parameters.Clear()
            rs.Parameters.AddWithValue("@section", cboLimit.Text)
            rs.Parameters.AddWithValue("@machine", cboMachine.Text)

            LoadRecordsOnLVSQL(lvList, "SELECT `id`, `test_name`, `si_range`, `conv_range`, `low_value`, `high_value`, `gender`, `classification`, `age_begin`, `age_end`, `test_code` FROM `reference_range` WHERE (`section` = @section AND `machine` = @machine) ORDER BY `id`", 10)
        Catch ex As Exception
            MessageBox.Show(ex.Message, "System Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub btnBack_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnClose.ItemClick
        Me.Close()
        Me.Dispose()
    End Sub

    Private Sub frmRangeList_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        cboLimit.Properties.Items.Clear()
        cboMachine.Properties.Items.Clear()

        Connect()
        rs.Connection = conn
        rs.CommandText = "SELECT `test_name` FROM `testtype` WHERE test_name NOT LIKE 'All'"
        reader = rs.ExecuteReader
        While reader.Read()
            cboLimit.Properties.Items.Add(reader(0).ToString)
        End While
        Disconnect()

        Connect()
        rs.Connection = conn
        rs.CommandText = "SELECT `name`, `serial_no` FROM `machines`"
        reader = rs.ExecuteReader
        While reader.Read()
            cboMachine.Properties.Items.Add(reader(0).ToString & "_" & reader(1).ToString)
        End While
        Disconnect()
    End Sub

    Private Sub cboLimit_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cboLimit.SelectedIndexChanged, cboMachine.SelectedIndexChanged
        LoadRecords()
    End Sub

    Private Sub btnDelete_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnDelete.ItemClick
        Try
            DeleteRecord("`reference_range`", "id", lvList.FocusedItem.SubItems(0).Text)
            LoadRecords()
        Catch ex As Exception
            MessageBox.Show("No Record Selected", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
        End Try
    End Sub

    Private Sub btnRefresh_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnRefresh.ItemClick
        LoadRecords()
    End Sub

    Private Sub btnEdit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnEdit.ItemClick
        If Me.lvList.SelectedItems.Count > 0 Then
            frmRangeAE.id = lvList.FocusedItem.SubItems(0).Text
            frmRangeAE.cboSection.Text = cboLimit.Text
            frmRangeAE.cboTestName.Text = lvList.FocusedItem.SubItems(1).Text
            frmRangeAE.txtSIRange.Text = lvList.FocusedItem.SubItems(2).Text
            'frmRangeA.txtConvRange.Text = lvList.FocusedItem.SubItems(3).Text
            frmRangeAE.txtLow.Text = lvList.FocusedItem.SubItems(4).Text
            frmRangeAE.txtHigh.Text = lvList.FocusedItem.SubItems(5).Text
            frmRangeAE.cboGender.Text = lvList.FocusedItem.SubItems(6).Text
            frmRangeAE.cboClassifocation.Text = lvList.FocusedItem.SubItems(7).Text
            frmRangeAE.txtAgeBegin.Text = lvList.FocusedItem.SubItems(8).Text
            frmRangeAE.txtAgeEnd.Text = lvList.FocusedItem.SubItems(9).Text
            frmRangeAE.txtTestCode.Text = lvList.FocusedItem.SubItems(10).Text
            frmRangeAE.cboMachine.Text = cboMachine.Text
            Connect()
            rs.Connection = conn
            rs.CommandText = "SELECT `id`, `specimen`, `test_code`, `channel`, `convertion_factor`, `convertion_multiplier`, `instrument` FROM `specimen` WHERE (`specimen` = '" + lvList.FocusedItem.SubItems(1).Text + "' AND `test_code` = '" + lvList.FocusedItem.SubItems(10).Text + "' AND `instrument` = '" + cboMachine.Text + "')"
            reader = rs.ExecuteReader
            While reader.Read()
                frmRangeAE.txtConvFactor.Text = reader(4).ToString
                frmRangeAE.txtConvertionMultiplier.Text = reader(5).ToString
                frmRangeAE.txtChannel.Text = reader(3).ToString
            End While
            Disconnect()

            frmRangeAE.AutoLoadSection()
            frmRangeAE.AutoLoadMachine()
            frmRangeAE.btnSave.Text = "&Update"
            frmRangeAE.ShowDialog()
        End If
    End Sub

    Private Sub btnAddRefRange_ItemClick(ByVal sender As System.Object, ByVal e As DevExpress.XtraBars.ItemClickEventArgs) Handles btnAddRefRange.ItemClick
        frmRangeAE.cboSection.Text = cboLimit.Text
        frmRangeAE.cboMachine.Text = cboMachine.Text
        frmRangeAE.ShowDialog()
    End Sub

    Private Sub btnCopy_ItemClick(ByVal sender As Object, ByVal e As DevExpress.XtraBars.ItemClickEventArgs) Handles btnCopy.ItemClick
        frmRangeCopy.ShowDialog()
    End Sub
End Class