Imports DevExpress.XtraGrid.Views.Grid
Imports MySql.Data.MySqlClient

Public Class frmSpecimen

    Public Sub LoadRecords()
        Try
            Dim SQL As String = "SELECT `id` AS SequenceNo, `packagename` AS PackageName, `packagecode` AS HISCode, `section` AS `Section`, `subsection` AS SubSection FROM `packages` WHERE `section` = @Section ORDER BY `id`"

            GridView.Columns.Clear()
            GridView.Appearance.HeaderPanel.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center
            GridView.Appearance.HeaderPanel.FontStyleDelta = FontStyle.Bold

            Connect()
            rs.Connection = conn
            rs.CommandType = CommandType.Text
            rs.CommandText = SQL
            rs.Parameters.Clear()
            rs.Parameters.AddWithValue("@Section", cboSection.Text)

            Dim adapter As New MySqlDataAdapter(rs)

            Dim myTable As DataTable = New DataTable
            adapter.Fill(myTable)

            dtList.DataSource = myTable

            GridView.Columns("Section").Visible = False
            GridView.Columns("SubSection").Visible = False
            'GridView.Columns("Section").Visible = False
            'GridView.Columns("PatientType").Visible = False

            ' Make the grid read-only. 
            GridView.OptionsBehavior.Editable = False
            ' Prevent the focused cell from being highlighted. 
            GridView.OptionsSelection.EnableAppearanceFocusedCell = False
            ' Draw a dotted focus rectangle around the entire row. 
            GridView.FocusRectStyle = DrawFocusRectStyle.RowFullFocus

            For x As Integer = 0 To GridView.RowCount - 1 Step 1
                If GridView.GetRowCellValue(x, GridView.Columns("status")) = "Enable" Then
                    GridView.SelectRow(x)
                Else

                End If
            Next
            Disconnect()

        Catch ex As Exception
            MessageBox.Show(ex.Message, " ThenSystem Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Public Sub LoadRecordsTest(ByVal Package As String)
        Try
            Dim SQL As String = "SELECT `id` AS SequenceNo, `specimen` AS Specimen, `specimen_type` AS SpecimenType, `test_code` AS TestCode, `si_unit` AS SIUnit, `conventional_unit` AS ConventionalUnit, `section` AS `Section`, `test_name` AS SubSection, `order_no` AS DisplayNo, `test_group` AS TestGroup, `status` AS `Status`, `his_code` AS HISCode, `his_field` AS HisField FROM `default_specimen` WHERE `section` = @Section AND `his_code` = @PackageName ORDER BY `order_no`"

            GridViewTest.Columns.Clear()
            GridViewTest.Appearance.HeaderPanel.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center
            GridViewTest.Appearance.HeaderPanel.FontStyleDelta = FontStyle.Bold

            Connect()
            rs.Connection = conn
            rs.CommandType = CommandType.Text
            rs.CommandText = SQL
            rs.Parameters.Clear()
            rs.Parameters.AddWithValue("@Section", cboSection.Text)
            rs.Parameters.AddWithValue("@PackageName", Package)

            Dim adapter As New MySqlDataAdapter(rs)

            Dim myTable As DataTable = New DataTable
            adapter.Fill(myTable)

            dtTestList.DataSource = myTable

            GridViewTest.Columns("Section").Visible = False
            GridViewTest.Columns("SubSection").Visible = False
            GridViewTest.Columns("Status").Visible = False
            'GridView.Columns("PatientType").Visible = False

            ' Make the grid read-only. 
            GridViewTest.OptionsBehavior.Editable = False
            ' Prevent the focused cell from being highlighted. 
            GridViewTest.OptionsSelection.EnableAppearanceFocusedCell = False
            ' Draw a dotted focus rectangle around the entire row. 
            GridViewTest.FocusRectStyle = DrawFocusRectStyle.RowFullFocus

            GridViewTest.OptionsSelection.MultiSelect = True
            GridViewTest.OptionsSelection.MultiSelectMode = GridMultiSelectMode.CheckBoxRowSelect

            For x As Integer = 0 To GridViewTest.RowCount - 1 Step 1
                If GridViewTest.GetRowCellValue(x, GridViewTest.Columns("Status")) = "Enable" Then
                    GridViewTest.SelectRow(x)
                Else

                End If
            Next
            Disconnect()

        Catch ex As Exception
            MessageBox.Show(ex.Message, "System Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub frmTestType_FormClosed(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
        Me.Dispose()
    End Sub

    Private Sub frmMachineList_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        LoadRecords()
        AutoLoadTestName()
    End Sub

    Private Sub SimpleButton2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnClose.ItemClick
        Me.Close()
    End Sub

    Private Sub txtSearch_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtSearch.TextChanged
        'Try
        '    SearchRecordsOnLV(lvList, "default_specimen", 12, "specimen", txtSearch.Text, "specimen")
        'Catch ex As Exception
        '    MessageBox.Show(ex.Message, "System Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        'End Try
    End Sub

    Private Sub btnDelete_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnDelete.ItemClick
        Try
            DeleteRecord("specimen", "id", GridView.GetFocusedRowCellValue(GridView.Columns("SequenceNo")))
            LoadRecords()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "System Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub btnEdit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnEdit.ItemClick
        Dim selectedRows() As Integer = GridView.GetSelectedRows()
        For Each rowHandle As Integer In selectedRows
            If rowHandle >= 0 Then
                frmSpecimenPAE.id = GridView.GetRowCellValue(rowHandle, GridView.Columns("SequenceNo"))
                frmSpecimenPAE.txtPackageName.Text = GridView.GetRowCellValue(rowHandle, GridView.Columns("PackageName"))
                frmSpecimenPAE.txtHISCode.Text = GridView.GetRowCellValue(rowHandle, GridView.Columns("HISCode"))
                frmSpecimenPAE.cboTestName.Text = GridView.GetRowCellValue(rowHandle, GridView.Columns("Section"))
                frmSpecimenPAE.btnSave.Text = "&Update"
                frmSpecimenPAE.ShowDialog()
            End If
        Next
    End Sub

    Private Sub btnRefresh_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnRefresh.ItemClick
        LoadRecords()
    End Sub

    Private Sub btnAdd_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnAdd.ItemClick
        frmSpecimenPAE.cboTestName.Text = cboSection.Text

        frmSpecimenPAE.btnSave.Text = "&Save"
        frmSpecimenPAE.ShowDialog()
    End Sub

    Public Sub AutoLoadTestName()
        cboSection.Properties.Items.Clear()
        Connect()
        rs.Connection = conn
        rs.CommandText = "Select * FROM `testtype` WHERE `test_name` Not Like 'All' ORDER BY `test_name`"
        reader = rs.ExecuteReader
        While reader.Read
            cboSection.Properties.Items.Add(reader(1).ToString)
        End While
        Disconnect()
    End Sub

    Private Sub btnUpdateStatus_Click(sender As Object, e As EventArgs) Handles btnUpdateStatus.Click
        Dim selectedRows() As Integer = GridViewTest.GetSelectedRows()
        For Each rowHandle As Integer In selectedRows
            If rowHandle >= 0 Then
                If (GridViewTest.IsRowSelected(rowHandle)) Then
                    UpdateRecordwthoutMSG("UPDATE `default_specimen` SET `status` = 'Enable' WHERE `id` = '" & GridViewTest.GetRowCellValue(rowHandle, GridViewTest.Columns("SequenceNo")) & "'")
                Else

                End If
            End If
        Next

        If GridViewTest.GetFocusedRowCellValue(GridViewTest.Columns("Status")) = "Enable" Then
            UpdateRecordwthoutMSG("UPDATE `default_specimen` SET `status` = 'Disable' WHERE `id` = '" & GridViewTest.GetFocusedRowCellValue(GridViewTest.Columns("SequenceNo")) & "'")
        Else
            UpdateRecordwthoutMSG("UPDATE `default_specimen` SET `status` = 'Enable' WHERE `id` = '" & GridViewTest.GetFocusedRowCellValue(GridViewTest.Columns("SequenceNo")) & "'")
        End If
    End Sub

    Private Sub cboSection_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cboSection.SelectedIndexChanged
        LoadRecords()
    End Sub

    Private Sub GridView_RowClick(sender As Object, e As RowClickEventArgs) Handles GridView.RowClick
        LoadRecordsTest(GridView.GetRowCellValue(e.RowHandle, GridView.Columns("HISCode")))
    End Sub

    Private Sub btnAddTest_Click(sender As Object, e As EventArgs) Handles btnAddTest.Click
        frmSpecimenAE.cboTestName.Text = cboSection.Text
        frmSpecimenAE.txtHISCode.Text = GridView.GetFocusedRowCellValue(GridView.Columns("HISCode"))
        frmSpecimenAE.btnSave.Text = "&Save"
        frmSpecimenAE.ShowDialog()
    End Sub

    Private Sub btnEditTest_Click(sender As Object, e As EventArgs) Handles btnEditTest.Click
        frmSpecimenAE.ID = GridViewTest.GetFocusedRowCellValue(GridViewTest.Columns("SequenceNo"))
        frmSpecimenAE.cboTestName.Text = GridViewTest.GetFocusedRowCellValue(GridViewTest.Columns("Section"))
        frmSpecimenAE.cboTest.Text = GridViewTest.GetFocusedRowCellValue(GridViewTest.Columns("Specimen"))
        frmSpecimenAE.cboSpecimen.Text = GridViewTest.GetFocusedRowCellValue(GridViewTest.Columns("SpecimenType"))
        frmSpecimenAE.cboUnit.Text = GridViewTest.GetFocusedRowCellValue(GridViewTest.Columns("SIUnit"))
        frmSpecimenAE.cboConv_Unit.Text = GridViewTest.GetFocusedRowCellValue(GridViewTest.Columns("ConventionalUnit"))
        frmSpecimenAE.txtTestCode.Text = GridViewTest.GetFocusedRowCellValue(GridViewTest.Columns("TestCode"))
        frmSpecimenAE.txtHISCode.Text = GridViewTest.GetFocusedRowCellValue(GridViewTest.Columns("HISCode"))
        frmSpecimenAE.txtOrderNo.Text = GridViewTest.GetFocusedRowCellValue(GridViewTest.Columns("DisplayNo"))
        frmSpecimenAE.cboTestGroup.Text = GridViewTest.GetFocusedRowCellValue(GridViewTest.Columns("TestGroup"))

        frmSpecimenAE.btnSave.Text = "&Update"
        frmSpecimenAE.ShowDialog()
    End Sub
End Class