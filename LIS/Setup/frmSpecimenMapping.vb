Imports DevExpress.XtraGrid.Views.Grid
Imports MySql.Data.MySqlClient

Public Class frmSpecimenMapping

    Public Sub LoadRecords()
        Try
            Dim SQL As String = "SELECT `id` AS SequenceNo, `specimen` AS Specimen, `description` AS Description, `specimen_type` AS SpecimenType, `channel` AS Channel, `test_code` AS TestCode, `si_unit` AS SIUnit, `conventional_unit` AS ConventionalUnit, `convertion_factor` AS ConvertionFactor, `convertion_multiplier` AS DecimalPlaces, `order_no` AS DisplayNo, `test_group` AS TestGroup, `instrument` AS Instrument, `section` AS `Section`, `test_name` AS SubSection, `status` AS `Status` FROM `specimen` WHERE `section` = @Section AND `instrument` = @Instrument ORDER BY `order_no`"

            GridView.Columns.Clear()
            GridView.Appearance.HeaderPanel.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center
            GridView.Appearance.HeaderPanel.FontStyleDelta = FontStyle.Bold

            Connect()
            rs.Connection = conn
            rs.CommandType = CommandType.Text
            rs.CommandText = SQL
            rs.Parameters.Clear()
            rs.Parameters.AddWithValue("@Section", cboSection.Text)
            rs.Parameters.AddWithValue("@Instrument", cboMachine.Text)

            Dim adapter As New MySqlDataAdapter(rs)

            Dim myTable As DataTable = New DataTable
            adapter.Fill(myTable)

            dtList.DataSource = myTable

            GridView.Columns("Section").Visible = False
            GridView.Columns("SubSection").Visible = False
            GridView.Columns("Status").Visible = False
            GridView.Columns("Instrument").Visible = False

            ' Make the grid read-only. 
            GridView.OptionsBehavior.Editable = False
            ' Prevent the focused cell from being highlighted. 
            GridView.OptionsSelection.EnableAppearanceFocusedCell = False
            ' Draw a dotted focus rectangle around the entire row. 
            GridView.FocusRectStyle = DrawFocusRectStyle.RowFullFocus

            GridView.OptionsSelection.MultiSelect = True
            GridView.OptionsSelection.MultiSelectMode = GridMultiSelectMode.CheckBoxRowSelect

            For x As Integer = 0 To GridView.RowCount - 1 Step 1
                If GridView.GetRowCellValue(x, GridView.Columns("Status")) = "Enable" Then
                    GridView.SelectRow(x)
                Else

                End If
            Next
            Disconnect()

        Catch ex As Exception
            MessageBox.Show(ex.Message, " ThenSystem Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
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

    Private Sub btnDelete_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnDelete.ItemClick
        Try
            DeleteRecord("specimen", "id", GridView.GetFocusedRowCellValue(GridView.Columns("SequenceNo")))
            LoadRecords()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "System Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub btnEdit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnEdit.ItemClick
        'frmSpecimenMappingAE.AutoLoadCombo()
        'frmSpecimenMappingAE.AutoLoadComboSpecimen()
        frmSpecimenMappingAE.name = GridView.GetFocusedRowCellValue(GridView.Columns("SequenceNo"))
        frmSpecimenMappingAE.cboTest.Text = GridView.GetFocusedRowCellValue(GridView.Columns("Specimen"))
        frmSpecimenMappingAE.txtDescription.Text = GridView.GetFocusedRowCellValue(GridView.Columns("Description")).ToString
        frmSpecimenMappingAE.cboSpecimen.Text = GridView.GetFocusedRowCellValue(GridView.Columns("SpecimenType")).ToString
        frmSpecimenMappingAE.txtChannel.Text = GridView.GetFocusedRowCellValue(GridView.Columns("Channel")).ToString
        frmSpecimenMappingAE.txtTestCode.Text = GridView.GetFocusedRowCellValue(GridView.Columns("TestCode")).ToString
        frmSpecimenMappingAE.cboUnit.Text = GridView.GetFocusedRowCellValue(GridView.Columns("SIUnit")).ToString
        frmSpecimenMappingAE.cboConv_Unit.Text = GridView.GetFocusedRowCellValue(GridView.Columns("ConventionalUnit")).ToString
        frmSpecimenMappingAE.cboConvetionFactor.Text = GridView.GetFocusedRowCellValue(GridView.Columns("ConvertionFactor")).ToString
        frmSpecimenMappingAE.txtMultiplier.Text = GridView.GetFocusedRowCellValue(GridView.Columns("DecimalPlaces")).ToString
        frmSpecimenMappingAE.cboTestName.Text = GridView.GetFocusedRowCellValue(GridView.Columns("Section")).ToString
        frmSpecimenMappingAE.txtOrderNo.Text = GridView.GetFocusedRowCellValue(GridView.Columns("DisplayNo")).ToString
        frmSpecimenMappingAE.cboTestGroup.Text = GridView.GetFocusedRowCellValue(GridView.Columns("TestGroup")).ToString
        frmSpecimenMappingAE.cboInstrument.Text = GridView.GetFocusedRowCellValue(GridView.Columns("Instrument")).ToString
        frmSpecimenMappingAE.btnSave.Text = "&Update"
        frmSpecimenMappingAE.btnNew.Enabled = False
        frmSpecimenMappingAE.ShowDialog()
    End Sub

    Private Sub btnRefresh_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnRefresh.ItemClick
        LoadRecords()
    End Sub

    Private Sub btnAdd_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnAdd.ItemClick
        frmSpecimenMappingAE.AutoLoadCombo()
        frmSpecimenMappingAE.AutoLoadComboSpecimen()
        frmSpecimenMappingAE.cboTestName.Text = cboSection.Text
        frmSpecimenMappingAE.cboInstrument.Text = cboMachine.Text
        frmSpecimenMappingAE.btnSave.Text = "&Save"
        frmSpecimenMappingAE.btnNew.Enabled = True
        frmSpecimenMappingAE.ShowDialog()
    End Sub

    Public Sub AutoLoadTestName()
        cboSection.Properties.Items.Clear()
        Connect()
        rs.Connection = conn
        rs.CommandText = "SELECT * FROM `testtype` WHERE `test_name` NOT LIKE 'All' ORDER BY `test_name`"
        reader = rs.ExecuteReader
        While reader.Read
            cboSection.Properties.Items.Add(reader(1).ToString)
        End While
        Disconnect()
    End Sub

    Public Sub AutoLoadMachines()
        cboMachine.Properties.Items.Clear()
        Connect()
        rs.Connection = conn
        rs.CommandText = "SELECT `name`, `serial_no` FROM `Machines` WHERE `test_name` = '" & cboSection.Text & "' ORDER BY `name`"
        reader = rs.ExecuteReader
        While reader.Read
            cboMachine.Properties.Items.Add(reader(0).ToString & "_" & reader(1).ToString)
        End While
        Disconnect()
    End Sub

    Private Sub lvList_DrawColumnHeader(ByVal sender As Object, ByVal e As System.Windows.Forms.DrawListViewColumnHeaderEventArgs)
        Using sf As StringFormat =
            New StringFormat()

            ' Store the column text alignment, letting it default
            ' to Left if it has not been set to Center or Right.
            Select Case (e.Header.TextAlign)

                Case HorizontalAlignment.Center
                    sf.Alignment = StringAlignment.Center

                Case HorizontalAlignment.Right
                    sf.Alignment = StringAlignment.Far

            End Select

            ' Draw the standard header background.
            e.DrawBackground()

            ' Draw the header text.
            Using headerFont As Font =
                New Font("Helvetica", 10, FontStyle.Bold)

                e.Graphics.DrawString(e.Header.Text, headerFont,
                    Brushes.Black, e.Bounds, sf)
            End Using
        End Using
        Return
    End Sub

    Private Sub cboLimit_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cboSection.SelectedIndexChanged
        AutoLoadMachines()
    End Sub

    Private Sub cboMachine_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cboMachine.SelectedIndexChanged
        LoadRecords()
    End Sub

    Private Sub btnUpdate_ItemClick(ByVal sender As Object, ByVal e As DevExpress.XtraBars.ItemClickEventArgs) Handles btnUpdate.ItemClick
        Dim selectedRows() As Integer = GridView.GetSelectedRows()
        For Each rowHandle As Integer In selectedRows
            If rowHandle >= 0 Then
                If (GridView.IsRowSelected(rowHandle)) Then
                    UpdateRecordwthoutMSG("UPDATE `specimen` SET `status` = 'Enable' WHERE `id` = '" & GridView.GetRowCellValue(rowHandle, GridView.Columns("SequenceNo")) & "'")
                Else

                End If
            End If
        Next

        If GridView.GetFocusedRowCellValue(GridView.Columns("Status")) = "Enable" Then
            UpdateRecordwthoutMSG("UPDATE `specimen` SET `status` = 'Disable' WHERE `id` = '" & GridView.GetFocusedRowCellValue(GridView.Columns("SequenceNo")) & "'")
        Else
            UpdateRecordwthoutMSG("UPDATE `specimen` SET `status` = 'Enable' WHERE `id` = '" & GridView.GetFocusedRowCellValue(GridView.Columns("SequenceNo")) & "'")
        End If
        LoadRecords()
    End Sub

    Private Sub btnCopy_ItemClick(ByVal sender As Object, ByVal e As DevExpress.XtraBars.ItemClickEventArgs) Handles btnCopy.ItemClick
        frmSpecimenCopy.ShowDialog()
    End Sub

End Class