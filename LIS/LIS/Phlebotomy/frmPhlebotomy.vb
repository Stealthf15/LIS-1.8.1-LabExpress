Imports System.Drawing.Printing
Imports DevExpress.Xpo
Imports DevExpress.XtraGrid.Views.Grid
Imports DevExpress.XtraPrinting.BarCode

Public Class frmPhlebotomy

    Public Sub LoadRecords()
        Try
            GridView.Columns.Clear()
            GridView.Appearance.HeaderPanel.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center
            GridView.Appearance.HeaderPanel.FontStyleDelta = FontStyle.Bold

            Dim SQL As String = "SELECT 
                        `id` AS ID, `status` AS `Status`, `sample_id` AS SampleID, `patient_id` AS PatientID, `patient_name`AS PatientName, 
                        `test` AS Request, `bdate` AS DateOfBirth, `sex` AS Sex, DATE_FORMAT(`date`, '%m/%d/%Y') AS DateReceived, `time` AS TimeReceived, 
                        `testtype` AS Section, `sub_section` AS SubSection, `main_id` AS RefID 
                        FROM `tmpWorklist` WHERE (`status` = 'Checked-In' OR `status` = 'Rejected' OR `status` = 'Cancelled' OR `status` = 'Processing' OR `status` = 'For Verification' OR `status` = 'Verified' OR `status` = 'Validated') 
                        AND (`date` BETWEEN @DateFrom and @DateTo) ORDER BY `status` ASC"

            Dim command As New MySql.Data.MySqlClient.MySqlCommand(SQL, conn)

            command.Parameters.Clear()
            command.Parameters.Add("@DateFrom", MySql.Data.MySqlClient.MySqlDbType.DateTime).Value = Format(dtFrom1.Value, "yyyy-MM-dd")
            command.Parameters.Add("@DateTo", MySql.Data.MySqlClient.MySqlDbType.DateTime).Value = Format(dtTo1.Value, "yyyy-MM-dd")

            Dim adapter As New MySql.Data.MySqlClient.MySqlDataAdapter(command)

            Dim myTable As DataTable = New DataTable
            adapter.Fill(myTable)

            dtList.DataSource = myTable

            GridView.Columns("RefID").Visible = False
            GridView.Columns("Section").Visible = False
            GridView.Columns("SubSection").Visible = False

            ' Make the grid read-only. 
            GridView.OptionsBehavior.Editable = False
            ' Prevent the focused cell from being highlighted. 
            GridView.OptionsSelection.EnableAppearanceFocusedCell = False
            ' Draw a dotted focus rectangle around the entire row. 
            GridView.FocusRectStyle = DrawFocusRectStyle.RowFullFocus

            GridView.OptionsSelection.MultiSelect = True
            GridView.OptionsSelection.MultiSelectMode = GridMultiSelectMode.CheckBoxRowSelect

        Catch ex As Exception
            MessageBox.Show(ex.Message, "System Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub GridView_RowCellStyle(sender As Object, e As RowCellStyleEventArgs) Handles GridView.RowCellStyle
        Dim view As GridView = TryCast(sender, GridView)
        If (e.Column.FieldName = "ID") Or (e.Column.FieldName = "Status") Then
            If view.GetRowCellValue(e.RowHandle, "Status") = "Warding" Then
                e.Appearance.BackColor = Color.Orange
                e.Appearance.BackColor2 = Color.Orange
                e.Appearance.ForeColor = Color.White
            ElseIf view.GetRowCellValue(e.RowHandle, "Status") = "Rejected" Then
                e.Appearance.BackColor = Color.Crimson
                e.Appearance.BackColor2 = Color.Crimson
                e.Appearance.ForeColor = Color.White
            ElseIf view.GetRowCellValue(e.RowHandle, "Status") = "Cancelled" Then
                e.Appearance.BackColor = Color.Gray
                e.Appearance.BackColor2 = Color.Gray
                e.Appearance.ForeColor = Color.White
            ElseIf view.GetRowCellValue(e.RowHandle, "Status") = "Processing" Then
                e.Appearance.BackColor = Color.Gold
                e.Appearance.BackColor2 = Color.Gold
                e.Appearance.ForeColor = Color.Black
            ElseIf view.GetRowCellValue(e.RowHandle, "Status") = "Result Received" Then
                e.Appearance.BackColor = Color.LightGreen
                e.Appearance.BackColor2 = Color.LightGreen
                e.Appearance.ForeColor = Color.Black
            ElseIf view.GetRowCellValue(e.RowHandle, "Status") = "Validated" Then
                e.Appearance.BackColor = Color.Green
                e.Appearance.BackColor2 = Color.Green
                e.Appearance.ForeColor = Color.Black
            ElseIf view.GetRowCellValue(e.RowHandle, "Status") = "Printed" Then
                e.Appearance.BackColor = Color.ForestGreen
                e.Appearance.BackColor2 = Color.ForestGreen
                e.Appearance.ForeColor = Color.White
            ElseIf view.GetRowCellValue(e.RowHandle, "Status") = "Verified" Then
                e.Appearance.BackColor = Color.DarkCyan
                e.Appearance.BackColor2 = Color.DarkCyan
                e.Appearance.ForeColor = Color.White
            ElseIf view.GetRowCellValue(e.RowHandle, "Status") = "For Verification" Then
                e.Appearance.BackColor = Color.Tan
                e.Appearance.BackColor2 = Color.Tan
                e.Appearance.ForeColor = Color.Black
            ElseIf view.GetRowCellValue(e.RowHandle, "Status") = "Checked-In" Then
                e.Appearance.BackColor = Color.CornflowerBlue
                e.Appearance.BackColor2 = Color.CornflowerBlue
                e.Appearance.ForeColor = Color.White
            End If
        End If
    End Sub

    Public Sub LoadRecordsFilterWard()
        Try
            GridView.Columns.Clear()
            GridView.Appearance.HeaderPanel.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center
            GridView.Appearance.HeaderPanel.FontStyleDelta = FontStyle.Bold

            Dim SQL As String = "SELECT 
                        `id` AS ID, `status` AS `Status`, `sample_id` AS SampleID, `patient_id` AS PatientID, `patient_name`AS PatientName, 
                        `test` AS Request, `bdate` AS DateOfBirth, `sex` AS Sex, DATE_FORMAT(`date`, '%m/%d/%Y') AS DateReceived, `time` AS TimeReceived, 
                        `testtype` AS Section, `sub_section` AS SubSection, `main_id` AS RefID 
                        FROM `tmpWorklist` WHERE (`status` = 'Ordered' OR `status` = 'Rejected' OR `status` = 'Cancelled' OR `status` = 'Warding') 
                        AND (`date` BETWEEN @DateFrom AND @DateTo) AND `dept` = @Search ORDER BY `id` DESC"

            Dim command As New MySql.Data.MySqlClient.MySqlCommand(SQL, conn)

            command.Parameters.Clear()
            command.Parameters.Add("@DateFrom", MySql.Data.MySqlClient.MySqlDbType.DateTime).Value = Format(dtFrom1.Value, "yyyy-MM-dd")
            command.Parameters.Add("@DateTo", MySql.Data.MySqlClient.MySqlDbType.DateTime).Value = Format(dtTo1.Value, "yyyy-MM-dd")
            command.Parameters.AddWithValue("@Search", cboWard.Text)

            Dim adapter As New MySql.Data.MySqlClient.MySqlDataAdapter(command)

            Dim myTable As DataTable = New DataTable
            adapter.Fill(myTable)

            dtList.DataSource = myTable

            ' Make the grid read-only. 
            GridView.OptionsBehavior.Editable = False
            ' Prevent the focused cell from being highlighted. 
            GridView.OptionsSelection.EnableAppearanceFocusedCell = False
            ' Draw a dotted focus rectangle around the entire row. 
            GridView.FocusRectStyle = DrawFocusRectStyle.RowFullFocus

            GridView.OptionsSelection.MultiSelect = True
            GridView.OptionsSelection.MultiSelectMode = GridMultiSelectMode.CheckBoxRowSelect

        Catch ex As Exception
            MessageBox.Show(ex.Message, "System Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub frmNewOrder_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        LoadRecords()
        LoadWard()
    End Sub

    Private Sub LoadWard()
        Connect()
        rs.Connection = conn
        rs.CommandType = CommandType.Text
        rs.CommandText = "SELECT DISTINCT `dept` FROM tmpworklist"
        reader = rs.ExecuteReader
        While reader.Read
            cboWard.Properties.Items.Add(reader(0).ToString)
        End While
        Disconnect()
    End Sub


    Private Sub btnAdd_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnAdd.ItemClick
        Dim selectedRows() As Integer = GridView.GetSelectedRows()
        For Each rowHandle As Integer In selectedRows
            If rowHandle >= 0 Then
                Dim Result As DialogResult = MessageBox.Show("You're about to Check-In Patient " & GridView.GetRowCellValue(rowHandle, GridView.Columns("PatientName")) & "." & vbCrLf & vbCrLf & "Do you want to continue to print Barcode Sticker " & GridView.GetRowCellValue(rowHandle, GridView.Columns("SampleID")) & "?", "System Message", MessageBoxButtons.YesNoCancel, MessageBoxIcon.Question)

                If (Result = DialogResult.Yes) Then
                    Try
                        PrintBarcode(GridView.GetRowCellValue(rowHandle, GridView.Columns("Request")),
                                     GridView.GetRowCellValue(rowHandle, GridView.Columns("SampleID")),
                                     GridView.GetRowCellValue(rowHandle, GridView.Columns("PatientID")),
                                     GridView.GetRowCellValue(rowHandle, GridView.Columns("PatientName")),
                                     GridView.GetRowCellValue(rowHandle, GridView.Columns("DateOfBirth")),
                                     GridView.GetRowCellValue(rowHandle, GridView.Columns("Sex")),
                                     GridView.GetRowCellValue(rowHandle, GridView.Columns("Section")),
                                     GridView.GetRowCellValue(rowHandle, GridView.Columns("SubSection")))

                    Catch ex As Exception
                        MessageBox.Show("Error in connection on printer. " + ex.Message, "Barcode Printing Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    End Try
                    GoTo A
                ElseIf (Result = DialogResult.No) Then
                    GoTo A
                ElseIf (Result = DialogResult.Cancel) Then
                    Exit Sub
                End If

A:

                rs.Parameters.Clear()
                rs.Parameters.AddWithValue("@mainID", GridView.GetRowCellValue(rowHandle, GridView.Columns("RefID")))
                rs.Parameters.AddWithValue("@SampleID", GridView.GetRowCellValue(rowHandle, GridView.Columns("SampleID")))
                rs.Parameters.AddWithValue("@status", "Checked-In")
                rs.Parameters.AddWithValue("@Section", GridView.GetRowCellValue(rowHandle, GridView.Columns("Section")))
                rs.Parameters.AddWithValue("@SubSection", GridView.GetRowCellValue(rowHandle, GridView.Columns("SubSection")))
                rs.Parameters.AddWithValue("@time_checked_in", Now)

                UpdateRecordwthoutMSG("UPDATE `tmpWorklist` SET " _
                    & "`sample_id` = @SampleID," _
                    & "`main_id` = @mainID," _
                    & "`status` = @status" _
                    & " WHERE main_id = @mainID AND `testtype` = @Section AND `sub_section` = @SubSection"
                    )

                UpdateRecordwthoutMSG("UPDATE `additional_info` SET " _
                    & "`sample_id` = @SampleID," _
                    & "`accession_no` = @mainID" _
                    & " WHERE sample_id = @mainID AND `section` = @Section AND `sub_section` = @SubSection"
                    )

                'UpdateRecordwthoutMSG("UPDATE `tmpresult` SET " _
                '    & "`sample_id` = @SampleID" _
                '    & " WHERE sample_id = @mainID AND section = @Section AND sub_section = @SubSection"
                '    )

                'UpdateRecordwthoutMSG("UPDATE `patient_order` SET " _
                '    & "`sample_id` = @SampleID" _
                '    & " WHERE sample_id = @mainID AND section = @Section AND sub_section = @SubSection"
                '    )

                'UpdateRecordwthoutMSG("UPDATE `lis_order` SET " _
                '    & "`sample_id` = @SampleID" _
                '    & " WHERE sample_id = @mainID AND section = @Section AND sub_section = @SubSection"
                '    )

                Connect()
                rs.Connection = conn
                rs.CommandType = CommandType.Text
                rs.CommandText = "SELECT `sample_id` FROM `specimen_tracking` WHERE `sample_id` = @SampleID"
                reader = rs.ExecuteReader
                reader.Read()
                If reader.HasRows Then
                    Disconnect()
                    UpdateRecordwthoutMSG("UPDATE `specimen_tracking` SET " _
                        & "`sample_id` = @SampleID," _
                        & "`extracted` = @time_checked_in" _
                        & " WHERE sample_id = @mainID AND `section` = @Section AND `sub_section` = @SubSection"
                        )
                Else
                    Disconnect()
                    SaveRecordwthoutMSG("INSERT INTO `specimen_tracking` (`sample_id`, `extracted`, `section`, `sub_section`) VALUES " _
                        & "(" _
                        & "@SampleID," _
                        & "@time_checked_in," _
                        & "@Section," _
                        & "@SubSection" _
                        & ")"
                        )
                End If
                Disconnect()
                'Log activity
                SpecimenActivity("z_logs_specimen", GridView.GetRowCellValue(rowHandle, GridView.Columns("SampleID")), GridView.GetRowCellValue(rowHandle, GridView.Columns("PatientID")), GridView.GetRowCellValue(rowHandle, GridView.Columns("PatientName")), CurrUser, "Checked-In Specimen", "", GridView.GetRowCellValue(rowHandle, GridView.Columns("Request")), GridView.GetRowCellValue(rowHandle, GridView.Columns("Section")), GridView.GetRowCellValue(rowHandle, GridView.Columns("SubSection")))
            End If
        Next rowHandle
        LoadRecords()

    End Sub

    Private Sub btnRefresh_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnRefresh.ItemClick
        LoadRecords()
    End Sub

    Private Sub btnClose_ItemClick(ByVal sender As Object, ByVal e As DevExpress.XtraBars.ItemClickEventArgs) Handles btnClose.ItemClick
        Me.Close()
        Me.Dispose()
    End Sub

    Private Sub btnWarding_ItemClick(ByVal sender As Object, ByVal e As DevExpress.XtraBars.ItemClickEventArgs) Handles btnAddOrder.ItemClick
        Dim selectedRows() As Integer = GridView.GetSelectedRows()
        For Each rowHandle As Integer In selectedRows
            If rowHandle >= 0 Then
                Dim Result As DialogResult = MessageBox.Show("You're about to Check-In Patient " & GridView.GetRowCellValue(rowHandle, GridView.Columns("PatientName")) & "." & vbCrLf & vbCrLf & "Do you want to continue to print Barcode Sticker " & GridView.GetRowCellValue(rowHandle, GridView.Columns("SampleID")) & "?", "System Message", MessageBoxButtons.YesNoCancel, MessageBoxIcon.Question)

                If (Result = DialogResult.Yes) Then
                    Try
                        PrintBarcode(GridView.GetRowCellValue(rowHandle, GridView.Columns("Request")),
                                     GridView.GetRowCellValue(rowHandle, GridView.Columns("SampleID")),
                                     GridView.GetRowCellValue(rowHandle, GridView.Columns("PatientID")),
                                     GridView.GetRowCellValue(rowHandle, GridView.Columns("PatientName")),
                                     GridView.GetRowCellValue(rowHandle, GridView.Columns("DateOfBirth")),
                                     GridView.GetRowCellValue(rowHandle, GridView.Columns("Sex")),
                                     GridView.GetRowCellValue(rowHandle, GridView.Columns("Section")),
                                     GridView.GetRowCellValue(rowHandle, GridView.Columns("SubSection")))
                    Catch ex As Exception
                        MessageBox.Show("Error in connection on printer. " + ex.Message, "Barcode Printing Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    End Try
                    GoTo A
                ElseIf (Result = DialogResult.No) Then
                    GoTo A
                ElseIf (Result = DialogResult.Cancel) Then
                    Exit Sub
                End If

A:

                rs.Parameters.Clear()
                rs.Parameters.AddWithValue("@mainID", GridView.GetRowCellValue(rowHandle, GridView.Columns("RefID")))
                rs.Parameters.AddWithValue("@SampleID", GridView.GetRowCellValue(rowHandle, GridView.Columns("SampleID")))
                rs.Parameters.AddWithValue("@status", "Warding")
                rs.Parameters.AddWithValue("@Section", GridView.GetRowCellValue(rowHandle, GridView.Columns("Section")))
                rs.Parameters.AddWithValue("@SubSection", GridView.GetRowCellValue(rowHandle, GridView.Columns("SubSection")))
                rs.Parameters.AddWithValue("@time_warding", Now)

                UpdateRecordwthoutMSG("UPDATE `tmpWorklist` SET " _
                    & "`sample_id` = @SampleID," _
                    & "`main_id` = @mainID," _
                    & "`status` = @status" _
                    & " WHERE main_id = @mainID AND `testtype` = @Section AND `sub_section` = @SubSection"
                    )

                UpdateRecordwthoutMSG("UPDATE `additional_info` SET " _
                    & "`sample_id` = @SampleID," _
                    & "`accession_no` = @mainID" _
                    & " WHERE sample_id = @mainID AND `section` = @Section AND `sub_section` = @SubSection"
                    )

                'UpdateRecordwthoutMSG("UPDATE `tmpresult` SET " _
                '    & "`sample_id` = @SampleID" _
                '    & " WHERE sample_id = @mainID AND section = @Section AND sub_section = @SubSection"
                '    )

                'UpdateRecordwthoutMSG("UPDATE `patient_order` SET " _
                '    & "`sample_id` = @SampleID" _
                '    & " WHERE sample_id = @mainID AND section = @Section AND sub_section = @SubSection"
                '    )

                'UpdateRecordwthoutMSG("UPDATE `lis_order` SET " _
                '    & "`sample_id` = @SampleID" _
                '    & " WHERE sample_id = @mainID AND section = @Section AND sub_section = @SubSection"
                '    )

                Connect()
                rs.Connection = conn
                rs.CommandType = CommandType.Text
                rs.CommandText = "SELECT `sample_id` FROM `specimen_tracking` WHERE `sample_id` = @SampleID"
                reader = rs.ExecuteReader
                reader.Read()
                If reader.HasRows Then
                    Disconnect()
                    UpdateRecordwthoutMSG("UPDATE `specimen_tracking` SET " _
                        & "`sample_id` = @SampleID," _
                        & "`warding` = @time_warding" _
                        & " WHERE sample_id = @mainID AND `section` = @Section AND `sub_section` = @SubSection"
                        )
                Else
                    Disconnect()
                    SaveRecordwthoutMSG("INSERT INTO `specimen_tracking` (`sample_id`, `warding`, `section`, `sub_section`) VALUES " _
                        & "(" _
                        & "@SampleID," _
                        & "@time_warding," _
                        & "@Section," _
                        & "@SubSection" _
                        & ")"
                        )
                End If
                Disconnect()
                'Log activity
                SpecimenActivity("z_logs_specimen", GridView.GetRowCellValue(rowHandle, GridView.Columns("SampleID")), GridView.GetRowCellValue(rowHandle, GridView.Columns("PatientID")), GridView.GetRowCellValue(rowHandle, GridView.Columns("PatientName")), CurrUser, "Ward Specimen", "", GridView.GetRowCellValue(rowHandle, GridView.Columns("Request")), GridView.GetRowCellValue(rowHandle, GridView.Columns("Section")), GridView.GetRowCellValue(rowHandle, GridView.Columns("SubSection")))
            End If
        Next rowHandle
        LoadRecords()
    End Sub

    Private Sub frm_Activated(sender As Object, e As EventArgs) Handles Me.Activated
        'MainFOrm.aceFecal.Appearance.Normal.BackColor = Color.FromArgb(6, 31, 71)
        MainFOrm.accPhlebotomy.Appearance.Normal.BackColor = Color.FromArgb(16, 110, 190)
        MainFOrm.accPhlebotomy.Appearance.Normal.ForeColor = Color.FromArgb(255, 255, 255)
    End Sub

    Private Sub frm_FormClosed(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
        MainFOrm.lblTitle.Text = ""
        MainFOrm.accPhlebotomy.Appearance.Normal.BackColor = Color.FromArgb(240, 240, 240)
        MainFOrm.accPhlebotomy.Appearance.Normal.ForeColor = Color.FromArgb(27, 41, 62)
        Me.Dispose()
    End Sub

    Private Sub btnSearch_Click(sender As Object, e As EventArgs) Handles btnSearch.Click
        LoadRecords()
    End Sub

    Private Sub cboWard_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cboWard.SelectedIndexChanged
        LoadRecordsFilterWard()
    End Sub

    Private Sub btnDelete_ItemClick(sender As Object, e As DevExpress.XtraBars.ItemClickEventArgs) Handles btnDelete.ItemClick
        Dim selectedRows() As Integer = GridView.GetSelectedRows()
        For Each rowHandle As Integer In selectedRows
            If rowHandle >= 0 Then
                If MessageBox.Show("Are you sure you want to reject " & GridView.GetRowCellValue(rowHandle, GridView.Columns("PatientName")) & " order?", "Confirm Reject", MessageBoxButtons.YesNo, MessageBoxIcon.Question) = vbYes Then
                    frmRejectOrder.ID = GridView.GetRowCellValue(rowHandle, GridView.Columns("ID"))
                    frmRejectOrder.sID = GridView.GetRowCellValue(rowHandle, GridView.Columns("SampleID"))
                    frmRejectOrder.pID = GridView.GetRowCellValue(rowHandle, GridView.Columns("PatientID"))
                    frmRejectOrder.pName = GridView.GetRowCellValue(rowHandle, GridView.Columns("PatientName"))
                    frmRejectOrder.pTest = GridView.GetRowCellValue(rowHandle, GridView.Columns("Request"))
                    frmRejectOrder.pSection = GridView.GetRowCellValue(rowHandle, GridView.Columns("Section"))
                    frmRejectOrder.pSubSection = GridView.GetRowCellValue(rowHandle, GridView.Columns("SubSection"))
                    frmRejectOrder.ShowDialog()
                End If
            End If
        Next rowHandle

        LoadRecords()
    End Sub

    Private Sub btnCancel_ItemClick(ByVal sender As Object, ByVal e As DevExpress.XtraBars.ItemClickEventArgs) Handles btnCancel.ItemClick
        Dim selectedRows() As Integer = GridView.GetSelectedRows()
        For Each rowHandle As Integer In selectedRows
            If rowHandle >= 0 Then
                If MessageBox.Show("Are you sure you want to cancel " & GridView.GetRowCellValue(rowHandle, GridView.Columns("PatientName")) & " order?", "Confirm Cancel", MessageBoxButtons.YesNo, MessageBoxIcon.Question) = vbYes Then
                    frmCancelOR.ID = GridView.GetRowCellValue(rowHandle, GridView.Columns("ID"))
                    frmCancelOR.sID = GridView.GetRowCellValue(rowHandle, GridView.Columns("SampleID"))
                    frmCancelOR.pID = GridView.GetRowCellValue(rowHandle, GridView.Columns("PatientID"))
                    frmCancelOR.pName = GridView.GetRowCellValue(rowHandle, GridView.Columns("PatientName"))
                    frmCancelOR.pTest = GridView.GetRowCellValue(rowHandle, GridView.Columns("Request"))
                    frmCancelOR.pSection = GridView.GetRowCellValue(rowHandle, GridView.Columns("Section"))
                    frmCancelOR.pSubSection = GridView.GetRowCellValue(rowHandle, GridView.Columns("SubSection"))
                    frmCancelOR.ShowDialog()
                End If
            End If
        Next rowHandle

        LoadRecords()
    End Sub

    Private Sub btnSearch_KeyPress(sender As Object, e As KeyPressEventArgs) Handles btnSearch.KeyPress
        If e.KeyChar = Chr(13) Then
            LoadRecordsFilterWard()
        End If
    End Sub

    Private Sub NotifyMe()
        ToastNotificationsManager.ShowNotification(ToastNotificationsManager.Notifications(0))
    End Sub

    Private Sub txtSearch_SelectedIndexChanged(sender As Object, e As EventArgs) Handles txtSearch.TextChanged
        Try
            GridView.Columns.Clear()
            GridView.Appearance.HeaderPanel.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center
            GridView.Appearance.HeaderPanel.FontStyleDelta = FontStyle.Bold

            If rgSelect.SelectedIndex = 0 Then
                Dim SQL As String = "SELECT 
                        `id` AS ID, `status` AS `Status`, `sample_id` AS SampleID, `patient_id` AS PatientID, `patient_name`AS PatientName, 
                        `test` AS Request, `bdate` AS DateOfBirth, `sex` AS Sex, DATE_FORMAT(`date`, '%m/%d/%Y') AS DateReceived, `time` AS TimeReceived, 
                        `testtype` AS Section, `sub_section` AS SubSection, `main_id` AS RefID 
                        FROM `tmpWorklist` WHERE (`sample_id` LIKE '" & txtSearch.Text & "%') AND (`status` = 'Checked-In' OR `status` = 'Rejected' OR `status` = 'Cancelled' OR `status` = 'Processing' OR `status` = 'For Verification' OR `status` = 'Verified' OR `status` = 'Validated') 
                        AND (`date` BETWEEN @DateFrom and @DateTo) ORDER BY `id` DESC"

                Dim command As New MySql.Data.MySqlClient.MySqlCommand(SQL, conn)

                command.Parameters.Clear()
                command.Parameters.Add("@DateFrom", MySql.Data.MySqlClient.MySqlDbType.DateTime).Value = Format(dtFrom1.Value, "yyyy-MM-dd")
                command.Parameters.Add("@DateTo", MySql.Data.MySqlClient.MySqlDbType.DateTime).Value = Format(dtTo1.Value, "yyyy-MM-dd")

                Dim adapter As New MySql.Data.MySqlClient.MySqlDataAdapter(command)

                Dim myTable As DataTable = New DataTable
                adapter.Fill(myTable)

                dtList.DataSource = myTable

                GridView.Columns("RefID").Visible = False
                GridView.Columns("Section").Visible = False
                GridView.Columns("SubSection").Visible = False

                ' Make the grid read-only. 
                GridView.OptionsBehavior.Editable = False
                ' Prevent the focused cell from being highlighted. 
                GridView.OptionsSelection.EnableAppearanceFocusedCell = False
                ' Draw a dotted focus rectangle around the entire row. 
                GridView.FocusRectStyle = DrawFocusRectStyle.RowFullFocus

                GridView.OptionsSelection.MultiSelect = True
                GridView.OptionsSelection.MultiSelectMode = GridMultiSelectMode.CheckBoxRowSelect
            ElseIf rgSelect.SelectedIndex = 1 Then
                Dim SQL As String = "SELECT 
                        `id` AS ID, `status` AS `Status`, `sample_id` AS SampleID, `patient_id` AS PatientID, `patient_name`AS PatientName, 
                        `test` AS Request, `bdate` AS DateOfBirth, `sex` AS Sex, DATE_FORMAT(`date`, '%m/%d/%Y') AS DateReceived, `time` AS TimeReceived, 
                        `testtype` AS Section, `sub_section` AS SubSection, `main_id` AS RefID 
                        FROM `tmpWorklist` WHERE (`patient_id` LIKE '" & txtSearch.Text & "%') AND (`status` = 'Checked-In' OR `status` = 'Rejected' OR `status` = 'Cancelled' OR `status` = 'Processing' OR `status` = 'For Verification' OR `status` = 'Verified' OR `status` = 'Validated') 
                        AND (`date` BETWEEN @DateFrom and @DateTo) ORDER BY `id` DESC"

                Dim command As New MySql.Data.MySqlClient.MySqlCommand(SQL, conn)

                command.Parameters.Clear()
                command.Parameters.Add("@DateFrom", MySql.Data.MySqlClient.MySqlDbType.DateTime).Value = Format(dtFrom1.Value, "yyyy-MM-dd")
                command.Parameters.Add("@DateTo", MySql.Data.MySqlClient.MySqlDbType.DateTime).Value = Format(dtTo1.Value, "yyyy-MM-dd")

                Dim adapter As New MySql.Data.MySqlClient.MySqlDataAdapter(command)

                Dim myTable As DataTable = New DataTable
                adapter.Fill(myTable)

                dtList.DataSource = myTable

                GridView.Columns("RefID").Visible = False
                GridView.Columns("Section").Visible = False
                GridView.Columns("SubSection").Visible = False

                ' Make the grid read-only. 
                GridView.OptionsBehavior.Editable = False
                ' Prevent the focused cell from being highlighted. 
                GridView.OptionsSelection.EnableAppearanceFocusedCell = False
                ' Draw a dotted focus rectangle around the entire row. 
                GridView.FocusRectStyle = DrawFocusRectStyle.RowFullFocus

                GridView.OptionsSelection.MultiSelect = True
                GridView.OptionsSelection.MultiSelectMode = GridMultiSelectMode.CheckBoxRowSelect
            ElseIf rgSelect.SelectedIndex = 2 Then
                Dim SQL As String = "SELECT 
                        `id` AS ID, `status` AS `Status`, `sample_id` AS SampleID, `patient_id` AS PatientID, `patient_name`AS PatientName, 
                        `test` AS Request, `bdate` AS DateOfBirth, `sex` AS Sex, DATE_FORMAT(`date`, '%m/%d/%Y') AS DateReceived, `time` AS TimeReceived, 
                        `testtype` AS Section, `sub_section` AS SubSection, `main_id` AS RefID 
                        FROM `tmpWorklist` WHERE (`patient_name` LIKE '" & txtSearch.Text & "%') AND (`status` = 'Checked-In' OR `status` = 'Rejected' OR `status` = 'Cancelled' OR `status` = 'Processing' OR `status` = 'For Verification' OR `status` = 'Verified' OR `status` = 'Validated') 
                        AND (`date` BETWEEN @DateFrom and @DateTo) ORDER BY `id` DESC"

                Dim command As New MySql.Data.MySqlClient.MySqlCommand(SQL, conn)

                command.Parameters.Clear()
                command.Parameters.Add("@DateFrom", MySql.Data.MySqlClient.MySqlDbType.DateTime).Value = Format(dtFrom1.Value, "yyyy-MM-dd")
                command.Parameters.Add("@DateTo", MySql.Data.MySqlClient.MySqlDbType.DateTime).Value = Format(dtTo1.Value, "yyyy-MM-dd")

                Dim adapter As New MySql.Data.MySqlClient.MySqlDataAdapter(command)

                Dim myTable As DataTable = New DataTable
                adapter.Fill(myTable)

                dtList.DataSource = myTable

                GridView.Columns("RefID").Visible = False
                GridView.Columns("Section").Visible = False
                GridView.Columns("SubSection").Visible = False

                ' Make the grid read-only. 
                GridView.OptionsBehavior.Editable = False
                ' Prevent the focused cell from being highlighted. 
                GridView.OptionsSelection.EnableAppearanceFocusedCell = False
                ' Draw a dotted focus rectangle around the entire row. 
                GridView.FocusRectStyle = DrawFocusRectStyle.RowFullFocus

                GridView.OptionsSelection.MultiSelect = True
                GridView.OptionsSelection.MultiSelectMode = GridMultiSelectMode.CheckBoxRowSelect
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message, "System Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub GridView_RowClick(sender As Object, e As RowClickEventArgs) Handles GridView.RowClick
        'LoadRecordsOnLVSQL(lvTest, "SELECT * FROM `patient_order` WHERE `sample_id` = '" & GridView.GetFocusedRowCellValue(GridView.Columns("RefID")) & "' AND testtype = '" & GridView.GetFocusedRowCellValue(GridView.Columns("Section")) & "' AND sub_section = '" & GridView.GetFocusedRowCellValue(GridView.Columns("SubSection")) & "'", 3)

        Dim SQL As String = "SELECT `sample_id` AS `RefID`, `test_name` AS `TestName`, `testtype` AS `Section`, `sub_section` AS `SubSection` FROM `patient_order` WHERE `sample_id` = '" & GridView.GetFocusedRowCellValue(GridView.Columns("RefID")) & "' AND testtype = '" & GridView.GetFocusedRowCellValue(GridView.Columns("Section")) & "' AND sub_section = '" & GridView.GetFocusedRowCellValue(GridView.Columns("SubSection")) & "'"

        Dim command As New MySql.Data.MySqlClient.MySqlCommand(SQL, conn)

        Dim adapter As New MySql.Data.MySqlClient.MySqlDataAdapter(command)

        Dim myTable As DataTable = New DataTable
        adapter.Fill(myTable)

        dtOrderList.DataSource = myTable

        GridViewList.Columns("RefID").Visible = False
        GridViewList.Columns("Section").Visible = False
        GridViewList.Columns("SubSection").Visible = False

        ' Make the grid read-only. 
        GridViewList.OptionsBehavior.Editable = False
        ' Prevent the focused cell from being highlighted. 
        GridViewList.OptionsSelection.EnableAppearanceFocusedCell = False
        ' Draw a dotted focus rectangle around the entire row. 
        GridViewList.FocusRectStyle = DrawFocusRectStyle.RowFullFocus

        Connect()
        rs.Connection = conn
        rs.CommandType = CommandType.Text
        rs.CommandText = "SELECT `comment` FROM `lab_comment` WHERE `sample_id` = '" & GridView.GetFocusedRowCellValue(GridView.Columns("RefID")) & "'"
        reader = rs.ExecuteReader
        reader.Read()
        If reader.HasRows Then
            Me.txtComment.Text = reader(0).ToString
        End If
        Disconnect()
    End Sub


    'Private Sub GridView_FocusedRowChanged(ByVal sender As Object, ByVal e As DevExpress.XtraGrid.Views.Base.FocusedRowChangedEventArgs) Handles GridView.FocusedRowChanged
    '    Dim view As GridView = TryCast(sender, GridView)
    '    view.LayoutChanged()
    'End Sub

    'Private Sub GridView_CalcRowHeight(ByVal sender As Object, ByVal e As RowHeightEventArgs) Handles GridView.CalcRowHeight
    '    Dim view As GridView = TryCast(sender, GridView)
    '    If e.RowHandle = view.FocusedRowHandle Then
    '        e.RowHeight = 50
    '    End If
    'End Sub

End Class