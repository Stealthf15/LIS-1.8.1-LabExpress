Imports System.IO
Imports System.IO.Ports
Imports System.Threading
Imports System.Xml
Imports System.Text.RegularExpressions
Imports System.Drawing.Printing
Imports DevExpress.XtraPrinting.BarCode
Imports DevExpress.XtraGrid.Views.Grid

Public Class frmFecalWorklist

    Public mainID As String = ""
    Public Section As String = ""
    Public SubSection As String = ""
    Public Patient_ID As String = ""
    Public Validation As String = ""
    Dim arrImage() As Byte
    Dim bDate As String
    Dim age As String
    Dim list_Test As String
    Dim patientGender As String

    'For Worklist
    Dim SequenceNo, Status, SampleID, PatientID, PatientName, PatientDoB, PatientAge, PatientSex, Test, ResultDate, ResultTime As String

    'For Result
    Dim ResultName, ResultValue, ResultFlag, ResultOrder, ResultCode, ResultUnit, ResultUnitConv, ResultID, ResultGroup As String

    Private PrintDocType As String = "Barcode"
    Private StrPrinterName As String = My.Settings.BCPrinterName

    Public Sub LoadRecords()
        Try
            GridView.Columns.Clear()
            GridView.Appearance.HeaderPanel.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center
            GridView.Appearance.HeaderPanel.FontStyleDelta = FontStyle.Bold

            Dim SQL As String = "SELECT
                        `tmpWorklist`.`id` AS SequenceNo,
                        `tmpWorklist`.`status` AS Status,
                        `tmpWorklist`.`sample_id` AS SampleID,
                        `tmpWorklist`.`patient_id` AS PatientID,
                        `tmpWorklist`.`patient_name` AS PatientName, 
                        `tmpWorklist`.`test` AS Request,
                        `tmpWorklist`.`bdate` AS DateOfBirth,
                        `tmpWorklist`.`sex` AS Sex,
                        `tmpWorklist`.`age` AS Age,
                        `tmpWorklist`.`dept` AS RoomWard,
                        `tmpWorklist`.`physician` AS Physician,
                        `tmpWorklist`.`medtech` AS PerformedBy,
                        `tmpWorklist`.`verified_by` AS ReleasedBy,
                        DATE_FORMAT(`tmpWorklist`.`date`, '%m/%d/%Y') AS DateReceived,
                        `tmpWorklist`.`time` AS TimeReceived,
                        `tmpWorklist`.`testtype` AS Section,
                        `tmpWorklist`.`sub_section` AS SubSection,
                        DATE_FORMAT(`specimen_tracking`.`extracted`, '%m/%d/%Y %r') AS DateCheckedIn, 
                        `tmpWorklist`.`main_id` AS RefID,
                        `tmpWorklist`.`patient_type` AS PatientType
                        FROM `tmpWorklist` 
                        LEFT JOIN `specimen_tracking` ON
	                        `specimen_tracking`.`sample_id` = `tmpWorklist`.`main_id`
                        WHERE (`tmpWorklist`.`status` = 'Result Received' OR `tmpWorklist`.`status` = 'Validated' OR `tmpWorklist`.`status` = 'Processing' OR `tmpWorklist`.`status` = 'For Verification' OR `tmpWorklist`.`status` = 'Verified') 
                        AND (`tmpWorklist`.`testtype` = `specimen_tracking`.`section`)
                        AND (`tmpWorklist`.`sub_section` = `specimen_tracking`.`sub_section`)
                        AND (`tmpWorklist`.`testtype` = 'Fecalysis')
                        AND (`tmpWorklist`.`location` = @Location)
                        ORDER BY `tmpWorklist`.`status` DESC, `tmpWorklist`.`main_id`"

            Dim command As New MySql.Data.MySqlClient.MySqlCommand(SQL, conn)

            command.Parameters.Clear()
            command.Parameters.AddWithValue("@Location", cboLocation.Text)

            Dim adapter As New MySql.Data.MySqlClient.MySqlDataAdapter(command)

            Dim myTable As DataTable = New DataTable
            adapter.Fill(myTable)

            dtList.DataSource = myTable

            GridView.Columns("PerformedBy").Visible = False
            GridView.Columns("ReleasedBy").Visible = False
            GridView.Columns("PatientType").Visible = False
            GridView.Columns("RefID").Visible = False
            GridView.Columns("Section").Visible = False
            GridView.Columns("SubSection").Visible = False
            GridView.Columns("Age").Visible = False

            ' Make the grid read-only. 
            GridView.OptionsBehavior.Editable = False
            ' Prevent the focused cell from being highlighted. 
            GridView.OptionsSelection.EnableAppearanceFocusedCell = False
            ' Draw a dotted focus rectangle around the entire row. 
            GridView.FocusRectStyle = DrawFocusRectStyle.RowFullFocus

        Catch ex As Exception
            MessageBox.Show(ex.Message, "System Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub GridView_RowCellStyle(sender As Object, e As RowCellStyleEventArgs) Handles GridView.RowCellStyle
        Dim view As GridView = TryCast(sender, GridView)
        If (e.Column.FieldName = "SequenceNo") Or (e.Column.FieldName = "Status") Then
            If view.GetRowCellValue(e.RowHandle, "Status") = "Processing" Then
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
            Else
                e.Appearance.BackColor = Color.Gray
                e.Appearance.BackColor2 = Color.Gray
                e.Appearance.ForeColor = Color.White
            End If
        End If
    End Sub

    Public Sub LoadRecordsCompleted()
        Try
            GridCompleted.Columns.Clear()
            GridCompleted.Appearance.HeaderPanel.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center
            GridCompleted.Appearance.HeaderPanel.FontStyleDelta = FontStyle.Bold

            Dim SQL As String = "SELECT
                        `order`.`id` AS SequenceNo,
                        `order`.`status` AS Status,
                        `order`.`sample_id` AS SampleID,
                        `order`.`patient_id` AS PatientID,
                        `order`.`patient_name` AS PatientName, 
                        `order`.`test` AS Request,
                        `order`.`bdate` AS DateOfBirth,
                        `order`.`sex` AS Sex,
                        `order`.`age` AS Age,
                        `order`.`dept` AS RoomWard,
                        `order`.`physician` AS Physician,
                        `order`.`medtech` AS PerformedBy,
                        `order`.`verified_by` AS ReleasedBy,
                        DATE_FORMAT(`order`.`date`, '%m/%d/%Y') AS DateReceived,
                        `order`.`time` AS TimeReceived,
                        `order`.`testtype` AS Section,
                        `order`.`sub_section` AS SubSection,
                        DATE_FORMAT(`specimen_tracking`.`extracted`, '%m/%d/%Y %r') AS DateCheckedIn, 
                        DATE_FORMAT(str_to_date(`order`.`dt_released`, '%Y-%m-%d %r' ), '%m/%d/%Y %r' ) AS DateReleased,
                        `order`.`main_id` AS RefID,
                        `order`.`patient_type` AS PatientType
                        FROM `order` 
                        LEFT JOIN `specimen_tracking` ON
	                        `specimen_tracking`.`sample_id` = `order`.`main_id`
                        WHERE (`order`.`status` = 'Printed' OR `order`.`status` = 'Validated' OR `order`.`status` = 'Released') 
                        AND (`order`.`testtype` = `specimen_tracking`.`section`)
                        AND (`order`.`sub_section` = `specimen_tracking`.`sub_section`)
                        AND (`order`.`testtype` = 'Fecalysis')
                        AND (`order`.`location` = @Location)
                        AND (DATE(DATE_FORMAT(`order`.`dt_released`, '%Y-%m-%d')) BETWEEN @Date1 AND @Date2)
                        ORDER BY `order`.`main_id` DESC"

            Dim command As New MySql.Data.MySqlClient.MySqlCommand(SQL, conn)

            command.Parameters.Clear()
            command.Parameters.Add("@Date1", MySql.Data.MySqlClient.MySqlDbType.DateTime).Value = Format(dtFrom1.Value, "yyyy-MM-dd")
            command.Parameters.Add("@Date2", MySql.Data.MySqlClient.MySqlDbType.DateTime).Value = Format(dtTo1.Value, "yyyy-MM-dd")
            command.Parameters.AddWithValue("@Location", cboLocation.Text)

            Dim adapter As New MySql.Data.MySqlClient.MySqlDataAdapter(command)

            Dim myTable As DataTable = New DataTable
            adapter.Fill(myTable)

            dtCompleted.DataSource = myTable

            'GridCompleted.Columns("PerformedBy").Visible = False
            'GridCompleted.Columns("ReleasedBy").Visible = False
            GridCompleted.Columns("PatientType").Visible = False
            GridCompleted.Columns("RefID").Visible = False
            GridCompleted.Columns("Section").Visible = False
            GridCompleted.Columns("SubSection").Visible = False
            'GridCompleted.Columns("Age").Visible = False

            ' Make the grid read-only. 
            GridCompleted.OptionsBehavior.Editable = False
            ' Prevent the focused cell from being highlighted. 
            GridCompleted.OptionsSelection.EnableAppearanceFocusedCell = False
            ' Draw a dotted focus rectangle around the entire row. 
            GridCompleted.FocusRectStyle = DrawFocusRectStyle.RowFullFocus

        Catch ex As Exception
            MessageBox.Show(ex.Message, "System Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub GridComplete_RowCellStyle(sender As Object, e As RowCellStyleEventArgs) Handles GridCompleted.RowCellStyle
        Dim view As GridView = TryCast(sender, GridView)
        If (e.Column.FieldName = "SequenceNo") Or (e.Column.FieldName = "Status") Then
            If view.GetRowCellValue(e.RowHandle, "Status") = "Processing" Then
                e.Appearance.BackColor = Color.ForestGreen
                e.Appearance.BackColor2 = Color.ForestGreen
                e.Appearance.ForeColor = Color.White
            ElseIf view.GetRowCellValue(e.RowHandle, "Status") = "Released" Then
                e.Appearance.BackColor = Color.ForestGreen
                e.Appearance.BackColor2 = Color.ForestGreen
                e.Appearance.ForeColor = Color.White
            ElseIf view.GetRowCellValue(e.RowHandle, "Status") = "Validated" Then
                e.Appearance.BackColor = Color.ForestGreen
                e.Appearance.BackColor2 = Color.ForestGreen
                e.Appearance.ForeColor = Color.White
            ElseIf view.GetRowCellValue(e.RowHandle, "Status") = "Printed" Then
                e.Appearance.BackColor = Color.ForestGreen
                e.Appearance.BackColor2 = Color.ForestGreen
                e.Appearance.ForeColor = Color.White
            Else
                e.Appearance.BackColor = Color.Gray
                e.Appearance.BackColor2 = Color.Gray
                e.Appearance.ForeColor = Color.White
            End If
        End If
    End Sub

    Private Sub txtSearch_TextChanged(sender As Object, e As EventArgs) Handles txtSearch.TextChanged
        Try
            Dim SQL As String

            GridView.Columns.Clear()
            GridView.Appearance.HeaderPanel.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center
            GridView.Appearance.HeaderPanel.FontStyleDelta = FontStyle.Bold

            If rgSelect.SelectedIndex = 0 Then
                'Search SampleID
                SQL = "SELECT
                        `tmpWorklist`.`id` AS SequenceNo,
                        `tmpWorklist`.`status` AS Status,
                        `tmpWorklist`.`sample_id` AS SampleID,
                        `tmpWorklist`.`patient_id` AS PatientID,
                        `tmpWorklist`.`patient_name` AS PatientName, 
                        `tmpWorklist`.`test` AS Request,
                        `tmpWorklist`.`bdate` AS DateOfBirth,
                        `tmpWorklist`.`sex` AS Sex,
                        `tmpWorklist`.`age` AS Age,
                        `tmpWorklist`.`dept` AS RoomWard,
                        `tmpWorklist`.`physician` AS Physician,
                        `tmpWorklist`.`medtech` AS PerformedBy,
                        `tmpWorklist`.`verified_by` AS ReleasedBy,
                        DATE_FORMAT(`tmpWorklist`.`date`, '%m/%d/%Y') AS DateReceived,
                        `tmpWorklist`.`time` AS TimeReceived,
                        `tmpWorklist`.`testtype` AS Section,
                        `tmpWorklist`.`sub_section` AS SubSection,
                        DATE_FORMAT(`specimen_tracking`.`extracted`, '%m/%d/%Y %r') AS DateCheckedIn, 
                        `tmpWorklist`.`main_id` AS RefID,
                        `tmpWorklist`.`patient_type` AS PatientType
                        FROM `tmpWorklist` 
                        LEFT JOIN `specimen_tracking` ON
	                        `specimen_tracking`.`sample_id` = `tmpWorklist`.`main_id`
                        WHERE (`tmpWorklist`.`status` = 'Result Received' OR `tmpWorklist`.`status` = 'Validated' OR `tmpWorklist`.`status` = 'Processing' OR `tmpWorklist`.`status` = 'For Verification' OR `tmpWorklist`.`status` = 'Verified')
                        AND (`tmpworklist`.`sample_id` LIKE '" & txtSearch.Text & "%')
                        AND (`tmpWorklist`.`testtype` = `specimen_tracking`.`section`)
                        AND (`tmpWorklist`.`sub_section` = `specimen_tracking`.`sub_section`)
                        AND (`tmpWorklist`.`testtype` = 'Fecalysis')
                        AND (`tmpWorklist`.`location` = @Location)
                        ORDER BY `tmpWorklist`.`main_id` DESC"

                Dim command As New MySql.Data.MySqlClient.MySqlCommand(SQL, conn)

                command.Parameters.Clear()
                command.Parameters.AddWithValue("@Location", cboLocation.Text)

                Dim adapter As New MySql.Data.MySqlClient.MySqlDataAdapter(command)

                Dim myTable As DataTable = New DataTable
                adapter.Fill(myTable)

                dtList.DataSource = myTable

                GridView.Columns("PerformedBy").Visible = False
                GridView.Columns("ReleasedBy").Visible = False
                GridView.Columns("PatientType").Visible = False
                GridView.Columns("RefID").Visible = False
                GridView.Columns("Section").Visible = False
                GridView.Columns("SubSection").Visible = False
                GridView.Columns("Age").Visible = False

                ' Make the grid read-only. 
                GridView.OptionsBehavior.Editable = False
                ' Prevent the focused cell from being highlighted. 
                GridView.OptionsSelection.EnableAppearanceFocusedCell = False
                ' Draw a dotted focus rectangle around the entire row. 
                GridView.FocusRectStyle = DrawFocusRectStyle.RowFullFocus

                GridView.OptionsSelection.MultiSelect = True
                GridView.OptionsSelection.MultiSelectMode = GridMultiSelectMode.RowSelect
            ElseIf rgSelect.SelectedIndex = 1 Then
                'Search PatientID
                SQL = "SELECT
                        `tmpWorklist`.`id` AS SequenceNo,
                        `tmpWorklist`.`status` AS Status,
                        `tmpWorklist`.`sample_id` AS SampleID,
                        `tmpWorklist`.`patient_id` AS PatientID,
                        `tmpWorklist`.`patient_name` AS PatientName, 
                        `tmpWorklist`.`test` AS Request,
                        `tmpWorklist`.`bdate` AS DateOfBirth,
                        `tmpWorklist`.`sex` AS Sex,
                        `tmpWorklist`.`age` AS Age,
                        `tmpWorklist`.`dept` AS RoomWard,
                        `tmpWorklist`.`physician` AS Physician,
                        `tmpWorklist`.`medtech` AS PerformedBy,
                        `tmpWorklist`.`verified_by` AS ReleasedBy,
                        DATE_FORMAT(`tmpWorklist`.`date`, '%m/%d/%Y') AS DateReceived,
                        `tmpWorklist`.`time` AS TimeReceived,
                        `tmpWorklist`.`testtype` AS Section,
                        `tmpWorklist`.`sub_section` AS SubSection,
                        DATE_FORMAT(`specimen_tracking`.`extracted`, '%m/%d/%Y %r') AS DateCheckedIn, 
                        `tmpWorklist`.`main_id` AS RefID,
                        `tmpWorklist`.`patient_type` AS PatientType
                        FROM `tmpWorklist` 
                        LEFT JOIN `specimen_tracking` ON
	                        `specimen_tracking`.`sample_id` = `tmpWorklist`.`main_id`
                        WHERE (`tmpWorklist`.`status` = 'Result Received' OR `tmpWorklist`.`status` = 'Validated' OR `tmpWorklist`.`status` = 'Processing' OR `tmpWorklist`.`status` = 'For Verification' OR `tmpWorklist`.`status` = 'Verified')
                        AND (`tmpworklist`.`patient_id` LIKE '" & txtSearch.Text & "%')
                        AND (`tmpWorklist`.`testtype` = `specimen_tracking`.`section`)
                        AND (`tmpWorklist`.`sub_section` = `specimen_tracking`.`sub_section`)
                        AND (`tmpWorklist`.`testtype` = 'Fecalysis')
                        AND (`tmpWorklist`.`location` = @Location)
                        ORDER BY `tmpWorklist`.`main_id` DESC"

                Dim command As New MySql.Data.MySqlClient.MySqlCommand(SQL, conn)

                command.Parameters.Clear()
                command.Parameters.AddWithValue("@Location", cboLocation.Text)

                Dim adapter As New MySql.Data.MySqlClient.MySqlDataAdapter(command)

                Dim myTable As DataTable = New DataTable
                adapter.Fill(myTable)

                dtList.DataSource = myTable

                GridView.Columns("PerformedBy").Visible = False
                GridView.Columns("ReleasedBy").Visible = False
                GridView.Columns("PatientType").Visible = False
                GridView.Columns("RefID").Visible = False
                GridView.Columns("Section").Visible = False
                GridView.Columns("SubSection").Visible = False
                GridView.Columns("Age").Visible = False

                ' Make the grid read-only. 
                GridView.OptionsBehavior.Editable = False
                ' Prevent the focused cell from being highlighted. 
                GridView.OptionsSelection.EnableAppearanceFocusedCell = False
                ' Draw a dotted focus rectangle around the entire row. 
                GridView.FocusRectStyle = DrawFocusRectStyle.RowFullFocus

                GridView.OptionsSelection.MultiSelect = True
                GridView.OptionsSelection.MultiSelectMode = GridMultiSelectMode.RowSelect
            ElseIf rgSelect.SelectedIndex = 2 Then
                'Search Patient Name
                SQL = "SELECT
                        `tmpWorklist`.`id` AS SequenceNo,
                        `tmpWorklist`.`status` AS Status,
                        `tmpWorklist`.`sample_id` AS SampleID,
                        `tmpWorklist`.`patient_id` AS PatientID,
                        `tmpWorklist`.`patient_name` AS PatientName, 
                        `tmpWorklist`.`test` AS Request,
                        `tmpWorklist`.`bdate` AS DateOfBirth,
                        `tmpWorklist`.`sex` AS Sex,
                        `tmpWorklist`.`age` AS Age,
                        `tmpWorklist`.`dept` AS RoomWard,
                        `tmpWorklist`.`physician` AS Physician,
                        `tmpWorklist`.`medtech` AS PerformedBy,
                        `tmpWorklist`.`verified_by` AS ReleasedBy,
                        DATE_FORMAT(`tmpWorklist`.`date`, '%m/%d/%Y') AS DateReceived,
                        `tmpWorklist`.`time` AS TimeReceived,
                        `tmpWorklist`.`testtype` AS Section,
                        `tmpWorklist`.`sub_section` AS SubSection,
                        DATE_FORMAT(`specimen_tracking`.`extracted`, '%m/%d/%Y %r') AS DateCheckedIn, 
                        `tmpWorklist`.`main_id` AS RefID,
                        `tmpWorklist`.`patient_type` AS PatientType
                        FROM `tmpWorklist` 
                        LEFT JOIN `specimen_tracking` ON
	                        `specimen_tracking`.`sample_id` = `tmpWorklist`.`main_id`
                        WHERE (`tmpWorklist`.`status` = 'Result Received' OR `tmpWorklist`.`status` = 'Validated' OR `tmpWorklist`.`status` = 'Processing' OR `tmpWorklist`.`status` = 'For Verification' OR `tmpWorklist`.`status` = 'Verified')
                        AND (`tmpworklist`.`patient_name` LIKE '" & txtSearch.Text & "%')
                        AND (`tmpWorklist`.`testtype` = `specimen_tracking`.`section`)
                        AND (`tmpWorklist`.`sub_section` = `specimen_tracking`.`sub_section`)
                        AND (`tmpWorklist`.`testtype` = 'Fecalysis')
                        AND (`tmpWorklist`.`location` = @Location)
                        ORDER BY `tmpWorklist`.`main_id` DESC"

                Dim command As New MySql.Data.MySqlClient.MySqlCommand(SQL, conn)

                command.Parameters.Clear()
                command.Parameters.AddWithValue("@Location", cboLocation.Text)

                Dim adapter As New MySql.Data.MySqlClient.MySqlDataAdapter(command)

                Dim myTable As DataTable = New DataTable
                adapter.Fill(myTable)

                dtList.DataSource = myTable

                GridView.Columns("PerformedBy").Visible = False
                GridView.Columns("ReleasedBy").Visible = False
                GridView.Columns("PatientType").Visible = False
                GridView.Columns("RefID").Visible = False
                GridView.Columns("Section").Visible = False
                GridView.Columns("SubSection").Visible = False
                GridView.Columns("Age").Visible = False

                ' Make the grid read-only. 
                GridView.OptionsBehavior.Editable = False
                ' Prevent the focused cell from being highlighted. 
                GridView.OptionsSelection.EnableAppearanceFocusedCell = False
                ' Draw a dotted focus rectangle around the entire row. 
                GridView.FocusRectStyle = DrawFocusRectStyle.RowFullFocus

                GridView.OptionsSelection.MultiSelect = True
                GridView.OptionsSelection.MultiSelectMode = GridMultiSelectMode.RowSelect
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message, "System Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub txtSearch1_TextChanged(sender As Object, e As EventArgs) Handles txtSearch1.TextChanged
        Try
            GridCompleted.Columns.Clear()
            GridCompleted.Appearance.HeaderPanel.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center
            GridCompleted.Appearance.HeaderPanel.FontStyleDelta = FontStyle.Bold

            If rgSelect1.SelectedIndex = 0 Then
                Dim SQL As String = "SELECT
                        `order`.`id` AS SequenceNo,
                        `order`.`status` AS Status,
                        `order`.`sample_id` AS SampleID,
                        `order`.`patient_id` AS PatientID,
                        `order`.`patient_name` AS PatientName, 
                        `order`.`test` AS Request,
                        `order`.`bdate` AS DateOfBirth,
                        `order`.`sex` AS Sex,
                        `order`.`age` AS Age,
                        `order`.`dept` AS RoomWard,
                        `order`.`physician` AS Physician,
                        `order`.`medtech` AS PerformedBy,
                        `order`.`verified_by` AS ReleasedBy,
                        DATE_FORMAT(`order`.`date`, '%m/%d/%Y') AS DateReceived,
                        `order`.`time` AS TimeReceived,
                        `order`.`testtype` AS Section,
                        `order`.`sub_section` AS SubSection,
                        DATE_FORMAT(`specimen_tracking`.`extracted`, '%m/%d/%Y %r') AS DateCheckedIn, 
                        DATE_FORMAT(str_to_date(`order`.`dt_released`, '%Y-%m-%d %r' ), '%m/%d/%Y %r' ) AS DateReleased,
                        `order`.`main_id` AS RefID,
                        `order`.`patient_type` AS PatientType
                        FROM `order` 
                        LEFT JOIN `specimen_tracking` ON
	                        `specimen_tracking`.`sample_id` = `order`.`main_id`
                        WHERE (`order`.`status` = 'Printed' OR `order`.`status` = 'Validated' OR `order`.`status` = 'Released') 
                        AND (`order`.`sample_id` LIKE '" & txtSearch1.Text & "%')
                        AND (`order`.`testtype` = `specimen_tracking`.`section`)
                        AND (`order`.`sub_section` = `specimen_tracking`.`sub_section`)
                        AND (`order`.`testtype` = 'Fecalysis')
                        AND (`order`.`location` = @Location)
                        AND (DATE(DATE_FORMAT(`order`.`dt_released`, '%Y-%m-%d')) BETWEEN @Date1 AND @Date2)
                        ORDER BY `order`.`main_id` DESC"

                Dim command As New MySql.Data.MySqlClient.MySqlCommand(SQL, conn)

                command.Parameters.Clear()
                command.Parameters.Add("@Date1", MySql.Data.MySqlClient.MySqlDbType.DateTime).Value = Format(dtFrom1.Value, "yyyy-MM-dd")
                command.Parameters.Add("@Date2", MySql.Data.MySqlClient.MySqlDbType.DateTime).Value = Format(dtTo1.Value, "yyyy-MM-dd")
                command.Parameters.AddWithValue("@Location", cboLocation1.Text)

                Dim adapter As New MySql.Data.MySqlClient.MySqlDataAdapter(command)

                Dim myTable As DataTable = New DataTable
                adapter.Fill(myTable)

                dtCompleted.DataSource = myTable

                'GridCompleted.Columns("PerformedBy").Visible = False
                'GridCompleted.Columns("ReleasedBy").Visible = False
                GridCompleted.Columns("PatientType").Visible = False
                GridCompleted.Columns("RefID").Visible = False
                GridCompleted.Columns("Section").Visible = False
                GridCompleted.Columns("SubSection").Visible = False
                'GridCompleted.Columns("Age").Visible = False

                ' Make the grid read-only. 
                GridCompleted.OptionsBehavior.Editable = False
                ' Prevent the focused cell from being highlighted. 
                GridCompleted.OptionsSelection.EnableAppearanceFocusedCell = False
                ' Draw a dotted focus rectangle around the entire row. 
                GridCompleted.FocusRectStyle = DrawFocusRectStyle.RowFullFocus

                GridCompleted.OptionsSelection.MultiSelect = True
                GridCompleted.OptionsSelection.MultiSelectMode = GridMultiSelectMode.RowSelect
            ElseIf rgSelect1.SelectedIndex = 1 Then
                Dim SQL As String = "SELECT
                        `order`.`id` AS SequenceNo,
                        `order`.`status` AS Status,
                        `order`.`sample_id` AS SampleID,
                        `order`.`patient_id` AS PatientID,
                        `order`.`patient_name` AS PatientName, 
                        `order`.`test` AS Request,
                        `order`.`bdate` AS DateOfBirth,
                        `order`.`sex` AS Sex,
                        `order`.`age` AS Age,
                        `order`.`dept` AS RoomWard,
                        `order`.`physician` AS Physician,
                        `order`.`medtech` AS PerformedBy,
                        `order`.`verified_by` AS ReleasedBy,
                        DATE_FORMAT(`order`.`date`, '%m/%d/%Y') AS DateReceived,
                        `order`.`time` AS TimeReceived,
                        `order`.`testtype` AS Section,
                        `order`.`sub_section` AS SubSection,
                        DATE_FORMAT(`specimen_tracking`.`extracted`, '%m/%d/%Y %r') AS DateCheckedIn, 
                        DATE_FORMAT(str_to_date(`order`.`dt_released`, '%Y-%m-%d %r' ), '%m/%d/%Y %r' ) AS DateReleased,
                        `order`.`main_id` AS RefID,
                        `order`.`patient_type` AS PatientType
                        FROM `order` 
                        LEFT JOIN `specimen_tracking` ON
	                        `specimen_tracking`.`sample_id` = `order`.`main_id`
                        WHERE (`order`.`status` = 'Printed' OR `order`.`status` = 'Validated' OR `order`.`status` = 'Released') 
                        AND (`order`.`patient_id` LIKE '" & txtSearch1.Text & "%')
                        AND (`order`.`testtype` = `specimen_tracking`.`section`)
                        AND (`order`.`sub_section` = `specimen_tracking`.`sub_section`)
                        AND (`order`.`testtype` = 'Fecalysis')
                        AND (`order`.`location` = @Location)
                        AND (DATE(DATE_FORMAT(`order`.`dt_released`, '%Y-%m-%d')) BETWEEN @Date1 AND @Date2)
                        ORDER BY `order`.`main_id` DESC"

                Dim command As New MySql.Data.MySqlClient.MySqlCommand(SQL, conn)

                command.Parameters.Clear()
                command.Parameters.Add("@Date1", MySql.Data.MySqlClient.MySqlDbType.DateTime).Value = Format(dtFrom1.Value, "yyyy-MM-dd")
                command.Parameters.Add("@Date2", MySql.Data.MySqlClient.MySqlDbType.DateTime).Value = Format(dtTo1.Value, "yyyy-MM-dd")
                command.Parameters.AddWithValue("@Location", cboLocation1.Text)

                Dim adapter As New MySql.Data.MySqlClient.MySqlDataAdapter(command)

                Dim myTable As DataTable = New DataTable
                adapter.Fill(myTable)

                dtCompleted.DataSource = myTable

                'GridCompleted.Columns("PerformedBy").Visible = False
                'GridCompleted.Columns("ReleasedBy").Visible = False
                GridCompleted.Columns("PatientType").Visible = False
                GridCompleted.Columns("RefID").Visible = False
                GridCompleted.Columns("Section").Visible = False
                GridCompleted.Columns("SubSection").Visible = False
                'GridCompleted.Columns("Age").Visible = False

                ' Make the grid read-only. 
                GridCompleted.OptionsBehavior.Editable = False
                ' Prevent the focused cell from being highlighted. 
                GridCompleted.OptionsSelection.EnableAppearanceFocusedCell = False
                ' Draw a dotted focus rectangle around the entire row. 
                GridCompleted.FocusRectStyle = DrawFocusRectStyle.RowFullFocus

                GridCompleted.OptionsSelection.MultiSelect = True
                GridCompleted.OptionsSelection.MultiSelectMode = GridMultiSelectMode.RowSelect
            ElseIf rgSelect1.SelectedIndex = 2 Then
                Dim SQL As String = "SELECT
                        `order`.`id` AS SequenceNo,
                        `order`.`status` AS Status,
                        `order`.`sample_id` AS SampleID,
                        `order`.`patient_id` AS PatientID,
                        `order`.`patient_name` AS PatientName, 
                        `order`.`test` AS Request,
                        `order`.`bdate` AS DateOfBirth,
                        `order`.`sex` AS Sex,
                        `order`.`age` AS Age,
                        `order`.`dept` AS RoomWard,
                        `order`.`physician` AS Physician,
                        `order`.`medtech` AS PerformedBy,
                        `order`.`verified_by` AS ReleasedBy,
                        DATE_FORMAT(`order`.`date`, '%m/%d/%Y') AS DateReceived,
                        `order`.`time` AS TimeReceived,
                        `order`.`testtype` AS Section,
                        `order`.`sub_section` AS SubSection,
                        DATE_FORMAT(`specimen_tracking`.`extracted`, '%m/%d/%Y %r') AS DateCheckedIn, 
                        DATE_FORMAT(str_to_date(`order`.`dt_released`, '%Y-%m-%d %r' ), '%m/%d/%Y %r' ) AS DateReleased,
                        `order`.`main_id` AS RefID,
                        `order`.`patient_type` AS PatientType
                        FROM `order` 
                        LEFT JOIN `specimen_tracking` ON
	                        `specimen_tracking`.`sample_id` = `order`.`main_id`
                        WHERE (`order`.`status` = 'Printed' OR `order`.`status` = 'Validated' OR `order`.`status` = 'Released') 
                        AND (`order`.`patient_name` LIKE '" & txtSearch1.Text & "%')
                        AND (`order`.`testtype` = `specimen_tracking`.`section`)
                        AND (`order`.`sub_section` = `specimen_tracking`.`sub_section`)
                        AND (`order`.`testtype` = 'Fecalysis')
                        AND (`order`.`location` = @Location)
                        AND (DATE(DATE_FORMAT(`order`.`dt_released`, '%Y-%m-%d')) BETWEEN @Date1 AND @Date2)
                        ORDER BY `order`.`main_id` DESC"

                Dim command As New MySql.Data.MySqlClient.MySqlCommand(SQL, conn)

                command.Parameters.Clear()
                command.Parameters.Add("@Date1", MySql.Data.MySqlClient.MySqlDbType.DateTime).Value = Format(dtFrom1.Value, "yyyy-MM-dd")
                command.Parameters.Add("@Date2", MySql.Data.MySqlClient.MySqlDbType.DateTime).Value = Format(dtTo1.Value, "yyyy-MM-dd")
                command.Parameters.AddWithValue("@Location", cboLocation1.Text)

                Dim adapter As New MySql.Data.MySqlClient.MySqlDataAdapter(command)

                Dim myTable As DataTable = New DataTable
                adapter.Fill(myTable)

                dtCompleted.DataSource = myTable

                'GridCompleted.Columns("PerformedBy").Visible = False
                'GridCompleted.Columns("ReleasedBy").Visible = False
                GridCompleted.Columns("PatientType").Visible = False
                GridCompleted.Columns("RefID").Visible = False
                GridCompleted.Columns("Section").Visible = False
                GridCompleted.Columns("SubSection").Visible = False
                'GridCompleted.Columns("Age").Visible = False

                ' Make the grid read-only. 
                GridCompleted.OptionsBehavior.Editable = False
                ' Prevent the focused cell from being highlighted. 
                GridCompleted.OptionsSelection.EnableAppearanceFocusedCell = False
                ' Draw a dotted focus rectangle around the entire row. 
                GridCompleted.FocusRectStyle = DrawFocusRectStyle.RowFullFocus

                GridCompleted.OptionsSelection.MultiSelect = True
                GridCompleted.OptionsSelection.MultiSelectMode = GridMultiSelectMode.RowSelect
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message, "System Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub frmNewOrder_FormClosed(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
        Me.Dispose()
    End Sub

    Private Sub frmNewOrder_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Me.cboLocation.Properties.Items.Clear()
        Me.cboLocation1.Properties.Items.Clear()

        Connect()
        rs.Connection = conn
        rs.CommandType = CommandType.Text
        rs.CommandText = "SELECT `location` FROM `location` ORDER BY `id`"
        reader = rs.ExecuteReader
        While reader.Read
            Me.cboLocation.Properties.Items.Add(reader(0))
            Me.cboLocation1.Properties.Items.Add(reader(0))
        End While
        Disconnect()

        'Load Location Automatically
        cboLocation.Text = CurrDept
        cboLocation1.Text = CurrDept

        LoadRecords()
        LoadRecordsCompleted()

        DisablePermission()
    End Sub

    Private Sub btnView_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnView.ItemClick
        If XTab.SelectedTabPageIndex = 0 Then
            'If Not GridView.GetFocusedRowCellValue(GridView.Columns("Status")) = "Processing" Then
            '    DisplayResult()
            'End If
            DisplayResult()
        ElseIf XTab.SelectedTabPageIndex = 1 Then
            DisplayResultCompleted()
        End If
    End Sub

    Private Sub btnRefresh_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnRefresh.ItemClick
        LoadRecords()
        LoadRecordsCompleted()
    End Sub

    Private Sub btnClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnClose.ItemClick
        Me.Close()
        Me.Dispose()
    End Sub

    Private Sub dtTo_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles dtTo.ValueChanged, dtFrom.ValueChanged
        LoadRecords()
    End Sub

    Private Sub cboLimit_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cboLocation.SelectedIndexChanged
        Try
            LoadRecords()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "System Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub cboLimit1_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cboLocation1.SelectedIndexChanged
        Try
            LoadRecordsCompleted()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "System Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub GridView_DoubleClick(sender As Object, e As EventArgs) Handles GridView.DoubleClick
        'If Not GridView.GetFocusedRowCellValue(GridView.Columns("Status")) = "Processing" Then
        '    DisplayResult()
        'End If
        DisplayResult()
    End Sub

    Private Sub GridCompleted_DoubleClick(sender As Object, e As EventArgs) Handles GridCompleted.DoubleClick
        DisplayResultCompleted()
    End Sub


    Private Sub btnSearch_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSearch.Click
        LoadRecordsCompleted()
    End Sub

    Private Sub XTab_SelectedPageChanged(ByVal sender As Object, ByVal e As DevExpress.XtraTab.TabPageChangedEventArgs) Handles XTab.SelectedPageChanged
        If XTab.SelectedTabPageIndex = 0 Then
            btnDelete.Enabled = True
            btnView.Enabled = True
            btnRefresh.Enabled = True
            btnBarcode.Enabled = True
            btnPrint.Enabled = False
            lblCountQueue.Text = "Record Count: " & GridView.RowCount
        ElseIf XTab.SelectedTabPageIndex = 1 Then
            btnView.Enabled = True
            btnRefresh.Enabled = True
            btnDelete.Enabled = False
            btnBarcode.Enabled = False
            btnPrint.Enabled = True
            btnReject.Enabled = False
            lblCountQueue.Text = "Record Count: " & GridCompleted.RowCount
        End If
    End Sub

    Private Sub tmLoadNew_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tmLoadNew.Tick
        LoadNewResult()
    End Sub

    Private Sub LoadNewResult()
        Try
            Connect()
            rs.Connection = conn
            rs.CommandType = CommandType.Text
            rs.CommandText = "SELECT `status` FROM `tmpresultstatus` WHERE `status` = 'New Result'"
            reader = rs.ExecuteReader
            reader.Read()
            If reader.HasRows Then
                Disconnect()
                LoadRecords()
                UpdateResultStatus()
            End If
            Disconnect()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
        End Try
    End Sub

    Private Sub UpdateResultStatus()
        Connect()
        rs.Connection = conn
        rs.CommandType = CommandType.Text
        rs.CommandText = "UPDATE `tmpresultstatus` SET `status` = 'No Result'"
        rs.ExecuteNonQuery()
        Disconnect()
    End Sub

    Private Sub btnPrint_ItemClick(ByVal sender As Object, ByVal e As DevExpress.XtraBars.ItemClickEventArgs) Handles btnPrint.ItemClick
        Try
            If XTab.SelectedTabPageIndex = 0 Then

            ElseIf XTab.SelectedTabPageIndex = 1 Then
                Using myRDLCPrinter As New RDLCPrinterPrintReleased(GridCompleted.GetFocusedRowCellValue(GridCompleted.Columns("RefID")).ToString,
                                                                    GridCompleted.GetFocusedRowCellValue(GridCompleted.Columns("Section")).ToString,
                                                                    GridCompleted.GetFocusedRowCellValue(GridCompleted.Columns("SubSection")).ToString,
                                                                    "", My.Settings.DefaultPrinter)
                    If My.Settings.SaveAsPDF Then
                        Dim byteViewer As Byte() = myRDLCPrinter.LocalReport.Render("PDF")
                        Dim saveFileDialog1 As New SaveFileDialog()
                        saveFileDialog1.Filter = "*PDF files (*.pdf)|*.pdf"
                        saveFileDialog1.FilterIndex = 2
                        saveFileDialog1.RestoreDirectory = True
                        Dim newFile As New FileStream(My.Settings.PDFLocation & GridCompleted.GetFocusedRowCellValue(GridCompleted.Columns("SampleID")).ToString &
                                                      "_" & GridCompleted.GetFocusedRowCellValue(GridCompleted.Columns("PatientName")).ToString & ".pdf", FileMode.Create)
                        newFile.Write(byteViewer, 0, byteViewer.Length)
                        newFile.Close()

                        myRDLCPrinter.Print(1)
                    Else
                        myRDLCPrinter.Print(1)
                    End If
                End Using
                'Activity Logs
                ActivityLogs(GridCompleted.GetFocusedRowCellValue(GridCompleted.Columns("SampleID")).ToString,
                             GridCompleted.GetFocusedRowCellValue(GridCompleted.Columns("PatientID")).ToString,
                             GridCompleted.GetFocusedRowCellValue(GridCompleted.Columns("PatientName")).ToString,
                             CurrUser,
                             "Print Result",
                             GridCompleted.GetFocusedRowCellValue(GridCompleted.Columns("Request")).ToString,
                             GridCompleted.GetFocusedRowCellValue(GridCompleted.Columns("Section")).ToString,
                             GridCompleted.GetFocusedRowCellValue(GridCompleted.Columns("SubSection")).ToString)
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message, "System Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Exit Sub
        End Try
    End Sub

    Private Sub btnDelete_ItemClick(ByVal sender As Object, ByVal e As DevExpress.XtraBars.ItemClickEventArgs) Handles btnDelete.ItemClick
        If MessageBox.Show("Are you sure you want to reject " & GridView.GetFocusedRowCellValue(GridView.Columns("PatientName")).ToString & " order?", "Confirm Reject", MessageBoxButtons.YesNo, MessageBoxIcon.Question) = vbYes Then
            frmRejectOrder.ID = GridView.GetFocusedRowCellValue(GridView.Columns("SequenceNo")).ToString
            frmRejectOrder.sID = GridView.GetFocusedRowCellValue(GridView.Columns("SampleID")).ToString
            frmRejectOrder.pID = GridView.GetFocusedRowCellValue(GridView.Columns("PatientID")).ToString
            frmRejectOrder.pName = GridView.GetFocusedRowCellValue(GridView.Columns("PatientName")).ToString
            frmRejectOrder.pTest = GridView.GetFocusedRowCellValue(GridView.Columns("Request")).ToString
            frmRejectOrder.pSection = GridView.GetFocusedRowCellValue(GridView.Columns("Section")).ToString
            frmRejectOrder.pSubSection = GridView.GetFocusedRowCellValue(GridView.Columns("SubSection")).ToString
            frmRejectOrder.ShowDialog()
        End If
        LoadRecords()
    End Sub

    Private Sub btnReject_ItemClick(ByVal sender As Object, ByVal e As DevExpress.XtraBars.ItemClickEventArgs) Handles btnReject.ItemClick
        If MessageBox.Show("Are you sure you want to cancel " & GridView.GetFocusedRowCellValue(GridView.Columns("PatientName")) & " order?", "Confirm Cancel", MessageBoxButtons.YesNo, MessageBoxIcon.Question) = vbYes Then
            frmCancelOR.ID = GridView.GetFocusedRowCellValue(GridView.Columns("SequenceNo")).ToString
            frmCancelOR.sID = GridView.GetFocusedRowCellValue(GridView.Columns("SampleID")).ToString
            frmCancelOR.pID = GridView.GetFocusedRowCellValue(GridView.Columns("PatientID")).ToString
            frmCancelOR.pName = GridView.GetFocusedRowCellValue(GridView.Columns("PatientName")).ToString
            frmCancelOR.pTest = GridView.GetFocusedRowCellValue(GridView.Columns("Request")).ToString
            frmCancelOR.pSection = GridView.GetFocusedRowCellValue(GridView.Columns("Section")).ToString
            frmCancelOR.pSubSection = GridView.GetFocusedRowCellValue(GridView.Columns("SubSection")).ToString
            frmCancelOR.ShowDialog()
        End If
        LoadRecords()
    End Sub

    Private Sub btnBarcode_ItemClick(ByVal sender As Object, ByVal e As DevExpress.XtraBars.ItemClickEventArgs) Handles btnBarcode.ItemClick

        PrintBarcode(GridView.GetFocusedRowCellValue(GridView.Columns("Request")).ToString,
                            GridView.GetFocusedRowCellValue(GridView.Columns("SampleID")).ToString,
                            GridView.GetFocusedRowCellValue(GridView.Columns("PatientID")).ToString,
                            GridView.GetFocusedRowCellValue(GridView.Columns("PatientName")).ToString,
                            GridView.GetFocusedRowCellValue(GridView.Columns("DateOfBirth")).ToString,
                            GridView.GetFocusedRowCellValue(GridView.Columns("Sex")).ToString,
                            GridView.GetFocusedRowCellValue(GridView.Columns("Section")).ToString,
                            GridView.GetFocusedRowCellValue(GridView.Columns("SubSection")).ToString)
        'Activity Logs
        ActivityLogs(GridView.GetFocusedRowCellValue(GridView.Columns("SampleID")).ToString,
                            GridView.GetFocusedRowCellValue(GridView.Columns("PatientID")).ToString,
                            GridView.GetFocusedRowCellValue(GridView.Columns("PatientName")).ToString,
                            CurrUser,
                            "Reprint Barcode",
                            GridView.GetFocusedRowCellValue(GridView.Columns("Request")).ToString,
                            GridView.GetFocusedRowCellValue(GridView.Columns("Section")).ToString,
                            GridView.GetFocusedRowCellValue(GridView.Columns("SubSection")).ToString)

    End Sub

    Private Sub frm_Activated(sender As Object, e As EventArgs) Handles Me.Activated
        'MainFOrm.aceFecal.Appearance.Normal.BackColor = Color.FromArgb(6, 31, 71)
        MainFOrm.aceFecal.Appearance.Normal.BackColor = Color.FromArgb(16, 110, 190)
        MainFOrm.aceFecal.Appearance.Normal.ForeColor = Color.FromArgb(255, 255, 255)
    End Sub

    Private Sub frm_FormClosed(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
        MainFOrm.lblTitle.Text = ""
        MainFOrm.aceFecal.Appearance.Normal.BackColor = Color.FromArgb(240, 240, 240)
        MainFOrm.aceFecal.Appearance.Normal.ForeColor = Color.FromArgb(27, 41, 62)
        Me.Dispose()
    End Sub

#Region "DisplayResults"
    '############################################-----For On-Queue Orders-----############################################
    Private Sub DisplayResult()
        'On Error Resume Next
        frmFecalNew.cboPathologist.Properties.Items.Clear()
        frmFecalNew.cboMedTech.Properties.Items.Clear()
        frmFecalNew.cboVerify.Properties.Items.Clear()

        If GridView.GetFocusedRowCellValue(GridView.Columns("Status")).ToString = "For Verification" Then
            frmFecalNew.btnVerification.Caption = "Verify"
        End If

        '###########################---Load Basic Patient Details---######################################################
        frmFecalNew.status = GridView.GetFocusedRowCellValue(GridView.Columns("Status")).ToString
        frmFecalNew.mainID = GridView.GetFocusedRowCellValue(GridView.Columns("RefID")).ToString
        frmFecalNew.Section = GridView.GetFocusedRowCellValue(GridView.Columns("Section")).ToString
        frmFecalNew.SubSection = GridView.GetFocusedRowCellValue(GridView.Columns("SubSection")).ToString
        frmFecalNew.PatientID = GridView.GetFocusedRowCellValue(GridView.Columns("PatientID")).ToString
        frmFecalNew.txtSampleID.Text = GridView.GetFocusedRowCellValue(GridView.Columns("SampleID")).ToString
        frmFecalNew.txtPatientID.Text = GridView.GetFocusedRowCellValue(GridView.Columns("PatientID")).ToString
        frmFecalNew.txtName.Text = GridView.GetFocusedRowCellValue(GridView.Columns("PatientName")).ToString
        frmFecalNew.cboRequest.Text = GridView.GetFocusedRowCellValue(GridView.Columns("Request")).ToString
        frmFecalNew.dtReceived.Value = GridView.GetFocusedRowCellValue(GridView.Columns("DateReceived")).ToString
        frmFecalNew.tmTimeReceived.Text = GridView.GetFocusedRowCellValue(GridView.Columns("TimeReceived")).ToString
        frmFecalNew.cboSex.Text = GridView.GetFocusedRowCellValue(GridView.Columns("Sex")).ToString
        frmFecalNew.dtBDate.Text = GridView.GetFocusedRowCellValue(GridView.Columns("DateOfBirth")).ToString
        frmFecalNew.cboPatientType.Text = GridView.GetFocusedRowCellValue(GridView.Columns("PatientType")).ToString
        frmFecalNew.cboPhysician.Text = GridView.GetFocusedRowCellValue(GridView.Columns("Physician")).ToString
        frmFecalNew.cboRoom.Text = GridView.GetFocusedRowCellValue(GridView.Columns("RoomWard")).ToString
        frmFecalNew.cboMedTech.Text = GridView.GetFocusedRowCellValue(GridView.Columns("PerformedBy")).ToString
        frmFecalNew.cboVerify.Text = GridView.GetFocusedRowCellValue(GridView.Columns("ReleasedBy")).ToString
        frmFecalNew.cboPhysician.Text = GridView.GetFocusedRowCellValue(GridView.Columns("Physician")).ToString

        'For Age computation
        Dim Age As String = ""
        If Not GridView.GetFocusedRowCellValue(GridView.Columns("DateOfBirth")).ToString = "" Then
            Age = GetBDate(GridView.GetFocusedRowCellValue(GridView.Columns("DateOfBirth")).ToString)
            frmFecalNew.txtAge.Text = Age.Split(" ").GetValue(0)
            frmFecalNew.txtClass.Text = Age.Split(" ").GetValue(1)
        End If
        '######################################----END-----###############################################################

        'Parameters 
        rs.Parameters.Clear()
        rs.Parameters.AddWithValue("@PatientID", GridView.GetFocusedRowCellValue(GridView.Columns("PatientID")).ToString)
        rs.Parameters.AddWithValue("@MainID", GridView.GetFocusedRowCellValue(GridView.Columns("RefID")).ToString)
        rs.Parameters.AddWithValue("@Section", GridView.GetFocusedRowCellValue(GridView.Columns("Section")).ToString)
        rs.Parameters.AddWithValue("@SubSection", GridView.GetFocusedRowCellValue(GridView.Columns("SubSection")).ToString)
        rs.Parameters.AddWithValue("@CurrUser", CurrUser)
        '###########################---Load Address and Contact No.---####################################################
        Connect()
        rs.Connection = conn
        rs.CommandType = CommandType.Text
        rs.CommandText = "SELECT `address`, `contact_no`, `civil_status` FROM `patient_info` WHERE `patient_id` = @PatientID"
        reader = rs.ExecuteReader
        reader.Read()
        If reader.HasRows Then
            frmFecalNew.txtAddress.Text = reader(0).ToString
            frmFecalNew.txtContact.Text = reader(1).ToString
            frmFecalNew.cboCS.Text = reader(2).ToString
        End If
        Disconnect()
        '######################################----END-----###############################################################

        '###########################---Load Additional Info---####################################################
        Connect()
        rs.Connection = conn
        rs.CommandType = CommandType.Text
        rs.CommandText = "SELECT `accession_no`, `or_no`, `cs_no` FROM `additional_info` WHERE `sample_id` = @MainID AND section = @Section AND sub_section = @SubSection"
        reader = rs.ExecuteReader
        reader.Read()
        If reader.HasRows Then
            frmFecalNew.txtAccession.Text = reader(0).ToString
            frmFecalNew.txtORNo.Text = reader(1).ToString
            frmFecalNew.txtChargeSlip.Text = reader(2).ToString
        End If
        Disconnect()
        '######################################----END-----###############################################################

        '###########################---Load remarks and comments---####################################################
        Connect()
        rs.Connection = conn
        rs.CommandType = CommandType.Text
        rs.CommandText = "SELECT `remarks`, `diagnosis` FROM `patient_remarks` WHERE `sample_id` = @MainID AND section = @Section AND sub_section = @SubSection"
        reader = rs.ExecuteReader
        reader.Read()
        If reader.HasRows Then
            frmFecalNew.txtRemarks.Text = reader(0).ToString
            frmFecalNew.txtComment.Text = reader(1).ToString
        End If
        Disconnect()
        '######################################----END-----###############################################################

        '##############################------To enable necessary buttons-------#######################################################
        'If GridView.GetFocusedRowCellValue(GridView.Columns("Status")) = "Validated" Then
        '    frmFecalNew.btnPrint.Enabled = False
        '    frmFecalNew.btnValidate.Enabled = True
        '    frmFecalNew.btnPrintNow.Enabled = True
        'End If
        '############################----------End------------##############################################################

        'frmFecalNew.GetBDate()

        '###########################---Load Email Address---####################################################
        Connect()
        rs.Connection = conn
        rs.CommandType = CommandType.Text
        rs.CommandText = "SELECT `email_address` FROM `email_details` WHERE `sample_id` = @MainID AND section = @Section AND sub_section = @SubSection"
        reader = rs.ExecuteReader
        reader.Read()
        If reader.HasRows Then
            frmFecalNew.txtEmail.Text = reader(0).ToString
        End If
        Disconnect()
        '######################################----END-----###############################################################

        frmFecalNew.cboPathologist.SelectedIndex = 0

        'Activity Logs
        ActivityLogs(GridView.GetFocusedRowCellValue(GridView.Columns("SampleID")).ToString,
                             GridView.GetFocusedRowCellValue(GridView.Columns("PatientID")).ToString,
                             GridView.GetFocusedRowCellValue(GridView.Columns("PatientName")).ToString,
                             CurrUser,
                             "View Result",
                             GridView.GetFocusedRowCellValue(GridView.Columns("Request")).ToString,
                             GridView.GetFocusedRowCellValue(GridView.Columns("Section")).ToString,
                             GridView.GetFocusedRowCellValue(GridView.Columns("SubSection")).ToString)

        frmFecalNew.ShowDialog()
    End Sub
    '############################################--------------END------------############################################

    '############################################-----For Completed Orders-----############################################
    Private Sub DisplayResultCompleted()

        frmFecalOrdered.cboPathologist.Properties.Items.Clear()
        frmFecalOrdered.cboMedTech.Properties.Items.Clear()
        frmFecalOrdered.cboVerify.Properties.Items.Clear()

        '###########################---Load Basic Patient Details---######################################################
        frmFecalOrdered.mainID = GridCompleted.GetFocusedRowCellValue(GridCompleted.Columns("RefID")).ToString
        frmFecalOrdered.Section = GridCompleted.GetFocusedRowCellValue(GridCompleted.Columns("Section")).ToString
        frmFecalOrdered.SubSection = GridCompleted.GetFocusedRowCellValue(GridCompleted.Columns("SubSection")).ToString
        frmFecalOrdered.PatientID = GridCompleted.GetFocusedRowCellValue(GridCompleted.Columns("PatientID")).ToString
        frmFecalOrdered.txtSampleID.Text = GridCompleted.GetFocusedRowCellValue(GridCompleted.Columns("SampleID")).ToString
        frmFecalOrdered.txtPatientID.Text = GridCompleted.GetFocusedRowCellValue(GridCompleted.Columns("PatientID")).ToString
        frmFecalOrdered.txtName.Text = GridCompleted.GetFocusedRowCellValue(GridCompleted.Columns("PatientName")).ToString
        frmFecalOrdered.cboRequest.Text = GridCompleted.GetFocusedRowCellValue(GridCompleted.Columns("Request")).ToString
        frmFecalOrdered.dtReceived.Text = GridCompleted.GetFocusedRowCellValue(GridCompleted.Columns("DateReceived")).ToString
        frmFecalOrdered.tmTimeReceived.Text = GridCompleted.GetFocusedRowCellValue(GridCompleted.Columns("TimeReceived")).ToString
        frmFecalOrdered.tmTimeReleased.Text = GridCompleted.GetFocusedRowCellValue(GridCompleted.Columns("DateReleased")).ToString
        frmFecalOrdered.cboSex.Text = GridCompleted.GetFocusedRowCellValue(GridCompleted.Columns("Sex")).ToString
        frmFecalOrdered.dtBDate.Text = GridCompleted.GetFocusedRowCellValue(GridCompleted.Columns("DateOfBirth")).ToString
        frmFecalOrdered.cboPatientType.Text = GridCompleted.GetFocusedRowCellValue(GridCompleted.Columns("PatientType")).ToString
        frmFecalOrdered.cboPhysician.Text = GridCompleted.GetFocusedRowCellValue(GridCompleted.Columns("Physician")).ToString
        frmFecalOrdered.cboRoom.Text = GridCompleted.GetFocusedRowCellValue(GridCompleted.Columns("RoomWard")).ToString
        frmFecalOrdered.cboMedTech.Text = GridCompleted.GetFocusedRowCellValue(GridCompleted.Columns("PerformedBy")).ToString
        frmFecalOrdered.cboVerify.Text = GridCompleted.GetFocusedRowCellValue(GridCompleted.Columns("ReleasedBy")).ToString
        frmFecalOrdered.cboPhysician.Text = GridCompleted.GetFocusedRowCellValue(GridCompleted.Columns("Physician")).ToString

        'For Age computation
        Dim Age As String = ""
        If Not GridCompleted.GetFocusedRowCellValue(GridCompleted.Columns("DateOfBirth")).ToString = "" Then
            Age = GetBDate(GridCompleted.GetFocusedRowCellValue(GridCompleted.Columns("DateOfBirth")).ToString)
            frmFecalOrdered.txtAge.Text = Age.Split(" ").GetValue(0)
            frmFecalOrdered.txtClass.Text = Age.Split(" ").GetValue(1)
        End If
        '######################################----END-----###############################################################

        'Parameters 
        rs.Parameters.Clear()
        rs.Parameters.AddWithValue("@PatientID", GridCompleted.GetFocusedRowCellValue(GridCompleted.Columns("PatientID")).ToString)
        rs.Parameters.AddWithValue("@MainID", GridCompleted.GetFocusedRowCellValue(GridCompleted.Columns("RefID")).ToString)
        rs.Parameters.AddWithValue("@Section", GridCompleted.GetFocusedRowCellValue(GridCompleted.Columns("Section")).ToString)
        rs.Parameters.AddWithValue("@SubSection", GridCompleted.GetFocusedRowCellValue(GridCompleted.Columns("SubSection")).ToString)
        rs.Parameters.AddWithValue("@CurrUser", CurrUser)


        '###########################---Load Address and Contact No.---####################################################
        Connect()
        rs.Connection = conn
        rs.CommandType = CommandType.Text
        rs.CommandText = "SELECT `address`, `contact_no`, `civil_status` FROM `patient_info` WHERE `patient_id` = @PatientID"
        reader = rs.ExecuteReader
        reader.Read()
        If reader.HasRows Then
            frmFecalOrdered.txtAddress.Text = reader(0).ToString
            frmFecalOrdered.txtContact.Text = reader(1).ToString
            frmFecalOrdered.cboCS.Text = reader(2).ToString
        End If
        Disconnect()
        '######################################----END-----###############################################################

        '###########################---Load Additional Info---####################################################
        Connect()
        rs.Connection = conn
        rs.CommandType = CommandType.Text
        rs.CommandText = "SELECT `accession_no`, `or_no`, `cs_no` FROM `additional_info` WHERE `sample_id` = @MainID AND section = @Section AND sub_section = @SubSection"
        reader = rs.ExecuteReader
        reader.Read()
        If reader.HasRows Then
            frmFecalOrdered.txtAccession.Text = reader(0).ToString
            frmFecalOrdered.txtORNo.Text = reader(1).ToString
            frmFecalOrdered.txtChargeSlip.Text = reader(2).ToString
        End If
        Disconnect()
        '######################################----END-----###############################################################

        '###########################---Load remarks and comments---####################################################
        Connect()
        rs.Connection = conn
        rs.CommandType = CommandType.Text
        rs.CommandText = "SELECT `remarks`, `diagnosis` FROM `patient_remarks` WHERE `sample_id` = @MainID AND section = @Section AND sub_section = @SubSection"
        reader = rs.ExecuteReader
        reader.Read()
        If reader.HasRows Then
            frmFecalOrdered.txtRemarks.Text = reader(0).ToString
            frmFecalOrdered.txtComment.Text = reader(1).ToString
        End If
        Disconnect()
        '######################################----END-----###############################################################

        'Activity Logs
        ActivityLogs(GridCompleted.GetFocusedRowCellValue(GridCompleted.Columns("SampleID")).ToString,
                        GridCompleted.GetFocusedRowCellValue(GridCompleted.Columns("PatientID")).ToString,
                        GridCompleted.GetFocusedRowCellValue(GridCompleted.Columns("PatientName")).ToString,
                        CurrUser,
                        "View Archived Result",
                        GridCompleted.GetFocusedRowCellValue(GridCompleted.Columns("Request")).ToString,
                        GridCompleted.GetFocusedRowCellValue(GridCompleted.Columns("Section")).ToString,
                        GridCompleted.GetFocusedRowCellValue(GridCompleted.Columns("SubSection")).ToString)
        frmFecalOrdered.ShowDialog()
    End Sub
    '############################################--------------END------------############################################

#End Region

    Public Sub DisablePermission()
        Connect()
        rs.Connection = conn
        rs.CommandText = "SELECT * FROM `user_permission` WHERE `user_id` = '" & CurrEmail & "'"
        reader = rs.ExecuteReader
        While reader.Read
            If reader(3).ToString = 0 Then
                If btnBarcode.Caption = reader(2).ToString Then
                    btnBarcode.Visibility = DevExpress.XtraBars.BarItemVisibility.Never
                ElseIf btnPrint.Caption = reader(2).ToString Then
                    btnPrint.Visibility = DevExpress.XtraBars.BarItemVisibility.Never
                ElseIf btnView.Caption = reader(2).ToString Then
                    btnView.Visibility = DevExpress.XtraBars.BarItemVisibility.Never
                ElseIf btnDelete.Caption = reader(2).ToString Then
                    btnDelete.Visibility = DevExpress.XtraBars.BarItemVisibility.Never
                ElseIf btnReject.Caption = reader(2).ToString Then
                    btnReject.Visibility = DevExpress.XtraBars.BarItemVisibility.Never
                End If
            ElseIf reader(3).ToString = 1 Then
                If btnBarcode.Caption = reader(2).ToString Then
                    btnBarcode.Visibility = DevExpress.XtraBars.BarItemVisibility.Always
                ElseIf btnPrint.Caption = reader(2).ToString Then
                    btnPrint.Visibility = DevExpress.XtraBars.BarItemVisibility.Always
                ElseIf btnView.Caption = reader(2).ToString Then
                    btnView.Visibility = DevExpress.XtraBars.BarItemVisibility.Always
                ElseIf btnDelete.Caption = reader(2).ToString Then
                    btnDelete.Visibility = DevExpress.XtraBars.BarItemVisibility.Always
                ElseIf btnReject.Caption = reader(2).ToString Then
                    btnReject.Visibility = DevExpress.XtraBars.BarItemVisibility.Always
                End If
            End If
        End While
        Disconnect()
    End Sub
End Class