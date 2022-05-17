Imports System.Drawing.Printing
Imports DevExpress.XtraGrid.Views.Grid
Imports DevExpress.XtraPrinting.BarCode
Imports System.Data.SqlClient

Public Class frmiHOMIS

    Dim LastID As Integer
    Dim SampleID As String

    Public mainID As String = ""
    Dim arrImage() As Byte
    Dim Mode As String
    Dim BDate As String
    Dim Age As String

    Dim Section As String
    Dim SubSection As String

    Private PrintDocType As String = "Barcode"
    'Private StrPrinterName As String = "EPSON L220 Series"
    Private StrPrinterName As String = My.Settings.BCPrinterName

    Dim procreslt As String = "002"
    Dim Test As String
    Dim HISCode As String

    Dim PatientOrder As String
    Dim Test_Name As String
    Dim LIS_TestCode As String
    Dim TestCode As String
    Dim HIS_TestCode As String
    Dim Order_No As String
    Dim Test_Group As String
    Dim Unit As String
    Dim Unit_Conv As String
    Dim ResultID As String
    Dim HIS_Field As String
    Dim Order_TestCode As String

    Public Sub LoadRecordsiHomis()
        Me.GridView.Columns.Clear()
        'Load Records from HIS
        Dim SQL As String = "SELECT TOP 1000 " _
            & "hperson.hpercode AS HospitalNo, " _
            & "hperson.patlast + ', ' + hperson.patfirst AS PatientName, " _
            & "hperson.patbdate AS DateOfBirth, " _
            & "hperson.patsex AS Sex, " _
            & "hdocord.dopriority AS Priority, " _
            & "hprocm.procdesc AS Test," _
            & "('DR. ' + hpersonal.lastname + ', ' + hpersonal.firstname) AS RequestingPhysician, " _
            & "hdocord.pcchrgcod AS ChargeSlip, " _
            & "henctr.toecode AS Location, " _
            & "hdocord.dodate AS DateTimeRequested, " _
            & "CASE " _
            & "WHEN hprocm.procreslt = 002 THEN 'Hematology' " _
            & "WHEN hprocm.procreslt = 008 THEN 'Chemistry' " _
            & "WHEN hprocm.procreslt = 009 THEN 'Hematology' " _
            & "WHEN hprocm.procreslt = 007 THEN 'Fecalysis' " _
            & "WHEN hprocm.procreslt = 001 THEN 'Urinalysis' " _
            & "WHEN hprocm.procreslt = 006 THEN 'ImmunoSero' " _
            & "WHEN hprocm.procreslt = 003 THEN 'Blood Culture' " _
            & "WHEN hprocm.procreslt = 012 THEN 'Blood Bank' " _
            & "END AS Section, " _
            & "hdocord.docointkey AS MainID, " _
            & "hdocord.proccode AS HISCode" _
            & " FROM " _
            & "hperson, " _
            & "hdocord, " _
            & "hprocm, " _
            & "hpersonal, " _
            & "hprovider, " _
            & "henctr" _
            & " WHERE " _
            & "(hperson.hpercode = hdocord.hpercode AND hdocord.proccode = hprocm.proccode)" _
            & " AND " _
            & "(hdocord.licno = hprovider.licno AND hprovider.employeeid = hpersonal.employeeid)" _
            & " AND " _
            & "hdocord.enccode = henctr.enccode" _
            & " AND " _
            & "hdocord.estatus = 'P'" _
            & " AND " _
            & "hprocm.procreslt = @Section" _
            & " AND " _
            & "hdocord.pcchrgcod NOT LIKE ''" _
            & " AND " _
            & "(CONVERT(varchar, hdocord.dodate, 101) BETWEEN @DateFrom AND @DateTo)" _
            & "ORDER BY hdocord.dodate DESC"

        GridView.Columns.Clear()
        GridView.Appearance.HeaderPanel.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center
        GridView.Appearance.HeaderPanel.FontStyleDelta = FontStyle.Bold

        ConnectSQL()
        rsSQL.Connection = connSQL
        rsSQL.CommandType = CommandType.Text
        rsSQL.CommandText = SQL
        rsSQL.Parameters.Clear()
        rsSQL.Parameters.AddWithValue("@Section", procreslt)
        rsSQL.Parameters.Add("@DateFrom", SqlDbType.DateTime).Value = Format(dtFrom1.Value, "yyyy-MM-dd")
        rsSQL.Parameters.Add("@DateTo", SqlDbType.DateTime).Value = Format(dtTo1.Value, "yyyy-MM-dd")

        Dim adapter As New SqlDataAdapter(rsSQL)

        Dim myTable As DataTable = New DataTable
        adapter.Fill(myTable)

        dtList.DataSource = myTable

        GridView.Columns("MainID").Visible = False
        'GridView.Columns("HISCode").Visible = False
        GridView.Columns("Section").Visible = False
        'GridView.Columns("PatientType").Visible = False

        ' Make the grid read-only. 
        GridView.OptionsBehavior.Editable = False
        ' Prevent the focused cell from being highlighted. 
        GridView.OptionsSelection.EnableAppearanceFocusedCell = False
        ' Draw a dotted focus rectangle around the entire row. 
        GridView.FocusRectStyle = DrawFocusRectStyle.RowFullFocus

        GridView.OptionsSelection.MultiSelect = True
        GridView.OptionsSelection.MultiSelectMode = GridMultiSelectMode.CheckBoxRowSelect

        DisconnectSQL()
    End Sub

    Private Sub frmNewOrder_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        LoadRecordsiHomis()
    End Sub

    Private Sub btnAdd_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnAdd.ItemClick

        GetLastID()

        Dim Result As DialogResult = MessageBox.Show("You're about to Check-In Patient " & GridView.GetFocusedRowCellValue(GridView.Columns("PatientName")) & "." & vbCrLf & vbCrLf & "Do you want to continue to print Barcode Sticker " & SampleID & "?", "System Message", MessageBoxButtons.YesNoCancel, MessageBoxIcon.Question)

        If (Result = DialogResult.Yes) Then
            Try
                Connect1()
                rs1.Parameters.Clear()
                rs1.Parameters.AddWithValue("@HISCode", GridView.GetFocusedRowCellValue(GridView.Columns("HISCode")))
                rs1.Connection = conn1
                rs1.CommandType = CommandType.Text
                rs1.CommandText = "SELECT `test_name` FROM `default_specimen` WHERE `his_code` = @HISCode"
                reader1 = rs1.ExecuteReader
                reader1.Read()
                If reader1.HasRows Then
                    HISCode = reader1(0).ToString
                End If
                Disconnect1()

                PrintBarcode(GridView.GetFocusedRowCellValue(GridView.Columns("Test")),
                                     SampleID,
                                     GridView.GetFocusedRowCellValue(GridView.Columns("HospitalNo")),
                                     GridView.GetFocusedRowCellValue(GridView.Columns("PatientName")),
                                     GridView.GetFocusedRowCellValue("DateofBirth"),
                                     GridView.GetFocusedRowCellValue(GridView.Columns("Sex")),
                                     GridView.GetFocusedRowCellValue(GridView.Columns("Section")),
                                     HISCode)

            Catch ex As Exception
                Disconnect()
                MessageBox.Show("Error in connection on printer. " + ex.Message, "Barcode Printing Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Exit Sub
            End Try
            GoTo A
        ElseIf (Result = DialogResult.No) Then
            GoTo A
        ElseIf (Result = DialogResult.Cancel) Then
            Exit Sub
        End If

A:

        Connect3()
        rs3.Connection = conn3
        rs3.CommandType = CommandType.Text
        rs3.CommandText = "INSERT INTO `tmpWorklist` (`sample_id`, `patient_id`, `patient_name`, `sex`, `bdate`, `age`, `physician`, `dept`, `status`, `main_id`, `date`, `time`, `barcode`, `testtype`, `sub_section`, `test`, `patient_type`, `TYPE`, `location`) VALUES " _
                        & "(" _
                        & "@sample_id," _
                        & "@PATIENT_ID," _
                        & "@NAME," _
                        & "@sex," _
                        & "@bdate," _
                        & "@AGE," _
                        & "@PHYSICIAN," _
                        & "@dept," _
                        & "@status," _
                        & "@MainID," _
                        & "@DATE," _
                        & "@time," _
                        & "@Barcode," _
                        & "@TestType," _
                        & "@SubSection," _
                        & "@Test," _
                        & "@Patient_Type," _
                        & "@TYPE," _
                        & "@location" _
                        & ")"

        Dim selectedRows() As Integer = GridView.GetSelectedRows()
        For Each rowHandle As Integer In selectedRows
            If rowHandle >= 0 Then
                Connect1()
                rs1.Parameters.Clear()
                rs1.Parameters.AddWithValue("@HISCode", GridView.GetRowCellValue(rowHandle, GridView.Columns("HISCode")))
                rs1.Connection = conn1
                rs1.CommandType = CommandType.Text
                rs1.CommandText = "SELECT `test_name` FROM `default_specimen` WHERE `his_code` = @HISCode"
                reader1 = rs1.ExecuteReader
                reader1.Read()
                If reader1.HasRows Then
                    HISCode = reader1(0).ToString
                End If
                Disconnect1()

                Test &= GridView.GetRowCellValue(rowHandle, GridView.Columns("Test")) & ", "
                'IIf(IsNothing(GridView.GetRowCellValue(rowHandle, GridView.Columns("DateOfBirth"))), vbNull, Format(CDate(GridView.GetRowCellValue(rowHandle, GridView.Columns("DateOfBirth"))), "MM/dd/yyyy"))

                rs3.Parameters.Clear()
                rs3.Parameters.AddWithValue("@sample_id", SampleID)
                rs3.Parameters.AddWithValue("@NAME", GridView.GetRowCellValue(rowHandle, GridView.Columns("PatientName")))
                rs3.Parameters.AddWithValue("@AGE", "")
                rs3.Parameters.AddWithValue("@bdate", Format(CDate(GridView.GetRowCellValue(rowHandle, GridView.Columns("DateOfBirth"))), "MM/dd/yyyy"))
                rs3.Parameters.AddWithValue("@PHYSICIAN", GridView.GetRowCellValue(rowHandle, GridView.Columns("RequestingPhysician")))
                rs3.Parameters.AddWithValue("@DATE", Format(CDate(GridView.GetRowCellValue(rowHandle, GridView.Columns("DateTimeRequested"))), "yyyy-MM-dd"))
                rs3.Parameters.AddWithValue("@TIME", Format(CDate(GridView.GetRowCellValue(rowHandle, GridView.Columns("DateTimeRequested"))), "hh:mm:ss tt"))
                rs3.Parameters.AddWithValue("@PATIENT_ID", GridView.GetRowCellValue(rowHandle, GridView.Columns("HospitalNo")))
                rs3.Parameters.AddWithValue("@dept", "")
                rs3.Parameters.AddWithValue("@status", "Checked-In")
                rs3.Parameters.AddWithValue("@address", "")
                rs3.Parameters.AddWithValue("@contact", "")
                rs3.Parameters.AddWithValue("@Barcode", "")
                rs3.Parameters.AddWithValue("@TestType", GridView.GetRowCellValue(rowHandle, GridView.Columns("Section")))
                rs3.Parameters.AddWithValue("@Test", Test.TrimEnd(", "))
                rs3.Parameters.AddWithValue("@Patient_Type", "")
                rs3.Parameters.AddWithValue("@TYPE", "")
                rs3.Parameters.AddWithValue("@location", GridView.GetRowCellValue(rowHandle, GridView.Columns("Location")))
                rs3.Parameters.AddWithValue("@SubSection", HISCode)

                rs3.Parameters.AddWithValue("@MainID", SampleID)

                If GridView.GetRowCellValue(rowHandle, GridView.Columns("Sex")) = "M" Then
                    rs3.Parameters.AddWithValue("@sex", "Male")
                ElseIf GridView.GetRowCellValue(rowHandle, GridView.Columns("Sex")) = "F" Then
                    rs3.Parameters.AddWithValue("@sex", "Female")
                End If

            End If

            'Find Record in SBSILIS Database that match the Test Code
            Connect1()
                rs1.Connection = conn1
                rs1.CommandType = CommandType.Text
                rs1.CommandText = "SELECT his_code, test_code, specimen, order_no, test_group, si_unit, conventional_unit, his_field, test_name FROM `default_specimen` WHERE `his_code` = @HISCode AND `status` = 'Enable'"
                reader1 = rs1.ExecuteReader
                While reader1.Read()
                    If reader1.HasRows Then
                        HIS_TestCode = reader1(0).ToString
                        TestCode = reader1(1).ToString
                        test_name = reader1(2).ToString
                        Order_No = reader1(3).ToString
                        Test_Group = reader1(4).ToString
                        unit = reader1(5).ToString
                        unit_conv = reader1(6).ToString
                        HIS_Field = reader1(7).ToString
                        Order_TestCode = reader1(8).ToString

                        'Save Specimen to Result According to Number of Test in the list
                        Connect2()
                        rs2.Parameters.Clear()
                        rs2.Parameters.AddWithValue("@PATIENT_ID", GridView.GetRowCellValue(rowHandle, GridView.Columns("HospitalNo")))
                        rs2.Parameters.AddWithValue("@TestType", GridView.GetRowCellValue(rowHandle, GridView.Columns("Section")))
                        rs2.Parameters.AddWithValue("@HISMainID", GridView.GetRowCellValue(rowHandle, GridView.Columns("MainID")))
                        rs2.Parameters.AddWithValue("@MainID", SampleID)
                        rs2.Parameters.AddWithValue("@SubSection", HISCode)

                        rs2.Connection = conn2
                        rs2.CommandType = CommandType.Text
                        rs2.CommandText = "INSERT INTO `tmpResult` (`universal_id`, `measurement`, `test_code`, `sample_id`, `date`, `patient_id`, `order_no`, `units`, `unit_conv`, `instrument`, `status`, `test_group`, `his_code`, `his_mainid`, `section`, `sub_section`) VALUES " _
                                & "(" _
                                & "'" & Test_Name & "', '', '" & TestCode & "', @MainID, NOW(), @patient_id, '" & Order_No & "', '" & Unit & "', '" & Unit_Conv & "', 'Other_Test', 0, '" & Test_Group & "', '" & HIS_Field & "', @HISMainID, @TestType, @SubSection" _
                                & ")"
                        rs2.ExecuteNonQuery()
                        Disconnect2()
                    End If
                End While
                Disconnect1()
            'End of Save Result

            'Check existing Patient Order
            Connect1()
            rs1.Connection = conn1
            rs1.CommandType = CommandType.Text
            rs1.CommandText = "SELECT `sample_id`, `test_code` FROM `patient_order` WHERE `sample_id` = '" & SampleID & "' AND `test_code` = '" & Order_TestCode & "'"
            reader1 = rs1.ExecuteReader
            If reader1.HasRows Then

            Else
                'Save Patient Order According to Number of Test in the list
                Connect2()
                rs2.Connection = conn2
                rs2.CommandType = CommandType.Text
                rs2.CommandText = "INSERT INTO `patient_order` (`test_name`, `test_code`, `patient_id`, `sample_id`, `testtype`, `mode`, `his_mainID`, `sub_section`) VALUES (" _
                        & "'" & Order_TestCode & "'," _
                        & "'" & Order_TestCode & "'," _
                        & "'" & GridView.GetRowCellValue(rowHandle, GridView.Columns("HospitalNo")) & "'," _
                        & "'" & SampleID & "'," _
                        & "'" & GridView.GetRowCellValue(rowHandle, GridView.Columns("Section")) & "'," _
                        & "'" & GridView.GetRowCellValue(rowHandle, GridView.Columns("Priority")) & "'," _
                        & "'" & GridView.GetRowCellValue(rowHandle, GridView.Columns("MainID")) & "'," _
                        & "'" & HISCode & "')"
                rs2.ExecuteNonQuery()
                Disconnect2()
                'End of Patient Order
            End If
            Disconnect1()
            'End of Find

            'Check existing Specimen Tracking
            Connect1()
            rs1.Connection = conn1
            rs1.CommandType = CommandType.Text
            rs1.CommandText = "SELECT * FROM `specimen_tracking` WHERE sample_id = '" & SampleID & "' AND section = '" & cboSection.Text & "' AND sub_section = '" & HISCode & "'"
            reader1 = rs1.ExecuteReader
            reader1.Read()
            If reader1.HasRows Then

            Else
                Connect2()
                rs2.Parameters.Clear()
                rs2.Parameters.AddWithValue("@DateReceived", Format(CDate(GridView.GetRowCellValue(rowHandle, GridView.Columns("DateTimeRequested"))), "yyyy-MM-dd hh:mm:ss tt"))
                rs2.Parameters.AddWithValue("@DateCheckedIn", Format(CDate(Now), "yyyy-MM-dd HH:mm:ss"))
                rs2.Connection = conn2
                rs2.CommandType = CommandType.Text
                rs2.CommandText = "INSERT INTO `specimen_tracking` (`sample_id`, `received`, `extracted`, `section`, `sub_section`) VALUES ('" & SampleID & "', @DateReceived, @DateCheckedIn, '" & cboSection.Text & "', '" & HISCode & "')"
                rs2.ExecuteNonQuery()
                Disconnect2()
            End If
            Disconnect1()

            'Update iHOMIS 'estatus' to 'S'
            Dim SQL_Update = "UPDATE hdocord SET estatus = 'S' WHERE docointkey = '" & GridView.GetRowCellValue(rowHandle, GridView.Columns("MainID")) & "'"
            ConnectSQL()
            rsSQL.Connection = connSQL
            rsSQL.CommandType = CommandType.Text
            rsSQL.CommandText = SQL_Update
            rsSQL.ExecuteNonQuery()
            DisconnectSQL()
            'end update iHOMIS.
        Next
        rs3.ExecuteNonQuery()
        Disconnect3()

        'Save Last ID for Reference in the next Sample to check in
        Connect()
        rs.Connection = conn
        rs.CommandType = CommandType.Text
        rs.CommandText = "INSERT INTO `last_id` VALUES ('" & LastID + 1 & "')"
        rs.ExecuteNonQuery()
        Disconnect()
        'End of Last Id Checked-In

        MessageBox.Show("Patient order successfully Checked-In", "Check-In", MessageBoxButtons.OK, MessageBoxIcon.Information)
        LoadRecordsiHomis()

    End Sub

    Private Sub btnRefresh_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnRefresh.ItemClick
        LoadRecordsiHomis()
    End Sub

    Private Sub btnClose_ItemClick(ByVal sender As Object, ByVal e As DevExpress.XtraBars.ItemClickEventArgs) Handles btnClose.ItemClick
        Me.Close()
        Me.Dispose()
    End Sub

    Private Sub frm_Activated(sender As Object, e As EventArgs) Handles Me.Activated
        'MainFOrm.aceFecal.Appearance.Normal.BackColor = Color.FromArgb(6, 31, 71)
        MainFOrm.acciHOMIS.Appearance.Normal.BackColor = Color.FromArgb(16, 110, 190)
        MainFOrm.acciHOMIS.Appearance.Normal.ForeColor = Color.FromArgb(255, 255, 255)
    End Sub

    Private Sub frm_FormClosed(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
        MainFOrm.lblTitle.Text = ""
        MainFOrm.acciHOMIS.Appearance.Normal.BackColor = Color.FromArgb(240, 240, 240)
        MainFOrm.acciHOMIS.Appearance.Normal.ForeColor = Color.FromArgb(27, 41, 62)
        Me.Dispose()
    End Sub

    Private Sub btnSearch_Click(sender As Object, e As EventArgs) Handles btnSearch.Click
        LoadRecordsiHomis()
    End Sub

    Private Sub cboSection_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cboSection.SelectedIndexChanged
        If cboSection.SelectedItem = "Hematology" Then
            procreslt = "002"
        ElseIf cboSection.SelectedItem = "Chemistry" Then
            procreslt = "008"
        ElseIf cboSection.SelectedItem = "Hematology" Then
            procreslt = "009"
        ElseIf cboSection.SelectedItem = "Fecalysis" Then
            procreslt = "007"
        ElseIf cboSection.SelectedItem = "Urinalysis" Then
            procreslt = "001"
        ElseIf cboSection.SelectedItem = "Blood Culture" Then
            procreslt = "003"
        ElseIf cboSection.SelectedItem = "Blood Bank" Then
            procreslt = "012"
        ElseIf cboSection.SelectedItem = "ImmunoSero" Then
            procreslt = "006"
        End If
        LoadRecordsiHomis()
    End Sub

    Private Sub GetLastID()
        Dim MYSQL As String = "SELECT * FROM LAST_ID ORDER BY lastidno DESC"
        Connect()
        rs.Connection = conn
        rs.CommandType = CommandType.Text
        rs.CommandText = MYSQL
        reader = rs.ExecuteReader
        reader.Read()
        If reader.HasRows Then
            LastID = reader(0).ToString
            If reader(0).ToString > 0 And reader(0).ToString <= 9 Then
                SampleID = "000000" & reader(0).ToString + 1
            ElseIf reader(0).ToString > 9 And reader(0).ToString <= 99 Then
                SampleID = "00000" & reader(0).ToString + 1
            ElseIf reader(0).ToString > 99 And reader(0).ToString <= 999 Then
                SampleID = "0000" & reader(0).ToString + 1
            ElseIf reader(0).ToString > 999 And reader(0).ToString <= 9999 Then
                SampleID = "000" & reader(0).ToString + 1
            ElseIf reader(0).ToString > 9999 And reader(0).ToString <= 99999 Then
                SampleID = "00" & reader(0).ToString + 1
            ElseIf reader(0).ToString > 99999 And reader(0).ToString <= 999999 Then
                SampleID = "0" & reader(0).ToString + 1
            ElseIf reader(0).ToString > 999999 And reader(0).ToString <= 9999999 Then
                SampleID = reader(0).ToString + 1
            End If
        Else
            LastID = 1
            SampleID = "000000" & 1
        End If
        Disconnect()
    End Sub

End Class