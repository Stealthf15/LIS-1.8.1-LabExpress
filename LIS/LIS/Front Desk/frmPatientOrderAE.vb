Imports System.Drawing.Printing
Imports DevExpress.XtraEditors
Imports DevExpress.XtraEditors.Repository
Imports DevExpress.XtraGrid.Views.Grid
Imports DevExpress.XtraPrinting.BarCode
Imports MySql.Data.MySqlClient
Public Class frmPatientOrderAE

    Dim LastID As Integer
    Dim SampleID As String

    Dim itemDict As New Dictionary(Of String, String)
    Dim duplicateItems As New List(Of String)
    Dim Count As Integer = 0

    Dim arrImage() As Byte

    Dim BarcodeID As String = ""
    Dim SpecimenType As String = ""

    Private PrintDocType As String = "Barcode"
    'Private StrPrinterName As String = "EPSON L220 Series"
    Private StrPrinterName As String = My.Settings.BCPrinterName

    Private Sub cboTestName_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cboTestName.SelectedIndexChanged
        lvTest.Items.Clear()

        Connect()
        rs.Connection = conn
        rs.CommandType = CommandType.Text
        rs.CommandText = "SELECT DISTINCT `his_code` FROM `default_specimen` WHERE `section` = '" & cboTestName.Text & "'"
        reader = rs.ExecuteReader
        While reader.Read
            Dim iItem As New ListViewItem(reader(0).ToString, 0)
            lvTest.Items.Add(iItem)
        End While
        Disconnect()

        chAll.Checked = False
        lvList.Items.Clear()
    End Sub


    Private Sub lvTest_SelectedIndexChanged(sender As Object, e As EventArgs) Handles lvTest.SelectedIndexChanged
        chAll.Checked = False
        Dim a As Integer = 0
        lvList.Items.Clear()

        Connect()
        rs.Connection = conn
        rs.CommandType = CommandType.Text
        rs.CommandText = "SELECT `specimen`, `test_code`, `order_no`, `section`, `test_name`, `specimen_type` FROM `default_specimen` WHERE his_code = '" & lvTest.FocusedItem.SubItems(0).Text & "' AND `status` = 'Enable'"
        reader = rs.ExecuteReader
        While reader.Read
            a = a + 1
            Dim iItem As New ListViewItem((a).ToString, 0)
            iItem.SubItems.Add(reader(0).ToString)
            iItem.SubItems.Add(reader(1).ToString)
            iItem.SubItems.Add(reader(2).ToString)
            iItem.SubItems.Add(reader(3).ToString)
            iItem.SubItems.Add(reader(4).ToString)
            iItem.SubItems.Add(reader(5).ToString)
            lvList.Items.Add(iItem)
        End While
        Disconnect()

        chAll.Checked = True
    End Sub

    Private Sub btnAdd_Click(sender As Object, e As EventArgs) Handles btnAdd.Click
        If lvList.CheckedItems.Count = 0 Then
            MessageBox.Show("No test select. Please tick one or more test in the Test Group Details.", "Empty Test", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            Exit Sub
        End If
        'FindListGetTest()
        FindListNoDups()

        For x As Integer = 0 To Me.lvList.CheckedItems.Count - 1
            Dim iItem As New ListViewItem(lvList.CheckedItems(x).Text, 0)
            iItem.SubItems.Add(lvList.CheckedItems(x).SubItems(1).Text)
            iItem.SubItems.Add(lvList.CheckedItems(x).SubItems(2).Text)
            iItem.SubItems.Add(lvList.CheckedItems(x).SubItems(4).Text)
            iItem.SubItems.Add(SampleID)

            iItem.SubItems.Add(lvList.CheckedItems(x).SubItems(3).Text)
            iItem.SubItems.Add(lvList.CheckedItems(x).SubItems(5).Text)
            iItem.SubItems.Add(lvTest.FocusedItem.Text)
            iItem.SubItems.Add(lvList.CheckedItems(x).SubItems(6).Text)
            lvOrder.Items.Add(iItem)
        Next

        'FindListViewCodes()

        MessageBox.Show(cboTestName.Text & " request for " & txtName.Text & " was successfuly saved.", "Create Order", MessageBoxButtons.OK, MessageBoxIcon.Information)
        'lvList.Items.Clear()
        chAll.CheckState = CheckState.Unchecked
    End Sub

    Private Sub FindListNoDups()
        Dim a As Integer = 0

        For x As Integer = 0 To Me.lvList.CheckedItems.Count - 1 Step 1

            If Not codeExists(lvList.CheckedItems(x).SubItems(4).Text) Then
                Dim iItems As New ListViewItem(lvList.Items(x).SubItems(4).Text)
                'For y As Integer = 0 To Me.lv.Items.Count - 1
                '    a = a + 1
                '    iItems.SubItems.Add(a)

                'Next
                iItems.SubItems.Add(lvList.Items(x).SubItems(4).Text)
                iItems.SubItems.Add(lvList.Items(x).SubItems(5).Text)
                iItems.SubItems.Add(lvList.Items(x).SubItems(6).Text)
                lv.Items.Add(iItems)
            Else
                'Exit Sub
            End If
        Next

    End Sub

    Private Sub FindListGetTest()
        If lv.Items.Count > 0 Then
            Dim iItems As ListViewItem
            For x As Integer = 0 To Me.lv.Items.Count - 1 Step 1
                For y As Integer = 0 To Me.lvList.CheckedItems.Count - 1 Step 1
                    If lvList.CheckedItems(y).SubItems(3).Text = lv.Items(x).SubItems(1).Text Then

                    Else
                        iItems = New ListViewItem((lv.Items.Count + 1).ToString, 0)
                        iItems.SubItems.Add(lvList.CheckedItems(y).SubItems(3).Text)
                        iItems.SubItems.Add(lvList.CheckedItems(y).SubItems(4).Text)
                        lv.Items.Add(iItems)
                    End If
                Next
            Next

            'For x As Integer = 0 To Me.lvList.CheckedItems.Count - 1 Step 1
            '    iItem.SubItems.Add(lvList.CheckedItems(x).SubItems(4).Text)
            '    cboRequest.Text = lvList.CheckedItems(x).SubItems(4).Text
            'Next
            'lv.Items.Add(iItems)
        ElseIf lv.Items.Count < 1 Then
            Dim iItem As New ListViewItem((1).ToString, 0)
            'iItem.SubItems.Add(cboTestName.Text)
            'For x As Integer = 0 To Me.lvList.CheckedItems.Count - 1 Step 1
            '    iItem.SubItems.Add(lvList.CheckedItems(x).SubItems(4).Text)
            '    cboRequest.Text = lvList.CheckedItems(x).SubItems(4).Text
            'Next
            'lv.Items.Add(iItem)

            For x As Integer = 0 To Me.lv.Items.Count - 1 Step 1
                For y As Integer = 0 To Me.lvList.CheckedItems.Count - 1 Step 1
                    If lvList.CheckedItems(y).SubItems(3).Text = lv.Items(x).SubItems(1).Text Then

                    Else
                        iItem = New ListViewItem((lv.Items.Count + 1).ToString, 0)
                        iItem.SubItems.Add(lvList.CheckedItems(y).SubItems(3).Text)
                        iItem.SubItems.Add(lvList.CheckedItems(y).SubItems(4).Text)
                        lv.Items.Add(iItem)
                    End If
                Next
            Next

        End If
    End Sub

    Private Sub FindListViewCodes()
        For x As Integer = 0 To Me.lvList.CheckedItems.Count - 1 Step 1
            If Not codeExists(lvList.CheckedItems(x).SubItems(4).Text) Then
                Dim iItems As New ListViewItem(lvList.CheckedItems(x).SubItems(2).Text, 0)
                For y = 0 To Me.lv.Items.Count - 1 Step 1
                    If cboTestName.Text = lv.Items(y).SubItems(1).Text Then
                        'iItems.SubItems.Add(SampleID & "-" & lv.Items(y).SubItems(0).Text)
                        iItems.SubItems.Add(SampleID)
                        iItems.SubItems.Add(lv.Items(y).SubItems(1).Text)
                        iItems.SubItems.Add(lv.Items(y).SubItems(2).Text)
                    End If
                Next
                lvDup.Items.Add(iItems)
            Else
                Exit Sub
            End If
        Next
    End Sub

    Private Sub chAll_CheckedChanged(sender As Object, e As EventArgs) Handles chAll.CheckedChanged
        If chAll.CheckState = CheckState.Checked Then
            For x As Integer = lvList.Items.Count - 1 To 0 Step -1
                lvList.Items(x).Checked = True
            Next
        Else
            For x As Integer = lvList.Items.Count - 1 To 0 Step -1
                lvList.Items(x).Checked = False
            Next
        End If
    End Sub

    Private Sub SaveBarcodeDetails()
        For x As Integer = 0 To Me.lv.Items.Count - 1 Step 1
            bcCode.Symbology = New Code128Generator With {.CharacterSet = Code128Charset.CharsetAuto}
            bcCode.Text = SampleID
            Dim bmp = New Bitmap(bcCode.Width, bcCode.Height)

            bcCode.DrawToBitmap(bmp, bcCode.ClientRectangle)
            picCode.Image = bmp

            Dim myMS As New IO.MemoryStream
            If Not IsNothing(picCode.Image) Then
                picCode.Image.Save(myMS, Imaging.ImageFormat.Jpeg)
                arrImage = myMS.GetBuffer
            Else
                arrImage = Nothing
            End If

            rs.Parameters.Clear()
            rs.Parameters.AddWithValue("@SampleID", SampleID)
            rs.Parameters.AddWithValue("@SpecimenType", lv.Items(x).SubItems(3).Text)
            rs.Parameters.AddWithValue("@Code", "")
            rs.Parameters.AddWithValue("@Name", txtName.Text)
            rs.Parameters.AddWithValue("@Age", txtAge.Text)
            rs.Parameters.AddWithValue("@Sex", cboSex.Text)
            rs.Parameters.AddWithValue("@Test", "")
            rs.Parameters.AddWithValue("@DateReceived", Now)
            rs.Parameters.AddWithValue("@Section", lv.Items(x).SubItems(0).Text)
            rs.Parameters.AddWithValue("@SubSection", lv.Items(x).SubItems(2).Text)
            rs.Parameters.AddWithValue("@Barcode", arrImage)

            Connect()
            rs.Connection = conn
            rs.CommandType = CommandType.Text
            rs.CommandText = "INSERT INTO `barcode_details` (`sample_id`, `specimen_type`, `code`, `name`, `age`, `sex`, `test`, `datereceived`, `section`, `subsection`, `barcode`) VALUES " _
              & "(" _
              & "@SampleID," _
              & "@SpecimenType," _
              & "@Code," _
              & "@Name," _
              & "@Age," _
              & "@Sex," _
              & "@Test," _
              & "@DateReceived," _
              & "@Section," _
              & "@SubSection," _
              & "@SampleID" _
              & ")"
            rs.ExecuteNonQuery()
            Disconnect()
        Next

    End Sub

    Private Sub btnSave_ItemClick(sender As Object, e As DevExpress.XtraBars.ItemClickEventArgs) Handles btnSave.ItemClick
        If Me.txtName.Text = "" Or lvOrder.Items.Count = 0 Then
            MessageBox.Show("Unable to save orders." & vbNewLine & vbNewLine & "No order in order list.", "System Warning", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            Exit Sub
        End If
        SaveBarcodeDetails()
        SaveWorklist()

        Dim Message As DialogResult

        Message = MessageBox.Show("Patient order successfully save." & vbNewLine & vbNewLine & "Do you want to add another order? ", "Success", MessageBoxButtons.YesNo, MessageBoxIcon.Question)
        If Message = vbYes Then
            For Each ctrl As Control In Me.GroupPatient.Controls
                If TypeOf ctrl Is TextEdit Or TypeOf ctrl Is ComboBoxEdit Or TypeOf ctrl Is DateEdit Then
                    ctrl.Text = ""
                End If
            Next

            For Each ctrl As Control In Me.GroupAdditional.Controls
                If TypeOf ctrl Is TextEdit Or TypeOf ctrl Is ComboBoxEdit Or TypeOf ctrl Is DateEdit Then
                    ctrl.Text = ""
                End If
            Next

            lvList.Items.Clear()
            'lvTest.Items.Clear()
            lvOrder.Items.Clear()
            lv.Clear()
            cboTestName.Text = ""
            '#########################Fixed Clear Patient Details in Patient Order 10/11/2019#############################
            cboPType.Text = "Out Patient"
            cboMode.Text = "Routine"
            '#########################Fixed Clear Patient Details in Patient Order 10/11/2019#############################
            LoadSampleID()
            txtSampleID.Focus()
            'AutoPatientID()
            frmPatientOrder.LoadRecords()
        Else
            Me.Close()
        End If
    End Sub

    Private Sub btnSavePrint_ItemClick(sender As Object, e As DevExpress.XtraBars.ItemClickEventArgs) Handles btnSavePrint.ItemClick
        If Me.txtName.Text = "" Or lvOrder.Items.Count = 0 Then
            MessageBox.Show("Unable to save orders." & vbNewLine & vbNewLine & "No order in order list.", "System Warning", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            Exit Sub
        End If

        SaveWorklistBCode()
        SaveBarcodeDetails()
        Dim Message As DialogResult

        Message = MessageBox.Show("Patient order successfully save." & vbNewLine & vbNewLine & "Do you want to add another order? ", "Success", MessageBoxButtons.YesNo, MessageBoxIcon.Question)
        If Message = vbYes Then
            For Each ctrl As Control In Me.GroupPatient.Controls
                If TypeOf ctrl Is TextEdit Or TypeOf ctrl Is ComboBoxEdit Or TypeOf ctrl Is DateEdit Then
                    ctrl.Text = ""
                End If
            Next

            For Each ctrl As Control In Me.GroupAdditional.Controls
                If TypeOf ctrl Is TextEdit Or TypeOf ctrl Is ComboBoxEdit Or TypeOf ctrl Is DateEdit Then
                    ctrl.Text = ""
                End If
            Next

            lvList.Items.Clear()
            'lvTest.Items.Clear()
            lvOrder.Items.Clear()
            lv.Clear()
            cboTestName.Text = ""
            '#########################Fixed Clear Patient Details in Patient Order 10/11/2019#############################
            cboPType.Text = "Out Patient"
            cboMode.Text = "Routine"
            '#########################Fixed Clear Patient Details in Patient Order 10/11/2019#############################
            txtSampleID.Focus()
            LoadSampleID()
            'AutoPatientID()
            frmPatientOrder.LoadRecords()
        Else
            Me.Close()
        End If
    End Sub

    Private Sub btnClear_ItemClick(ByVal sender As Object, ByVal e As DevExpress.XtraBars.ItemClickEventArgs) Handles btnClear.ItemClick
        For Each ctrl As Control In Me.GroupAdditional.Controls
            If TypeOf ctrl Is TextEdit Or TypeOf ctrl Is ComboBoxEdit Or TypeOf ctrl Is DateEdit Then
                ctrl.Text = ""
            End If
        Next

        For Each ctrl As Control In Me.GroupPatient.Controls
            If TypeOf ctrl Is TextEdit Or TypeOf ctrl Is ComboBoxEdit Or TypeOf ctrl Is DateEdit Then
                ctrl.Text = ""
            End If
        Next

        lvList.Items.Clear()
        lvOrder.Items.Clear()
        lv.Items.Clear()
        'lvTest.Items.Clear()
        cboTestName.Text = ""
        '#########################Fixed Clear Patient Details in Patient Order 10/11/2019#############################
        cboPType.Text = "Out Patient"
        cboMode.Text = "Routine"
        '#########################Fixed Clear Patient Details in Patient Order 10/11/2019#############################
        txtSampleID.Focus()
        LoadSampleID()
        'AutoPatientID()
    End Sub

    Private Sub LoadSampleID()
        Connect()
        rs.Connection = conn
        rs.CommandType = CommandType.Text
        rs.CommandText = "Select `lastidno` FROM `last_id` ORDER BY `lastidno` DESC LIMIT 1"
        reader = rs.ExecuteReader
        reader.Read()
        If reader.HasRows Then
            LastID = reader(0)
            If reader(0) > 0 And reader(0) <= 9 Then
                SampleID = "000000" & reader(0).ToString + 1
            ElseIf reader(0) > 9 And reader(0) <= 99 Then
                SampleID = "00000" & reader(0).ToString + 1
            ElseIf reader(0) > 99 And reader(0) <= 999 Then
                SampleID = "0000" & reader(0).ToString + 1
            ElseIf reader(0) > 999 And reader(0) <= 9999 Then
                SampleID = "000" & reader(0).ToString + 1
            ElseIf reader(0) > 9999 And reader(0) <= 99999 Then
                SampleID = "00" & reader(0).ToString + 1
            ElseIf reader(0) > 99999 And reader(0) <= 999999 Then
                SampleID = "0" & reader(0).ToString + 1
            ElseIf reader(0) > 999999 And reader(0) <= 9999999 Then
                SampleID = reader(0).ToString + 1
            End If
        Else
            LastID = 1
            SampleID = "000000" & reader(0).ToString + 1
        End If
        Disconnect()
        txtSampleID.Text = SampleID
    End Sub

    Private Sub btnSearch_Click(sender As Object, e As EventArgs) Handles btnSearch.Click
        frmPatientList.Type = "Order"
        frmPatientList.ShowDialog()
        LoadSampleID()
    End Sub

    Private Sub SaveOrders()
        'Try
        For y As Integer = 0 To lvOrder.Items.Count - 1 Step 1
            Connect()
            rs.Connection = conn
            rs.CommandType = CommandType.Text
            rs.CommandText = ("INSERT INTO `lis_order` (`sample_id`, `mode`, `order_code`, `specimen_type`, `section`, `sub_section`) VALUES " _
                        & "(" _
                        & "'" & SampleID & "'," _
                        & "'" & cboMode.Text.Substring(0) & "'," _
                        & "'" & lvOrder.Items(y).SubItems(2).Text & "'," _
                        & "'1'," _
                        & "'" & lvOrder.Items(y).SubItems(3).Text & "'," _
                        & "'" & lvOrder.Items(y).SubItems(6).Text & "'" _
                        & ")"
                        )
            rs.ExecuteNonQuery()
            Disconnect()
        Next

        For x As Integer = 0 To lvOrder.Items.Count - 1 Step 1
            GetTimeReceived(SampleID, Format(Now, "yyyy-MM-dd HH:mm:ss"), lvOrder.Items(x).SubItems(3).Text, lvOrder.Items(x).SubItems(6).Text)

            Connect()
            rs.Connection = conn
            rs.CommandType = CommandType.Text
            rs.CommandText = ("INSERT INTO `patient_order` (`test_name`, `mode`, `sample_id`, `test_code`, `patient_id`, `testtype`, `sub_section`, `his_code`) VALUES " _
                    & "(" _
                    & "'" & lvOrder.Items(x).SubItems(1).Text & "'," _
                    & "'" & cboMode.Text.Substring(0) & "'," _
                    & "'" & lvOrder.Items(x).SubItems(4).Text & "'," _
                    & "'" & lvOrder.Items(x).SubItems(2).Text & "'," _
                    & "'" & txtPatientID.Text & "'," _
                    & "'" & lvOrder.Items(x).SubItems(3).Text & "'," _
                    & "'" & lvOrder.Items(x).SubItems(6).Text & "'," _
                    & "'" & lvOrder.Items(x).SubItems(7).Text & "'" _
                    & ")"
                    )
            rs.ExecuteNonQuery()
            Disconnect()

            Connect()
            rs.Connection = conn
            rs.CommandType = CommandType.Text
            rs.CommandText = "SELECT `specimen`, `test_code`, `order_no`, `test_group`, `si_unit`, `conventional_unit`, `his_code` FROM `default_specimen` where `test_code` = '" & lvOrder.Items(x).SubItems(2).Text & "' AND his_code = '" & lvOrder.Items(x).SubItems(7).Text & "' AND `status` = 'Enable'"
            reader = rs.ExecuteReader

            reader.Read()
            If reader.HasRows Then
                test_name = reader(0).ToString
                test_code = reader(1).ToString
                order_no = reader(2).ToString
                test_group = reader(3).ToString
                unit = reader(4).ToString
                unit_conv = reader(5).ToString

            End If
            Disconnect()

            Connect()
            rs.Connection = conn
            rs.CommandType = CommandType.Text
            rs.CommandText = ("INSERT INTO `tmpResult` (`universal_id`, `measurement`, `test_code`, `sample_id`, `date`, `patient_id`, `order_no`, `units`, `unit_conv`, `instrument`, `status`, `test_group`, `section`, `sub_section`) VALUES " _
                    & "(" _
                    & "'" & test_name & "', '', '" & test_code & "', '" & lvOrder.Items(x).SubItems(4).Text & "', NOW(), '" & txtPatientID.Text & "', '" & order_no & "', '" & unit & "', '" & unit_conv & "', 'Other_Test', 0, '" & test_group & "', '" & lvOrder.Items(x).SubItems(3).Text & "', '" & lvOrder.Items(x).SubItems(6).Text & "'" _
                    & ")"
                    )
            rs.ExecuteNonQuery()
            Disconnect()
        Next
        'Catch ex As Exception
        '    'Disconnect()
        '    'MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
        '    'Exit Sub
        'End Try
    End Sub

    Private Sub SaveWorklist()
        Try
            'Save the last Sample ID for next patient increament by 1
            Connect()
            rs.Connection = conn
            rs.CommandType = CommandType.Text
            rs.CommandText = ("INSERT INTO `last_id` VALUES ('" & LastID + 1 & "')")
            rs.ExecuteNonQuery()
            Disconnect()

            For x As Integer = 0 To lv.Items.Count - 1 Step 1
                If Me.txtName.Text = "" Or Me.txtAge.Text = "" Or txtPatientID.Text = "" Or cboSex.Text = "" Then
                    MessageBox.Show("Please verify the data carefully.", "Verification Failed", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
                    Exit Sub
                End If

                bcCode.Symbology = New Code128Generator With {.CharacterSet = Code128Charset.CharsetAuto}
                bcCode.Text = SampleID
                Dim bmp = New Bitmap(bcCode.Width, bcCode.Height)

                bcCode.DrawToBitmap(bmp, bcCode.ClientRectangle)
                picCode.Image = bmp

                Dim myMS As New IO.MemoryStream
                If Not IsNothing(picCode.Image) Then
                    picCode.Image.Save(myMS, Imaging.ImageFormat.Jpeg)
                    arrImage = myMS.GetBuffer
                Else
                    arrImage = Nothing
                End If

                rs.Parameters.Clear()
                rs.Parameters.AddWithValue("@sample_id", SampleID)
                rs.Parameters.AddWithValue("@NAME", txtName.Text)
                rs.Parameters.AddWithValue("@AGE", txtAge.Text)
                rs.Parameters.AddWithValue("@bdate", dtBDate.Text)
                rs.Parameters.AddWithValue("@PHYSICIAN", cboPhysician.Text)
                rs.Parameters.AddWithValue("@DATE", Now)
                rs.Parameters.AddWithValue("@TIME", Now.ToLongTimeString)
                rs.Parameters.AddWithValue("@PATIENT_ID", txtPatientID.Text)
                rs.Parameters.AddWithValue("@dept", cboRoom.Text)
                rs.Parameters.AddWithValue("@sex", cboSex.Text)
                rs.Parameters.AddWithValue("@status", "Checked-In")
                rs.Parameters.AddWithValue("@address", txtAddress.Text)
                rs.Parameters.AddWithValue("@civilstatus", cboStatus.Text)
                rs.Parameters.AddWithValue("@contact", txtContact.Text)
                rs.Parameters.AddWithValue("@Barcode", arrImage)
                rs.Parameters.AddWithValue("@TestType", lv.Items(x).SubItems(1).Text)
                rs.Parameters.AddWithValue("@Test", lv.Items(x).SubItems(2).Text)
                rs.Parameters.AddWithValue("@Patient_Type", cboPType.Text)
                rs.Parameters.AddWithValue("@TYPE", txtClass.Text)
                'Check This
                rs.Parameters.AddWithValue("@location", CurrDept)

                'rs.Parameters.AddWithValue("@MainID", SampleID & "-" & lv.Items(x).SubItems(0).Text)
                rs.Parameters.AddWithValue("@MainID", SampleID)
                rs.Parameters.AddWithValue("@accession_no", txtAccession.Text)
                rs.Parameters.AddWithValue("@OR_No", txtORNo.Text)
                rs.Parameters.AddWithValue("@ChargeSlip", txtChargeSlip.Text)
                rs.Parameters.AddWithValue("@Instrument", "Other_Test")
                rs.Parameters.AddWithValue("@SubSection", lv.Items(x).SubItems(2).Text)
                rs.Parameters.AddWithValue("@EmailAddress", txtEmail.Text)

                Connect()
                rs.Connection = conn
                rs.CommandType = CommandType.Text
                rs.CommandText = "SELECT * FROM patient_info WHERE `patient_id` LIKE @PATIENT_ID"
                reader = rs.ExecuteReader
                reader.Read()
                If reader.HasRows Then
                    Disconnect()
                Else
                    Disconnect()
                    Connect()
                    rs.Connection = conn
                    rs.CommandType = CommandType.Text
                    rs.CommandText = ("INSERT INTO patient_info (patient_id, name, sex, date_of_birth, age, civil_status, address, contact_no, `date`) VALUES " _
                                & "(" _
                                & "@PATIENT_ID," _
                                & "@NAME," _
                                & "@sex," _
                                & "@bdate," _
                                & "@AGE," _
                                & "@civilstatus," _
                                & "@address," _
                                & "@contact," _
                                & "@DATE" _
                                & ")"
                                )
                    rs.ExecuteNonQuery()
                    Disconnect()
                End If

                Connect()
                rs.Connection = conn
                rs.CommandType = CommandType.Text
                rs.CommandText = "SELECT * FROM `tmpWorklist` WHERE `sample_id` = @sample_id AND testtype = @TestType AND sub_section = @SubSection"
                reader = rs.ExecuteReader
                reader.Read()
                If reader.HasRows Then
                    Disconnect()
                    Connect()
                    rs.Connection = conn
                    rs.CommandType = CommandType.Text
                    rs.CommandText = ("UPDATE `tmpWorklist` SET " _
                                & "`patient_name` = @NAME," _
                                & "`sex` = @sex," _
                                & "`bdate` = @bdate," _
                                & "`age` = @AGE," _
                                & "`physician` = @PHYSICIAN," _
                                & "`dept` = @dept," _
                                & "`test` = @Test," _
                                & "`patient_type` = @Patient_Type" _
                                & " WHERE main_id = @sample_id AND testtype = @TestType AND sub_section = @SubSection")
                    rs.ExecuteNonQuery()
                    Disconnect()
                Else
                    Disconnect()
                    Connect()
                    rs.Connection = conn
                    rs.CommandType = CommandType.Text
                    rs.CommandText = ("INSERT INTO `tmpWorklist` (`sample_id`, `patient_id`, `patient_name`, `sex`, `bdate`, `age`, `physician`, `dept`, `status`, `main_id`, `date`, `time`, `barcode`, `testtype`, `test`, `patient_type`, `TYPE`, `location`, `instrument`, `sub_section`) VALUES " _
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
                                & "@Test," _
                                & "@Patient_Type," _
                                & "@TYPE," _
                                & "@location," _
                                & "@Instrument," _
                                & "@SubSection" _
                                & ")"
                                )
                    rs.ExecuteNonQuery()
                    Disconnect()
                End If
                Disconnect()

                Connect()
                rs.Connection = conn
                rs.CommandType = CommandType.Text
                rs.CommandText = ("INSERT INTO `additional_info` (`accession_no`, `or_no`, `cs_no`, `sample_id`, `section`, `sub_section`) VALUES " _
                            & "(" _
                            & "@accession_no," _
                            & "@OR_No," _
                            & "@ChargeSlip," _
                            & "@mainID," _
                            & "@TESTTYPE," _
                            & "@SubSection" _
                            & ")"
                            )
                rs.ExecuteNonQuery()
                Disconnect()

                '------------------Save Email Details------------------------------
                Connect()
                rs.Connection = conn
                rs.CommandType = CommandType.Text
                rs.CommandText = ("INSERT INTO `email_details` (`sample_id`, `email_address`, `section`, `sub_section`) VALUES " _
                            & "(" _
                            & "@mainID," _
                            & "@EmailAddress," _
                            & "@TESTTYPE," _
                            & "@SubSection" _
                            & ")"
                            )
                rs.ExecuteNonQuery()
                Disconnect()
                '------------------Save Email Details------------------------------
            Next

            SaveOrders()

            frmPatientOrder.LoadRecords()
        Catch ex As Exception
            Disconnect()
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            Exit Sub
        End Try
    End Sub

    Private Sub SaveWorklistBCode()
        Try
            'Save the last Sample ID for next patient increament by 1
            Connect()
            rs.Connection = conn
            rs.CommandType = CommandType.Text
            rs.CommandText = ("INSERT INTO `last_id` VALUES ('" & LastID + 1 & "')")
            rs.ExecuteNonQuery()
            Disconnect()

            For x As Integer = 0 To lv.Items.Count - 1 Step 1
                If Me.txtName.Text = "" Or Me.txtAge.Text = "" Or txtPatientID.Text = "" Or cboSex.Text = "" Then
                    MessageBox.Show("Please verify the data carefully.", "Verification Failed", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
                    Exit Sub
                End If

                'bcCode.Symbology = New Code128Generator With {.CharacterSet = Code128Charset.CharsetAuto}

                'bcCode.Text = SampleID

                'BarcodeID = SampleID

                'Dim bmp = New Bitmap(bcCode.Width, bcCode.Height)
                'bcCode.DrawToBitmap(bmp, bcCode.ClientRectangle)
                'picCode.Image = bmp

                'Dim myMS As New IO.MemoryStream
                'If Not IsNothing(picCode.Image) Then
                '    picCode.Image.Save(myMS, Imaging.ImageFormat.Jpeg)
                '    arrImage = myMS.GetBuffer
                'Else
                '    arrImage = Nothing
                'End If

                'SpecimenType = lv.Items(x).SubItems(3).Text

                'If My.Settings.PrintBarcode Then
                '    'Paper Size
                '    Dim papersize As New PaperSize("2x1 Inch Label Size", 200, 100)
                '    PrintDocument.PrinterSettings.PrinterName = StrPrinterName
                '    PrintDocument.PrintController = New StandardPrintController

                '    PrintDocument.DefaultPageSettings.PaperSize = papersize
                '    PrintDocument.DefaultPageSettings.Margins = New Margins(0, 0, 0, 0)
                '    PrintDocument.OriginAtMargins = False
                '    PrintDocument.DefaultPageSettings.Landscape = False

                '    'Printing
                '    PrintPreviewDialog.Document = PrintDocument
                '    PrintPreviewDialog.Document.Print()
                '    'PrintPreviewDialog.ShowDialog()
                'End If

                Try
                    PrintBarcode(lv.Items(x).SubItems(3).Text,
                                     SampleID,
                                     txtPatientID.Text,
                                     txtName.Text,
                                     dtBDate.Text,
                                     cboSex.Text,
                                     lv.Items(x).SubItems(1).Text,
                                     lv.Items(x).SubItems(2).Text)
                Catch ex As Exception
                    MessageBox.Show("Error in connection on printer. " + ex.Message, "Barcode Printing Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                End Try

                rs.Parameters.Clear()
                rs.Parameters.AddWithValue("@sample_id", SampleID)
                rs.Parameters.AddWithValue("@NAME", txtName.Text)
                rs.Parameters.AddWithValue("@AGE", txtAge.Text)
                rs.Parameters.AddWithValue("@bdate", dtBDate.Text)
                rs.Parameters.AddWithValue("@PHYSICIAN", cboPhysician.Text)
                rs.Parameters.AddWithValue("@DATE", Now)
                rs.Parameters.AddWithValue("@TIME", Now.ToLongTimeString)
                rs.Parameters.AddWithValue("@PATIENT_ID", txtPatientID.Text)
                rs.Parameters.AddWithValue("@dept", cboRoom.Text)
                rs.Parameters.AddWithValue("@sex", cboSex.Text)
                rs.Parameters.AddWithValue("@status", "Checked-In")
                rs.Parameters.AddWithValue("@address", txtAddress.Text)
                rs.Parameters.AddWithValue("@civilstatus", cboStatus.Text)
                rs.Parameters.AddWithValue("@contact", txtContact.Text)
                rs.Parameters.AddWithValue("@Barcode", arrImage)
                rs.Parameters.AddWithValue("@TestType", lv.Items(x).SubItems(1).Text)
                rs.Parameters.AddWithValue("@Test", lv.Items(x).SubItems(2).Text)
                rs.Parameters.AddWithValue("@Patient_Type", cboPType.Text)
                rs.Parameters.AddWithValue("@TYPE", txtClass.Text)
                'Check This
                rs.Parameters.AddWithValue("@location", CurrDept)

                'rs.Parameters.AddWithValue("@MainID", SampleID & "-" & lv.Items(x).SubItems(0).Text)
                rs.Parameters.AddWithValue("@MainID", SampleID)
                rs.Parameters.AddWithValue("@accession_no", txtAccession.Text)
                rs.Parameters.AddWithValue("@OR_No", txtORNo.Text)
                rs.Parameters.AddWithValue("@ChargeSlip", txtChargeSlip.Text)
                rs.Parameters.AddWithValue("@Instrument", "Other_Test")
                rs.Parameters.AddWithValue("@SubSection", lv.Items(x).SubItems(2).Text)
                rs.Parameters.AddWithValue("@EmailAddress", txtEmail.Text)

                Connect()
                rs.Connection = conn
                rs.CommandType = CommandType.Text
                rs.CommandText = "SELECT * FROM patient_info WHERE `patient_id` LIKE @PATIENT_ID"
                reader = rs.ExecuteReader
                reader.Read()
                If reader.HasRows Then
                    Disconnect()
                Else
                    Disconnect()
                    Connect()
                    rs.Connection = conn
                    rs.CommandText = ("INSERT INTO patient_info (patient_id, name, sex, date_of_birth, age, civil_status, address, contact_no, `date`) VALUES " _
                                & "(" _
                                & "@PATIENT_ID," _
                                & "@NAME," _
                                & "@sex," _
                                & "@bdate," _
                                & "@AGE," _
                                & "@civilstatus," _
                                & "@address," _
                                & "@contact," _
                                & "@DATE" _
                                & ")"
                                )
                    rs.ExecuteNonQuery()
                    Disconnect()
                End If

                Connect()
                rs.Connection = conn
                rs.CommandType = CommandType.Text
                rs.CommandText = "SELECT * FROM `tmpWorklist` WHERE `sample_id` = @sample_id AND testtype = @TestType AND sub_section = @SubSection"
                reader = rs.ExecuteReader
                reader.Read()
                If reader.HasRows Then
                    Disconnect()
                    Connect()
                    rs.Connection = conn
                    rs.CommandType = CommandType.Text
                    rs.CommandText = ("UPDATE `tmpWorklist` SET " _
                                & "`patient_name` = @NAME," _
                                & "`sex` = @sex," _
                                & "`bdate` = @bdate," _
                                & "`age` = @AGE," _
                                & "`physician` = @PHYSICIAN," _
                                & "`dept` = @dept," _
                                & "`test` = @Test," _
                                & "`patient_type` = @Patient_Type" _
                                & " WHERE main_id = @sample_id AND testtype = @TestType AND sub_section = @SubSection")
                    rs.ExecuteNonQuery()
                    Disconnect()
                Else
                    Disconnect()
                    Connect()
                    rs.Connection = conn
                    rs.CommandType = CommandType.Text
                    rs.CommandText = ("INSERT INTO `tmpWorklist` (`sample_id`, `patient_id`, `patient_name`, `sex`, `bdate`, `age`, `physician`, `dept`, `status`, `main_id`, `date`, `time`, `barcode`, `testtype`, `test`, `patient_type`, `TYPE`, `location`, `instrument`, `sub_section`) VALUES " _
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
                                & "@Test," _
                                & "@Patient_Type," _
                                & "@TYPE," _
                                & "@location," _
                                & "@Instrument," _
                                & "@SubSection" _
                                & ")"
                                )
                    rs.ExecuteNonQuery()
                    Disconnect()
                End If
                Disconnect()

                Connect()
                rs.Connection = conn
                rs.CommandType = CommandType.Text
                rs.CommandText = ("INSERT INTO `additional_info` (`accession_no`, `or_no`, `cs_no`, `sample_id`, `section`, `sub_section`) VALUES " _
                            & "(" _
                            & "@accession_no," _
                            & "@OR_No," _
                            & "@ChargeSlip," _
                            & "@mainID," _
                            & "@TESTTYPE," _
                            & "@SubSection" _
                            & ")"
                            )
                rs.ExecuteNonQuery()
                Disconnect()

                '------------------Save Email Details------------------------------
                Connect()
                rs.Connection = conn
                rs.CommandType = CommandType.Text
                rs.CommandText = ("INSERT INTO `email_details` (`sample_id`, `email_address`, `section`, `sub_section`) VALUES " _
                            & "(" _
                            & "@mainID," _
                            & "@EmailAddress," _
                            & "@TESTTYPE," _
                            & "@SubSection" _
                            & ")"
                            )
                rs.ExecuteNonQuery()
                Disconnect()
                '------------------Save Email Details------------------------------
            Next

            SaveOrders()

            frmPatientOrder.LoadRecords()
        Catch ex As Exception
            Disconnect()
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            Exit Sub
        End Try
    End Sub

    Private Sub GetTimeReceived(SampleID As String, Time As String, Section As String, SubSection As String)
        Try
            rs.Parameters.Clear()
            rs.Parameters.AddWithValue("@SampleID", SampleID)
            rs.Parameters.AddWithValue("@Time", Time)
            rs.Parameters.AddWithValue("@Section", Section)
            rs.Parameters.AddWithValue("@SubSection", SubSection)
            Connect()
            rs.Connection = conn
            rs.CommandType = CommandType.Text
            rs.CommandText = "SELECT `sample_id` FROM `specimen_tracking` WHERE sample_id = @SampleID AND section = @Section AND sub_section = @SubSection"
            reader = rs.ExecuteReader
            reader.Read()
            If reader.HasRows Then
                Disconnect()
                Exit Sub
            Else
                Disconnect()

                SaveRecordwthoutMSG("INSERT INTO `specimen_tracking` (`sample_id`, `received`, `extracted`, `section`, `sub_section`) VALUES " _
                    & "(" _
                    & "@SampleID, @Time, @Time, @Section, @SubSection" _
                    & ")"
                    )
            End If
            Disconnect()
        Catch ex As Exception
            Disconnect()
            MessageBox.Show(ex.Message)
            Exit Sub
        End Try
    End Sub

    Public Sub AutoLoadRoom()
        Try
            Me.cboRoom.Properties.Items.Clear()
            Connect()
            rs.Connection = conn
            rs.CommandType = CommandType.Text
            rs.CommandText = "SELECT * FROM `department` ORDER BY `dept_name`"
            reader = rs.ExecuteReader
            While reader.Read
                Me.cboRoom.Properties.Items.Add(reader(1).ToString)
            End While
            Disconnect()
        Catch ex As Exception
            Disconnect()
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            Exit Sub
        End Try
    End Sub

    Public Sub AutoLoadDoctor()
        Try
            Me.cboPhysician.Properties.Items.Clear()
            Connect()
            rs.Connection = conn
            rs.CommandType = CommandType.Text
            rs.CommandText = "SELECT name FROM `requestor` ORDER BY `name`"
            reader = rs.ExecuteReader
            While reader.Read
                Me.cboPhysician.Properties.Items.Add(reader(0))
            End While
            Disconnect()
        Catch ex As Exception
            Disconnect()
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            Exit Sub
        End Try
    End Sub

    Private Sub btnRemove_Click(sender As Object, e As EventArgs) Handles btnRemove.Click
        lvOrder.FocusedItem.Remove()
    End Sub

    Private Sub btnDelete_Click(sender As Object, e As EventArgs) Handles btnDelete.Click
        lvOrder.Items.Clear()
    End Sub

    Public Sub AutoLoadTestName()
        Try
            Me.cboTestName.Properties.Items.Clear()
            Connect()
            rs.Connection = conn
            rs.CommandType = CommandType.Text
            rs.CommandText = "SELECT `test_name` FROM `testtype` WHERE `test_name` NOT LIKE 'All' ORDER BY `test_name`"
            reader = rs.ExecuteReader
            While reader.Read
                Me.cboTestName.Properties.Items.Add(reader(0).ToString)
            End While
            Disconnect()
        Catch ex As Exception
            Disconnect()
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            Exit Sub
        End Try
    End Sub

    Public Sub AutoPatientID()
        Try
            Connect()
            rs.Connection = conn
            rs.CommandType = CommandType.Text
            rs.CommandText = "SELECT (id) as id FROM `patient_info` ORDER BY `id` DESC"
            reader = rs.ExecuteReader
            reader.Read()
            If reader.HasRows Then
                If reader(0).ToString > 0 And reader(0).ToString <= 9 Then
                    txtPatientID.Text = Now.Year & "-" & Now.Day & "-" & "000000" & reader(0).ToString + 1
                ElseIf reader(0).ToString > 9 And reader(0).ToString <= 99 Then
                    txtPatientID.Text = Now.Year & "-" & Now.Day & "-" & "00000" & reader(0).ToString + 1
                ElseIf reader(0).ToString > 99 And reader(0).ToString <= 999 Then
                    txtPatientID.Text = Now.Year & "-" & Now.Day & "-" & "0000" & reader(0).ToString + 1
                ElseIf reader(0).ToString > 999 And reader(0).ToString <= 9999 Then
                    txtPatientID.Text = Now.Year & "-" & Now.Day & "-" & "000" & reader(0).ToString + 1
                ElseIf reader(0).ToString > 9999 And reader(0).ToString <= 99999 Then
                    txtPatientID.Text = Now.Year & "-" & Now.Day & "-" & "00" & reader(0).ToString + 1
                ElseIf reader(0).ToString > 99999 And reader(0).ToString <= 999999 Then
                    txtPatientID.Text = Now.Year & "-" & Now.Day & "-" & "0" & reader(0).ToString + 1
                ElseIf reader(0).ToString > 999999 And reader(0).ToString <= 9999999 Then
                    txtPatientID.Text = reader(0).ToString + 1
                End If
            Else
                txtPatientID.Text = Now.Year & "-" & Now.Day & "-" & "000000" & 1
            End If
            Disconnect()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Auto Patient ID", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Disconnect()
            Exit Sub
        End Try
    End Sub

    Dim Specimen As String

    'Private Sub PrintDocument_PrintPage(ByVal sender As System.Object, ByVal e As System.Drawing.Printing.PrintPageEventArgs) Handles PrintDocument.PrintPage
    '    'DEDVMH Barcode Format
    '    'Try
    '    '    Dim bm As New Bitmap(Me.bcCode.Width, Me.bcCode.Height)

    '    '    bcCode.DrawToBitmap(bm, Me.bcCode.ClientRectangle)

    '    '    'Font and Size of Font
    '    '    Dim myFont As Font = New Font("Tahoma", 5.8, FontStyle.Bold)
    '    '    Dim FontArrow As Font = New Font("Tahoma", 12, FontStyle.Bold)
    '    '    Dim myBrush As Brush = Brushes.Black
    '    '    Dim ID As String = SampleID & " - " & SpecimenType
    '    '    Dim Name As String = txtName.Text
    '    '    Dim Test As String = cboRequest.Text
    '    '    Dim dateAge As String = txtAge.Text
    '    '    Dim gender As String = cboSex.Text

    '    '    e.Graphics.InterpolationMode = Drawing2D.InterpolationMode.HighQualityBilinear
    '    '    e.Graphics.SmoothingMode = Drawing2D.SmoothingMode.HighQuality
    '    '    e.Graphics.DrawString(Name, myFont, myBrush, 10, 5)
    '    '    e.Graphics.DrawString(dateAge, myFont, myBrush, 145, 5)
    '    '    e.Graphics.DrawString(gender, myFont, myBrush, 160, 5)
    '    '    e.Graphics.DrawString("→", FontArrow, myBrush, 170, 30)
    '    '    e.Graphics.DrawImage(bm, 10, 15, 100, 40)
    '    '    e.Graphics.DrawString(ID, myFont, myBrush, 30, 60)
    '    '    e.Graphics.DrawString(Test, myFont, myBrush, 15, 70)
    '    '    e.Graphics.DrawString(Now, myFont, myBrush, 15, 80)

    '    'Catch ex As Exception
    '    '    Disconnect()
    '    '    MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
    '    '    Exit Sub
    '    'End Try

    '    Try
    '        Dim bm As New Bitmap(Me.bcCode.Width, Me.bcCode.Height)

    '        bcCode.DrawToBitmap(bm, Me.bcCode.ClientRectangle)

    '        'Font and Size of Font
    '        Dim myFont As Font = New Font("Arial", 7, FontStyle.Regular)

    '        Dim myBrush As Brush = Brushes.Black
    '        Dim Name As String = txtName.Text
    '        Dim ID As String = BarcodeID
    '        'Dim Test As String = lvList.FocusedItem.SubItems(5).Text
    '        Dim DateTime As String = Now

    '        If Test = "CBC" Then
    '            Specimen = "WB"
    '        ElseIf Test = "Coagulation" Then
    '            Specimen = "Plasma"
    '        End If

    '        e.Graphics.InterpolationMode = Drawing2D.InterpolationMode.HighQualityBilinear
    '        e.Graphics.SmoothingMode = Drawing2D.SmoothingMode.HighQuality

    '        e.Graphics.DrawString(ID, myFont, myBrush, 5, 5)
    '        e.Graphics.DrawString(Specimen, myFont, myBrush, 80, 5)
    '        e.Graphics.DrawString(DateTime, myFont, myBrush, 110, 5)
    '        e.Graphics.DrawImage(bm, 10, 15, 100, 40)
    '        e.Graphics.DrawString(Name, myFont, myBrush, 5, 57)
    '        'e.Graphics.DrawString(Test, myFont, myBrush, 5, 59)
    '    Catch ex As Exception
    '        Disconnect()
    '        MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
    '        Exit Sub
    '    End Try
    'End Sub

    Private Sub frmResultNew_FormClosed(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
        Me.Dispose()
    End Sub

    Private Sub frmNewOrderAE_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        LoadSampleID()
        AutoLoadRoom()
        AutoLoadDoctor()
        AutoLoadTestName()


        Connect()
        rs.Connection = conn
        rs.CommandType = CommandType.Text
        rs.CommandText = "SELECT DISTINCT `his_code` FROM `default_specimen`"
        reader = rs.ExecuteReader
        While reader.Read
            Dim iItem As New ListViewItem(reader(0).ToString, 0)
            lvTest.Items.Add(iItem)
        End While
        Disconnect()
    End Sub

    Private Sub btnClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnClose.ItemClick
        Me.Close()
        Me.Dispose()
    End Sub

    Private Sub txtAge_KeyPress(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtAge.KeyPress
        If InStr("1234567890.", e.KeyChar) = 0 And Not Chr(AscW(e.KeyChar)) = vbBack Then
            e.KeyChar = ChrW(0)
        End If
    End Sub

    Private Sub dtBDate_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles dtBDate.EditValueChanged
        GetBDate()
    End Sub

    Public Sub GetBDate()
        Try
            If dtBDate.DateTime = Nothing Then
                Exit Sub
            End If

            Dim Birthday As Date = dtBDate.DateTime
            Dim endDate As Date = Date.Now

            Dim timeSpan As TimeSpan = endDate.Subtract(Birthday)
            Dim totalDays As Integer = timeSpan.TotalDays
            Dim totalYears As Integer = Math.Truncate(totalDays / 365)
            Dim totalMonths As Integer = Math.Truncate((totalDays Mod 365) / 30)
            Dim remainingDays As Integer = Math.Truncate((totalDays Mod 365) Mod 30)

            If totalDays <= 61 Then
                txtClass.Text = "Day(s)"
                txtAge.Text = totalDays

            ElseIf totalDays >= 62 And totalDays <= 364 Then
                txtClass.Text = "Month(s)"
                txtAge.Text = totalMonths

            ElseIf totalDays >= 365 Then
                txtClass.Text = "Year(s)"
                txtAge.Text = totalYears
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Exception", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
        End Try
    End Sub

    Private Function codeExists(ByVal text As String) As Boolean
        For Each lvi As ListViewItem In lv.Items
            If lvi.Text.Equals(text) Then Return True
        Next
        Return False
    End Function

    Private Sub SimpleButton1_Click(sender As Object, e As EventArgs) Handles SimpleButton1.Click
        frmPackages.ShowDialog()
    End Sub
End Class
