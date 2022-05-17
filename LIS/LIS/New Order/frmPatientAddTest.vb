Public Class frmPatientAddTest
    Public _patientID As String = ""
    Public _sampleID As String = ""
    Public _test As String = "" 'Sub Section and Test

    Public Sub AutoLoadTestName()
        Connect()
        rs.Connection = conn
        rs.CommandText = "SELECT `testtype` FROM `tmpworklist` WHERE `sample_id` = '" & _sampleID & "' AND `patient_id` = '" & _patientID & "'"
        reader = rs.ExecuteReader
        While reader.Read
            cboTestName.Text = reader(0).ToString()
        End While
        Disconnect()
    End Sub
    Private Sub AutoLoadTest()
        Try
            rs.Parameters.Clear()
            rs.Parameters.AddWithValue("@TestName", cboTestName.Text)
            LoadRecordsOnLVSQL(lvTest, "SELECT `id`, `specimen`, `test_code`, `his_code` FROM `default_specimen` WHERE `section` = @TestName", 3)
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Warning", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            Exit Sub
        End Try
    End Sub
    Private Function AutoLoadMode(ByVal mode As String)
        Connect()
        rs.Connection = conn
        rs.CommandText = "SELECT `mode` FROM `patient_order` WHERE `sample_id` = '" & _sampleID & "' AND `patient_id` = '" & _patientID & "'"
        reader = rs.ExecuteReader
        If reader.Read Then
            mode = reader(0).ToString()
        End If
        Disconnect()
        Return mode
    End Function

    Private Sub btnClose_ItemClick(ByVal sender As System.Object, ByVal e As DevExpress.XtraBars.ItemClickEventArgs) Handles btnClose.ItemClick
        Me.Dispose()
    End Sub

    Private Sub frmPatientAddTest_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        AutoLoadTestName()
        AutoLoadTest()
        LoadPatientOrder()
    End Sub

    'Private Sub lvTest_ItemChecked(ByVal sender As Object, ByVal e As System.Windows.Forms.ItemCheckedEventArgs) Handles lvTest.ItemChecked
    '    Dim Test As String

    '    txtTest.Text = ""

    '    Disconnect1()
    '    Connect1()
    '    rs1.Connection = conn1
    '    rs1.Parameters.Clear()
    '    rs1.Parameters.AddWithValue("@sample_id", _sampleID)
    '    rs1.CommandText = "SELECT `test` FROM tmpworklist WHERE `sample_id` LIKE @sample_id"
    '    reader1 = rs1.ExecuteReader
    '    If reader1.Read() Then
    '        'testdata = reader1(0).ToString
    '        For x As Integer = 0 To lvTest.CheckedItems.Count - 1 Step 1
    '            txtTest.Text = ""
    '            'txtTest.Text = testdata
    '            Test &= lvTest.CheckedItems(x).SubItems(2).Text & ", "
    '            txtTest.Text &= Test.TrimEnd(", ")
    '        Next
    '    End If
    '    Disconnect1()
    'End Sub


    Private Sub SaveOrders()

        rs.Parameters.Clear()
        rs.Parameters.AddWithValue("@sample_id", _sampleID)
        rs.Parameters.AddWithValue("@PATIENT_ID", _patientID)
        rs.Parameters.AddWithValue("@status", "Checked-In")
        rs.Parameters.AddWithValue("@TestType", cboTestName.Text)
        'rs.Parameters.AddWithValue("@Test", txtTest.Text)
        rs.Parameters.AddWithValue("@Test", _test)

        rs.Parameters.AddWithValue("@MainID", SampleID.TrimEnd)
        Connect()
        rs.Connection = conn
        rs.CommandText = "SELECT * FROM `tmpWorklist` WHERE `sample_id` LIKE @sample_id"
        reader = rs.ExecuteReader
        reader.Read()
        If reader.HasRows Then
            Disconnect()
            UpdateRecordwthoutMSG("UPDATE `tmpWorklist` SET " _
                                & "`test` = @Test" _
                                & " WHERE main_id = @sample_id")
        End If


        'frmOrders.LoadRecords()
    End Sub

    Private Sub LoadPatientOrder()
        'Dim test As String
        'Dim test_Code As String
        'Dim testtype As String
        'Dim his_Code As String
        'Dim specimen As String


        'Connect2()
        'rs2.Parameters.Clear()
        'rs2.Connection = conn2
        'rs2.CommandText = "SELECT `test_name`, `test_code`, `testtype`, `his_code` FROM `patient_order` WHERE `sample_id` = @MainID"
        'rs2.Parameters.AddWithValue("@MainID", _sampleID)
        'reader2 = rs2.ExecuteReader
        'While reader2.Read
        '    test = reader2(0)
        '    test_Code = reader2(1)
        '    testtype = reader2(2) 'To Get the Testtype of Sample
        '    his_Code = reader2(3)


        '    AutoLoadTest() 'Load the Check Box List based on Test type initialised

        '    Connect1()
        '    rs1.Connection = conn1
        '    rs1.Parameters.Clear()
        '    rs1.Parameters.AddWithValue("@test", test)
        '    rs1.Parameters.AddWithValue("@test_code", test_Code)
        '    rs1.Parameters.AddWithValue("@section", testtype)
        '    rs1.Parameters.AddWithValue("@his_code", his_Code)
        '    rs1.CommandText = "SELECT `specimen`, `test_code`, `his_code`, `section`, `test_group` FROM `default_specimen` WHERE `test_code` = @test_code AND `his_code` = @his_code AND `specimen` = @test"
        '    reader1 = rs1.ExecuteReader
        '    If reader1.Read Then
        '        specimen = reader1(1).ToString
        '        lvTest.FindItemWithText(reader1(1).ToString).Checked = True
        '    End If
        '    Disconnect1()
        'End While
        'Disconnect1()
        'Disconnect2()
    End Sub

    Private Sub btnAddTest_ItemClick(ByVal sender As System.Object, ByVal e As DevExpress.XtraBars.ItemClickEventArgs) Handles btnAddTest.ItemClick
        Dim OrderTestCode As String
        Dim mode As String = ""

        mode = AutoLoadMode(mode)
        SaveOrders()
        If Me.lvTest.CheckedItems.Count > 0 Then
            rs.Parameters.Clear()
            rs.Parameters.AddWithValue("@MainID", _sampleID)
            rs.Parameters.AddWithValue("@patient_id", _patientID)
            rs.Parameters.AddWithValue("@Section", cboTestName.Text)
            rs.Parameters.AddWithValue("@test", _test)
            rs.Parameters.AddWithValue("@date", Now)

            For x As Integer = lvTest.CheckedItems.Count - 1 To 0 Step -1
                Connect()
                rs.Connection = conn
                rs.CommandText = "SELECT * FROM `patient_order` WHERE `test_name` = '" & lvTest.CheckedItems(x).SubItems(1).Text & "' AND `sample_id` = @MainID"
                reader = rs.ExecuteReader
                reader.Read()
                If reader.HasRows Then
                    Disconnect()
                Else
                    Disconnect()

                    Connect()
                    rs.Connection = conn
                    rs.CommandText = "SELECT `specimen`, `test_code`, `order_no`, `test_group`, `si_unit`, `conventional_unit`, `test_code` FROM `specimen` where `test_code` = '" & lvTest.CheckedItems(x).SubItems(2).Text.Split("^").GetValue(0).ToString & "'"
                    reader = rs.ExecuteReader
                    reader.Read()
                    If reader.HasRows Then
                        test_name = reader(0).ToString
                        test_code = reader(1).ToString
                        order_no = reader(2).ToString
                        test_group = reader(3).ToString
                        unit = reader(4).ToString
                        unit_conv = reader(5).ToString
                        OrderTestCode = reader(6).ToString
                    End If
                    Disconnect()

                    SaveRecordwthoutMSG("INSERT INTO `patient_order` (`test_name`, `test_code`, `patient_id`, `sample_id`, `testtype`, `mode`, `sub_section`, `his_code`) VALUES (" _
                        & "'" & lvTest.CheckedItems(x).SubItems(1).Text() & "'," _
                        & "'" & OrderTestCode & "'," _
                        & "'" & _patientID & "'," _
                        & "@MainID," _
                        & "'" & cboTestName.Text & "'," _
                        & "'" & mode & "'," _
                        & "@test," _
                        & "'" & lvTest.CheckedItems(x).SubItems(3).Text() & "')")

                    SaveRecordwthoutMSG("INSERT INTO `tmpResult` (`universal_id`, `measurement`, `test_code`, `sample_id`, `date`, `patient_id`, `order_no`, `units`, `unit_conv`, `instrument`, `status`, `test_group`, `section`, `sub_section`) VALUES " _
                        & "(" _
                        & "'" & test_name & "', '', '" & test_code & "', @mainID, NOW(), @patient_id, '" & order_no & "', '" & unit & "', '" & unit_conv & "', 'Other_Test', 0, '" & test_group & "', @Section, '" & _test & "'" _
                        & ")"
                        )

                    SaveRecordwthoutMSG("INSERT INTO `lis_order` (`sample_id`, `order_code`, `mode`, `specimen_type`, `section`, `sub_section`) VALUES (" _
                        & "'" & _sampleID & "'," _
                        & "'" & test_code & "'," _
                        & "'" & mode & "'," _
                        & "1," _
                        & "@Section," _
                        & "'" & _test & "')")
                End If
                Disconnect()
                lvTest.CheckedItems(x).Checked = False
            Next
        End If

        MessageBox.Show("Patient Order Successfully Added", "Add Order Successful", MessageBoxButtons.OK, MessageBoxIcon.Information)
        Me.Dispose()
        frmPatientOrder.LoadRecords()
        'frmNewOrder.LoadRecords()
        'frmNewOrder.LoadRecordsCompleted()
    End Sub

    Dim blnAscending As Boolean = True
    Private Sub lvList_ColumnClick(ByVal sender As Object, ByVal e As System.Windows.Forms.ColumnClickEventArgs) Handles lvTest.ColumnClick
        lvTest.SelectedItems.Clear()

        ' Set the ListViewItemSorter property to a new ListViewItemComparer 
        ' object. Setting this property immediately sorts the 
        ' ListView using the ListViewItemComparer object. 

        'Change this based on which direction you want

        If blnAscending Then
            Me.lvTest.ListViewItemSorter = New ListViewItemComparer(e.Column, True)
            blnAscending = False
        Else
            Me.lvTest.ListViewItemSorter = New ListViewItemComparer(e.Column, False)
            blnAscending = True
        End If
    End Sub

    Private Sub txtSearch_TextChanged(sender As Object, e As EventArgs) Handles txtSearch.TextChanged
        Try
            lvTest.ListViewItemSorter = New MyComparer
            lvTest.Sorting = SortOrder.Descending
            lvTest.Sort()
            lvTest.Items.Clear()
            rs.Parameters.Clear()
            rs.Parameters.AddWithValue("@TestName", cboTestName.Text)
            LoadRecordsOnLVSQL(lvTest, "SELECT `id`, `specimen`, `test_code`, `his_code` FROM `default_specimen` WHERE 
                                        `specimen` LIKE '" & txtSearch.Text & "%' OR `test_code` LIKE '" & txtSearch.Text & "%' 
                                        OR `his_code` LIKE '" & txtSearch.Text & "%' AND `section` LIKE @TestName", 3)
        Catch ex As Exception
            MessageBox.Show(ex.Message, "System Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub
End Class