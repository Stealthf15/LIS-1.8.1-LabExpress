Imports System.Text.RegularExpressions

Public Class frmUrinAddTest
    Public TypeResult As String = ""
    Public mainID As String = ""
    Public patientID As String = ""
    Public Section As String = ""
    Public SubSection As String = ""

    Public Sub LoadRecords()
        Try
            lvList.Items.Clear()
            Connect()
            rs.Connection = conn
            rs.CommandText = "SELECT * FROM `default_specimen` WHERE `Section` = 'Urinalysis' AND `status` = 'Enable' ORDER BY `order_no`"
            reader = rs.ExecuteReader
            While reader.Read()
                If reader.HasRows Then
                    iItem = New ListViewItem(reader(0).ToString, 0)
                    For x As Integer = 1 To 12 Step 1
                        iItem.SubItems.Add(reader(x).ToString())
                    Next
                    lvList.Items.Add(iItem)
                End If
            End While
            Disconnect()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "System Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub frmAddTestSemi_FormClosed(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
        Me.Dispose()
    End Sub

    Private Sub btnClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnClose.ItemClick
        Me.Close()
    End Sub

    Private Sub frmResultsNew_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        LoadRecords()
    End Sub

    Private Sub txtSearch_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtSearch.TextChanged
        Try
            SearchRecordsOnLV(lvList, "default_specimen", 12, "specimen", txtSearch.Text, "specimen")
        Catch ex As Exception
            MessageBox.Show(ex.Message, "System Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub btnAddTest_ItemClick(ByVal sender As Object, ByVal e As DevExpress.XtraBars.ItemClickEventArgs) Handles btnAddTest.ItemClick
        If TypeResult = "New" Then
            For x = Me.lvList.CheckedItems.Count - 1 To 0 Step -1
                rs.Parameters.Clear()
                rs.Parameters.AddWithValue("@UniversalID", lvList.CheckedItems(x).SubItems(1).Text)
                rs.Parameters.AddWithValue("@Measurement", "")
                rs.Parameters.AddWithValue("@sample_id", mainID)
                rs.Parameters.AddWithValue("@patient_id", patientID)
                rs.Parameters.AddWithValue("@test_code", lvList.CheckedItems(x).SubItems(4).Text)
                rs.Parameters.AddWithValue("@his_testcode", lvList.CheckedItems(x).SubItems(12).Text)
                rs.Parameters.AddWithValue("@si_Unit", lvList.CheckedItems(x).SubItems(7).Text)
                rs.Parameters.AddWithValue("@conv_unit", lvList.CheckedItems(x).SubItems(8).Text)
                rs.Parameters.AddWithValue("@section", lvList.CheckedItems(x).SubItems(9).Text)
                rs.Parameters.AddWithValue("@order_no", lvList.CheckedItems(x).SubItems(10).Text)
                rs.Parameters.AddWithValue("@test_group", lvList.CheckedItems(x).SubItems(11).Text)
                rs.Parameters.AddWithValue("@date", Now)
                rs.Parameters.AddWithValue("@SubSection", SubSection)

                SaveRecordwthoutMSG("INSERT INTO `tmpresult` (`universal_id`, `measurement`,`test_code`, `sample_id`, `date`, `patient_id`, `order_no`, `test_group`, `units`, `unit_conv`, `instrument`, `his_test_code`, `section`, `sub_section`) VALUES " _
                            & "(" _
                            & "@UniversalID, @Measurement, @test_code, @sample_id, @date, @patient_id, @order_no, @test_group, @si_unit, @conv_unit, 'Other_Test', @his_testcode, @section, @SubSection" _
                            & ")"
                            )

                frmUrinNew.LoadTest()
            Next
            Me.Close()
        ElseIf TypeResult = "Old" Then
            For x = Me.lvList.CheckedItems.Count - 1 To 0 Step -1
                rs.Parameters.Clear()
                rs.Parameters.AddWithValue("@UniversalID", lvList.CheckedItems(x).SubItems(1).Text)
                rs.Parameters.AddWithValue("@Measurement", "")
                rs.Parameters.AddWithValue("@sample_id", mainID)
                rs.Parameters.AddWithValue("@patient_id", patientID)
                rs.Parameters.AddWithValue("@test_code", lvList.CheckedItems(x).SubItems(4).Text)
                rs.Parameters.AddWithValue("@his_testcode", lvList.CheckedItems(x).SubItems(12).Text)
                rs.Parameters.AddWithValue("@si_Unit", lvList.CheckedItems(x).SubItems(7).Text)
                rs.Parameters.AddWithValue("@conv_unit", lvList.CheckedItems(x).SubItems(8).Text)
                rs.Parameters.AddWithValue("@section", lvList.CheckedItems(x).SubItems(9).Text)
                rs.Parameters.AddWithValue("@order_no", lvList.CheckedItems(x).SubItems(10).Text)
                rs.Parameters.AddWithValue("@test_group", lvList.CheckedItems(x).SubItems(11).Text)
                rs.Parameters.AddWithValue("@date", Now)
                rs.Parameters.AddWithValue("@SubSection", SubSection)

                SaveRecordwthoutMSG("INSERT INTO `result` (`universal_id`, `measurement`,`test_code`, `sample_id`, `date`, `patient_id`, `order_no`, `test_group`, `units`, `unit_conv`, `instrument`, `his_test_code`, `section`, `sub_section`) VALUES " _
                            & "(" _
                            & "@UniversalID, @Measurement, @test_code, @sample_id, @date, @patient_id, @order_no, @test_group, @si_unit, @conv_unit, 'Other_Test', @his_testcode, @section, @SubSection" _
                            & ")"
                            )

                frmUrinOrdered.LoadTest()
            Next
            Me.Close()
        End If
    End Sub

    Private Sub cboLimit_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)
        LoadRecords()
    End Sub
End Class

'Dim C_Chol As Double
'Dim C_HDL As Double
'Dim Trigly As Double
'Dim LDL As Double

'If lvList.CheckedItems(x).SubItems(2).Text = "C_LDL" Then
'    Connect()
'    rs.Connection = conn
'    rs.CommandText = "SELECT `measurement` FROM `tmpresult` WHERE `sample_id` LIKE @sample_id AND `test_code` LIKE 'C_Chol'"
'    reader = rs.ExecuteReader
'    reader.Read()
'    If reader.HasRows Then
'        C_Chol = reader(0)
'    End If
'    Disconnect()

'    Connect()
'    rs.Connection = conn
'    rs.CommandText = "SELECT `measurement` FROM `tmpresult` WHERE `sample_id` LIKE @sample_id AND `test_code` LIKE 'C_HDL'"
'    reader = rs.ExecuteReader
'    reader.Read()
'    If reader.HasRows Then
'        C_HDL = reader(0)
'    End If
'    Disconnect()

'    Connect()
'    rs.Connection = conn
'    rs.CommandText = "SELECT `measurement` FROM `tmpresult` WHERE `sample_id` LIKE @sample_id AND `test_code` LIKE 'Trigly'"
'    reader = rs.ExecuteReader
'    reader.Read()
'    If reader.HasRows Then
'        Trigly = reader(0)
'    End If
'    Disconnect()
'    LDL = C_Chol - C_HDL - (Trigly / 2.2)

'    SaveRecordwthoutMSG("INSERT INTO `tmpResult` (`universal_id`, `measurement`, `test_code`, `sample_id`, `date`, `patient_id`, `order_no`, `test_group`, `units`, `unit_conv`, `instrument`) VALUES " _
'        & "(" _
'        & "@UniversalID, '" & LDL & "', @test_code, @sample_id, @date, @patient_id, @order_no, @test_group, @si_unit, @conv_unit, 'Other_Test'" _
'        & ")"
'        )
'Else
'    SaveRecordwthoutMSG("INSERT INTO `tmpResult` (`universal_id`, `measurement`, `test_code`, `sample_id`, `date`, `patient_id`, `order_no`, `test_group`, `units`, `unit_conv`, `instrument`) VALUES " _
'        & "(" _
'        & "@UniversalID, @Measurement, @test_code, @sample_id, @date, @patient_id, @order_no, @test_group, @si_unit, @conv_unit, 'Other_Test'" _
'        & ")"
'        )
'End If

'Dim C_Chol As Double
'Dim C_HDL As Double
'Dim Trigly As Double
'Dim LDL As Double

'If lvList.CheckedItems(x).SubItems(2).Text = "C_LDL" Then
'    Connect()
'    rs.Connection = conn
'    rs.CommandText = "SELECT `measurement` FROM `result` WHERE `sample_id` LIKE @sample_id AND `test_code` LIKE 'C_Chol'"
'    reader = rs.ExecuteReader
'    reader.Read()
'    If reader.HasRows Then
'        C_Chol = reader(0)
'    End If
'    Disconnect()

'    Connect()
'    rs.Connection = conn
'    rs.CommandText = "SELECT `measurement` FROM `result` WHERE `sample_id` LIKE @sample_id AND `test_code` LIKE 'C_HDL'"
'    reader = rs.ExecuteReader
'    reader.Read()
'    If reader.HasRows Then
'        C_HDL = reader(0)
'    End If
'    Disconnect()

'    Connect()
'    rs.Connection = conn
'    rs.CommandText = "SELECT `measurement` FROM `result` WHERE `sample_id` LIKE @sample_id AND `test_code` LIKE 'Trigly'"
'    reader = rs.ExecuteReader
'    reader.Read()
'    If reader.HasRows Then
'        Trigly = reader(0)
'    End If
'    Disconnect()
'    LDL = C_Chol - C_HDL - (Trigly / 2.2)

'    SaveRecordwthoutMSG("INSERT INTO `result` (`universal_id`, `measurement`, `test_code`, `sample_id`, `date`, `patient_id`, `order_no`, `test_group`, `units`, `unit_conv`, `instrument`) VALUES " _
'        & "(" _
'        & "@UniversalID, '" & LDL & "', @test_code, @sample_id, @date, @patient_id, @order_no, @test_group, @si_unit, @conv_unit, 'Other_Test'" _
'        & ")"
'        )
'Else
'    SaveRecordwthoutMSG("INSERT INTO `result` (`universal_id`, `measurement`, `test_code`, `sample_id`, `date`, `patient_id`, `order_no`, `test_group`, `units`, `unit_conv`, `instrument`) VALUES " _
'        & "(" _
'        & "@UniversalID, @Measurement, @test_code, @sample_id, @date, @patient_id, @order_no, @test_group, @si_unit, @conv_unit, 'Other_Test'" _
'        & ")"
'        )
'End If