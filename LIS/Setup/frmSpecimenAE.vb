Public Class frmSpecimenAE

    Public ID As String

    Private Sub btnBack_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnClose.Click
        Me.Close()
    End Sub

    Private Sub btnSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSave.Click
        If Me.cboTest.Text = "" Or Me.cboTestName.Text = "" Then
            MessageBox.Show("Please fill all required information.", "Empty String", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            Exit Sub
        End If

        rs.Parameters.Clear()
        rs.Parameters.AddWithValue("@name", ID)
        rs.Parameters.AddWithValue("@testname", cboTest.Text)
        rs.Parameters.AddWithValue("@specimen_type", cboSpecimen.Text)
        rs.Parameters.AddWithValue("@test_code", txtTestCode.Text)
        rs.Parameters.AddWithValue("@HIS_code", txtHISCode.Text)
        rs.Parameters.AddWithValue("@HIS_Field", txtHISField.Text)
        rs.Parameters.AddWithValue("@si_unit", cboUnit.Text)
        rs.Parameters.AddWithValue("@conventional_unit", cboConv_Unit.Text)
        rs.Parameters.AddWithValue("@test_name", cboTestName.Text)
        rs.Parameters.AddWithValue("@order_no", txtOrderNo.Text)
        rs.Parameters.AddWithValue("@test_group", cboTestGroup.Text)
        rs.Parameters.AddWithValue("@status", "Enable")

        If Me.btnSave.Text = "&Save" Then
            SaveRecord("INSERT INTO `default_specimen` (`specimen`, `specimen_type`, `test_code`, `his_code`, `his_field`, `si_unit`, `conventional_unit`, `section`, `test_name`, `order_no`, `test_group`, `status`) VALUES (@testname, @specimen_type, @test_code, @HIS_code, @HIS_Field, @si_unit, @conventional_unit, @test_name, @HIS_Code, @order_no, @test_group, @status)")
            frmSpecimen.LoadRecordsTest(txtHISCode.Text)
            Me.Close()
        Else
            UpdateRecord("UPDATE `default_specimen` SET `specimen` = @testname, `specimen_type` = @specimen_type, `test_code` = @test_code, `his_code` = @HIS_code, `his_field` = @HIS_Field, `si_unit` = @si_unit, `conventional_unit` = @conventional_unit, `test_name` = @test_name, `order_no` = @order_no, `test_group` = @test_group WHERE `id` = @name")
            frmSpecimen.LoadRecordsTest(txtHISCode.Text)
            Me.Close()
        End If
    End Sub

    Public Sub AutoLoadCombo()
        cboTestName.Properties.Items.Clear()
        Connect()
        rs.Connection = conn
        rs.CommandText = "SELECT `test_name` FROM `testtype` WHERE `test_name` NOT LIKE 'All' ORDER BY `TEST_name`"
        reader = rs.ExecuteReader
        While reader.Read
            cboTestName.Properties.Items.Add(reader(0).ToString)
        End While
        Disconnect()
    End Sub

    Public Sub AutoLoadComboSpecimen()
        cboSpecimen.Properties.Items.Clear()
        Connect()
        rs.Connection = conn
        rs.CommandText = "SELECT * FROM `specimen_type` ORDER BY `name`"
        reader = rs.ExecuteReader
        While reader.Read
            cboSpecimen.Properties.Items.Add(reader(1).ToString)
        End While
        Disconnect()
    End Sub

    Public Sub AutoLoadComboTest()
        rs.Parameters.Clear()
        rs.Parameters.AddWithValue("@section", cboTestName.Text)

        cboTest.Properties.Items.Clear()
        Connect()
        rs.Connection = conn
        rs.CommandText = "SELECT `specimen` FROM `default_specimen` WHERE `test_name` = @section ORDER BY `specimen`"
        reader = rs.ExecuteReader
        While reader.Read
            cboTest.Properties.Items.Add(reader(0).ToString)
        End While
        Disconnect()
    End Sub

    Private Sub txtChannel_KeyPress(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtOrderNo.KeyPress
        If InStr("1234567890.", e.KeyChar) = 0 And Not Chr(AscW(e.KeyChar)) = vbBack Then
            e.KeyChar = ChrW(0)
        End If
    End Sub

    Private Sub frmTestTypeAE_FormClosed(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
        Me.Dispose()
    End Sub

    Private Sub frmTestTypeAE_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        AutoLoadComboTest()
    End Sub

    Private Sub cboTest_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cboTest.SelectedIndexChanged
        Connect()
        rs.Connection = conn
        rs.CommandText = "SELECT `test_code`, `si_unit`, `conventional_unit`, `order_no`, `test_group` FROM `default_specimen` WHERE `specimen` LIKE '" & Me.cboTest.Text & "'"
        reader = rs.ExecuteReader
        reader.Read()
        If reader.HasRows Then
            txtTestCode.Text = reader(0).ToString
            cboUnit.Text = reader(1).ToString
            cboConv_Unit.Text = reader(2).ToString
            txtOrderNo.Text = reader(3).ToString
            cboTestGroup.Text = reader(4).ToString
        End If
        Disconnect()
    End Sub

    Private Sub cboTestName_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cboTestName.SelectedIndexChanged
        AutoLoadComboTest()
    End Sub
End Class