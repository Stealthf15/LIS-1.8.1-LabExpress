Public Class frmSpecimenAE

    Public name As String

    Private Sub btnBack_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnClose.Click
        Me.Close()
    End Sub

    Private Sub btnSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSave.Click
        If Me.cboTest.Text = "" Or Me.cboTestName.Text = "" Then
            MessageBox.Show("Please fill all required information.", "Empty String", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            Exit Sub
        End If

        rs.Parameters.Clear()
        rs.Parameters.AddWithValue("@name", name)
        rs.Parameters.AddWithValue("@testname", cboTest.Text)
        rs.Parameters.AddWithValue("@test_code", txtTestCode.Text)
        rs.Parameters.AddWithValue("@si_unit", cboUnit.Text)
        rs.Parameters.AddWithValue("@conventional_unit", cboConv_Unit.Text)
        rs.Parameters.AddWithValue("@test_name", cboTestName.Text)
        rs.Parameters.AddWithValue("@order_no", txtorder.Text)
        rs.Parameters.AddWithValue("@test_group", cboTestGroup.Text)

        If Me.btnSave.Text = "&Save" Then
            SaveRecord("INSERT INTO `default_specimen` (`specimen`, `test_code`, `si_unit`, `conventional_unit`, `section`, `order_no`, `test_group`) VALUES (@testname, @test_code, @si_unit, @conventional_unit, @test_name, @order_no, @test_group)")

            frmSpecimen.LoadRecords()
            Me.Close()
        Else
            UpdateRecord("UPDATE `default_specimen` SET `specimen` = @testname, `test_code` = @test_code, `si_unit` = @si_unit, `conventional_unit` = @conventional_unit, `section` = @test_name, `order_no` = @order_no, `test_group` = @test_group WHERE `id` = @name")
            UpdateRecordwthoutMSG("UPDATE `specimen` SET `specimen` = @testname, `test_code` = @test_code, `si_unit` = @si_unit, `conventional_unit` = @conventional_unit, `test_name` = @test_name, `order_no` = @order_no, `test_group` = @test_group WHERE `test_code` = @test_code")
            frmSpecimen.LoadRecords()
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

    Public Sub AutoLoadComboUnit()
        cboUnit.Properties.Items.Clear()
        Connect()
        rs.Connection = conn
        rs.CommandText = "SELECT * FROM `units` ORDER BY `Unit`"
        reader = rs.ExecuteReader
        While reader.Read
            cboUnit.Properties.Items.Add(reader(1).ToString)
            cboConv_Unit.Properties.Items.Add(reader(1).ToString)
        End While
        Disconnect()
    End Sub

    Private Sub frmTestTypeAE_FormClosed(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
        Me.Dispose()
    End Sub

End Class