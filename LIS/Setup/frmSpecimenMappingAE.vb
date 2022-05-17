Imports DevExpress.XtraEditors

Public Class frmSpecimenMappingAE

    Public name As String

    Private Sub btnBack_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnClose.Click
        Me.Close()
    End Sub

    Private Sub btnSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSave.Click
        If Me.cboTest.Text = "" Or Me.cboTestName.Text = "" Or txtChannel.Text = "" Then
            MessageBox.Show("Please fill all required information.", "Empty String", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            Exit Sub
        End If

        rs.Parameters.Clear()
        rs.Parameters.AddWithValue("@name", name)
        rs.Parameters.AddWithValue("@testname", cboTest.Text)
        rs.Parameters.AddWithValue("@description", txtDescription.Text)
        rs.Parameters.AddWithValue("@specimen_type", cboSpecimen.Text)
        rs.Parameters.AddWithValue("@channel", txtChannel.Text)
        rs.Parameters.AddWithValue("@test_code", txtTestCode.Text)
        'rs.Parameters.AddWithValue("@his_code", txtHISCode.Text)
        'rs.Parameters.AddWithValue("@his_field", txtHISField.Text)
        rs.Parameters.AddWithValue("@si_unit", cboUnit.Text)
        rs.Parameters.AddWithValue("@conventional_unit", cboConv_Unit.Text)
        rs.Parameters.AddWithValue("@convertion_factor", cboConvetionFactor.Text)
        rs.Parameters.AddWithValue("@convertion_multiplier", txtMultiplier.Text)
        rs.Parameters.AddWithValue("@test_name", cboTestName.Text)
        rs.Parameters.AddWithValue("@order_no", txtOrderNo.Text)
        rs.Parameters.AddWithValue("@test_group", cboTestGroup.Text)
        rs.Parameters.AddWithValue("@instrument", cboInstrument.Text)
        rs.Parameters.AddWithValue("@status", "Enable")

        If Me.btnSave.Text = "&Save" Then
            SaveRecord("INSERT INTO `specimen` (`specimen`, `description`, `specimen_type`, `channel`, `test_code`, `si_unit`, `conventional_unit`, `convertion_factor`, `convertion_multiplier`, `section`, `order_no`, `test_group`, `status`, `instrument`) VALUES (@testname, @description, @specimen_type, @channel, @test_code, @si_unit, @conventional_unit, @convertion_factor, @convertion_multiplier, @test_name, @order_no, @test_group, @status, @instrument)")
            Connect()
            rs.Connection = conn
            rs.CommandText = "SELECT `test_code` FROM `conversion_factor` WHERE `test_code` = @test_code AND `instrument` = @instrument"
            reader = rs.ExecuteReader
            reader.Read()
            If reader.HasRows Then
                Disconnect()
                UpdateRecordwthoutMSG("UPDATE `conversion_factor` SET `test_code` = @test_code, `convertion_factor` = @convertion_factor, `convertion_multiplier` = @convertion_multiplier WHERE `test_code` = @test_code AND `instrument` = @instrument")
            Else
                Disconnect()
                SaveRecordwthoutMSG("INSERT INTO `conversion_factor` (`test_code`, `convertion_factor`, `convertion_multiplier`, `instrument`) VALUES (@test_code, @convertion_factor, @convertion_multiplier, @instrument)")
            End If
            Disconnect()

            frmSpecimenMapping.LoadRecords()
            Me.Close()
        Else
            UpdateRecordwthoutMSG("UPDATE `specimen` SET `specimen` = @testname, `description` = @description, `specimen_type` = @specimen_type, `channel` = @channel, `test_code` = @test_code, `si_unit` = @si_unit, `conventional_unit` = @conventional_unit, `convertion_factor` = @convertion_factor, `convertion_multiplier` = @convertion_multiplier, `section` = @test_name, `order_no` = @order_no, `test_group` = @test_group, `instrument` = @instrument WHERE `id` = @name")
            UpdateRecord("UPDATE `conversion_factor` SET `test_code` = @test_code, `convertion_factor` = @convertion_factor, `convertion_multiplier` = @convertion_multiplier WHERE `test_code` = @test_code AND `instrument` = @instrument")
            frmSpecimenMapping.LoadRecords()
            Me.Close()
        End If
    End Sub

    Private Sub btnNew_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnNew.Click
        If Me.cboTest.Text = "" Or Me.cboTestName.Text = "" Or txtChannel.Text = "" Then
            MessageBox.Show("Please fill all required information.", "Empty String", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            Exit Sub
        End If

        rs.Parameters.Clear()
        rs.Parameters.AddWithValue("@name", name)
        rs.Parameters.AddWithValue("@testname", cboTest.Text)
        rs.Parameters.AddWithValue("@description", txtDescription.Text)
        rs.Parameters.AddWithValue("@specimen_type", cboSpecimen.Text)
        rs.Parameters.AddWithValue("@channel", txtChannel.Text)
        rs.Parameters.AddWithValue("@test_code", txtTestCode.Text)
        'rs.Parameters.AddWithValue("@his_code", txtHISCode.Text)
        'rs.Parameters.AddWithValue("@his_field", txtHISField.Text)
        rs.Parameters.AddWithValue("@si_unit", cboUnit.Text)
        rs.Parameters.AddWithValue("@conventional_unit", cboConv_Unit.Text)
        rs.Parameters.AddWithValue("@convertion_factor", cboConvetionFactor.Text)
        rs.Parameters.AddWithValue("@convertion_multiplier", txtMultiplier.Text)
        rs.Parameters.AddWithValue("@test_name", cboTestName.Text)
        rs.Parameters.AddWithValue("@order_no", txtOrderNo.Text)
        rs.Parameters.AddWithValue("@test_group", cboTestGroup.Text)
        rs.Parameters.AddWithValue("@instrument", cboInstrument.Text)
        rs.Parameters.AddWithValue("@status", "Enable")

        If Me.btnSave.Text = "&Save" Then
            SaveRecord("INSERT INTO `specimen` (`specimen`, `description`, `specimen_type`, `channel`, `test_code`, `si_unit`, `conventional_unit`, `convertion_factor`, `convertion_multiplier`, `section`, `order_no`, `test_group`, `status`, `instrument`) VALUES (@testname, @description, @specimen_type, @channel, @test_code, @si_unit, @conventional_unit, @convertion_factor, @convertion_multiplier, @test_name, @order_no, @test_group, @status, @instrument)")
            Connect()
            rs.Connection = conn
            rs.CommandText = "SELECT `test_code` FROM `conversion_factor` WHERE `test_code` = @test_code AND `instrument` = @instrument"
            reader = rs.ExecuteReader
            reader.Read()
            If reader.HasRows Then
                Disconnect()
                UpdateRecordwthoutMSG("UPDATE `conversion_factor` SET `test_code` = @test_code, `convertion_factor` = @convertion_factor, `convertion_multiplier` = @convertion_multiplier WHERE `test_code` = @test_code AND `instrument` = @instrument")
            Else
                Disconnect()
                SaveRecordwthoutMSG("INSERT INTO `conversion_factor` (`test_code`, `convertion_factor`, `convertion_multiplier`, `instrument`) VALUES (@test_code, @convertion_factor, @convertion_multiplier, @instrument)")
            End If
            Disconnect()

            frmSpecimenMapping.LoadRecords()
            Me.Close()
        Else
            UpdateRecordwthoutMSG("UPDATE `specimen` SET `specimen` = @testname, `description` = @description, `specimen_type` = @specimen_type, `channel` = @channel, `test_code` = @test_code, `si_unit` = @si_unit, `conventional_unit` = @conventional_unit, `convertion_factor` = @convertion_factor, `convertion_multiplier` = @convertion_multiplier, `section` = @test_name, `order_no` = @order_no, `test_group` = @test_group, `instrument` = @instrument WHERE `id` = @name")
            UpdateRecord("UPDATE `conversion_factor` SET `test_code` = @test_code, `convertion_factor` = @convertion_factor, `convertion_multiplier` = @convertion_multiplier WHERE `test_code` = @test_code AND `instrument` = @instrument")
            frmSpecimenMapping.LoadRecords()
            Me.Close()
        End If

        For Each ctrl As Control In Me.Controls
            If TypeOf ctrl Is ComboBoxEdit Or TypeOf ctrl Is TextEdit Then
                If ctrl.Tag = "Section" Or ctrl.Tag = "Instrument" Then

                Else
                    ctrl.Text = Nothing
                End If
            End If
        Next

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

    Public Sub AutoLoadComboInstrument()
        cboInstrument.Properties.Items.Clear()
        Connect()
        rs.Connection = conn
        rs.CommandText = "SELECT * FROM `machines` WHERE `test_name` = '" & cboTestName.Text & "' ORDER BY `name` "
        reader = rs.ExecuteReader
        While reader.Read
            cboInstrument.Properties.Items.Add(reader(1).ToString & "_" & reader(3).ToString)
        End While
        Disconnect()
    End Sub

    Private Sub txtChannel_KeyPress(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtOrderNo.KeyPress, cboConvetionFactor.KeyPress, txtMultiplier.KeyPress
        If InStr("1234567890.", e.KeyChar) = 0 And Not Chr(AscW(e.KeyChar)) = vbBack Then
            e.KeyChar = ChrW(0)
        End If
    End Sub

    Private Sub frmTestTypeAE_FormClosed(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
        Me.Dispose()
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
        AutoLoadComboInstrument()
    End Sub
End Class