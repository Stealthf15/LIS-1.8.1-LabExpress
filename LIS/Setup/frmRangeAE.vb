Imports DevExpress.XtraEditors

Public Class frmRangeAE

    Public name, id As String

    Private Sub btnBack_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnClose.Click
        Me.Close()
    End Sub

    Private Sub btnSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSave.Click
        rs.Parameters.Clear()
        rs.Parameters.AddWithValue("@id", id)
        rs.Parameters.AddWithValue("@section", cboSection.Text)
        rs.Parameters.AddWithValue("@test_name", cboTestName.Text)
        rs.Parameters.AddWithValue("@conv_factor", txtConvFactor.Text)
        rs.Parameters.AddWithValue("@test_code", txtTestCode.Text)
        rs.Parameters.AddWithValue("@si_range", txtSIRange.Text)
        rs.Parameters.AddWithValue("@conv_range", txtConvRange.Text)
        rs.Parameters.AddWithValue("@low_value", txtLow.Text)
        rs.Parameters.AddWithValue("@high_value", txtHigh.Text)
        rs.Parameters.AddWithValue("@age_begin", txtAgeBegin.Text)
        rs.Parameters.AddWithValue("@age_end", txtAgeEnd.Text)
        rs.Parameters.AddWithValue("@channel", txtChannel.Text)
        rs.Parameters.AddWithValue("@gender", cboGender.Text)
        rs.Parameters.AddWithValue("@classification", cboClassifocation.Text)
        rs.Parameters.AddWithValue("@machine", cboMachine.Text)
        If btnSave.Text = "&Save" Then
            If cboGender.Text = "Both" Then
                SaveRecordwthoutMSG("INSERT INTO `reference_range` (`test_name`, `si_range`, `conv_range`, `low_value`, `high_value`, `gender`, `classification`, `age_begin`, `age_end`, `test_code`, `machine`, `channel`, `section`) VALUES " _
                    & "(" _
                    & "@test_name," _
                    & "@si_range," _
                    & "@conv_range," _
                    & "@low_value," _
                    & "@high_value," _
                    & "'Male'," _
                    & "@classification," _
                    & "@age_begin," _
                    & "@age_end," _
                    & "@test_code," _
                    & "@machine," _
                    & "@channel," _
                    & "@section" _
                    & ")"
                    )

                SaveRecord("INSERT INTO `reference_range` (`test_name`, `si_range`, `conv_range`, `low_value`, `high_value`, `gender`, `classification`, `age_begin`, `age_end`, `test_code`, `machine`, `channel`, `section`) VALUES " _
                    & "(" _
                    & "@test_name," _
                    & "@si_range," _
                    & "@conv_range," _
                    & "@low_value," _
                    & "@high_value," _
                    & "'Female'," _
                    & "@classification," _
                    & "@age_begin," _
                    & "@age_end," _
                    & "@test_code," _
                    & "@machine," _
                    & "@channel," _
                    & "@section" _
                    & ")"
                    )
            Else
                SaveRecord("INSERT INTO `reference_range` (`test_name`, `si_range`, `conv_range`, `low_value`, `high_value`, `gender`, `classification`, `age_begin`, `age_end`, `test_code`, `machine`, `channel`, `section`) VALUES " _
                    & "(" _
                    & "@test_name," _
                    & "@si_range," _
                    & "@conv_range," _
                    & "@low_value," _
                    & "@high_value," _
                    & "@gender," _
                    & "@classification," _
                    & "@age_begin," _
                    & "@age_end," _
                    & "@test_code," _
                    & "@machine," _
                    & "@channel," _
                    & "@section" _
                    & ")"
                    )
            End If
        Else
            SaveRecord("UPDATE `reference_range` SET " _
                    & "`test_name` = @test_name," _
                    & "`si_range` = @si_range," _
                    & "`conv_range` = @conv_range," _
                    & "`low_value` = @low_value," _
                    & "`high_value` = @high_value," _
                    & "`gender` = @gender," _
                    & "`classification` = @classification," _
                    & "`age_begin` = @age_begin," _
                    & "`age_end` = @age_end," _
                    & "`test_code` = @test_code," _
                    & "`machine` = @machine," _
                    & "`channel` = @channel," _
                    & "`section` = @section" _
                    & " WHERE `id` = @id"
                    )
        End If

        frmRangeList.LoadRecords()

        Me.Close()
        Me.Dispose()
    End Sub

    Private Sub txtChannel_KeyPress(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyPressEventArgs)
        If InStr("1234567890.-", e.KeyChar) = 0 And Not Chr(AscW(e.KeyChar)) = vbBack Then
            e.KeyChar = ChrW(0)
        End If
    End Sub

    Private Sub frmRangeA_FormClosed(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
        Me.Dispose()
    End Sub

    Private Sub frmRangeA_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        'AutoLoadSection()
        'AutoLoadMachine()
    End Sub

    Public Sub AutoLoadSection()
        cboSection.Properties.Items.Clear()
        Connect()
        rs.Connection = conn
        rs.CommandText = "SELECT `test_name` FROM `testtype` WHERE `test_name` NOT LIKE 'All' ORDER BY `test_name`"
        reader = rs.ExecuteReader
        While reader.Read
            cboSection.Properties.Items.Add(reader(0).ToString)
        End While
        Disconnect()
    End Sub

    Public Sub AutoLoadMachine()
        cboMachine.Properties.Items.Clear()
        rs.Parameters.Clear()
        rs.Parameters.AddWithValue("@section", cboSection.Text)

        Connect()
        rs.Connection = conn
        rs.CommandText = "SELECT `name`, `serial_no` FROM `machines` WHERE `test_name` = @section ORDER BY `name`"
        reader = rs.ExecuteReader
        While reader.Read
            cboMachine.Properties.Items.Add(reader(0).ToString & "_" & reader(1).ToString)
        End While
        Disconnect()
    End Sub

    Private Sub cboSection_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cboSection.SelectedIndexChanged, cboSection.TextChanged
        AutoLoadMachine()
    End Sub

    Private Sub AutoLoadTest()
        cboTestName.Properties.Items.Clear()
        rs.Parameters.Clear()
        rs.Parameters.AddWithValue("@section", cboSection.Text)
        rs.Parameters.AddWithValue("@instrument", cboMachine.Text)

        Connect()
        rs.Connection = conn
        rs.CommandText = "SELECT `specimen` FROM `specimen` WHERE `instrument` = @instrument AND `section` = @section ORDER BY `specimen`"
        reader = rs.ExecuteReader
        While reader.Read
            cboTestName.Properties.Items.Add(reader(0).ToString)
        End While
        Disconnect()
    End Sub

    Private Sub cboMachine_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cboMachine.SelectedIndexChanged, cboMachine.TextChanged
        AutoLoadTest()
    End Sub

    Private Sub cboTestName_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cboTestName.SelectedIndexChanged, cboTestName.TextChanged
        rs.Parameters.Clear()
        rs.Parameters.AddWithValue("@test_name", cboTestName.Text)
        rs.Parameters.AddWithValue("@instrument", cboMachine.Text)

        Connect()
        rs.Connection = conn
        rs.CommandText = "SELECT `channel`, `test_code`, `convertion_factor`, `convertion_multiplier` FROM `specimen` WHERE (`specimen` = @test_name AND `instrument` = @instrument)"
        reader = rs.ExecuteReader
        reader.Read()
        If reader.HasRows Then
            txtChannel.Text = reader(0).ToString
            txtTestCode.Text = reader(1).ToString
            txtConvFactor.Text = reader(2).ToString
            txtConvertionMultiplier.Text = reader(3).ToString
        End If
        Disconnect()
    End Sub

    Private Sub txtLow_EditValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtLow.EditValueChanged, txtHigh.EditValueChanged, txtConvFactor.TextChanged, txtConvertionMultiplier.TextChanged
        txtConvRange.Text = (FormatNumber(Val(Me.txtLow.Text) / Val(Me.txtConvFactor.Text), Val(txtConvertionMultiplier.Text)) & "-" & FormatNumber(Val(Me.txtHigh.Text) / Val(Me.txtConvFactor.Text), Val(txtConvertionMultiplier.Text))).ToString
    End Sub

    Private Sub SimpleButton1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles SimpleButton1.Click
        rs.Parameters.Clear()
        rs.Parameters.AddWithValue("@id", id)
        rs.Parameters.AddWithValue("@section", cboSection.Text)
        rs.Parameters.AddWithValue("@test_name", cboTestName.Text)
        rs.Parameters.AddWithValue("@conv_factor", txtConvFactor.Text)
        rs.Parameters.AddWithValue("@test_code", txtTestCode.Text)
        rs.Parameters.AddWithValue("@si_range", txtSIRange.Text)
        rs.Parameters.AddWithValue("@conv_range", txtConvRange.Text)
        rs.Parameters.AddWithValue("@low_value", txtLow.Text)
        rs.Parameters.AddWithValue("@high_value", txtHigh.Text)
        rs.Parameters.AddWithValue("@age_begin", txtAgeBegin.Text)
        rs.Parameters.AddWithValue("@age_end", txtAgeEnd.Text)
        rs.Parameters.AddWithValue("@channel", txtChannel.Text)
        rs.Parameters.AddWithValue("@gender", cboGender.Text)
        rs.Parameters.AddWithValue("@classification", cboClassifocation.Text)
        rs.Parameters.AddWithValue("@machine", cboMachine.Text)
        If btnSave.Text = "&Save" Then
            If cboGender.Text = "Both" Then
                SaveRecordwthoutMSG("INSERT INTO `reference_range` (`test_name`, `si_range`, `conv_range`, `low_value`, `high_value`, `gender`, `classification`, `age_begin`, `age_end`, `test_code`, `machine`, `channel`, `section`) VALUES " _
                    & "(" _
                    & "@test_name," _
                    & "@si_range," _
                    & "@conv_range," _
                    & "@low_value," _
                    & "@high_value," _
                    & "'Male'," _
                    & "@classification," _
                    & "@age_begin," _
                    & "@age_end," _
                    & "@test_code," _
                    & "@machine," _
                    & "@channel," _
                    & "@section" _
                    & ")"
                    )

                SaveRecord("INSERT INTO `reference_range` (`test_name`, `si_range`, `conv_range`, `low_value`, `high_value`, `gender`, `classification`, `age_begin`, `age_end`, `test_code`, `machine`, `channel`, `section`) VALUES " _
                    & "(" _
                    & "@test_name," _
                    & "@si_range," _
                    & "@conv_range," _
                    & "@low_value," _
                    & "@high_value," _
                    & "'Female'," _
                    & "@classification," _
                    & "@age_begin," _
                    & "@age_end," _
                    & "@test_code," _
                    & "@machine," _
                    & "@channel," _
                    & "@section" _
                    & ")"
                    )
            Else
                SaveRecord("INSERT INTO `reference_range` (`test_name`, `si_range`, `conv_range`, `low_value`, `high_value`, `gender`, `classification`, `age_begin`, `age_end`, `test_code`, `machine`, `channel`, `section`) VALUES " _
                    & "(" _
                    & "@test_name," _
                    & "@si_range," _
                    & "@conv_range," _
                    & "@low_value," _
                    & "@high_value," _
                    & "@gender," _
                    & "@classification," _
                    & "@age_begin," _
                    & "@age_end," _
                    & "@test_code," _
                    & "@machine," _
                    & "@channel," _
                    & "@section" _
                    & ")"
                    )
            End If
        Else
            SaveRecord("UPDATE `reference_range` SET " _
                    & "`test_name` = @test_name," _
                    & "`si_range` = @si_range," _
                    & "`conv_range` = @conv_range," _
                    & "`low_value` = @low_value," _
                    & "`high_value` = @high_value," _
                    & "`gender` = @gender," _
                    & "`classification` = @classification," _
                    & "`age_begin` = @age_begin," _
                    & "`age_end` = @age_end," _
                    & "`test_code` = @test_code," _
                    & "`machine` = @machine," _
                    & "`channel` = @channel," _
                    & "`section` = @section" _
                    & " WHERE `id` = @id"
                    )
        End If

        frmRangeList.LoadRecords()

        'For Each ctrl As Control In Me.XtraTabPage.Controls
        '    If TypeOf ctrl Is ComboBoxEdit Or TypeOf ctrl Is TextEdit Then
        '        If ctrl.Tag = "Section" Or ctrl.Tag = "Instrument" Then

        '        Else
        '            ctrl.Text = Nothing
        '        End If
        '    End If
        'Next
    End Sub
End Class