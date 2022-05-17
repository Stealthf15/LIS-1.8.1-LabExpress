Public Class frmLogin

    Public ID As String
    Private medtech_Name As String

    Private Sub btnBack_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnClose.Click
        End
    End Sub
    Private Sub loadMedtechName()
        Connect()
        rs.Connection = conn
        rs.CommandText = "SELECT CONCAT(fname, ' ', mname, ' ', lname, ', ', designation) AS `name` FROM `medtech` WHERE id = '" & CurrEmail & "' ORDER BY `name`"
        reader = rs.ExecuteReader
        While reader.Read
            medtech_Name = reader(0).ToString
        End While
        Disconnect()
    End Sub

    Private Sub btnSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSave.Click
        Try
            'Got this Encrypting code from StockOverFlow.com
            Dim bytes() As Byte = System.Text.Encoding.UTF8.GetBytes(txtPassword.Text)
            Dim hashOfBytes() As Byte = New System.Security.Cryptography.SHA1Managed().ComputeHash(bytes)
            Dim strHash As String = Convert.ToBase64String(hashOfBytes)

            rs.Parameters.Clear()
            rs.Parameters.AddWithValue("@username", txtUsername.Text)
            rs.Parameters.AddWithValue("@password", strHash)
            rs.Parameters.AddWithValue("@location", cboLocation.Text)

            Connect()
            rs.Connection = conn
            rs.CommandText = "SELECT * FROM `user_account` WHERE `Username` = @username AND `password` = @password AND `dept` = @location"
            reader = rs.ExecuteReader
            reader.Read()
            If (reader.HasRows = True) Then
                If (reader("username") = Me.txtUsername.Text) Then
                    CurrID = reader(0).ToString
                    CurrUser = reader(1).ToString
                    CurrPost = reader(7).ToString
                    CurrDept = reader(6).ToString
                    CurrType = reader(4).ToString
                    CurrEmail = reader(5).ToString
                    Disconnect()
                    loadMedtechName()
                    If chkNewWorklist.Checked = True Then
                        rs.Parameters.Clear()
                        rs.Parameters.Add("@Date", MySql.Data.MySqlClient.MySqlDbType.DateTime).Value = Format(Now, "yyyy-MM-dd")

                        '##################################------Archive Onqueue------###################################################################################
                        '-------added in 0.4.6.5----------
                        Connect()
                        rs.Connection = conn
                        rs.CommandType = CommandType.Text
                        rs.CommandText = "SELECT `status`, `sample_id`, `patient_id`, `main_id` FROM `tmpWorklist` WHERE (`status` = 'Result Received' OR `status` = 'Validated' OR `status` = 'Checked-In')"
                        reader = rs.ExecuteReader
                        While reader.Read
                            Connect1()
                            rs1.Connection = conn1
                            rs1.CommandText = "INSERT INTO `result` (`universal_id`, `flag`, `measurement`, `units`, `reference_range`, `value_conv`, `unit_conv`, `ref_conv`, `test_code`, `order_no`, `sample_id`, `test_group`, `date`, `patient_id`, `status`, `instrument`, `his_code`, `his_mainid`, `section`, `sub_section`, `print_status`) " _
                            & "SELECT `universal_id`, `flag`, `measurement`, `units`, `reference_range`, `value_conv`, `unit_conv`, `ref_conv`, `test_code`, `order_no`, `sample_id`, `test_group`, `date`, `patient_id`, `status`, `instrument`, `his_code`, `his_mainid`, `section`, `sub_section`, `print_status` FROM `tmpResult` WHERE `sample_id` = '" & reader(1).ToString & "' AND '" & reader(2).ToString & "'"
                            rs1.ExecuteNonQuery()
                            Disconnect1()

                            Connect1()
                            rs1.Connection = conn1
                            rs1.CommandText = "DELETE FROM `tmpResult` WHERE `sample_id` = '" & reader(1).ToString & "' AND '" & reader(2).ToString & "'"
                            rs1.ExecuteNonQuery()
                            Disconnect1()
                        End While
                        Disconnect()
                        '--------------------------------

                        Connect()
                        rs.Dispose()
                        rs.Connection = conn
                        rs.CommandText = "INSERT INTO `order` (`status`, `sample_id`, `patient_id`, `patient_name`, `age`, `sex`, `bdate`, `physician`, `dept`, `medtech`, `verified_by`, `date`, `time`, `main_id`, `dt_released`, `patient_type`, `test`, `testtype`, `barcode`, `instrument`, `type`, `location`, sub_section) " _
                            & "SELECT `status`, `sample_id`, `patient_id`, `patient_name`, `age`, `sex`, `bdate`, `physician`, `dept`, `medtech`, `verified_by`, `date`, `time`, CONCAT(`main_id`, '-', CURRENT_DATE()), `dt_released`, `patient_type`, `test`, `testtype`, `barcode`, `instrument`, `type`, `location`, `sub_section` FROM `tmpWorklist` WHERE `location` = '" & cboLocation.Text & "' AND (`status` = 'Result Received' OR `status` = 'Validated' OR `status` = 'Checked-In')"
                        rs.ExecuteNonQuery()
                        Disconnect()

                        Connect()
                        rs.Dispose()
                        rs.Connection = conn
                        rs.CommandText = "INSERT INTO `archived_order` (`sample_id`, `patient_id`, `date_archived`, `archived_by`) " _
                            & "SELECT `sample_id`, `patient_id`, CURRENT_TIMESTAMP(), '" & medtech_Name & "' FROM `tmpWorklist` WHERE `location` = '" & cboLocation.Text & "' AND (`status` = 'Result Received' OR `status` = 'Validated' OR `status` = 'Checked-In')"
                        rs.ExecuteNonQuery()
                        Disconnect()

                        'Connect()
                        'rs.Connection = conn
                        'rs.CommandText = "INSERT INTO `result` (`universal_id`, `flag`, `measurement`, `units`, `reference_range`, `value_conv`, `unit_conv`, `ref_conv`, `test_code`, `order_no`, `sample_id`, `test_group`, `date`, `patient_id`, `instrument`) " _
                        '    & "SELECT `universal_id`, `flag`, `measurement`, `units`, `reference_range`, `value_conv`, `unit_conv`, `ref_conv`, `test_code`, `order_no`, CONCAT(`sample_id`, '-', CURRENT_DATE()), `test_group`, `date`, `patient_id`, `instrument` FROM `tmpResult`"
                        'rs.ExecuteNonQuery()
                        'Disconnect()

                        Connect()
                        rs.Connection = conn
                        rs.CommandText = "DELETE FROM `tmpWorklist` WHERE `location` = '" & cboLocation.Text & "' AND (`status` = 'Result Received' OR `status` = 'Validated' OR `status` = 'Checked-In')"
                        rs.ExecuteNonQuery()
                        Disconnect()

                        'Connect()
                        'rs.Connection = conn
                        'rs.CommandText = "DELETE FROM `tmpResult`"
                        'rs.ExecuteNonQuery()
                        'Disconnect()
                        '##################################------Archive Onqueue------###################################################################################

                        '##################################------RESET AUTONUMBER-----###################################################################################
                        'Connect()
                        'rs.Connection = conn
                        'rs.CommandText = "DELETE FROM `last_id`"
                        'rs.ExecuteNonQuery()
                        'Disconnect()
                        '##################################------RESET AUTONUMBER-----###################################################################################
                    End If

                    'Timer for Realtime New and Pending Orders
                    MainFOrm.TimerNew.Enabled = True
                    MainFOrm.TimerPending.Enabled = True

                    MainFOrm.DisableControls()

                    MainFOrm.Show()
                    Me.Hide()
                    txtPassword.Text = ""
                    txtUsername.Text = ""
                Else
                    Disconnect()
                    MessageBox.Show("Invalid Username or Password.", "Authentication Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    Exit Sub
                End If
            Else
                Disconnect()
                MessageBox.Show("Invalid Username or Password.", "Authentication Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Exit Sub
            End If
            Disconnect()
        Catch ex As Exception
            frmLoginTech.ShowDialog()
            Exit Sub
        End Try
    End Sub

    Private Sub btnConfig_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        frmLoginTech.ShowDialog()
    End Sub

    Private Sub frmLogin_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        'On Error Resume Next
        'If My.Settings.isFirstRun Then
        '    frmSetup.ShowDialog()
        '    Me.Hide()
        'End If
        LoadLocation()
    End Sub

    Private Sub LoadLocation()
        Me.cboLocation.Properties.Items.Clear()
        Try
            Connect()
            rs.Connection = conn
            rs.CommandText = "SELECT `location` FROM `location` ORDER BY `id`"
            reader = rs.ExecuteReader
            While reader.Read
                Me.cboLocation.Properties.Items.Add(reader(0))
            End While
            Disconnect()
        Catch ex As Exception
            frmLoginTech.ShowDialog()
        End Try
    End Sub

    Private Sub lblDB_Click(sender As Object, e As EventArgs) Handles lblDB.Click
        frmLoginTech.ShowDialog()
    End Sub
End Class