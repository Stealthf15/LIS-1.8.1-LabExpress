Public Class frmChangeUser

    Public ID As String

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
                    txtPassword.Text = ""
                    txtUsername.Text = ""

                    For Each f As Form In frmMain.MdiChildren
                        f.Close()
                    Next

                    If chkNewWorklist.Checked = True Then
                        rs.Parameters.Clear()
                        rs.Parameters.Add("@Date", MySql.Data.MySqlClient.MySqlDbType.DateTime).Value = Format(Now, "yyyy-MM-dd")

                        '##################################------Archive Onqueue------###################################################################################
                        Connect()
                        rs.Dispose()
                        rs.Connection = conn
                        rs.CommandText = "INSERT INTO `order` (`status`, `sample_id`, `patient_id`, `patient_name`, `age`, `sex`, `bdate`, `physician`, `dept`, `medtech`, `verified_by`, `date`, `time`, `main_id`, `dt_released`, `patient_type`, `test`, `testtype`, `barcode`, `instrument`, `type`, `location`) " _
                            & "SELECT `status`, `sample_id`, `patient_id`, `patient_name`, `age`, `sex`, `bdate`, `physician`, `dept`, `medtech`, `verified_by`, `date`, `time`, CONCAT(`main_id`, '-', CURRENT_DATE()), `dt_released`, `patient_type`, `test`, `testtype`, `barcode`, `instrument`, `type`, `location` FROM `tmpWorklist` WHERE `location` = '" & cboLocation.Text & "'"
                        rs.ExecuteNonQuery()
                        Disconnect()

                        Connect()
                        rs.Connection = conn
                        rs.CommandText = "INSERT INTO `result` (`universal_id`, `flag`, `measurement`, `units`, `reference_range`, `value_conv`, `unit_conv`, `ref_conv`, `test_code`, `order_no`, `sample_id`, `test_group`, `date`, `patient_id`, `instrument`) " _
                            & "SELECT `universal_id`, `flag`, `measurement`, `units`, `reference_range`, `value_conv`, `unit_conv`, `ref_conv`, `test_code`, `order_no`, CONCAT(`sample_id`, '-', CURRENT_DATE()), `test_group`, `date`, `patient_id`, `instrument` FROM `tmpResult`"
                        rs.ExecuteNonQuery()
                        Disconnect()

                        Connect()
                        rs.Connection = conn
                        rs.CommandText = "DELETE FROM `tmpWorklist` WHERE `location` = '" & cboLocation.Text & "'"
                        rs.ExecuteNonQuery()
                        Disconnect()

                        Connect()
                        rs.Connection = conn
                        rs.CommandText = "DELETE FROM `tmpResult`"
                        rs.ExecuteNonQuery()
                        Disconnect()
                        '##################################------Archive Onqueue------###################################################################################

                        '##################################------RESET AUTONUMBER-----###################################################################################
                        Connect()
                        rs.Connection = conn
                        rs.CommandText = "DELETE FROM `last_id`"
                        rs.ExecuteNonQuery()
                        Disconnect()
                        '##################################------RESET AUTONUMBER-----###################################################################################

                        '##################################------Archive Completed------###################################################################################
                        'Connect()
                        'rs.Dispose()
                        'rs.Connection = conn
                        'rs.CommandText = "INSERT INTO `archive_worklist` (`status`, `sample_id`, `patient_id`, `patient_name`, `age`, `sex`, `bdate`, `physician`, `dept`, `medtech`, `verified_by`, `date`, `time`, `main_id`, `dt_released`, `patient_type`, `test`, `testtype`, `barcode`, `instrument`, `type`, `location`) " _
                        '    & "SELECT `status`, `sample_id`, `patient_id`, `patient_name`, `age`, `sex`, `bdate`, `physician`, `dept`, `medtech`, `verified_by`, `date`, `time`, `main_id`, `dt_released`, `patient_type`, `test`, `testtype`, `barcode`, `instrument`, `type`, `location` FROM `order` WHERE `date` < @Date"
                        'rs.ExecuteNonQuery()
                        'Disconnect()

                        'Connect()
                        'rs.Connection = conn
                        'rs.CommandText = "INSERT INTO `archive_result` (`universal_id`, `flag`, `measurement`, `units`, `reference_range`, `value_conv`, `unit_conv`, `ref_conv`, `test_code`, `order_no`, `sample_id`, `test_group`, `date`, `patient_id`, `instrument`) " _
                        '    & "SELECT `universal_id`, `flag`, `measurement`, `units`, `reference_range`, `value_conv`, `unit_conv`, `ref_conv`, `test_code`, `order_no`, `sample_id`, `test_group`, `date`, `patient_id`, `instrument` FROM `result` WHERE `date` < @Date"
                        'rs.ExecuteNonQuery()
                        'Disconnect()

                        'Connect()
                        'rs.Connection = conn
                        'rs.CommandText = "DELETE FROM `order` WHERE `date` < @Date"
                        'rs.ExecuteNonQuery()
                        'Disconnect()

                        'Connect()
                        'rs.Connection = conn
                        'rs.CommandText = "DELETE FROM `result` WHERE `date` < @Date"
                        'rs.ExecuteNonQuery()
                        'Disconnect()
                        '##################################------Archive Completed------###################################################################################
                    End If

                    Me.Close()
                    Me.Dispose()
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
            Disconnect()
            MessageBox.Show(ex.Message, "System Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Exit Sub
        End Try
    End Sub

    Private Sub frmChangeUser_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Me.txtUsername.Focus()

        Me.cboLocation.Properties.Items.Clear()
        Connect()
        rs.Connection = conn
        rs.CommandText = "SELECT `location` FROM `location` ORDER BY `id`"
        reader = rs.ExecuteReader
        While reader.Read
            Me.cboLocation.Properties.Items.Add(reader(0))
        End While
        Disconnect()
    End Sub
End Class