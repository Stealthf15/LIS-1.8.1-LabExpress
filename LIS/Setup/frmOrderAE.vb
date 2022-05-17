Public Class frmOrderAE

    Public name As String

    Private Sub btnBack_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnClose.Click
        Me.Close()
    End Sub

    Private Sub btnSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSave.Click
        If Me.cboTest.Text = "" Or Me.cboTestName.Text = "" Then
            MessageBox.Show("Please fill all required information.", "Empty String", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            Exit Sub
        End If

        If chkPackage.Checked = True Then
            rs.Parameters.Clear()
            rs.Parameters.AddWithValue("@name", name)
            rs.Parameters.AddWithValue("@testname", cboTest.Text)
            rs.Parameters.AddWithValue("@test_code", txtTestCode.Text)
            rs.Parameters.AddWithValue("@section", cboTestName.Text)
            rs.Parameters.AddWithValue("@package", txtPackge.Text)

            If chkDisplay.Checked = True Then
                rs.Parameters.AddWithValue("@status", 1)
            Else
                rs.Parameters.AddWithValue("@status", 0)
            End If

            If Me.btnSave.Text = "&Save" Then
                SaveRecord("INSERT INTO `default_order` (`package`, `test_code`, `section`, `status`) VALUES (@package, @test_code, @section, @status)")

                frmOrder.LoadRecords()
                Me.Close()
            Else
                UpdateRecord("UPDATE `default_order` SET `package` = @package, `test_code` = @test_code, `section` = @section, `status` = @status WHERE `id` = @name")
                frmOrder.LoadRecords()
                Me.Close()
            End If
        Else
            rs.Parameters.Clear()
            rs.Parameters.AddWithValue("@name", name)
            rs.Parameters.AddWithValue("@testname", cboTest.Text)
            rs.Parameters.AddWithValue("@test_code", txtTestCode.Text)
            rs.Parameters.AddWithValue("@section", cboTestName.Text)
            rs.Parameters.AddWithValue("@package", txtPackge.Text)

            If chkDisplay.Checked = True Then
                rs.Parameters.AddWithValue("@status", 1)
            Else
                rs.Parameters.AddWithValue("@status", 0)
            End If

            If Me.btnSave.Text = "&Save" Then
                SaveRecord("INSERT INTO `default_order` (`specimen`, `test_code`, `section`, `status`) VALUES (@testname, @test_code, @section, @status)")

                frmOrder.LoadRecords()
                Me.Close()
            Else
                UpdateRecord("UPDATE `default_order` SET `specimen` = @testname, `test_code` = @test_code, `section` = @section, `status` = @status WHERE `id` = @name")
                frmOrder.LoadRecords()
                Me.Close()
            End If
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

    Private Sub frmTestTypeAE_FormClosed(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
        Me.Dispose()
    End Sub

    Private Sub cboTestName_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cboTestName.SelectedIndexChanged
        cboTest.Properties.Items.Clear()
        Connect()
        rs.Connection = conn
        rs.CommandText = "SELECT `specimen` FROM `specimen` WHERE `test_name` = '" & cboTestName.Text & "' ORDER BY `specimen`"
        reader = rs.ExecuteReader
        While reader.Read
            cboTest.Properties.Items.Add(reader(0).ToString)
        End While
        Disconnect()
    End Sub

    Dim TestCode As String

    Private Sub cboTest_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cboTest.SelectedIndexChanged
        Connect()
        rs.Connection = conn
        rs.CommandText = "SELECT `test_code` FROM `specimen` WHERE `specimen` = '" & cboTest.Text & "'"
        reader = rs.ExecuteReader
        reader.Read()
        If reader.HasRows Then
            TestCode = reader(0).ToString
        End If
        Disconnect()
    End Sub

    Private Sub btnAddTest_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnAddTest.Click
        Me.txtTestCode.Text &= TestCode & "^" & cboTest.Text & "~"
    End Sub

    Private Sub chkPackage_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkPackage.CheckedChanged
        If chkPackage.Checked = True Then
            txtPackge.Enabled = True
        Else
            txtPackge.Enabled = False
        End If
        txtPackge.Text = ""
    End Sub
End Class