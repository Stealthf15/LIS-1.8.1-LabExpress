Public Class frmRangeCopy

    Public name As String

    Private Sub btnBack_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnClose.Click
        Me.Close()
    End Sub

    Private Sub btnSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSave.Click
        Try
            If Me.cboFrom.Text = "" Or Me.cboType.Text = "" Then
                MessageBox.Show("Please fill all required information.", "Empty String", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
                Exit Sub
            End If

            rs.Parameters.Clear()
            rs.Parameters.AddWithValue("@instrument_from", cboFrom.Text)
            rs.Parameters.AddWithValue("@instrument_to", cboTo.Text)

            SaveRecord("INSERT INTO `reference_range` (`test_name`, `si_range`, `conv_range`, `low_value`, `high_value`, `gender`, `classification`, `age_begin`, `age_end`, `test_code`, `machine`, `channel`, `section`) " _
                    & "SELECT `test_name`, `si_range`, `conv_range`, `low_value`, `high_value`, `gender`, `classification`, `age_begin`, `age_end`, `test_code`,  @instrument_to, `channel`, `section` FROM `default_reference_range` WHERE `machine` = @instrument_from"
                    )

            frmRangeList.LoadRecords()
            Me.Close()
        Catch ex As Exception
            Disconnect()
            MessageBox.Show(ex.Message, "Exception", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
        End Try
    End Sub

    Public Sub AutoLoadCombo()
        cboTo.Properties.Items.Clear()
        Connect()
        rs.Connection = conn
        rs.CommandText = "SELECT CONCAT(`name`, '_', `serial_no`) as `machine` FROM `machines` ORDER BY `machine`"
        reader = rs.ExecuteReader
        While reader.Read
            cboTo.Properties.Items.Add(reader(0).ToString)
        End While
        Disconnect()
    End Sub

    Private Sub frmTestTypeAE_FormClosed(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
        Me.Dispose()
    End Sub

    Private Sub frmSpecimenCopy_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        AutoLoadCombo()
    End Sub
End Class