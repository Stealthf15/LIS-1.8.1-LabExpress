Public Class frmSpecimenCopy

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

            SaveRecord("INSERT INTO `specimen` (`specimen`, `description`, `specimen_type`, `channel`, `test_code`, `si_unit`, `conventional_unit`, `convertion_factor`, `convertion_multiplier`, `test_name`, `order_no`, `test_group`, `status`, `instrument`) " _
                       & "SELECT `specimen`, `description`, `specimen_type`, `channel`, `test_code`, `si_unit`, `conventional_unit`, `convertion_factor`, `convertion_multiplier`, `test_name`, `order_no`, `test_group`, `status`, @instrument_to FROM `default_specimen_mapping` WHERE `instrument` = @instrument_from"
                       )

            SaveRecordwthoutMSG("INSERT INTO `conversion_factor` (`test_code`,  `convertion_factor`, `convertion_multiplier`, `instrument`) " _
                       & "SELECT `test_code`, `convertion_factor`, `convertion_multiplier`, @instrument_to FROM `default_specimen_mapping` WHERE `instrument` = @instrument_from"
                       )

            frmTestName.LoadRecords()

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