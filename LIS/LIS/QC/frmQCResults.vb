Imports System.Text.RegularExpressions

Public Class frmQCResults

    Public MedTechID As String = ""
    Public PathologistID As String = ""
    Public VerifyID As String = ""
    Public mainID As String
    Public Channel As String = ""
    Public MainSampleID As String = ""
    Public PatientID As String = ""
    Public RDate As Date

    Public FLAG As String = ""
    Public RANGE As String = ""

    Public Section As String = ""

    Public DefaultUnit As Integer = 0

    Public _runCount As Integer

    Dim Diff As Double

    Private Sub txtResult_KeyPress(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyPressEventArgs)
        If InStr("1234567890.", e.KeyChar) = 0 And Not Chr(AscW(e.KeyChar)) = vbBack Then
            e.KeyChar = ChrW(0)
        End If
    End Sub

    Private Sub frmResultNewChem_FormClosed(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
        Me.Dispose()
    End Sub

    Private Sub btnClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnClose.ItemClick
        Me.Close()
        Me.Dispose()
    End Sub

    Private Sub txtAge_KeyPress(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyPressEventArgs)
        If InStr("1234567890", e.KeyChar) = 0 And Not Chr(AscW(e.KeyChar)) = vbBack Then
            e.KeyChar = ChrW(0)
        End If
    End Sub

    Public Sub LoadTest()
        'On Error Resume Next
        Try
            If _runCount = "0" Then
                cboRunCount.Visible = False
                LabelControl2.Visible = False
            ElseIf _runCount = "1" Then
                dtResult.Columns.Clear()
                Dim SQL As String = "SELECT `universal_id`, `measurement`, `unit`, `ll`, `ul`, `instrument`, `test_code`, `id` FROM `QCResult` WHERE `control_id` = @MainID AND `instrument` = @instrument AND `date` = @date AND `run_count` = '1'"
                Connect()
                Dim command As New MySql.Data.MySqlClient.MySqlCommand(SQL, conn)

                command.Parameters.Clear()
                command.Parameters.AddWithValue("@MainID", MainSampleID)
                command.Parameters.AddWithValue("@instrument", frmQC.cboMachines.Text)
                command.Parameters.Add("@date", MySql.Data.MySqlClient.MySqlDbType.Date).Value = Format(dtFrom.Value, "yyyy-MM-dd")

                Dim adapter As New MySql.Data.MySqlClient.MySqlDataAdapter(command)

                Dim myTable As DataTable = New DataTable
                adapter.Fill(myTable)

                dtResult.Font = New Font("Tahoma", 9)
                dtResult.ForeColor = Color.Black
                dtResult.AutoGenerateColumns = False
                dtResult.DataSource = myTable
                Disconnect()

                Dim col As New DataGridViewImageColumn
                col.DataPropertyName = ""
                col.HeaderText = ""
                col.Name = "Image"
                col.Visible = True
                col.Image = ImageList.Images(0)
                col.Width = 20
                col.HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
                dtResult.Columns.Add(col)

                Dim col1 As New DataGridViewTextBoxColumn
                col1.DataPropertyName = "universal_id"
                col1.HeaderText = "Test"
                col1.Name = "Test"
                col1.Visible = True
                col1.Width = 170
                col1.HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
                dtResult.Columns.Add(col1)

                Dim col3 As New DataGridViewTextBoxColumn
                col3.DataPropertyName = "measurement"
                col3.HeaderText = "Result"
                col3.Name = "Result"
                col3.Visible = True
                col3.Width = 50
                col3.HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
                dtResult.Columns.Add(col3)

                Dim col4 As New DataGridViewTextBoxColumn
                col4.DataPropertyName = "unit"
                col4.HeaderText = "Units"
                col4.Name = "Units"
                col4.Visible = True
                col4.Width = 100
                col4.HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
                dtResult.Columns.Add(col4)

                Dim col5 As New DataGridViewTextBoxColumn
                col5.DataPropertyName = "ll"
                col5.HeaderText = "Lower Limit"
                col5.Name = "Low_Range"
                col5.Visible = True
                col5.Width = 110
                col5.HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
                dtResult.Columns.Add(col5)

                Dim col6 As New DataGridViewTextBoxColumn
                col6.DataPropertyName = "ul"
                col6.HeaderText = "Upper Limit"
                col6.Name = "High_Range"
                col6.Visible = True
                col6.Width = 110
                col6.HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
                dtResult.Columns.Add(col6)

                Dim col9 As New DataGridViewComboBoxColumn
                col9.DataPropertyName = "instrument"
                col9.HeaderText = "Instrument"
                col9.Name = "Instrument"

                Connect()
                rs.Connection = conn
                rs.CommandText = "SELECT CONCAT(`name`, '_', `serial_no`) as `machine` FROM `machines` WHERE `test_name` = '" & Section & "'"
                reader = rs.ExecuteReader
                While reader.Read
                    col9.Items.Add(reader(0).ToString)
                End While
                Disconnect()

                col9.Width = 150
                col9.HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
                dtResult.Columns.Add(col9)

                Dim col10 As New DataGridViewTextBoxColumn
                col10.DataPropertyName = "test_code"
                col10.HeaderText = "Test Code"
                col10.Name = "Test_Code"
                col10.Visible = True
                col10.Width = 100
                col10.HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
                dtResult.Columns.Add(col10)

                Dim col11 As New DataGridViewTextBoxColumn
                col11.DataPropertyName = "id"
                col11.HeaderText = "ID"
                col11.Name = "ID"
                col11.Visible = True
                col11.Width = 50
                col11.HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
                dtResult.Columns.Add(col11)

                Dim col12 As New DataGridViewTextBoxColumn
                col12.DataPropertyName = "test_group"
                col12.HeaderText = "Test Group"
                col12.Name = "Test_Group"
                col12.Visible = True
                col12.Width = 150
                col12.HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
                dtResult.Columns.Add(col12)
                cboRunCount.Visible = False
                LabelControl2.Visible = False
            Else
                dtResult.Columns.Clear()
                Dim SQL As String = "SELECT `universal_id`, `measurement`, `unit`, `ll`, `ul`, `instrument`, `test_code`, `id` FROM `QCResult` WHERE `control_id` = @MainID AND `instrument` = @instrument AND `date` = @date AND `run_count` = '" & cboRunCount.Text & "'"
                Connect()
                Dim command As New MySql.Data.MySqlClient.MySqlCommand(SQL, conn)

                command.Parameters.Clear()
                command.Parameters.AddWithValue("@MainID", MainSampleID)
                command.Parameters.AddWithValue("@instrument", frmQC.cboMachines.Text)
                command.Parameters.Add("@date", MySql.Data.MySqlClient.MySqlDbType.Date).Value = Format(dtFrom.Value, "yyyy-MM-dd")

                Dim adapter As New MySql.Data.MySqlClient.MySqlDataAdapter(command)

                Dim myTable As DataTable = New DataTable
                adapter.Fill(myTable)

                dtResult.Font = New Font("Tahoma", 9)
                dtResult.ForeColor = Color.Black
                dtResult.AutoGenerateColumns = False
                dtResult.DataSource = myTable
                Disconnect()

                Dim col As New DataGridViewImageColumn
                col.DataPropertyName = ""
                col.HeaderText = ""
                col.Name = "Image"
                col.Visible = True
                col.Image = ImageList.Images(0)
                col.Width = 20
                col.HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
                dtResult.Columns.Add(col)

                Dim col1 As New DataGridViewTextBoxColumn
                col1.DataPropertyName = "universal_id"
                col1.HeaderText = "Test"
                col1.Name = "Test"
                col1.Visible = True
                col1.Width = 170
                col1.HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
                dtResult.Columns.Add(col1)

                Dim col3 As New DataGridViewTextBoxColumn
                col3.DataPropertyName = "measurement"
                col3.HeaderText = "Result"
                col3.Name = "Result"
                col3.Visible = True
                col3.Width = 50
                col3.HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
                dtResult.Columns.Add(col3)

                Dim col4 As New DataGridViewTextBoxColumn
                col4.DataPropertyName = "unit"
                col4.HeaderText = "Units"
                col4.Name = "Units"
                col4.Visible = True
                col4.Width = 100
                col4.HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
                dtResult.Columns.Add(col4)

                Dim col5 As New DataGridViewTextBoxColumn
                col5.DataPropertyName = "ll"
                col5.HeaderText = "Lower Limit"
                col5.Name = "Low_Range"
                col5.Visible = True
                col5.Width = 110
                col5.HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
                dtResult.Columns.Add(col5)

                Dim col6 As New DataGridViewTextBoxColumn
                col6.DataPropertyName = "ul"
                col6.HeaderText = "Upper Limit"
                col6.Name = "High_Range"
                col6.Visible = True
                col6.Width = 110
                col6.HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
                dtResult.Columns.Add(col6)

                Dim col9 As New DataGridViewComboBoxColumn
                col9.DataPropertyName = "instrument"
                col9.HeaderText = "Instrument"
                col9.Name = "Instrument"

                Connect()
                rs.Connection = conn
                rs.CommandText = "SELECT CONCAT(`name`, '_', `serial_no`) as `machine` FROM `machines` WHERE `test_name` = '" & Section & "'"
                reader = rs.ExecuteReader
                While reader.Read
                    col9.Items.Add(reader(0).ToString)
                End While
                Disconnect()

                col9.Width = 150
                col9.HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
                dtResult.Columns.Add(col9)

                Dim col10 As New DataGridViewTextBoxColumn
                col10.DataPropertyName = "test_code"
                col10.HeaderText = "Test Code"
                col10.Name = "Test_Code"
                col10.Visible = True
                col10.Width = 100
                col10.HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
                dtResult.Columns.Add(col10)

                Dim col11 As New DataGridViewTextBoxColumn
                col11.DataPropertyName = "id"
                col11.HeaderText = "ID"
                col11.Name = "ID"
                col11.Visible = True
                col11.Width = 50
                col11.HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
                dtResult.Columns.Add(col11)

                Dim col12 As New DataGridViewTextBoxColumn
                col12.DataPropertyName = "test_group"
                col12.HeaderText = "Test Group"
                col12.Name = "Test_Group"
                col12.Visible = True
                col12.Width = 150
                col12.HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
                dtResult.Columns.Add(col12)

                cboRunCount.Visible = True
                LabelControl2.Visible = True
            End If
        Finally
            If conn.State = ConnectionState.Open Then
                conn.Close()
            End If
        End Try
    End Sub

    Private Sub frmResultsNew_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        AutoLoadQCRunCount()
        cboRunCount.Text = "1" 'Set Default Value in Loading of Result
        LoadTest()
    End Sub

    Private Sub btnPrint_ItemClick(ByVal sender As Object, ByVal e As DevExpress.XtraBars.ItemClickEventArgs) Handles btnPrint.ItemClick
        Using myRDLCPrinter As New RDLCPrinterQC(MainSampleID, "", frmQC.cboControl.Text, frmQC.cboMachines.Text, Now.ToShortDateString(), frmQC.cboLimit.Text, My.Settings.DefaultPrinter)
            myRDLCPrinter.Print(1)
        End Using
        Me.Dispose()
        Me.Close()
    End Sub

    Private Sub btnViewPrint_ItemClick(ByVal sender As Object, ByVal e As DevExpress.XtraBars.ItemClickEventArgs) Handles btnViewPrint.ItemClick
        'RPTresults.sample_id = MainSampleID
        PrintPreviewQC(MainSampleID, "QCResult", "", frmQC.cboControl.Text, frmQC.cboMachines.Text, Now.ToShortDateString(), frmQC.cboLimit.Text, RPTresults, RPTresults.ReportViewer1)
    End Sub

    Private Sub btnUpdate_ItemClick(ByVal sender As Object, ByVal e As DevExpress.XtraBars.ItemClickEventArgs) Handles btnUpdate.ItemClick
        For x = 0 To dtResult.RowCount - 1 Step 1
            Connect()
            rs.Connection = conn
            rs.CommandText = "UPDATE `control_result` SET `measurement` = '" & dtResult.Rows(x).Cells(2).Value & "' WHERE `id` = '" & dtResult.Rows(x).Cells(8).Value & "'"
            rs.ExecuteNonQuery()
            Disconnect()
        Next

        MessageBox.Show("Selected Result has been successfully update.", "Update", MessageBoxButtons.OK, MessageBoxIcon.Information)
    End Sub

    Private Sub btnReject_ItemClick(ByVal sender As Object, ByVal e As DevExpress.XtraBars.ItemClickEventArgs) Handles btnReject.ItemClick
        Dim rows As DataGridViewRow = dtResult.SelectedRows(0)

        rs.Parameters.Clear()
        rs.Parameters.AddWithValue("@id", rows.Cells(8).Value)

        DeleteRecordSQL("DELETE FROM `control_result` WHERE `id` = @id")
        LoadTest()
    End Sub

    Private Sub dtFrom_ValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles dtFrom.ValueChanged
        cboRunCount.Properties.Items.Clear() 'Clear Run Count Drop down
        dtResult.Columns.Clear() 'Clear DT Result Window
        AutoLoadQCRunCount() 'Load Run Count number for drop down
        LoadTest() 'Load Test
    End Sub

    Private Sub cboRunCount_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cboRunCount.SelectedIndexChanged
        LoadTest()
    End Sub

    Public Sub AutoLoadQCRunCount()
        Try
            Connect()
            _runCount = 0
            cboRunCount.Properties.Items.Clear()
            rs.Parameters.Clear()
            rs.Parameters.AddWithValue("@MainID", MainSampleID)
            rs.Parameters.AddWithValue("@instrument", frmQC.cboMachines.Text)
            rs.Parameters.Add("@date", MySql.Data.MySqlClient.MySqlDbType.Date).Value = Format(dtFrom.Value, "yyyy-MM-dd")
            rs.Connection = conn
            rs.CommandText = "SELECT `run_count`FROM `qcresult` WHERE `control_id` = @MainID AND `instrument` = @instrument AND `date` = @date"
            reader = rs.ExecuteReader
            While reader.Read
                _runCount = reader(0).ToString
            End While

            For i As Integer = 1 To _runCount Step 1
                cboRunCount.Properties.Items.Add(i)
            Next
            Disconnect()
        Catch ex As Exception
            Disconnect()
            MessageBox.Show("Error on Previewing of Results. Error Code #0000039", "Previewing of Results Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub btnAddControl_ItemClick(sender As Object, e As DevExpress.XtraBars.ItemClickEventArgs) Handles btnAddControl.ItemClick
        frmQCResultsAE._section = Section
        frmQCResultsAE._controlName = MainSampleID
        frmQCResultsAE._runCount = cboRunCount.Text
        frmQCResultsAE.ShowDialog()
    End Sub
End Class