Imports System.IO
Imports System.Data
Imports System.Text
Imports System.Drawing.Imaging
Imports System.Drawing.Printing
Imports System.Collections.Generic
Imports Microsoft.Reporting.WinForms
Imports System.Reflection
Imports MySql.Data.MySqlClient

Module RDLCPrintPreview

    Public Sub PrintPreview(ByVal SampleID As String, ByVal WorkListTable As String, ByVal ResultTable As String, ByVal PrintStatus As String, ByVal Section As String, ByVal SubSection As String, ByVal Frm As Form, ByVal ReportViewerData As Microsoft.Reporting.WinForms.ReportViewer)
        Try
            'Clear all datasurce
            ReportViewerData.LocalReport.DataSources.Clear()

            ReportViewerData.LocalReport.ReportPath = Application.StartupPath & "\Reports\Result\" & SubSection & ".rdlc"

            'Select report according to section

            'IF REPORT HAS DATASET WE NEED THIS
            '###################----DataSet----#################################################################
            'Order Table
            Dim cn = New MySqlConnection(MyConnectionString)
            Dim SQL = "SELECT * FROM `" & WorkListTable & "` WHERE `main_id` = '" & SampleID & "' AND testtype = '" & Section & "' AND sub_section = '" & SubSection & "'"
            Dim adapter = New MySqlDataAdapter(SQL, cn)
            Dim rdlc_table = New DataTable()
            adapter.Fill(rdlc_table)
            Dim rdlc_rds = New ReportDataSource("DataSet1", rdlc_table)
            ReportViewerData.LocalReport.DataSources.Add(rdlc_rds)
            Disconnect()

            'Result Table
            Connect()
            Dim result_SQL = "SELECT * FROM `" & ResultTable & "` WHERE `sample_id` = '" & SampleID & "' AND `print_status` = '" & PrintStatus & "' And Section = '" & Section & "' AND sub_section = '" & SubSection & "' ORDER BY `order_no`"
            Dim result_adapter = New MySqlDataAdapter(result_SQL, conn)
            Dim result_rdlc_table = New DataTable()
            result_adapter.Fill(result_rdlc_table)
            Dim result_rdlc_rds = New ReportDataSource("DataSet2", result_rdlc_table)
            ReportViewerData.LocalReport.DataSources.Add(result_rdlc_rds)
            Disconnect()

            'Patient Remarks Table
            Connect()
            SQL = "SELECT * FROM `patient_remarks` WHERE `sample_id` LIKE '" & SampleID & "'"
            adapter = New MySqlDataAdapter(SQL, conn)
            rdlc_table = New DataTable()
            adapter.Fill(rdlc_table)
            rdlc_rds = New ReportDataSource("DataSet3", rdlc_table)
            ReportViewerData.LocalReport.DataSources.Add(rdlc_rds)
            Disconnect()

            'ViewOrderMedtech Table
            Connect()
            SQL = "SELECT * FROM `viewordermedtech` WHERE `sample_id` LIKE '" & SampleID & "'"
            adapter = New MySqlDataAdapter(SQL, conn)
            rdlc_table = New DataTable()
            adapter.Fill(rdlc_table)
            rdlc_rds = New ReportDataSource("DataSet4", rdlc_table)
            ReportViewerData.LocalReport.DataSources.Add(rdlc_rds)
            Disconnect()

            'ViewOrderPathologist Table
            Connect()
            SQL = "SELECT * FROM `vieworderpathologist` WHERE `sample_id` LIKE '" & SampleID & "'"
            adapter = New MySqlDataAdapter(SQL, conn)
            rdlc_table = New DataTable()
            adapter.Fill(rdlc_table)
            rdlc_rds = New ReportDataSource("DataSet5", rdlc_table)
            ReportViewerData.LocalReport.DataSources.Add(rdlc_rds)
            Disconnect()

            'Company Profile Table
            Connect()
            SQL = "SELECT * FROM `company_profile`"
            adapter = New MySqlDataAdapter(SQL, conn)
            rdlc_table = New DataTable()
            adapter.Fill(rdlc_table)
            rdlc_rds = New ReportDataSource("DataSet6", rdlc_table)
            ReportViewerData.LocalReport.DataSources.Add(rdlc_rds)
            Disconnect()

            'ViewOrderVerified Table
            Connect()
            SQL = "SELECT * FROM `vieworderverified` WHERE `sample_id` LIKE '" & SampleID & "'"
            adapter = New MySqlDataAdapter(SQL, conn)
            rdlc_table = New DataTable()
            adapter.Fill(rdlc_table)
            rdlc_rds = New ReportDataSource("DataSet8", rdlc_table)
            ReportViewerData.LocalReport.DataSources.Add(rdlc_rds)
            Disconnect()

            'Additional(Information)
            Connect()
            SQL = "SELECT * FROM `additional_info` WHERE `sample_id` LIKE '" & SampleID & "'"
            adapter = New MySqlDataAdapter(SQL, conn)
            rdlc_table = New DataTable()
            adapter.Fill(rdlc_table)
            rdlc_rds = New ReportDataSource("DataSet7", rdlc_table)
            ReportViewerData.LocalReport.DataSources.Add(rdlc_rds)
            Disconnect()

            'Patient Information
            Connect()
            SQL = "SELECT * FROM `patient_info` WHERE `sample_id` LIKE '" & SampleID & "'"
            adapter = New MySqlDataAdapter(SQL, conn)
            rdlc_table = New DataTable()
            adapter.Fill(rdlc_table)
            rdlc_rds = New ReportDataSource("DataSet9", rdlc_table)
            ReportViewerData.LocalReport.DataSources.Add(rdlc_rds)
            Disconnect()

            'Specimen Tracking
            Connect()
            SQL = "SELECT * FROM `specimen_tracking` WHERE `sample_id` LIKE '" & SampleID & "'"
            adapter = New MySqlDataAdapter(SQL, conn)
            rdlc_table = New DataTable()
            adapter.Fill(rdlc_table)
            rdlc_rds = New ReportDataSource("DataSet10", rdlc_table)
            ReportViewerData.LocalReport.DataSources.Add(rdlc_rds)
            Disconnect()

            '###################----End of DataSet----############################################################

            'IF REPORT HAS PARAMETERS WE NEED THIS
            '###################----Parameters----################################################################
            'Parameterized data to pass in report parameters

            Dim Print_Status As New ReportParameter("PrintStatus", PrintStatus)
            ReportViewerData.LocalReport.SetParameters(Print_Status)
            '######################----End of Parameters----###############################################################

            'Set display format
            ReportViewerData.SetDisplayMode(Microsoft.Reporting.WinForms.DisplayMode.PrintLayout)

            'Refresh Report
            ReportViewerData.RefreshReport()
            Frm.Show()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error Previewing Report", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Exit Sub
        End Try
    End Sub
End Module
