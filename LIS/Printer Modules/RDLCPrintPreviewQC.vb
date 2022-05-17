Imports System.IO
Imports System.Data
Imports System.Text
Imports System.Drawing.Imaging
Imports System.Drawing.Printing
Imports System.Collections.Generic
Imports Microsoft.Reporting.WinForms
Imports System.Reflection
Imports MySql.Data.MySqlClient

Module RDLCPrintPreviewQC

    Public Sub PrintPreviewQC(ByVal SampleID As String, ByVal ResultTable As String, ByVal PrintStatus As String, ByVal ControlName As String, ByVal MachineName As String, ByVal DateNow As String, ByVal Section As String, ByVal Frm As Form, ByVal ReportViewerData As Microsoft.Reporting.WinForms.ReportViewer)
        Try
            'Clear all datasurce
            ReportViewerData.LocalReport.DataSources.Clear()

            'Select report according to section
            ReportViewerData.LocalReport.ReportPath = Application.StartupPath & "\Reports\QC\ReportStandardQC.RDLC"

            'IF REPORT HAS DATASET WE NEED THIS
            '###################----DataSet----#################################################################
            'Order Table
            Dim cn = New MySqlConnection(MyConnectionString)
            Dim SQL = "SELECT * FROM `company_profile`"
            Dim adapter = New MySqlDataAdapter(SQL, cn)
            Dim rdlc_table = New DataTable()
            adapter.Fill(rdlc_table)
            Dim rdlc_rds = New ReportDataSource("DataSet6", rdlc_table)
            ReportViewerData.LocalReport.DataSources.Add(rdlc_rds)
            Disconnect()

            'Result Table
            Connect()
            Dim result_SQL = "SELECT * FROM `" & ResultTable & "` WHERE `control_id` = '" & SampleID & "' AND `date` = CURRENT_DATE()"
            Dim result_adapter = New MySqlDataAdapter(result_SQL, conn)
            Dim result_rdlc_table = New DataTable()
            result_adapter.Fill(result_rdlc_table)
            Dim result_rdlc_rds = New ReportDataSource("DataSet2", result_rdlc_table)
            ReportViewerData.LocalReport.DataSources.Add(result_rdlc_rds)
            Disconnect()
            '###################----End of DataSet----############################################################

            'IF REPORT HAS PARAMETERS WE NEED THIS
            '###################----Parameters----################################################################
            'Parameterized data to pass in report parameters
            Dim Print_Status As New ReportParameter("PrintStatus", PrintStatus)
            ReportViewerData.LocalReport.SetParameters(Print_Status)

            Dim Control_Name As New ReportParameter("ControlName", ControlName)
            ReportViewerData.LocalReport.SetParameters(Control_Name)

            Dim Machine_Name As New ReportParameter("Machine", MachineName)
            ReportViewerData.LocalReport.SetParameters(Machine_Name)

            Dim Date_Now As New ReportParameter("Date", DateNow)
            ReportViewerData.LocalReport.SetParameters(Date_Now)

            Dim SectionName As New ReportParameter("Section", Section)
            ReportViewerData.LocalReport.SetParameters(SectionName)

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
