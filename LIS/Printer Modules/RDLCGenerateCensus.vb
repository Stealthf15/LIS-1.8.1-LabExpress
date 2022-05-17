Imports System.IO
Imports System.IO.Ports
Imports System.Threading
Imports System.Threading.Thread
Imports System.Text.RegularExpressions

Imports System.Text
Imports DevExpress

Imports Microsoft.Reporting

Module RDLCGenerateCensus
    Public Function GenerateReportCensus(ByVal dtName As String, ByVal dtFrom As DateTimePicker, ByVal dtTo As DateTimePicker, ByVal cboType As DevExpress.XtraEditors.ComboBoxEdit, ByVal rdlcName As String, ByVal ReportViewerData As Microsoft.Reporting.WinForms.ReportViewer)
        'Clear all datasurce
        ReportViewerData.LocalReport.DataSources.Clear()

        'SQL Query to retrieve records
        Dim SQL As String = "SELECT * FROM `" & dtName & "` WHERE (`date` BETWEEN STR_TO_DATE('" & dtFrom.Value & "', '%m/%d/%Y') AND STR_TO_DATE('" & dtTo.Value & "', '%m/%d/%Y')) ORDER BY `date`"

        'Connect to Database for Dataset1
        Connect()
        Dim dt As DataTable = New DataTable
        Dim command As New MySql.Data.MySqlClient.MySqlCommand(SQL, conn)

        'Set Parameters for command
        command.Parameters.Add("@d1", MySql.Data.MySqlClient.MySqlDbType.Date).Value = dtFrom.Value.ToString
        command.Parameters.Add("@d2", MySql.Data.MySqlClient.MySqlDbType.Date).Value = dtTo.Value.ToString

        'Fill adapter set in Dataset1
        Dim adapter As New MySql.Data.MySqlClient.MySqlDataAdapter(command)
        adapter.Fill(dt)
        Disconnect()

        'Set Report source data for Dataset1
        Dim ReportDataSource1 As Microsoft.Reporting.WinForms.ReportDataSource = New Microsoft.Reporting.WinForms.ReportDataSource()
        ReportDataSource1.Name = "DataSet1"
        ReportDataSource1.Value = dt
        ReportViewerData.LocalReport.DataSources.Add(ReportDataSource1)

        'Connect to Database for Dataset2
        Connect()
        Dim dt1 As DataTable = New DataTable
        Dim command1 As New MySql.Data.MySqlClient.MySqlCommand("SELECT * FROM `company_profile`", conn)

        'Fill adapter set in Dataset2
        Dim adapter1 As New MySql.Data.MySqlClient.MySqlDataAdapter(command1)
        adapter1.Fill(dt1)
        Disconnect()

        'Set Report source data for Dataset2
        Dim ReportDataSource2 As Microsoft.Reporting.WinForms.ReportDataSource = New Microsoft.Reporting.WinForms.ReportDataSource()
        ReportDataSource2.Name = "DataSet2"
        ReportDataSource2.Value = dt1
        ReportViewerData.LocalReport.DataSources.Add(ReportDataSource2)
        ReportViewerData.LocalReport.ReportPath = Application.StartupPath & "\Reports\Census\" & rdlcName

        'Set Parameters for reporting
        Dim DateFrom As New Microsoft.Reporting.WinForms.ReportParameter("DateFrom", FormatDateTime(dtFrom.Value, DateFormat.ShortDate))
        Dim DateTo As New Microsoft.Reporting.WinForms.ReportParameter("DateTo", FormatDateTime(dtTo.Value, DateFormat.ShortDate))
        Dim PatientType As New Microsoft.Reporting.WinForms.ReportParameter("PatientType", cboType.Text)

        ReportViewerData.LocalReport.SetParameters(DateFrom)
        ReportViewerData.LocalReport.SetParameters(DateTo)
        ReportViewerData.LocalReport.SetParameters(PatientType)

        'Set display format
        ReportViewerData.SetDisplayMode(Microsoft.Reporting.WinForms.DisplayMode.PrintLayout)

        'Refresh Report
        ReportViewerData.RefreshReport()
        Return (rdlcName)

    End Function
End Module
