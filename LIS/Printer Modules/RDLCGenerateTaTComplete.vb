Imports System.IO
Imports System.IO.Ports
Imports System.Threading
Imports System.Threading.Thread
Imports System.Text.RegularExpressions

Imports System.Text
Imports DevExpress

Imports Microsoft.Reporting
Imports MySql.Data.MySqlClient

Module RDLCGenerateTaTComplete
    Public Function GenerateReportTaT(ByVal cboLimit As XtraEditors.ComboBoxEdit, ByVal cbomedtech As XtraEditors.ComboBoxEdit, ByVal dtFrom As DateTimePicker, ByVal dtTo As DateTimePicker, ByVal rdlcName As String, ByVal ReportViewerData As Microsoft.Reporting.WinForms.ReportViewer)
        'Clear all datasurce
        ReportViewerData.LocalReport.DataSources.Clear()
        Try
            If cbomedtech.Text = "All" And cboLimit.Text = "All" Then
                Connect()
                rs.Parameters.Clear()
                rs.Parameters.AddWithValue("@DateFrom", Format(dtFrom.Value, "yyyy-MM-dd"))
                rs.Parameters.AddWithValue("@DateTo", Format(dtTo.Value, "yyyy-MM-dd"))
                Dim sql As String
                sql = "Tat_Detailed"

                Dim dt As DataTable = New DataTable
                rs.Connection = conn
                rs.CommandType = CommandType.StoredProcedure
                rs.CommandText = sql
                Dim adapter As New MySqlDataAdapter(rs)

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
                ReportDataSource2.Name = "DataSet6"
                ReportDataSource2.Value = dt1
                ReportViewerData.LocalReport.DataSources.Add(ReportDataSource2)
                ReportViewerData.LocalReport.ReportPath = Application.StartupPath & "\Reports\TaT\" & rdlcName

                'Set Parameters for reporting
                Dim DateFrom As New Microsoft.Reporting.WinForms.ReportParameter("DateFrom", FormatDateTime(dtFrom.Value, DateFormat.ShortDate))
                Dim DateTo As New Microsoft.Reporting.WinForms.ReportParameter("DateTo", FormatDateTime(dtTo.Value, DateFormat.ShortDate))
                Dim Section As New Microsoft.Reporting.WinForms.ReportParameter("Section", cboLimit.Text)
                Dim Name As New Microsoft.Reporting.WinForms.ReportParameter("Name", cbomedtech.Text)
                Dim CurUser As New Microsoft.Reporting.WinForms.ReportParameter("CurrUser", CurrUser)
                ReportViewerData.LocalReport.SetParameters(DateFrom)
                ReportViewerData.LocalReport.SetParameters(DateTo)
                ReportViewerData.LocalReport.SetParameters(Section)
                ReportViewerData.LocalReport.SetParameters(Name)
                ReportViewerData.LocalReport.SetParameters(CurUser)

                'Set display format
                ReportViewerData.SetDisplayMode(Microsoft.Reporting.WinForms.DisplayMode.PrintLayout)

                'Refresh Report
                ReportViewerData.RefreshReport()
                Return (rdlcName)
            Else
                If cbomedtech.Text <> "All" And cboLimit.Text = "All" Then
                    Connect()
                    rs.Parameters.Clear()
                    rs.Parameters.AddWithValue("@DateFrom", Format(dtFrom.Value, "yyyy-MM-dd"))
                    rs.Parameters.AddWithValue("@DateTo", Format(dtTo.Value, "yyyy-MM-dd"))
                    rs.Parameters.AddWithValue("@MedTech", cbomedtech.Text)

                    Dim sql As String
                    sql = "SELECT
						*
				    FROM `viewtat`
					WHERE (DATE_FORMAT(`viewtat`.`ReleaseDateTime`, '%Y-%m-%d') BETWEEN @DateFrom AND @DateTo) AND `viewtat`.`PerformedBy` = @MedTech;"

                    Dim dt As DataTable = New DataTable
                    rs.Connection = conn
                    rs.CommandType = CommandType.Text
                    rs.CommandText = sql
                    Dim adapter As New MySqlDataAdapter(rs)

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
                    ReportDataSource2.Name = "DataSet6"
                    ReportDataSource2.Value = dt1
                    ReportViewerData.LocalReport.DataSources.Add(ReportDataSource2)
                    ReportViewerData.LocalReport.ReportPath = Application.StartupPath & "\Reports\TaT\" & rdlcName

                    'Set Parameters for reporting
                    Dim DateFrom As New Microsoft.Reporting.WinForms.ReportParameter("DateFrom", FormatDateTime(dtFrom.Value, DateFormat.ShortDate))
                    Dim DateTo As New Microsoft.Reporting.WinForms.ReportParameter("DateTo", FormatDateTime(dtTo.Value, DateFormat.ShortDate))
                    Dim Section As New Microsoft.Reporting.WinForms.ReportParameter("Section", cboLimit.Text)
                    Dim Name As New Microsoft.Reporting.WinForms.ReportParameter("Name", cbomedtech.Text)
                    Dim CurUser As New Microsoft.Reporting.WinForms.ReportParameter("CurrUser", CurrUser)
                    ReportViewerData.LocalReport.SetParameters(DateFrom)
                    ReportViewerData.LocalReport.SetParameters(DateTo)
                    ReportViewerData.LocalReport.SetParameters(Section)
                    ReportViewerData.LocalReport.SetParameters(Name)
                    ReportViewerData.LocalReport.SetParameters(CurUser)

                    'Set display format
                    ReportViewerData.SetDisplayMode(Microsoft.Reporting.WinForms.DisplayMode.PrintLayout)

                    'Refresh Report
                    ReportViewerData.RefreshReport()
                    Return (rdlcName)
                ElseIf cbomedtech.Text <> "All" And cboLimit.Text <> "All" Then
                    Connect()
                    rs.Parameters.Clear()
                    rs.Parameters.AddWithValue("@DateFrom", Format(dtFrom.Value, "yyyy-MM-dd"))
                    rs.Parameters.AddWithValue("@DateTo", Format(dtTo.Value, "yyyy-MM-dd"))
                    rs.Parameters.AddWithValue("@MedTech", cbomedtech.Text)
                    rs.Parameters.AddWithValue("@Section", cboLimit.Text)

                    Dim sql As String
                    sql = "SELECT
						*
				    FROM `viewtat`
					WHERE (DATE_FORMAT(`viewtat`.`ReleaseDateTime`, '%Y-%m-%d') BETWEEN @DateFrom AND @DateTo) AND `viewtat`.`PerformedBy` = @MedTech AND `viewtat`.`Section` = @Section;"

                    Dim dt As DataTable = New DataTable
                    rs.Connection = conn
                    rs.CommandType = CommandType.Text
                    rs.CommandText = sql
                    Dim adapter As New MySqlDataAdapter(rs)

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
                    Dim command1 As New MySql.Data.MySqlClient.MySqlCommand("Select * FROM `company_profile`", conn)

                    'Fill adapter set in Dataset2
                    Dim adapter1 As New MySql.Data.MySqlClient.MySqlDataAdapter(command1)
                    adapter1.Fill(dt1)
                    Disconnect()

                    'Set Report source data for Dataset2
                    Dim ReportDataSource2 As Microsoft.Reporting.WinForms.ReportDataSource = New Microsoft.Reporting.WinForms.ReportDataSource()
                    ReportDataSource2.Name = "DataSet6"
                    ReportDataSource2.Value = dt1
                    ReportViewerData.LocalReport.DataSources.Add(ReportDataSource2)
                    ReportViewerData.LocalReport.ReportPath = Application.StartupPath & "\Reports\TaT\" & rdlcName

                    'Set Parameters for reporting
                    Dim DateFrom As New Microsoft.Reporting.WinForms.ReportParameter("DateFrom", FormatDateTime(dtFrom.Value, DateFormat.ShortDate))
                    Dim DateTo As New Microsoft.Reporting.WinForms.ReportParameter("DateTo", FormatDateTime(dtTo.Value, DateFormat.ShortDate))
                    Dim Section As New Microsoft.Reporting.WinForms.ReportParameter("Section", cboLimit.Text)
                    Dim Name As New Microsoft.Reporting.WinForms.ReportParameter("Name", cbomedtech.Text)
                    Dim CurUser As New Microsoft.Reporting.WinForms.ReportParameter("CurrUser", CurrUser)
                    ReportViewerData.LocalReport.SetParameters(DateFrom)
                    ReportViewerData.LocalReport.SetParameters(DateTo)
                    ReportViewerData.LocalReport.SetParameters(Section)
                    ReportViewerData.LocalReport.SetParameters(Name)
                    ReportViewerData.LocalReport.SetParameters(CurUser)

                    'Set display format
                    ReportViewerData.SetDisplayMode(Microsoft.Reporting.WinForms.DisplayMode.PrintLayout)

                    'Refresh Report
                    ReportViewerData.RefreshReport()
                    Return (rdlcName)
                ElseIf cbomedtech.Text = "All" And cboLimit.Text <> "All" Then
                    Connect()
                    rs.Parameters.Clear()
                    rs.Parameters.AddWithValue("@DateFrom", Format(dtFrom.Value, "yyyy-MM-dd"))
                    rs.Parameters.AddWithValue("@DateTo", Format(dtTo.Value, "yyyy-MM-dd"))
                    rs.Parameters.AddWithValue("@Section", cboLimit.Text)

                    Dim sql As String
                    sql = "SELECT
						*
				    FROM `viewtat`
					WHERE (DATE_FORMAT(`viewtat`.`ReleaseDateTime`, '%Y-%m-%d') BETWEEN @DateFrom AND @DateTo) AND `viewtat`.`Section` = @Section;"

                    Dim dt As DataTable = New DataTable
                    rs.Connection = conn
                    rs.CommandType = CommandType.Text
                    rs.CommandText = sql
                    Dim adapter As New MySqlDataAdapter(rs)

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
                    Dim command1 As New MySql.Data.MySqlClient.MySqlCommand("Select * FROM `company_profile`", conn)

                    'Fill adapter set in Dataset2
                    Dim adapter1 As New MySql.Data.MySqlClient.MySqlDataAdapter(command1)
                    adapter1.Fill(dt1)
                    Disconnect()

                    'Set Report source data for Dataset2
                    Dim ReportDataSource2 As Microsoft.Reporting.WinForms.ReportDataSource = New Microsoft.Reporting.WinForms.ReportDataSource()
                    ReportDataSource2.Name = "DataSet6"
                    ReportDataSource2.Value = dt1
                    ReportViewerData.LocalReport.DataSources.Add(ReportDataSource2)
                    ReportViewerData.LocalReport.ReportPath = Application.StartupPath & "\Reports\TaT\" & rdlcName

                    'Set Parameters for reporting
                    Dim DateFrom As New Microsoft.Reporting.WinForms.ReportParameter("DateFrom", FormatDateTime(dtFrom.Value, DateFormat.ShortDate))
                    Dim DateTo As New Microsoft.Reporting.WinForms.ReportParameter("DateTo", FormatDateTime(dtTo.Value, DateFormat.ShortDate))
                    Dim Section As New Microsoft.Reporting.WinForms.ReportParameter("Section", cboLimit.Text)
                    Dim Name As New Microsoft.Reporting.WinForms.ReportParameter("Name", cbomedtech.Text)
                    Dim CurUser As New Microsoft.Reporting.WinForms.ReportParameter("CurrUser", CurrUser)
                    ReportViewerData.LocalReport.SetParameters(DateFrom)
                    ReportViewerData.LocalReport.SetParameters(DateTo)
                    ReportViewerData.LocalReport.SetParameters(Section)
                    ReportViewerData.LocalReport.SetParameters(Name)
                    ReportViewerData.LocalReport.SetParameters(CurUser)

                    'Set display format
                    ReportViewerData.SetDisplayMode(Microsoft.Reporting.WinForms.DisplayMode.PrintLayout)

                    'Refresh Report
                    ReportViewerData.RefreshReport()
                    Return (rdlcName)
                End If
            End If
        Catch ex As MySqlException
            Disconnect()
            MessageBox.Show(ex.Message, "Search Error", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
        Finally
            Disconnect()
        End Try

        Return (rdlcName)
    End Function
End Module
