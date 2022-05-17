Imports DevExpress.XtraGrid
Imports DevExpress.XtraGrid.Views.Grid
Imports MySql.Data.MySqlClient

Public Class frmCensus

    Public Sub AutoLoadTestName()
        cboLimit.Properties.Items.Clear()
        Connect()
        rs.Connection = conn
        rs.CommandType = CommandType.Text
        rs.CommandText = "SELECT `test_name` FROM `testtype` WHERE `test_name` NOT LIKE 'All' ORDER BY `test_name`"
        reader = rs.ExecuteReader
        While reader.Read
            cboLimit.Properties.Items.Add(reader(0).ToString)
        End While
        Disconnect()

        cboLimit.SelectedIndex = 0
    End Sub

    Private Sub frmWorkListHema_FormClosed(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
        Me.Dispose()
    End Sub

    Private Sub btnPrint_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnPrint.ItemClick
        'If cboLimit.Text = "Hematology" Then
        '    GenerateReportCensus("census_" & cboMachine.Text, dtFrom, dtTo, cboMachine, "ReportCensusHema.rdlc", RPTCensus.ReportViewer1)
        'ElseIf cboLimit.Text = "Chemistry" Then
        '    GenerateReportCensus("census_" & cboMachine.Text, dtFrom, dtTo, cboMachine, "ReportCensusChem.rdlc", RPTCensus.ReportViewer1)
        'End If
        'RPTCensus.ShowDialog()
    End Sub

    Private Sub btnClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnClose.ItemClick
        Me.Close()
    End Sub

    Private Sub btnSearch_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSearch.Click

        Dim sql As String
        'sql = "SELECT * FROM `" & "census" & "_" & cboMachine.Text & "` WHERE (`date` BETWEEN @d1 AND @d2)"
        sql = "Census_TestPerformed"

        GridView1.Columns.Clear()
        GridView1.Appearance.HeaderPanel.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center
        GridView1.Appearance.HeaderPanel.FontStyleDelta = FontStyle.Bold

        GridView1.Appearance.OddRow.BackColor = Color.FromArgb(226, 236, 247)
        GridView1.OptionsView.EnableAppearanceOddRow = True
        GridView1.Appearance.EvenRow.BackColor = Color.White
        GridView1.OptionsView.EnableAppearanceEvenRow = True

        Dim dt As DataTable = New DataTable
        Dim command As New MySqlCommand()

        Connect()
        command.Connection = conn
        command.CommandType = CommandType.StoredProcedure
        command.CommandText = sql

        command.Parameters.Clear()
        command.Parameters.Add("@DateFrom", MySqlDbType.Date).Value = Format(dtFrom.Value, "yyyy-MM-dd")
        command.Parameters.Add("@DateTo", MySqlDbType.Date).Value = Format(dtTo.Value, "yyyy-MM-dd")
        command.Parameters.AddWithValue("@Section", cboLimit.Text)

        Dim adapter As New MySqlDataAdapter(command)
        adapter.Fill(dt)
        dtResult.DataSource = (dt)
        Disconnect()

        GridView1.OptionsBehavior.Editable = False
        GridView1.OptionsSelection.EnableAppearanceFocusedCell = False
        GridView1.FocusRectStyle = DrawFocusRectStyle.RowFullFocus

        LoadTaTFrequency()
        LoadCensusTest()
        LoadPanelTest()
    End Sub

    Private Sub frmWorkSheet_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        AutoLoadTestName()
    End Sub

    Private Sub LoadTaTFrequency()
        Dim Count_Census As Double
        Try
            Chart1.Titles.Clear()
            Chart1.Titles.Add("Daily Census")
            Chart1.Titles(0).Font = New Font("Tahoma", 14, FontStyle.Bold)

            Chart1.Series.Clear()
            Chart1.Series.Add(cboLimit.Text)
            Chart1.Series(cboLimit.Text).MarkerStyle = DataVisualization.Charting.MarkerStyle.Circle
            Chart1.Series(cboLimit.Text).MarkerSize = 5
            Chart1.Series(cboLimit.Text).MarkerColor = Color.Black
            Chart1.Series(cboLimit.Text).IsValueShownAsLabel = True
            Chart1.Series(cboLimit.Text).BorderWidth = 3
            Chart1.Series(cboLimit.Text).Color = Color.SkyBlue
            Chart1.Series(cboLimit.Text).IsXValueIndexed = True
            Chart1.Series(cboLimit.Text).ChartType = System.Windows.Forms.DataVisualization.Charting.SeriesChartType.Column
            Chart1.Series(cboLimit.Text).Font = New Font("Tahoma", 8, FontStyle.Bold)

            Connect()
            rs.Parameters.Clear()
            rs.Parameters.Add("@DateFrom", MySqlDbType.Date).Value = Format(dtFrom.Value, "yyyy-MM-dd")
            rs.Parameters.Add("@DateTo", MySqlDbType.Date).Value = Format(dtTo.Value, "yyyy-MM-dd")
            rs.Parameters.AddWithValue("@Section", cboLimit.Text)

            rs.Connection = conn
            rs.CommandType = CommandType.StoredProcedure
            rs.CommandText = "Census_TestPerformed"
            reader = rs.ExecuteReader
            While reader.Read()
                Chart1.Series(cboLimit.Text).Points.AddXY(reader(2).ToString, reader(0).ToString)
            End While
            Disconnect()

            Chart1.Series(cboLimit.Text).ChartArea = "ChartArea1"

            Chart1.ChartAreas(0).AxisX.Minimum = 0
            Chart1.ChartAreas(0).AxisX.Interval = 1

            Chart1.ChartAreas(0).AxisX.MajorGrid.LineDashStyle = DataVisualization.Charting.ChartDashStyle.NotSet
            Chart1.ChartAreas(0).AxisY.MajorGrid.LineDashStyle = DataVisualization.Charting.ChartDashStyle.NotSet

            Chart1.ChartAreas(0).AxisX.LabelStyle.Angle = -45
        Catch ex As Exception
            Exit Sub
            MessageBox.Show(ex.Message, "Exception", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
        End Try
    End Sub

    Private Sub LoadCensusTest()
        Dim sql As String

        sql = "Census_Test"

        GridCensusTest.Columns.Clear()
        GridCensusTest.Appearance.HeaderPanel.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center
        GridCensusTest.Appearance.HeaderPanel.FontStyleDelta = FontStyle.Bold

        GridCensusTest.Appearance.OddRow.BackColor = Color.FromArgb(226, 236, 247)
        GridCensusTest.OptionsView.EnableAppearanceOddRow = True
        GridCensusTest.Appearance.EvenRow.BackColor = Color.White
        GridCensusTest.OptionsView.EnableAppearanceEvenRow = True

        Dim dt As DataTable = New DataTable
        Dim command As New MySqlCommand()

        Connect()
        command.Connection = conn
        command.CommandType = CommandType.StoredProcedure
        command.CommandText = sql

        command.Parameters.Clear()
        command.Parameters.AddWithValue("@Section", cboLimit.Text)
        command.Parameters.Add("@DateFrom", MySqlDbType.Date).Value = Format(dtFrom.Value, "yyyy-MM-dd")
        command.Parameters.Add("@DateTo", MySqlDbType.Date).Value = Format(dtTo.Value, "yyyy-MM-dd")

        Dim adapter As New MySqlDataAdapter(command)
        adapter.Fill(dt)
        dtCensusTest.DataSource = (dt)
        Disconnect()

        GridCensusTest.OptionsBehavior.Editable = False
        GridCensusTest.OptionsSelection.EnableAppearanceFocusedCell = False
        GridCensusTest.FocusRectStyle = DrawFocusRectStyle.RowFullFocus
    End Sub

    'Private Sub btnGraph_ItemClick(ByVal sender As Object, ByVal e As DevExpress.XtraBars.ItemClickEventArgs) Handles btnGraph.ItemClick
    '    If cboLimit.Text = "Hematology" Then
    '        GenerateReportCensus("census_" & cboMachine.Text, dtFrom, dtTo, cboMachine, "ReportCensusHemaChart.rdlc", RPTCensusChart.ReportViewer1)
    '    ElseIf cboLimit.Text = "Chemistry" Then
    '        GenerateReportCensus("census_" & cboMachine.Text, dtFrom, dtTo, cboMachine, "ReportCensusChemChart.rdlc", RPTCensusChart.ReportViewer1)
    '    End If
    '    RPTCensusChart.ShowDialog()
    'End Sub

    Private Sub LoadPanelTest()
        Dim sql As String

        sql = "Census_TestPanel"

        GridPanel.Columns.Clear()
        GridPanel.Appearance.HeaderPanel.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center
        GridPanel.Appearance.HeaderPanel.FontStyleDelta = FontStyle.Bold

        GridPanel.Appearance.OddRow.BackColor = Color.FromArgb(226, 236, 247)
        GridPanel.OptionsView.EnableAppearanceOddRow = True
        GridPanel.Appearance.EvenRow.BackColor = Color.White
        GridPanel.OptionsView.EnableAppearanceEvenRow = True

        Dim dt As DataTable = New DataTable
        Dim command As New MySqlCommand()

        Connect()
        command.Connection = conn
        command.CommandType = CommandType.StoredProcedure
        command.CommandText = sql

        command.Parameters.Clear()
        command.Parameters.AddWithValue("@Section", cboLimit.Text)
        command.Parameters.Add("@DateFrom", MySqlDbType.Date).Value = Format(dtFrom.Value, "yyyy-MM-dd")
        command.Parameters.Add("@DateTo", MySqlDbType.Date).Value = Format(dtTo.Value, "yyyy-MM-dd")

        Dim adapter As New MySqlDataAdapter(command)
        adapter.Fill(dt)
        dtPanel.DataSource = (dt)
        Disconnect()

        GridPanel.OptionsBehavior.Editable = False
        GridPanel.OptionsSelection.EnableAppearanceFocusedCell = False
        GridPanel.FocusRectStyle = DrawFocusRectStyle.RowFullFocus
    End Sub
End Class