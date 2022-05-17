Imports MySql.Data.MySqlClient

Public Class frmCensusGraph

    Private Sub LoadTaT()
        Chart1.Series.Clear()

        Dim LL, UL, Process, Count As Double
        Dim TargetLL, TargetUL, TargetProcess As Integer

        Connect()
        rs.Connection = conn
        rs.CommandText = "SELECT COUNT(time_diff) as `time` FROM `viewTimeDiff` WHERE `time_diff` <= 7600 AND `status` = 'Ordered'"
        reader = rs.ExecuteReader
        reader.Read()
        LL = Val(reader(0))
        Disconnect()

        Connect()
        rs.Connection = conn
        rs.CommandText = "SELECT COUNT(time_diff) as `time` FROM `viewTimeDiff` WHERE `time_diff` < 7600 AND `status` NOT LIKE 'Ordered'"
        reader = rs.ExecuteReader
        reader.Read()
        Process = Val(reader(0))
        Disconnect()

        Connect()
        rs.Connection = conn
        rs.CommandText = "SELECT COUNT(time_diff) as `time` FROM `viewTimeDiff` WHERE `time_diff` > 7600"
        reader = rs.ExecuteReader
        reader.Read()
        UL = Val(reader(0))
        Disconnect()

        Connect()
        rs.Connection = conn
        rs.CommandText = "SELECT COUNT(time_diff) as `time` FROM `viewTimeDiff`"
        reader = rs.ExecuteReader
        reader.Read()
        Count = Val(reader(0))
        Disconnect()

        Chart1.Titles.Clear()
        Chart1.Titles.Add("Turn Arround Task")
        Chart1.Titles(0).Font = New Font("Tahoma", 14, FontStyle.Bold)

        Chart1.Titles.Add("Hematology Section")
        Chart1.Titles(1).Font = New Font("Tahoma", 12, FontStyle.Bold)

        'Dim ca1 = Chart1.ChartAreas(0)
        'Dim rip1 As RectangleF = ca1.InnerPlotPosition.ToRectangleF()
        'Dim ty1 As title = Chart1.Titles.Add("ty1")
        'ty1.Text = "Y-Axis 1\nTitle"
        'ty1.ForeColor = Color.DarkSlateBlue
        'ty1.Position.X = rip1.Left
        'ty1.Position.Y = rip1.Y

        Chart1.Series.Add("LL")
        Chart1.Series("LL").IsVisibleInLegend = False
        Chart1.Series("LL").BorderWidth = 3
        Chart1.Series("LL").Color = Color.DarkGreen
        Chart1.Series("LL").ChartType = System.Windows.Forms.DataVisualization.Charting.SeriesChartType.Column
        Chart1.Series("LL").Font = New Font("Tahoma", 10, FontStyle.Bold)

        Chart1.Series.Add("Process")
        Chart1.Series("Process").IsVisibleInLegend = False
        Chart1.Series("Process").Color = Color.Pink
        Chart1.Series("Process").ChartType = System.Windows.Forms.DataVisualization.Charting.SeriesChartType.Column
        Chart1.Series("Process").Font = New Font("Tahoma", 10, FontStyle.Bold)

        Chart1.Series.Add("UL")
        Chart1.Series("UL").IsVisibleInLegend = False
        Chart1.Series("UL").BorderWidth = 3
        Chart1.Series("UL").Color = Color.Crimson
        Chart1.Series("UL").ChartType = System.Windows.Forms.DataVisualization.Charting.SeriesChartType.Column
        Chart1.Series("UL").Font = New Font("Tahoma", 10, FontStyle.Bold)

        Chart1.Series("LL").ChartArea = "ChartArea1"
        Chart1.Series("UL").ChartArea = "ChartArea1"

        TargetLL = Chart1.Series("LL").Points.AddXY(1, LL)
        Chart1.Series("LL").Points.Item(TargetLL).Label = "Within 2hrs. (" & LL & ")"

        TargetProcess = Chart1.Series("Process").Points.AddXY(1, Process)
        Chart1.Series("Process").Points.Item(TargetProcess).Label = "Processing (" & Process & ")"

        TargetUL = Chart1.Series("UL").Points.AddXY(1, UL)
        Chart1.Series("UL").Points.Item(TargetUL).Label = "More than 2hrs. (" & UL & ")"

        Chart1.ChartAreas(0).AxisX.Interval = 1
        Chart1.ChartAreas(0).AxisY.Interval = 1
        Chart1.ChartAreas(0).AxisY.Maximum = Count

        Chart1.ChartAreas(0).AxisX.LabelStyle.Enabled = False
        Chart1.ChartAreas(0).AxisX.MajorGrid.LineDashStyle = DataVisualization.Charting.ChartDashStyle.NotSet
        Chart1.ChartAreas(0).AxisY.MajorGrid.LineDashStyle = DataVisualization.Charting.ChartDashStyle.NotSet

    End Sub

    Private Sub btnClose_ItemClick(ByVal sender As Object, ByVal e As DevExpress.XtraBars.ItemClickEventArgs) Handles btnClose.ItemClick
        Me.Close()
        Me.Dispose()
    End Sub

    Private Sub Timer1_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer1.Tick
        LoadTaT()
    End Sub
End Class