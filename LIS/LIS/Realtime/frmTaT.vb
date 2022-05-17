Imports System.Windows.Forms.DataVisualization.Charting
Imports DevExpress.Utils
Imports DevExpress.XtraGrid.Views.Grid
Imports MySql.Data.MySqlClient

Public Class frmTaT

    Dim DelayedCount As Integer = 0
    Dim OnTimeCount As Integer = 0

    Private Sub btnClose_ItemClick(ByVal sender As Object, ByVal e As DevExpress.XtraBars.ItemClickEventArgs) Handles btnClose.ItemClick
        Me.Close()
        Me.Dispose()
    End Sub

    'Private Sub Timer1_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer1.Tick
    '    Try
    '        LoadTaT()
    '        LoadTaTRatio()
    '        LoadTaTFrequency()
    '        LoadStats()
    '        LoadStats()
    '    Catch ex As Exception
    '        Disconnect()
    '        Timer1.Stop()
    '        MessageBox.Show(ex.Message, "System Error", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
    '        frmDBUtils.ShowDialog()
    '    End Try
    'End Sub

    Private Sub LoadTaT()
        ChartTaT.Series.Clear()

        Dim monthNames As String() = {"Processing", "Completed"}
        Dim startOffset As Integer = 0
        Dim endOffset As Integer = 2

        For Each MonthName As String In monthNames
            ChartTaT.ChartAreas(0).AxisX.LabelStyle.Enabled = True
            Dim monthLabel = New CustomLabel(startOffset, endOffset, MonthName, 0, LabelMarkStyle.None)
            ChartTaT.ChartAreas(0).AxisX.CustomLabels.Add(monthLabel)
            startOffset = startOffset + 1
            endOffset = endOffset + 1
        Next

        Dim LL, UL, Process, Count, CountReleased, Released, Released_Over As Double
        Dim TargetLL, TargetUL, TargetProcess, TargetReleased, TargetReleased_Over As Integer

        Connect()
        rs.Connection = conn
        rs.CommandType = CommandType.StoredProcedure
        rs.CommandText = "tat_ongoing"
        rs.Parameters.Clear()
        rs.Parameters.AddWithValue("@dDate", Format(Now, "yyyy-MM-dd"))
        rs.Parameters.AddWithValue("@Stat", "Checked-In")
        reader = rs.ExecuteReader
        reader.Read()
        LL = Val(reader(0))
        Disconnect()

        Connect()
        rs.Connection = conn
        rs.CommandType = CommandType.StoredProcedure
        rs.CommandText = "tat_processing"
        rs.Parameters.Clear()
        rs.Parameters.AddWithValue("@dDate", Format(Now, "yyyy-MM-dd"))
        rs.Parameters.AddWithValue("@Stat", "Checked-In")
        reader = rs.ExecuteReader
        reader.Read()
        Process = Val(reader(0))
        Disconnect()

        Connect()
        rs.Connection = conn
        rs.CommandType = CommandType.StoredProcedure
        rs.CommandText = "tat_over"
        rs.Parameters.Clear()
        rs.Parameters.AddWithValue("@dDate", Format(Now, "yyyy-MM-dd"))
        reader = rs.ExecuteReader
        reader.Read()
        UL = Val(reader(0))
        Disconnect()

        Connect()
        rs.Connection = conn
        rs.CommandType = CommandType.StoredProcedure
        rs.CommandText = "tat_completed_within"
        rs.Parameters.Clear()
        rs.Parameters.AddWithValue("@dDate", Format(Now, "yyyy-MM-dd"))
        reader = rs.ExecuteReader
        reader.Read()
        Released = Val(reader(0))
        Disconnect()

        Connect()
        rs.Connection = conn
        rs.CommandType = CommandType.StoredProcedure
        rs.CommandText = "tat_completed_over"
        rs.Parameters.Clear()
        rs.Parameters.AddWithValue("@dDate", Format(Now, "yyyy-MM-dd"))
        reader = rs.ExecuteReader
        reader.Read()
        Released_Over = Val(reader(0))
        Disconnect()

        Connect1()
        rs1.Connection = conn1
        rs1.CommandText = "SELECT COUNT(time_diff) as `time` FROM `viewTimeDiff`"
        reader1 = rs1.ExecuteReader
        reader1.Read()
        Count = Val(reader1(0))
        Disconnect1()

        Connect1()
        rs1.Connection = conn1
        rs1.CommandText = "SELECT COUNT(time_diff) as `time` FROM `viewTimeDiff_Completed`"
        reader1 = rs1.ExecuteReader
        reader1.Read()
        CountReleased = Val(reader1(0))
        Disconnect1()

        ChartTaT.Titles.Clear()
        ChartTaT.Titles.Add("Turn Around Time")
        ChartTaT.Titles(0).Font = New Font("Tahoma", 14, FontStyle.Bold)

        ChartTaT.Series.Add("LL")
        ChartTaT.Series("LL")("PixelPointWidth") = 800
        ChartTaT.Series("LL").IsVisibleInLegend = False
        ChartTaT.Series("LL").BorderWidth = 3
        ChartTaT.Series("LL").Color = Color.DarkGreen
        ChartTaT.Series("LL").ChartType = System.Windows.Forms.DataVisualization.Charting.SeriesChartType.Column
        ChartTaT.Series("LL").Font = New Font("Tahoma", 10, FontStyle.Bold)

        ChartTaT.Series.Add("Process")
        ChartTaT.Series("Process")("PixelPointWidth") = 800
        ChartTaT.Series("Process").IsVisibleInLegend = False
        ChartTaT.Series("Process").Color = Color.Pink
        ChartTaT.Series("Process").ChartType = System.Windows.Forms.DataVisualization.Charting.SeriesChartType.Column
        ChartTaT.Series("Process").Font = New Font("Tahoma", 10, FontStyle.Bold)

        ChartTaT.Series.Add("UL")
        ChartTaT.Series("UL")("PixelPointWidth") = 800
        ChartTaT.Series("UL").IsVisibleInLegend = False
        ChartTaT.Series("UL").BorderWidth = 3
        ChartTaT.Series("UL").Color = Color.Crimson
        ChartTaT.Series("UL").ChartType = System.Windows.Forms.DataVisualization.Charting.SeriesChartType.Column
        ChartTaT.Series("UL").Font = New Font("Tahoma", 10, FontStyle.Bold)

        ChartTaT.Series.Add("Released")
        ChartTaT.Series("Released")("PixelPointWidth") = 800
        ChartTaT.Series("Released").IsVisibleInLegend = False
        ChartTaT.Series("Released").Color = Color.Green
        ChartTaT.Series("Released").ChartType = System.Windows.Forms.DataVisualization.Charting.SeriesChartType.Column
        ChartTaT.Series("Released").Font = New Font("Tahoma", 10, FontStyle.Bold)

        ChartTaT.Series.Add("Released_Over")
        ChartTaT.Series("Released_Over")("PixelPointWidth") = 800
        ChartTaT.Series("Released_Over").IsVisibleInLegend = False
        ChartTaT.Series("Released_Over").Color = Color.Crimson
        ChartTaT.Series("Released_Over").ChartType = System.Windows.Forms.DataVisualization.Charting.SeriesChartType.Column
        ChartTaT.Series("Released_Over").Font = New Font("Tahoma", 10, FontStyle.Bold)

        ChartTaT.Series("LL").ChartArea = "ChartArea1"
        ChartTaT.Series("UL").ChartArea = "ChartArea1"

        TargetLL = ChartTaT.Series("LL").Points.AddXY(1, LL)
        ChartTaT.Series("LL").Points.Item(TargetLL).Label = "Within TaT (" & LL & ")"

        TargetProcess = ChartTaT.Series("Process").Points.AddXY(1, Process)
        ChartTaT.Series("Process").Points.Item(TargetProcess).Label = "Processing (" & Process & ")"

        TargetUL = ChartTaT.Series("UL").Points.AddXY(1, UL)
        ChartTaT.Series("UL").Points.Item(TargetUL).Label = "Over TaT (" & UL & ")"

        TargetReleased = ChartTaT.Series("Released").Points.AddXY(1.5, Released)
        ChartTaT.Series("Released").Points.Item(TargetReleased).Label = "Released Within TaT (" & Released & ")"

        TargetReleased_Over = ChartTaT.Series("Released_Over").Points.AddXY(1.5, Released_Over)
        ChartTaT.Series("Released_Over").Points.Item(TargetReleased_Over).Label = "Released Over TaT (" & Released_Over & ")"

        ChartTaT.ChartAreas(0).AxisX.Interval = 1
        ChartTaT.ChartAreas(0).AxisY.Interval = 1
        'ChartTaT.ChartAreas(0).AxisY.Maximum = Count + CountReleased

        ChartTaT.ChartAreas(0).AxisX.MajorGrid.LineDashStyle = DataVisualization.Charting.ChartDashStyle.NotSet
        ChartTaT.ChartAreas(0).AxisY.MajorGrid.LineDashStyle = DataVisualization.Charting.ChartDashStyle.NotSet

    End Sub

    Private Sub LoadTaTRatio()
        ChartRatio.Series.Clear()

        Dim LL, UL As Double
        Dim TargetLL, TargetUL As Integer

        ChartRatio.Titles.Clear()
        ChartRatio.Titles.Add("Out Patient To In Patient Ratio")
        ChartRatio.Titles(0).Font = New Font("Tahoma", 14, FontStyle.Bold)

        ChartRatio.Titles.Add(Format(Now, "MM/dd/yyyy"))
        ChartRatio.Titles(1).Font = New Font("Tahoma", 12, FontStyle.Bold)

        ChartRatio.Series.Add("LL")
        ChartRatio.Series("LL")("PixelPointWidth") = 500
        ChartRatio.Series("LL").IsVisibleInLegend = False
        ChartRatio.Series("LL").BorderWidth = 3
        ChartRatio.Series("LL").Color = Color.DarkGreen
        ChartRatio.Series("LL").ChartType = System.Windows.Forms.DataVisualization.Charting.SeriesChartType.Bar
        ChartRatio.Series("LL").Font = New Font("Tahoma", 10, FontStyle.Bold)

        ChartRatio.Series.Add("UL")
        ChartRatio.Series("UL")("PixelPointWidth") = 500
        ChartRatio.Series("UL").IsVisibleInLegend = False
        ChartRatio.Series("UL").BorderWidth = 3
        ChartRatio.Series("UL").Color = Color.Crimson
        ChartRatio.Series("UL").ChartType = System.Windows.Forms.DataVisualization.Charting.SeriesChartType.Bar
        ChartRatio.Series("UL").Font = New Font("Tahoma", 10, FontStyle.Bold)


        ChartRatio.Series("LL").ChartArea = "ChartArea1"
        ChartRatio.Series("UL").ChartArea = "ChartArea1"

        Connect()
        rs.Parameters.Clear()
        rs.Parameters.AddWithValue("@dDate", Format(Now, "yyyy-MM-dd"))
        rs.Parameters.AddWithValue("@PType", "Out Patient")

        rs.Connection = conn
        rs.CommandType = CommandType.StoredProcedure
        rs.CommandText = "tat_ratio_processing"
        reader = rs.ExecuteReader
        While reader.Read()
            LL = Val(reader(0))

            TargetLL = ChartRatio.Series("LL").Points.AddXY(1, LL)
            ChartRatio.Series("LL").Points.Item(TargetLL).Label = "Out Patient (" & LL & ")"
        End While
        Disconnect()

        Connect()
        rs.Parameters.Clear()
        rs.Parameters.AddWithValue("@dDate", Format(Now, "yyyy-MM-dd"))
        rs.Parameters.AddWithValue("@PType", "In Patient")

        rs.Connection = conn
        rs.CommandType = CommandType.StoredProcedure
        rs.CommandText = "tat_ratio_processing"
        reader = rs.ExecuteReader
        While reader.Read()
            UL = Val(reader(0))
            TargetUL = ChartRatio.Series("UL").Points.AddXY(1, UL)
            ChartRatio.Series("UL").Points.Item(TargetUL).Label = "In Patient (" & UL & ")"
        End While
        Disconnect()

        ChartRatio.ChartAreas(0).AxisX.Interval = 1
        ChartRatio.ChartAreas(0).AxisY.Interval = 1
        ChartRatio.ChartAreas(0).AxisY.Maximum = LL + UL

        ChartRatio.ChartAreas(0).AxisX.LabelStyle.Enabled = False
        ChartRatio.ChartAreas(0).AxisX.MajorGrid.LineDashStyle = DataVisualization.Charting.ChartDashStyle.NotSet
        ChartRatio.ChartAreas(0).AxisY.MajorGrid.LineDashStyle = DataVisualization.Charting.ChartDashStyle.NotSet

    End Sub

    Private Sub LoadTaTFrequency()
        Dim Count_Within, Count_Over, Total As Double
        Try
            Try
                Connect()
                rs.Parameters.Clear()
                rs.Parameters.AddWithValue("@m", Format(Now, "MM"))
                rs.Connection = conn
                rs.CommandType = CommandType.StoredProcedure
                rs.CommandText = "TaT_Frequency_Within"
                reader = rs.ExecuteReader
                reader.Read()
                Count_Within = Val(reader(0))
                Disconnect()

                Connect()
                rs.Parameters.Clear()
                rs.Parameters.AddWithValue("@m", Format(Now, "MM"))
                rs.Connection = conn
                rs.CommandType = CommandType.StoredProcedure
                rs.CommandText = "TaT_Frequency_Over"
                reader = rs.ExecuteReader
                reader.Read()
                Count_Over = Val(reader(0))
                Disconnect()

                Total = Count_Within + Count_Over

            Catch ex As Exception
                Disconnect()
                ChartTaTFrequency.Titles.Clear()
                ChartTaTFrequency.Series.Clear()
            End Try


            ChartTaTFrequency.Titles.Clear()
            ChartTaTFrequency.Titles.Add("Laboratory TaT Frequency")
            ChartTaTFrequency.Titles.Add("Late vs On Time Batches")
            ChartTaTFrequency.Titles(0).Font = New Font("Tahoma", 14, FontStyle.Bold)
            ChartTaTFrequency.Titles(1).Font = New Font("Tahoma", 14, FontStyle.Bold)

            ChartTaTFrequency.Series.Clear()
            ChartTaTFrequency.Series.Add("Within TaT")
            ChartTaTFrequency.Series("Within TaT").MarkerStyle = DataVisualization.Charting.MarkerStyle.Circle
            ChartTaTFrequency.Series("Within TaT").MarkerSize = 5
            ChartTaTFrequency.Series("Within TaT").MarkerColor = Color.Black
            ChartTaTFrequency.Series("Within TaT").IsValueShownAsLabel = True
            ChartTaTFrequency.Series("Within TaT").BorderWidth = 3
            ChartTaTFrequency.Series("Within TaT").Color = Color.Green
            ChartTaTFrequency.Series("Within TaT").IsXValueIndexed = False
            ChartTaTFrequency.Series("Within TaT").ChartType = System.Windows.Forms.DataVisualization.Charting.SeriesChartType.Line
            ChartTaTFrequency.Series("Within TaT").Font = New Font("Tahoma", 8, FontStyle.Bold)

            ChartTaTFrequency.Series.Add("Over TaT")
            ChartTaTFrequency.Series("Over TaT").MarkerStyle = DataVisualization.Charting.MarkerStyle.Circle
            ChartTaTFrequency.Series("Over TaT").MarkerSize = 5
            ChartTaTFrequency.Series("Over TaT").MarkerColor = Color.Black
            ChartTaTFrequency.Series("Over TaT").IsValueShownAsLabel = True
            ChartTaTFrequency.Series("Over TaT").BorderWidth = 3
            ChartTaTFrequency.Series("Over TaT").Color = Color.Crimson
            ChartTaTFrequency.Series("Over TaT").IsXValueIndexed = False
            ChartTaTFrequency.Series("Over TaT").ChartType = System.Windows.Forms.DataVisualization.Charting.SeriesChartType.Line
            ChartTaTFrequency.Series("Over TaT").Font = New Font("Tahoma", 8, FontStyle.Bold)

            Connect()
            rs.Connection = conn
            rs.CommandType = CommandType.StoredProcedure
            rs.CommandText = "TaT_Frequency_Within"
            reader = rs.ExecuteReader
            While reader.Read
                ChartTaTFrequency.Series("Within TaT").Points.AddXY(reader(1).ToString, reader(0).ToString)
            End While
            Disconnect()

            Connect()
            rs.Connection = conn
            rs.CommandType = CommandType.StoredProcedure
            rs.CommandText = "TaT_Frequency_Over"
            reader = rs.ExecuteReader
            While reader.Read
                ChartTaTFrequency.Series("Over TaT").Points.AddXY(reader(1).ToString, reader(0).ToString)
            End While
            Disconnect()

            ChartTaTFrequency.Series("Within TaT").ChartArea = "ChartArea1"
            ChartTaTFrequency.Series("Over TaT").ChartArea = "ChartArea1"

            ChartTaTFrequency.ChartAreas(0).AxisX.Minimum = 0
            ChartTaTFrequency.ChartAreas(0).AxisX.Interval = 1

            ChartTaTFrequency.ChartAreas(0).AxisX.MajorGrid.LineDashStyle = DataVisualization.Charting.ChartDashStyle.NotSet
            ChartTaTFrequency.ChartAreas(0).AxisY.MajorGrid.LineDashStyle = DataVisualization.Charting.ChartDashStyle.NotSet

            ChartTaTFrequency.ChartAreas(0).AxisX.LabelStyle.Angle = -45
        Catch ex As Exception
            Exit Sub
            MessageBox.Show(ex.Message, "Exception", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
        End Try
    End Sub

    Private Sub LoadTaTFrequencyAll()
        Dim Count_Within, Count_Over, Total As Double
        Try
            Try
                Connect()
                rs.Parameters.Clear()
                rs.Parameters.AddWithValue("@m", Format(Now, "MM"))
                rs.Connection = conn
                rs.CommandType = CommandType.StoredProcedure
                rs.CommandText = "TaT_Frequency_Within"
                reader = rs.ExecuteReader
                reader.Read()
                Count_Within = Val(reader(0))
                Disconnect()

                Connect()
                rs.Parameters.Clear()
                rs.Parameters.AddWithValue("@m", Format(Now, "MM"))
                rs.Connection = conn
                rs.CommandType = CommandType.StoredProcedure
                rs.CommandText = "TaT_Frequency_Over"
                reader = rs.ExecuteReader
                reader.Read()
                Count_Over = Val(reader(0))
                Disconnect()

                Total = Count_Within + Count_Over

            Catch ex As Exception
                Disconnect()
                ChartTaTFrequencyAll.Titles.Clear()
                ChartTaTFrequencyAll.Series.Clear()
            End Try


            ChartTaTFrequencyAll.Titles.Clear()
            ChartTaTFrequencyAll.Titles.Add("Laboratory TaT Frequency")
            ChartTaTFrequencyAll.Titles.Add("Late vs On Time Batches")
            ChartTaTFrequencyAll.Titles(0).Font = New Font("Tahoma", 14, FontStyle.Bold)
            ChartTaTFrequencyAll.Titles(1).Font = New Font("Tahoma", 14, FontStyle.Bold)

            ChartTaTFrequencyAll.Series.Clear()
            ChartTaTFrequencyAll.Series.Add("Within TaT")
            ChartTaTFrequencyAll.Series("Within TaT").MarkerStyle = DataVisualization.Charting.MarkerStyle.Circle
            ChartTaTFrequencyAll.Series("Within TaT").MarkerSize = 5
            ChartTaTFrequencyAll.Series("Within TaT").MarkerColor = Color.Black
            ChartTaTFrequencyAll.Series("Within TaT").IsValueShownAsLabel = True
            ChartTaTFrequencyAll.Series("Within TaT").BorderWidth = 3
            ChartTaTFrequencyAll.Series("Within TaT").Color = Color.Green
            ChartTaTFrequencyAll.Series("Within TaT").IsXValueIndexed = False
            ChartTaTFrequencyAll.Series("Within TaT").ChartType = System.Windows.Forms.DataVisualization.Charting.SeriesChartType.Line
            ChartTaTFrequencyAll.Series("Within TaT").Font = New Font("Tahoma", 8, FontStyle.Bold)

            ChartTaTFrequencyAll.Series.Add("Over TaT")
            ChartTaTFrequencyAll.Series("Over TaT").MarkerStyle = DataVisualization.Charting.MarkerStyle.Circle
            ChartTaTFrequencyAll.Series("Over TaT").MarkerSize = 5
            ChartTaTFrequencyAll.Series("Over TaT").MarkerColor = Color.Black
            ChartTaTFrequencyAll.Series("Over TaT").IsValueShownAsLabel = True
            ChartTaTFrequencyAll.Series("Over TaT").BorderWidth = 3
            ChartTaTFrequencyAll.Series("Over TaT").Color = Color.Crimson
            ChartTaTFrequencyAll.Series("Over TaT").IsXValueIndexed = False
            ChartTaTFrequencyAll.Series("Over TaT").ChartType = System.Windows.Forms.DataVisualization.Charting.SeriesChartType.Line
            ChartTaTFrequencyAll.Series("Over TaT").Font = New Font("Tahoma", 8, FontStyle.Bold)

            Connect()
            rs.Connection = conn
            rs.CommandType = CommandType.StoredProcedure
            rs.CommandText = "TaT_Frequency_Within"
            reader = rs.ExecuteReader
            While reader.Read
                ChartTaTFrequencyAll.Series("Within TaT").Points.AddXY(reader(1).ToString, reader(0).ToString)
            End While
            Disconnect()

            Connect()
            rs.Connection = conn
            rs.CommandType = CommandType.StoredProcedure
            rs.CommandText = "TaT_Frequency_Over"
            reader = rs.ExecuteReader
            While reader.Read
                ChartTaTFrequencyAll.Series("Over TaT").Points.AddXY(reader(1).ToString, reader(0).ToString)
            End While
            Disconnect()

            ChartTaTFrequencyAll.Series("Within TaT").ChartArea = "ChartArea1"
            ChartTaTFrequencyAll.Series("Over TaT").ChartArea = "ChartArea1"

            ChartTaTFrequencyAll.ChartAreas(0).AxisX.Minimum = 0
            ChartTaTFrequencyAll.ChartAreas(0).AxisX.Interval = 1

            ChartTaTFrequencyAll.ChartAreas(0).AxisX.IsMarginVisible = False

            ChartTaTFrequencyAll.ChartAreas(0).AxisX.MajorGrid.LineDashStyle = DataVisualization.Charting.ChartDashStyle.NotSet
            ChartTaTFrequencyAll.ChartAreas(0).AxisY.MajorGrid.LineDashStyle = DataVisualization.Charting.ChartDashStyle.NotSet

            ChartTaTFrequencyAll.ChartAreas(0).AxisX.LabelStyle.Angle = -45
        Catch ex As Exception
            Exit Sub
            MessageBox.Show(ex.Message, "Exception", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
        End Try
    End Sub

    Private Sub LoadStats()
        Try
            Dim PIn_P, POut_P, PIn_C, POut_C As Integer
            Dim Hema_P, Hema_C, Chem_P, Chem_C As Integer

            'Chart Frequency------------------------------------------------------------------------------------------------------------------------------------------
            LoadTaTFrequencyAll()
            '--------------------------------------------------------------------------------------------------------------------------------------------------------

            'Section Out Patient On-Process and Completed------------------------------------------------------------------------------------------------------------
            Connect()
            rs.Parameters.Clear()
            rs.Parameters.AddWithValue("@dDate", Format(Now, "yyyy-MM-dd"))
            rs.Parameters.AddWithValue("@PType", "Out Patient")
            rs.Connection = conn
            rs.CommandType = CommandType.StoredProcedure
            rs.CommandText = "tat_ratio_total_processing"
            reader = rs.ExecuteReader
            While reader.Read()
                POut_P = Val(reader(0))
            End While
            Disconnect()

            Connect()
            rs.Parameters.Clear()
            rs.Parameters.AddWithValue("@dDate", Format(Now, "yyyy-MM-dd"))
            rs.Parameters.AddWithValue("@PType", "Out Patient")
            rs.Connection = conn
            rs.CommandType = CommandType.StoredProcedure
            rs.CommandText = "tat_ratio_total_completed"
            reader = rs.ExecuteReader
            While reader.Read()
                POut_C = Val(reader(0))
            End While
            Disconnect()
            '---------------------------------------------------------------------------------------------------------------------------------------------------------

            'Section In Patient On-Process and Completed------------------------------------------------------------------------------------------------------------
            Connect()
            rs.Parameters.Clear()
            rs.Parameters.AddWithValue("@dDate", Format(Now, "yyyy-MM-dd"))
            rs.Parameters.AddWithValue("@PType", "In Patient")
            rs.Connection = conn
            rs.CommandType = CommandType.StoredProcedure
            rs.CommandText = "tat_ratio_total_processing"
            reader = rs.ExecuteReader
            While reader.Read()
                PIn_P = Val(reader(0))
            End While
            Disconnect()

            Connect()
            rs.Parameters.Clear()
            rs.Parameters.AddWithValue("@dDate", Format(Now, "yyyy-MM-dd"))
            rs.Parameters.AddWithValue("@PType", "In Patient")
            rs.Connection = conn
            rs.CommandType = CommandType.StoredProcedure
            rs.CommandText = "tat_ratio_total_completed"
            reader = rs.ExecuteReader
            While reader.Read()
                PIn_C = Val(reader(0))
            End While
            Disconnect()
            '-------------------------------------------------------------------------------------------------------------------------------------------------------

            'Section Hematology On-Process and Completed------------------------------------------------------------------------------------------------------------
            Connect()
            rs.Parameters.Clear()
            rs.Parameters.AddWithValue("@dDate", Format(Now, "yyyy-MM-dd"))
            rs.Parameters.AddWithValue("@Sec", "Hematology")
            rs.Connection = conn
            rs.CommandType = CommandType.StoredProcedure
            rs.CommandText = "tat_ratio_sec_processing"
            reader = rs.ExecuteReader
            While reader.Read()
                Hema_P = Val(reader(0))
            End While
            Disconnect()

            Connect()
            rs.Parameters.Clear()
            rs.Parameters.AddWithValue("@dDate", Format(Now, "yyyy-MM-dd"))
            rs.Parameters.AddWithValue("@Sec", "Hematology")
            rs.Connection = conn
            rs.CommandType = CommandType.StoredProcedure
            rs.CommandText = "tat_ratio_sec_completed"
            reader = rs.ExecuteReader
            While reader.Read()
                Hema_C = Val(reader(0))
            End While
            Disconnect()
            '-----------------------------------------------------------------------------------------------------------------------------------------------------

            'Section Chemistry On-Process and Completed------------------------------------------------------------------------------------------------------------
            Connect()
            rs.Parameters.Clear()
            rs.Parameters.AddWithValue("@dDate", Format(Now, "yyyy-MM-dd"))
            rs.Parameters.AddWithValue("@Sec", "Chemistry")
            rs.Connection = conn
            rs.CommandType = CommandType.StoredProcedure
            rs.CommandText = "tat_ratio_sec_processing"
            reader = rs.ExecuteReader
            While reader.Read()
                Chem_P = Val(reader(0))
            End While
            Disconnect()

            Connect()
            rs.Parameters.Clear()
            rs.Parameters.AddWithValue("@dDate", Format(Now, "yyyy-MM-dd"))
            rs.Parameters.AddWithValue("@Sec", "Chemistry")
            rs.Connection = conn
            rs.CommandType = CommandType.StoredProcedure
            rs.CommandText = "tat_ratio_sec_completed"
            reader = rs.ExecuteReader
            While reader.Read()
                Chem_C = Val(reader(0))
            End While
            Disconnect()
            '-----------------------------------------------------------------------------------------------------------------------------------------------------

            OutPatient.Text = POut_C + POut_P
            InPatient.Text = PIn_C + PIn_P

            Hematology.Text = Hema_C + Hema_P
            Chemistry.Text = Chem_P + Chem_C

        Catch ex As Exception
            Disconnect()
            MessageBox.Show("No Data Shown", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Asterisk)
            ChartTaTFrequency.Titles.Clear()
            ChartTaTFrequency.Series.Clear()
            Exit Sub
        End Try
    End Sub

    Dim a As Integer = 0
    Private Sub tmSlide_Tick(sender As Object, e As EventArgs) Handles tmSlide.Tick
        a = a + 1
        If a > 0 And a <= 20 Then
            LoadTaTTracker()
            TabPane.SelectedTabPageIndex = 0
        End If
        If a > 20 And a <= 40 Then
            LoadStats()
            TabPane.SelectedTabPageIndex = 1
        End If
        If a > 40 And a <= 60 Then
            LoadTaTFrequency()
            TabPane.SelectedTabPageIndex = 2
        End If
        If a > 60 And a <= 80 Then
            LoadTaT()
            TabPane.SelectedTabPageIndex = 3
        End If
        If a > 80 And a <= 100 Then
            LoadTaTRatio()
            TabPane.SelectedTabPageIndex = 4
        End If
        If a > 100 Then
            a = 1
        End If
    End Sub

    Private Sub LoadTaTTracker()
        Try
            dtResult.RefreshDataSource()

            GridView.Columns.Clear()
            GridView.Appearance.HeaderPanel.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center
            GridView.Appearance.HeaderPanel.FontStyleDelta = FontStyle.Bold
            GridView.Appearance.HeaderPanel.Font = New Font("Tahoma", 12)

            GridView.Appearance.OddRow.BackColor = Color.FromArgb(226, 236, 247)
            GridView.OptionsView.EnableAppearanceOddRow = True
            GridView.Appearance.EvenRow.BackColor = Color.White
            GridView.OptionsView.EnableAppearanceEvenRow = True

            Connect()
            rs.Parameters.Clear()
            Dim sql As String
            sql = "Tat_Tracker"

            Dim dt As DataTable = New DataTable
            rs.Connection = conn
            rs.CommandType = CommandType.StoredProcedure
            rs.CommandText = sql
            Dim adapter As New MySqlDataAdapter(rs)

            adapter.Fill(dt)
            dtResult.DataSource = (dt)
            Disconnect()

            GridView.OptionsBehavior.Editable = False
            GridView.OptionsSelection.EnableAppearanceFocusedCell = False
            GridView.FocusRectStyle = DrawFocusRectStyle.None

            GridView.Columns("Section").Visible = False

            GridView.Columns("ReceivedDT").DisplayFormat.FormatType = FormatType.DateTime
            GridView.Columns("ReceivedDT").DisplayFormat.FormatString = "MM/dd/yyyy hh:mm:ss tt"
            GridView.Columns("ReceivedDT").AppearanceCell.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center

            GridView.Columns("CheckedInDT").DisplayFormat.FormatType = FormatType.DateTime
            GridView.Columns("CheckedInDT").DisplayFormat.FormatString = "MM/dd/yyyy hh:mm:ss tt"
            GridView.Columns("CheckedInDT").AppearanceCell.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center

            GridView.Columns("ProcessingDT").DisplayFormat.FormatType = FormatType.DateTime
            GridView.Columns("ProcessingDT").DisplayFormat.FormatString = "MM/dd/yyyy hh:mm:ss tt"
            GridView.Columns("ProcessingDT").AppearanceCell.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center

            GridView.Columns("ValidatedDT").DisplayFormat.FormatType = FormatType.DateTime
            GridView.Columns("ValidatedDT").DisplayFormat.FormatString = "MM/dd/yyyy hh:mm:ss tt"
            GridView.Columns("ValidatedDT").AppearanceCell.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center

            GridView.Columns("ReleasedDT").DisplayFormat.FormatType = FormatType.DateTime
            GridView.Columns("ReleasedDT").DisplayFormat.FormatString = "MM/dd/yyyy hh:mm:ss tt"
            GridView.Columns("ReleasedDT").AppearanceCell.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center

            GridView.Columns("EstimatedTime").DisplayFormat.FormatType = FormatType.DateTime
            GridView.Columns("EstimatedTime").DisplayFormat.FormatString = "MM/dd/yyyy hh:mm:ss tt"
            GridView.Columns("EstimatedTime").AppearanceCell.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center

            GridView.Columns("Status").AppearanceCell.TextOptions.HAlignment = HorzAlignment.Center

            GridView.Columns("DateReleased").Visible = False
            GridView.Columns("Status").Visible = False
            DelayedCount = 0
            OnTimeCount = 0
            For x As Integer = 0 To GridView.RowCount - 1
                If GridView.GetRowCellValue(x, "Status").ToString = "Delayed" Then
                    DelayedCount = DelayedCount + 1
                    lblDelayed.Text = "Delayed Sample: " & DelayedCount
                Else
                    OnTimeCount = OnTimeCount + 1
                    lblOnTime.Text = "On-Time Sample: " & OnTimeCount
                End If
            Next
        Catch ex As Exception
            Exit Sub
            MessageBox.Show(ex.Message, "Exception", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
        End Try
    End Sub

    'Private Sub GridView_RowCellStyle(sender As Object, e As RowCellStyleEventArgs) Handles GridView.RowCellStyle
    '    If GridView.GetRowCellValue(e.RowHandle, "Status").ToString = "Delayed" Then
    '        e.Appearance.BackColor = Color.Crimson
    '        e.Appearance.ForeColor = Color.White
    '    Else
    '        e.Appearance.BackColor = Color.ForestGreen
    '        e.Appearance.ForeColor = Color.White
    '    End If
    'End Sub

    Private Sub frmTaT_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        tmSlide.Start()
    End Sub

    Private Sub Timer2_Tick(sender As Object, e As EventArgs) Handles Timer2.Tick
        lblTime.Text = Now.ToLongTimeString
        lblDate.Text = Now.ToLongDateString
    End Sub

    Private Sub frmTaT_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles MyBase.KeyDown
        If e.KeyCode = Keys.Escape Then Me.Close()
    End Sub
End Class