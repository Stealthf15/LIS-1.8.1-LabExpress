Imports System.IO
Public Class frmQC
    Private _No_of_Run As String
    Private _run As String

    Public Ave, stdDevP1, stdDevP3, stdDevP2, stdDevN1, stdDevN2, stdDevN3, sd As Double
    Public month As String = ""

    Private Sub LoadData()
        rs.Parameters.Clear()
        rs.Parameters.AddWithValue("@TestName", cboLimit.Text)

        'FOR TONDO MED'
        'rs.Parameters.AddWithValue("@TestName", "Chemistry")

        rs.Parameters.Add("@Date_From", MySql.Data.MySqlClient.MySqlDbType.Date).Value = Format(dtFrom.Value, "yyyy-MM-dd")
        rs.Parameters.Add("@Date_To", MySql.Data.MySqlClient.MySqlDbType.Date).Value = Format(dtTo.Value, "yyyy-MM-dd")
        rs.Parameters.AddWithValue("@ControlID", cboControl.Text)
        rs.Parameters.AddWithValue("@LotNo", cboLot.Text)
        rs.Parameters.AddWithValue("@Instrument", cboMachines.Text)

        LoadRecordsOnLVSQL(lvList, "SELECT DISTINCT `universal_id`, `test_code`, `month`, `year` FROM `control_result` WHERE `test_type` = @TestName AND `sample_id` = @ControlID AND `instrument` = @Instrument AND `lot_no` = @LotNo AND (`date` BETWEEN @Date_From AND @Date_To)", 3)
    End Sub

    Private Sub frmQC_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        AutoLoadTestName()
    End Sub

    Private Sub lvList_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles lvList.SelectedIndexChanged
        'Dim Ave, stdDevP1, stdDevP3, stdDevP2, stdDevN1, stdDevN2, stdDevN3, sd As Double

        Dim unit As String = ""

        Try
            rs.Parameters.Clear()
            rs.Parameters.AddWithValue("@TestCode", lvList.FocusedItem.SubItems(1).Text)
            rs.Parameters.Add("@Date_From", MySql.Data.MySqlClient.MySqlDbType.Date).Value = Format(dtFrom.Value, "yyyy-MM-dd")
            rs.Parameters.Add("@Date_To", MySql.Data.MySqlClient.MySqlDbType.Date).Value = Format(dtTo.Value, "yyyy-MM-dd")
            rs.Parameters.AddWithValue("@ControlID", cboControl.Text)
            rs.Parameters.AddWithValue("@LotNo", cboLot.Text)
            rs.Parameters.AddWithValue("@Instrument", cboMachines.Text)
            Try
                Connect()
                rs.Connection = conn
                rs.CommandText = "SELECT `target`, `sd`, `ul`, `ll`, `plus_one`, `minus_one`, `plus_three`, `minus_three` FROM `control_target` WHERE `test_code` = @TestCode AND `control_id` = @ControlID  AND `instrument` = @Instrument ORDER BY `id`"
                reader = rs.ExecuteReader
                reader.Read()
                Ave = Val(reader(0))
                sd = Val(reader(1))

                stdDevP1 = Val(reader(4))
                stdDevN1 = Val(reader(5))

                stdDevP2 = Val(reader(2))
                stdDevN2 = Val(reader(3))

                stdDevP3 = Val(reader(6))
                stdDevN3 = Val(reader(7))
                Disconnect()
            Catch ex As Exception
                Disconnect()
                MessageBox.Show(ex.Message, "Warning", MessageBoxButtons.OK, MessageBoxIcon.Asterisk)
                Chart1.Titles.Clear()
                Chart1.Series.Clear()
                Exit Sub
            End Try

            Chart1.Titles.Clear()
            Chart1.Series.Clear()

            If lvList.FocusedItem.SubItems(2).Text = "1" Then
                Month = "JAN"
            ElseIf lvList.FocusedItem.SubItems(2).Text = "2" Then
                Month = "FEB"
            ElseIf lvList.FocusedItem.SubItems(2).Text = "3" Then
                Month = "MAR"
            ElseIf lvList.FocusedItem.SubItems(2).Text = "4" Then
                Month = "APR"
            ElseIf lvList.FocusedItem.SubItems(2).Text = "5" Then
                Month = "MAY"
            ElseIf lvList.FocusedItem.SubItems(2).Text = "6" Then
                Month = "JUN"
            ElseIf lvList.FocusedItem.SubItems(2).Text = "7" Then
                Month = "JUL"
            ElseIf lvList.FocusedItem.SubItems(2).Text = "8" Then
                Month = "AUG"
            ElseIf lvList.FocusedItem.SubItems(2).Text = "9" Then
                Month = "SEP"
            ElseIf lvList.FocusedItem.SubItems(2).Text = "10" Then
                Month = "OCT"
            ElseIf lvList.FocusedItem.SubItems(2).Text = "11" Then
                Month = "NOV"
            ElseIf lvList.FocusedItem.SubItems(2).Text = "12" Then
                Month = "DEC"
            End If

            Chart1.Titles.Add(lvList.FocusedItem.SubItems(0).Text & " (" & lvList.FocusedItem.SubItems(1).Text & ")")
            Chart1.Titles(0).Font = New Font("Tahoma", 12, FontStyle.Bold)

            Chart1.Series.Add("+3 SD " & stdDevP3)
            Chart1.Series("+3 SD " & stdDevP3).BorderWidth = 3
            Chart1.Series("+3 SD " & stdDevP3).Color = Color.Red
            Chart1.Series("+3 SD " & stdDevP3).ChartType = System.Windows.Forms.DataVisualization.Charting.SeriesChartType.Line
            Chart1.Series("+3 SD " & stdDevP3).Font = New Font("Tahoma", 10, FontStyle.Bold)

            Chart1.Series.Add("+2 SD " & stdDevP2)
            Chart1.Series("+2 SD " & stdDevP2).BorderWidth = 3
            Chart1.Series("+2 SD " & stdDevP2).Color = Color.ForestGreen
            Chart1.Series("+2 SD " & stdDevP2).ChartType = System.Windows.Forms.DataVisualization.Charting.SeriesChartType.Line
            Chart1.Series("+2 SD " & stdDevP2).Font = New Font("Tahoma", 10, FontStyle.Bold)

            Chart1.Series.Add("+1 SD " & stdDevP1)
            Chart1.Series("+1 SD " & stdDevP1).BorderWidth = 3
            Chart1.Series("+1 SD " & stdDevP1).Color = Color.Yellow
            Chart1.Series("+1 SD " & stdDevP1).ChartType = System.Windows.Forms.DataVisualization.Charting.SeriesChartType.Line
            Chart1.Series("+1 SD " & stdDevP1).Font = New Font("Tahoma", 10, FontStyle.Bold)

            Chart1.Series.Add("Mean " & Ave)
            Chart1.Series("Mean " & Ave).BorderWidth = 3
            Chart1.Series("Mean " & Ave).Color = Color.Gray
            Chart1.Series("Mean " & Ave).ChartType = System.Windows.Forms.DataVisualization.Charting.SeriesChartType.Line
            Chart1.Series("Mean " & Ave).Font = New Font("Tahoma", 10, FontStyle.Bold)

            Chart1.Series.Add("-1 SD " & stdDevN1)
            Chart1.Series("-1 SD " & stdDevN1).BorderWidth = 3
            Chart1.Series("-1 SD " & stdDevN1).Color = Color.Yellow
            Chart1.Series("-1 SD " & stdDevN1).ChartType = System.Windows.Forms.DataVisualization.Charting.SeriesChartType.Line
            Chart1.Series("-1 SD " & stdDevN1).Font = New Font("Tahoma", 10, FontStyle.Bold)

            Chart1.Series.Add("-2 SD " & stdDevN2)
            Chart1.Series("-2 SD " & stdDevN2).BorderWidth = 3
            Chart1.Series("-2 SD " & stdDevN2).Color = Color.ForestGreen
            Chart1.Series("-2 SD " & stdDevN2).ChartType = System.Windows.Forms.DataVisualization.Charting.SeriesChartType.Line
            Chart1.Series("-2 SD " & stdDevN2).Font = New Font("Tahoma", 10, FontStyle.Bold)

            Chart1.Series.Add("-3 SD " & stdDevN3)
            Chart1.Series("-3 SD " & stdDevN3).BorderWidth = 3
            Chart1.Series("-3 SD " & stdDevN3).Color = Color.Red
            Chart1.Series("-3 SD " & stdDevN3).ChartType = System.Windows.Forms.DataVisualization.Charting.SeriesChartType.Line
            Chart1.Series("-3 SD " & stdDevN3).Font = New Font("Tahoma", 10, FontStyle.Bold)

            Chart1.Series.Add("Control Value")
            Chart1.Series("Control Value").MarkerStyle = DataVisualization.Charting.MarkerStyle.Circle
            Chart1.Series("Control Value").MarkerSize = 1
            Chart1.Series("Control Value").MarkerColor = Color.DarkBlue
            Chart1.Series("Control Value").IsValueShownAsLabel = True
            Chart1.Series("Control Value").BorderWidth = 3
            Chart1.Series("Control Value").Color = Color.RoyalBlue
            Chart1.Series("Control Value").IsXValueIndexed = True
            Chart1.Series("Control Value").ChartType = System.Windows.Forms.DataVisualization.Charting.SeriesChartType.Line
            Chart1.Series("Control Value").Font = New Font("Tahoma", 6, FontStyle.Bold)

            Connect()
            rs.Connection = conn
            rs.CommandText = "SELECT `measurement`, DATE_FORMAT(`date`, '%d'), `unit` FROM `control_result` WHERE `test_code` = @TestCode AND `sample_id` = @ControlID AND `lot_no` = @LotNo AND `instrument` = @Instrument AND (`date` BETWEEN @Date_From AND @Date_To) ORDER BY `date`"
            reader = rs.ExecuteReader
            For x As Integer = 1 To 31 Step 1
                While reader.Read
                    Chart1.Series("Mean " & Ave).Points.AddXY(reader(1).ToString, Ave)
                    Chart1.Series("+1 SD " & stdDevP1).Points.AddXY(reader(1).ToString, stdDevP1)
                    Chart1.Series("-1 SD " & stdDevN1).Points.AddXY(reader(1).ToString, stdDevN1)

                    Chart1.Series("+2 SD " & stdDevP2).Points.AddXY(reader(1).ToString, stdDevP2)
                    Chart1.Series("-2 SD " & stdDevN2).Points.AddXY(reader(1).ToString, stdDevN2)

                    Chart1.Series("+3 SD " & stdDevP3).Points.AddXY(reader(1).ToString, stdDevP3)
                    Chart1.Series("-3 SD " & stdDevN3).Points.AddXY(reader(1).ToString, stdDevN3)
                    Chart1.Series("Control Value").Points.AddXY(reader(1).ToString, reader(0).ToString)

                    unit = reader(2)
                End While
            Next
            Disconnect()

            Chart1.Series("Control Value").ChartArea = "ChartArea1"
            Chart1.Series("Mean " & Ave).ChartArea = "ChartArea1"
            Chart1.Series("+1 SD " & stdDevP1).ChartArea = "ChartArea1"
            Chart1.Series("-1 SD " & stdDevN1).ChartArea = "ChartArea1"

            Chart1.Series("+2 SD " & stdDevP2).ChartArea = "ChartArea1"
            Chart1.Series("-2 SD " & stdDevN2).ChartArea = "ChartArea1"

            Chart1.Series("+3 SD " & stdDevP3).ChartArea = "ChartArea1"
            Chart1.Series("-3 SD " & stdDevN3).ChartArea = "ChartArea1"

            Chart1.ChartAreas(0).AxisX.Minimum = 1
            Chart1.ChartAreas(0).AxisX.Interval = 1

            Chart1.ChartAreas(0).AxisY.Minimum = Ave - (sd * 3)
            Chart1.ChartAreas(0).AxisY.Maximum = Ave + (sd * 3)
            Chart1.ChartAreas(0).AxisY.Interval = sd * 2
            'Chart1.ChartAreas(0).AxisY.IntervalOffset = 1

            Chart1.ChartAreas(0).AxisX.LabelStyle.Angle = -45
            Chart1.ChartAreas(0).AxisX.Title = "Days"
            Chart1.ChartAreas(0).AxisX.TitleFont = New Font("Tahoma", 10, FontStyle.Bold)

            Chart1.ChartAreas(0).AxisY.Title = unit
            Chart1.ChartAreas(0).AxisY.TitleFont = New Font("Tahoma", 10, FontStyle.Bold)

            Chart1.ChartAreas(0).AxisY2.Enabled = DataVisualization.Charting.AxisEnabled.True
            Chart1.ChartAreas(0).AxisY2.Minimum = Ave - (sd * 3)
            Chart1.ChartAreas(0).AxisY2.Maximum = Ave + (sd * 3)
            Chart1.ChartAreas(0).AxisY2.Interval = sd * 2

            Chart1.ChartAreas(0).AxisY.CustomLabels.Clear()
            Chart1.ChartAreas(0).AxisY.CustomLabels.Add(Ave, Ave - 0.01, Ave.ToString)
            Chart1.ChartAreas(0).AxisY.CustomLabels.Add(stdDevP1, stdDevP1 - 0.01, stdDevP1.ToString)
            Chart1.ChartAreas(0).AxisY.CustomLabels.Add(stdDevN1, stdDevN1 - 0.01, stdDevN1.ToString)
            Chart1.ChartAreas(0).AxisY.CustomLabels.Add(stdDevP2, stdDevP2 - 0.01, stdDevP2.ToString)
            Chart1.ChartAreas(0).AxisY.CustomLabels.Add(stdDevN2, stdDevN2 - 0.01, stdDevN2.ToString)
            Chart1.ChartAreas(0).AxisY.CustomLabels.Add(stdDevP3, stdDevP3 - 0.01, stdDevP3.ToString)
            Chart1.ChartAreas(0).AxisY.CustomLabels.Add(stdDevN3 + (stdDevN3 * 0.01), stdDevN3 + ((stdDevN3 * 0.01) - 0.01), stdDevN3.ToString)
            Chart1.ChartAreas(0).AxisY.LabelStyle.Font = New Font("Tahoma", 8, FontStyle.Regular)

            Chart1.ChartAreas(0).AxisY2.CustomLabels.Clear()
            Chart1.ChartAreas(0).AxisY2.CustomLabels.Add(stdDevP1, stdDevP1 - 0.01, "+1 SD")
            Chart1.ChartAreas(0).AxisY2.CustomLabels.Add(stdDevN1, stdDevN1 - 0.01, "-1 SD")
            Chart1.ChartAreas(0).AxisY2.CustomLabels.Add(stdDevP2, stdDevP2 - 0.01, "+2 SD")
            Chart1.ChartAreas(0).AxisY2.CustomLabels.Add(stdDevN2, stdDevN2 - 0.01, "-2 SD")
            Chart1.ChartAreas(0).AxisY2.CustomLabels.Add(stdDevP3, stdDevP3 - 0.01, "+3 SD")
            Chart1.ChartAreas(0).AxisY2.CustomLabels.Add(stdDevN3 + (stdDevN3 * 0.01), stdDevN3 + ((stdDevN3 * 0.01) - 0.01), "-3 SD")
            Chart1.ChartAreas(0).AxisY2.LabelStyle.Font = New Font("Tahoma", 8, FontStyle.Regular)

        Catch ex As Exception
            Exit Sub
            MessageBox.Show(ex.Message, "Exception", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
        End Try
    End Sub

    Private Sub cboLimit_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cboLimit.SelectedIndexChanged
        rs.Parameters.Clear()
        rs.Parameters.AddWithValue("@Section", cboLimit.Text)

        'FOR TONDO MED'
        'rs.Parameters.AddWithValue("@Section", "Chemistry")

        Me.cboControl.Properties.Items.Clear()
        Connect()
        rs.Connection = conn
        rs.CommandText = "SELECT DISTINCT `sample_id` FROM `control_result` WHERE `test_type` = @Section"
        reader = rs.ExecuteReader
        While reader.Read
            Me.cboControl.Properties.Items.Add(reader(0).ToString)
        End While
        Disconnect()
    End Sub

    Public Sub AutoLoadTestName()
        Me.cboLimit.Properties.Items.Clear()
        Connect()
        rs.Connection = conn
        rs.CommandText = "SELECT * FROM `testtype` WHERE `test_name` NOT LIKE 'All' ORDER BY `test_name`"
        reader = rs.ExecuteReader
        While reader.Read
            Me.cboLimit.Properties.Items.Add(reader(1).ToString)
        End While
        Disconnect()

        Me.cboMachines.Properties.Items.Clear()
        Connect()
        rs.Connection = conn
        rs.CommandText = "SELECT DISTINCT `instrument` FROM `control_setting` ORDER BY `instrument` DESC"
        reader = rs.ExecuteReader
        While reader.Read
            Me.cboMachines.Properties.Items.Add(reader(0).ToString)
        End While
        Disconnect()
    End Sub

    Private Sub btnClose_ItemClick(ByVal sender As Object, ByVal e As DevExpress.XtraBars.ItemClickEventArgs) Handles btnClose.ItemClick
        Me.Close()
        Me.Dispose()
    End Sub

    Private Sub btnPreview_ItemClick(ByVal sender As Object, ByVal e As DevExpress.XtraBars.ItemClickEventArgs) Handles btnPreview.ItemClick
        PrintDocument1.DefaultPageSettings.Landscape = True
        PrintDocument1.DefaultPageSettings.PaperSize = New Printing.PaperSize("First custom size", 850, 1300)
        Chart1.Printing.PrintDocument = PrintDocument1 ' this enables the adding of other material to the page on which the chart is printed
        Chart1.Printing.PrintPreview()
        'Chart1.Printing.Print(True)
    End Sub

    Private Sub PrintDocument1_PrintPage(ByVal sender As System.Object, ByVal e As System.Drawing.Printing.PrintPageEventArgs) Handles PrintDocument1.PrintPage
        Dim myFont As Font = New Font("Calibri", 12, FontStyle.Bold)
        Dim myFont2 As Font = New Font("Tahoma", 8, FontStyle.Regular)
        Dim myBrush As Brush = Brushes.Black
        Dim incYAxis As Integer = 155

        e.Graphics.DrawString("LAB EXPRESS MEDICAL DIAGNOSTICS", myFont, myBrush, 450, 65)
        e.Graphics.DrawString("Quality Control Graph " & month & " " & lvList.FocusedItem.SubItems(3).Text & " (" & cboMachines.Text & ")", myFont, myBrush, 375, 90)
        e.Graphics.DrawString("Control:     " & cboControl.Text, myFont, myBrush, 20, 115)
        e.Graphics.DrawString("Lot No:     " & cboLot.Text, myFont, myBrush, 305, 115)
        e.Graphics.DrawString("Date", myFont2, myBrush, 10, 150)
        e.Graphics.DrawString("Data", myFont2, myBrush, 60, 150)
        e.Graphics.DrawString("SDI", myFont2, myBrush, 110, 150)

        rs.Parameters.Clear()
        rs.Parameters.AddWithValue("@TestCode", lvList.FocusedItem.SubItems(1).Text)
        rs.Parameters.Add("@Date_From", MySql.Data.MySqlClient.MySqlDbType.Date).Value = Format(dtFrom.Value, "yyyy-MM-dd")
        rs.Parameters.Add("@Date_To", MySql.Data.MySqlClient.MySqlDbType.Date).Value = Format(dtTo.Value, "yyyy-MM-dd")
        rs.Parameters.AddWithValue("@ControlID", cboControl.Text)
        rs.Parameters.AddWithValue("@LotNo", cboLot.Text)
        rs.Parameters.AddWithValue("@Instrument", cboMachines.Text)

        Connect()
        rs.Connection = conn
        rs.CommandText = "SELECT `measurement`, DATE_FORMAT(`date`, '%e') FROM `control_result` WHERE `test_code` = @TestCode AND `sample_id` = @ControlID AND `lot_no` = @LotNo AND `instrument` = @Instrument AND (`date` BETWEEN @Date_From AND @Date_To) ORDER BY `date`"
        reader = rs.ExecuteReader
        For x As Integer = 1 To 100 Step 1
            While reader.Read
                e.Graphics.DrawString(Format((reader(0) - Ave) / sd, "0.00"), myFont2, myBrush, 110, incYAxis + 15)
                e.Graphics.DrawString(reader(0), myFont2, myBrush, 60, incYAxis + 15)
                e.Graphics.DrawString(reader(1), myFont2, myBrush, 15, incYAxis + 15)
                incYAxis = incYAxis + 15
            End While
        Next
        Disconnect()

        'Chart1.Printing.PrintPaint(e.Graphics, New Rectangle(150, 150, Chart1.Width - 600, Chart1.Height - 250)) ' draw the chart
        Chart1.Printing.PrintPaint(e.Graphics, New Rectangle(150, 150, Chart1.Width + 100, Chart1.Height + 150))
    End Sub

    Private Sub btnSearch_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSearch.Click
        LoadData()
        Chart1.Titles.Clear()
        Chart1.Series.Clear()
    End Sub

    Private Sub LabelControl2_Click(sender As Object, e As EventArgs) Handles LabelControl2.Click

    End Sub

    Private Sub LabelControl3_Click(sender As Object, e As EventArgs) Handles LabelControl3.Click

    End Sub

    Private Sub LabelControl5_Click(sender As Object, e As EventArgs) Handles LabelControl5.Click

    End Sub

    Private Sub cboLot_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cboLot.SelectedIndexChanged

    End Sub

    Private Sub btnView_ItemClick(ByVal sender As Object, ByVal e As DevExpress.XtraBars.ItemClickEventArgs) Handles btnView.ItemClick
        frmQCResults.MainSampleID = cboControl.Text
        frmQCResults.Section = cboLimit.Text
        frmQCResults.ShowDialog()
    End Sub

    Private Sub btnPrint_ItemClick(ByVal sender As Object, ByVal e As DevExpress.XtraBars.ItemClickEventArgs) Handles btnPrint.ItemClick
        Try
            'Dim dt As String = Date.Now.ToString("MM-dd-yyyy")
            'Dim Path As String = Application.StartupPath & "\LJGraph\" & lvList.FocusedItem.SubItems(1).Text & " - " & dt
            'Dim start As String = Application.StartupPath & "\LJGraph\" & Format(dtFrom.Value, "yyyy-MM-dd") & " to " & Format(dtTo.Value, "yyyy-MM-dd")

            'If Not Directory.Exists(start) Then
            '    Directory.CreateDirectory(start)
            '    If Directory.Exists(start) Then
            '        Me.Chart1.SaveImage(start & "\" & lvList.FocusedItem.SubItems(0).Text & ".jpg", Drawing.Imaging.ImageFormat.Jpeg)
            '        MessageBox.Show("LJ Graph successfully saved as Image.", "Save Image", MessageBoxButtons.OK, MessageBoxIcon.Information)
            '        Exit Sub
            '    End If
            'Else
            '    Me.Chart1.SaveImage(start & "\" & lvList.FocusedItem.SubItems(0).Text & ".jpg", Drawing.Imaging.ImageFormat.Jpeg)
            '    MessageBox.Show("LJ Graph successfully saved as Image.", "Save Image", MessageBoxButtons.OK, MessageBoxIcon.Information)
            '    Exit Sub
            '    'Me.Chart1.SaveImage(Application.StartupPath & "\LJGraph\" & lvList.FocusedItem.SubItems(0).Text & ".jpg", Drawing.Imaging.ImageFormat.Jpeg)
            '    'MessageBox.Show("LJ Graph successfully saved as Image.", "Save Image", MessageBoxButtons.OK, MessageBoxIcon.Information)
            '    'Exit Sub
            'End If

            PrintDocument1.DefaultPageSettings.Landscape = True
            PrintDocument1.DefaultPageSettings.PaperSize = New Printing.PaperSize("First custom size", 850, 1300)
            Chart1.Printing.PrintDocument = PrintDocument1 ' this enables the adding of other material to the page on which the chart is printed
            Chart1.Printing.Print(True)

        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error While Saving", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Exit Sub
        End Try
    End Sub

    Private Sub btnRefresh_ItemClick(sender As Object, e As DevExpress.XtraBars.ItemClickEventArgs) Handles btnRefresh.ItemClick
        LoadData()
        Chart1.Titles.Clear()
        Chart1.Series.Clear()
    End Sub

    Private Sub frm_Activated(sender As Object, e As EventArgs) Handles Me.Activated
        'MainFOrm.aceFecal.Appearance.Normal.BackColor = Color.FromArgb(6, 31, 71)
        MainFOrm.btnQC.Appearance.Normal.BackColor = Color.FromArgb(16, 110, 190)
        MainFOrm.btnQC.Appearance.Normal.ForeColor = Color.FromArgb(255, 255, 255)
    End Sub

    Private Sub frm_FormClosed(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
        MainFOrm.lblTitle.Text = ""
        MainFOrm.btnQC.Appearance.Normal.BackColor = Color.FromArgb(240, 240, 240)
        MainFOrm.btnQC.Appearance.Normal.ForeColor = Color.FromArgb(27, 41, 62)
        Me.Dispose()
    End Sub

    Private Sub cboControl_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cboControl.SelectedIndexChanged
        rs.Parameters.Clear()
        rs.Parameters.AddWithValue("@Control", cboControl.Text)

        Me.cboLot.Properties.Items.Clear()
        Connect()
        rs.Connection = conn
        rs.CommandText = "SELECT DISTINCT `lot_no` FROM `control_result` WHERE `sample_id` = @Control"
        reader = rs.ExecuteReader
        While reader.Read
            Me.cboLot.Properties.Items.Add(reader(0).ToString)
        End While
        Disconnect()
    End Sub

    Private Sub cboMachines_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cboMachines.SelectedIndexChanged
        rs.Parameters.Clear()
        rs.Parameters.AddWithValue("@Machine", cboMachines.Text)

        Me.cboControl.Properties.Items.Clear()
        Connect()
        rs.Connection = conn
        rs.CommandText = "SELECT DISTINCT `sample_id` FROM `control_result` WHERE `instrument` = @Machine"
        reader = rs.ExecuteReader
        While reader.Read
            Me.cboControl.Properties.Items.Add(reader(0).ToString)
        End While
        Disconnect()
    End Sub
End Class