Imports DevExpress.XtraGrid.Views.Grid
Imports MySql.Data.MySqlClient

Public Class frmTatComplete
    Public MedTechID As String = ""
    Public Sub AutoLoadTestName()
        cboLimit.Properties.Items.Clear()
        cboLimit.Properties.Items.Add("All")

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
        GenerateReportTaT(cboLimit, cboMedtech, dtFrom, dtTo, "ReportTaT.rdlc", RPTTaT.ReportViewer1)
        RPTTaT.ShowDialog()
    End Sub

    Private Sub btnClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnClose.ItemClick
        Me.Close()
    End Sub

    '########################################### SEARCHING USING DATES AND PATIENT TYPE ############################################################
    Private Sub btnSearch_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSearch.Click
        Try
            LoadTaTCount()
            LoadTaTDetailed()
        Catch ex As MySqlException
            Disconnect()
            MessageBox.Show(ex.Message, "Search Error", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            Exit Sub
        Finally
            Disconnect()
        End Try
    End Sub
    '################################################## END ###################################################################

    Private Sub frmWorkSheet_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        AutoLoadTestName()
        AutoLoadMedtech()
    End Sub

    Public Sub AutoLoadMedtech()
        cboMedtech.Properties.Items.Clear()
        cboMedtech.Properties.Items.Add("All")
        Connect()
        rs.Connection = conn
        rs.CommandType = CommandType.Text
        rs.CommandText = "SELECT CONCAT(fname, ' ', mname, ' ', lname, ', ', designation) AS `name` FROM `medtech` ORDER BY `name`"
        reader = rs.ExecuteReader
        While reader.Read
            cboMedtech.Properties.Items.Add(reader(0))
        End While
        Disconnect()
        cboMedtech.SelectedIndex = 0
    End Sub

    Private Sub LoadTaTCount()
        GridTaTQuant.Columns.Clear()
        GridTaTQuant.Appearance.HeaderPanel.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center
        GridTaTQuant.Appearance.HeaderPanel.FontStyleDelta = FontStyle.Bold

        GridTaTQuant.Appearance.OddRow.BackColor = Color.FromArgb(226, 236, 247)
        GridTaTQuant.OptionsView.EnableAppearanceOddRow = True
        GridTaTQuant.Appearance.EvenRow.BackColor = Color.White
        GridTaTQuant.OptionsView.EnableAppearanceEvenRow = True

        If cboMedtech.Text = "All" And cboLimit.Text = "All" Then
            Connect()
            rs.Parameters.Clear()
            rs.Parameters.AddWithValue("@DateFrom", Format(dtFrom.Value, "yyyy-MM-dd"))
            rs.Parameters.AddWithValue("@DateTo", Format(dtTo.Value, "yyyy-MM-dd"))

            Dim sql As String
            sql = "SELECT * FROM
		              `viewtat_permedtech` 
                    WHERE (DATE_FORMAT(`viewtat_permedtech`.`DateReleased`, '%Y-%m-%d') BETWEEN @DateFrom AND @DateTo)"

            Dim dt As DataTable = New DataTable
            rs.Connection = conn
            rs.CommandType = CommandType.Text
            rs.CommandText = sql
            Dim adapter As New MySqlDataAdapter(rs)

            adapter.Fill(dt)
            dtTaTQuant.DataSource = (dt)
            Disconnect()
        ElseIf cboMedtech.Text <> "All" And cboLimit.Text = "All" Then
            Connect()
            rs.Parameters.Clear()
            rs.Parameters.AddWithValue("@DateFrom", Format(dtFrom.Value, "yyyy-MM-dd"))
            rs.Parameters.AddWithValue("@DateTo", Format(dtTo.Value, "yyyy-MM-dd"))
            rs.Parameters.AddWithValue("@MedTech", cboMedtech.Text)

            Dim sql As String
            sql = "SELECT * FROM
		              `viewtat_permedtech` 
                    WHERE (DATE_FORMAT(`viewtat_permedtech`.`DateReleased`, '%Y-%m-%d') BETWEEN @DateFrom AND @DateTo) AND `viewtat_permedtech`.PerformedBy = @MedTech"

            Dim dt As DataTable = New DataTable
            rs.Connection = conn
            rs.CommandType = CommandType.Text
            rs.CommandText = sql
            Dim adapter As New MySqlDataAdapter(rs)

            adapter.Fill(dt)
            dtTaTQuant.DataSource = (dt)
            Disconnect()
        ElseIf cboMedtech.Text = "All" And cboLimit.Text <> "All" Then
            Connect()
            rs.Parameters.Clear()
            rs.Parameters.AddWithValue("@DateFrom", Format(dtFrom.Value, "yyyy-MM-dd"))
            rs.Parameters.AddWithValue("@DateTo", Format(dtTo.Value, "yyyy-MM-dd"))
            rs.Parameters.AddWithValue("@Section", cboLimit.Text)

            Dim sql As String
            sql = "SELECT * FROM
		              `viewtat_permedtech` 
                    WHERE (DATE_FORMAT(`viewtat_permedtech`.`DateReleased`, '%Y-%m-%d') BETWEEN @DateFrom AND @DateTo) AND `viewtat_permedtech`.Section = @Section"

            Dim dt As DataTable = New DataTable
            rs.Connection = conn
            rs.CommandType = CommandType.Text
            rs.CommandText = sql
            Dim adapter As New MySqlDataAdapter(rs)

            adapter.Fill(dt)
            dtTaTQuant.DataSource = (dt)
            Disconnect()
        ElseIf cboMedtech.Text <> "All" And cboLimit.Text <> "All" Then
            Connect()
            rs.Parameters.Clear()
            rs.Parameters.AddWithValue("@DateFrom", Format(dtFrom.Value, "yyyy-MM-dd"))
            rs.Parameters.AddWithValue("@DateTo", Format(dtTo.Value, "yyyy-MM-dd"))
            rs.Parameters.AddWithValue("@MedTech", cboMedtech.Text)
            rs.Parameters.AddWithValue("@Section", cboLimit.Text)

            Dim sql As String
            sql = "SELECT * FROM
		              `viewtat_permedtech` 
                    WHERE (DATE_FORMAT(`viewtat_permedtech`.`DateReleased`, '%Y-%m-%d') BETWEEN @DateFrom AND @DateTo) AND `viewtat_permedtech`.PerformedBy = @MedTech AND `viewtat_permedtech`.Section = @Section"

            Dim dt As DataTable = New DataTable
            rs.Connection = conn
            rs.CommandType = CommandType.Text
            rs.CommandText = sql
            Dim adapter As New MySqlDataAdapter(rs)

            adapter.Fill(dt)
            dtTaTQuant.DataSource = (dt)
            Disconnect()
        End If

        GridTaTQuant.OptionsBehavior.Editable = False
        GridTaTQuant.OptionsSelection.EnableAppearanceFocusedCell = False
        GridTaTQuant.FocusRectStyle = DrawFocusRectStyle.RowFullFocus

        GridTaTQuant.Columns("Count").Summary.Add(DevExpress.Data.SummaryItemType.Sum, "Count", "Sum={0}")

        lblTotalCount.Text = "Total Test Performed: " & GridTaTQuant.Columns("Count").SummaryItem.SummaryValue.ToString
    End Sub

    Private Sub LoadTaTDetailed()
        dtResult.RefreshDataSource()

        GridView.Columns.Clear()
        GridView.Appearance.HeaderPanel.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center
        GridView.Appearance.HeaderPanel.FontStyleDelta = FontStyle.Bold

        GridView.Appearance.OddRow.BackColor = Color.FromArgb(226, 236, 247)
        GridView.OptionsView.EnableAppearanceOddRow = True
        GridView.Appearance.EvenRow.BackColor = Color.White
        GridView.OptionsView.EnableAppearanceEvenRow = True

        If cboMedtech.Text = "All" And cboLimit.Text = "All" Then
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
            dtResult.DataSource = (dt)
            Disconnect()
        ElseIf cboMedtech.Text <> "All" And cboLimit.Text = "All" Then
            Connect()
            rs.Parameters.Clear()
            rs.Parameters.AddWithValue("@DateFrom", Format(dtFrom.Value, "yyyy-MM-dd"))
            rs.Parameters.AddWithValue("@DateTo", Format(dtTo.Value, "yyyy-MM-dd"))
            rs.Parameters.AddWithValue("@MedTech", cboMedtech.Text)

            Dim sql As String
            sql = "SELECT
						*
				    FROM `viewtat`
                    WHERE (DATE_FORMAT(`viewtat`.`ReleaseDateTime`, '%Y-%m-%d') BETWEEN @DateFrom AND @DateTo) AND `viewtat`.`PerformedBy` = @MedTech ;"

            Dim dt As DataTable = New DataTable
            rs.Connection = conn
            rs.CommandType = CommandType.Text
            rs.CommandText = sql
            Dim adapter As New MySqlDataAdapter(rs)

            adapter.Fill(dt)
            dtResult.DataSource = (dt)
            Disconnect()
        ElseIf cboMedtech.Text = "All" And cboLimit.Text <> "All" Then
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
            dtResult.DataSource = (dt)
            Disconnect()
        ElseIf cboMedtech.Text <> "All" And cboLimit.Text <> "All" Then
            Connect()
            rs.Parameters.Clear()
            rs.Parameters.AddWithValue("@DateFrom", Format(dtFrom.Value, "yyyy-MM-dd"))
            rs.Parameters.AddWithValue("@DateTo", Format(dtTo.Value, "yyyy-MM-dd"))
            rs.Parameters.AddWithValue("@MedTech", cboMedtech.Text)
            rs.Parameters.AddWithValue("@Section", cboLimit.Text)

            Dim sql As String
            sql = "SELECT
						*
				    FROM `viewtat`
					WHERE (DATE_FORMAT(`viewtat`.`ReleaseDateTime`, '%Y-%m-%d') BETWEEN @DateFrom AND @DateTo) AND `viewtat`.`Section` = @Section AND `viewtat`.`PerformedBy` = @MedTech;"

            Dim dt As DataTable = New DataTable
            rs.Connection = conn
            rs.CommandType = CommandType.Text
            rs.CommandText = sql
            Dim adapter As New MySqlDataAdapter(rs)

            adapter.Fill(dt)
            dtResult.DataSource = (dt)
            Disconnect()
        End If

        GridView.OptionsBehavior.Editable = False
        GridView.OptionsSelection.EnableAppearanceFocusedCell = False
        GridView.FocusRectStyle = DrawFocusRectStyle.RowFullFocus

        GridView.Columns("PatientName").Summary.Add(DevExpress.Data.SummaryItemType.Count, "PatientName", "Count={0}")

        lblDiffTime.Text = "Total No. Performed: " & GridView.Columns("PatientName").SummaryItem.SummaryValue.ToString

    End Sub

End Class