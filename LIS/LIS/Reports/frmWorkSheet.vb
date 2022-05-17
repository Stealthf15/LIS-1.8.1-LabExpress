Imports DevExpress.XtraGrid.Views.Grid
Imports MySql.Data.MySqlClient

Public Class frmWorkSheet
    Public MedTechID As String = ""
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
        '    GenerateReport(cboLimit.Text, dtFrom, dtTo, "ReportWorksheetHema.rdlc", RPTWorksheet.ReportViewer1)
        '    RPTWorksheet.ShowDialog()
        'ElseIf cboLimit.Text = "Chemistry" Then
        '    GenerateReport(cboLimit.Text, dtFrom, dtTo, "ReportWorksheetChem.rdlc", RPTWorksheet.ReportViewer1)
        '    RPTWorksheet.ShowDialog()
        'End If

        GenerateReport(cboLimit.Text, dtFrom, dtTo, cboLimit.Text & ".rdlc", RPTWorksheet.ReportViewer1)
        RPTWorksheet.ShowDialog()
    End Sub

    Private Sub btnClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnClose.ItemClick
        Me.Close()
    End Sub

    '########################################### SEARCHING USING DATES AND PATIENT TYPE ############################################################
    Private Sub btnSearch_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSearch.Click
        'Try
        GridView.Columns.Clear()

        Connect()
        rs.Parameters.Clear()
        rs.Parameters.AddWithValue("@DateFrom", Format(CDate(dtFrom.Value.ToShortDateString & " " & tmFrom.Value.ToLongTimeString), "yyyy-MM-dd hh:mm:ss tt"))
        rs.Parameters.AddWithValue("@DateTo", Format(CDate(dtTo.Value.ToShortDateString & " " & tmTo.Value.ToLongTimeString), "yyyy-MM-dd hh:mm:ss tt"))
        Dim sql As String
        sql = "worksheet_" & cboLimit.Text

        Dim dt As DataTable = New DataTable
        rs.Connection = conn
        rs.CommandType = CommandType.StoredProcedure
        rs.CommandText = sql
        Dim adapter As New MySqlDataAdapter(rs)

        adapter.Fill(dt)
        dtResult.DataSource = (dt)

        ' Make the grid read-only. 
        GridView.OptionsBehavior.Editable = False
        ' Prevent the focused cell from being highlighted. 
        GridView.OptionsSelection.EnableAppearanceFocusedCell = False
        ' Draw a dotted focus rectangle around the entire row. 
        GridView.FocusRectStyle = DrawFocusRectStyle.RowFullFocus
        GridView.OptionsView.BestFitMode = GridBestFitMode.Fast

        GridView.Columns("Instrument").Visible = False
        GridView.Columns("Section").Visible = False
        GridView.Columns("SubSection").Visible = False
        GridView.Columns("Date").Visible = False
        Disconnect()

        Me.lblCount.Text = "COUNT: " & Me.GridView.RowCount
    End Sub
    '################################################## END ###################################################################

    Private Sub frmWorkSheet_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        AutoLoadTestName()
        'AutoLoadMedtech()
        'AutoloadShift()

        tmFrom.Text = "01:00:00 AM"
        tmTo.Text = "11:59:00 PM"
    End Sub

End Class