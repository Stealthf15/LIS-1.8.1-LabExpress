Imports DevExpress.Utils
Imports DevExpress.XtraGrid.Views.Grid
Imports MySql.Data.MySqlClient

Public Class frmTaTRealtime

    Private Sub btnClose_ItemClick(ByVal sender As Object, ByVal e As DevExpress.XtraBars.ItemClickEventArgs) Handles btnClose.ItemClick
        Me.Close()
        Me.Dispose()

    End Sub

    Private Sub LoadTaTTracker()
        Try
            dtResult.RefreshDataSource()

            GridView.Columns.Clear()
            GridView.Appearance.HeaderPanel.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center
            GridView.Appearance.HeaderPanel.FontStyleDelta = FontStyle.Bold
            GridView.Appearance.HeaderPanel.Font = New Font("Tahoma", 10)

            GridView.Appearance.OddRow.BackColor = Color.FromArgb(226, 236, 247)
            GridView.OptionsView.EnableAppearanceOddRow = True
            GridView.Appearance.EvenRow.BackColor = Color.White
            GridView.OptionsView.EnableAppearanceEvenRow = True

            Connect()
            rs.Parameters.Clear()
            Dim sql As String
            sql = "Tat_Tracker_Detailed"

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

            GridView.Columns("DateRelease").Visible = False

            ' Make the grid read-only. 
            GridView.OptionsBehavior.Editable = False
            ' Prevent the focused cell from being highlighted. 
            GridView.OptionsSelection.EnableAppearanceFocusedCell = False
            ' Draw a dotted focus rectangle around the entire row. 
            GridView.FocusRectStyle = DrawFocusRectStyle.RowFullFocus

            GridView.OptionsView.BestFitMode = GridBestFitMode.Default
            GridView.OptionsView.ColumnAutoWidth = False

        Catch ex As Exception
            Exit Sub
            MessageBox.Show(ex.Message, "Exception", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
        End Try
    End Sub

    Private Sub GridView_RowCellStyle(sender As Object, e As RowCellStyleEventArgs) Handles GridView.RowCellStyle
        Dim view As GridView = TryCast(sender, GridView)
        If view.GetRowCellValue(e.RowHandle, "Status") = "Delayed" Then
            e.Appearance.BackColor = If(True, Color.Crimson, Color.LightSalmon)
            e.Appearance.ForeColor = Color.White
        End If
    End Sub

    Private Sub frmTaTRealtime_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        LoadTaTTracker()
    End Sub
End Class