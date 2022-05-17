Imports DevExpress.XtraGrid.Views.Grid

Public Class frmEmailPush

    Public Sub LoadRecords()
        Try
            GridView.Columns.Clear()
            GridView.Appearance.HeaderPanel.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center
            GridView.Appearance.HeaderPanel.FontStyleDelta = FontStyle.Bold

            Dim SQL As String = "email_push"

            Dim command As New MySql.Data.MySqlClient.MySqlCommand(SQL, conn)
            command.CommandType = CommandType.StoredProcedure
            command.Parameters.Clear()
            command.Parameters.Add("@DateFrom", MySql.Data.MySqlClient.MySqlDbType.DateTime).Value = Format(dtFrom1.Value, "yyyy-MM-dd")
            command.Parameters.Add("@DateTo", MySql.Data.MySqlClient.MySqlDbType.DateTime).Value = Format(dtTo1.Value, "yyyy-MM-dd")
            command.Parameters.AddWithValue("@Section", cboSection.Text)

            Dim adapter As New MySql.Data.MySqlClient.MySqlDataAdapter(command)

            Dim myTable As DataTable = New DataTable
            adapter.Fill(myTable)

            dtList.DataSource = myTable

            GridView.Columns("Status").Visible = False
            GridView.Columns("Section").Visible = False
            GridView.Columns("SubSection").Visible = False
            GridView.Columns("PDFLocation").Visible = False

            ' Make the grid read-only. 
            GridView.OptionsBehavior.Editable = False
            ' Prevent the focused cell from being highlighted. 
            GridView.OptionsSelection.EnableAppearanceFocusedCell = False
            ' Draw a dotted focus rectangle around the entire row. 
            GridView.FocusRectStyle = DrawFocusRectStyle.RowFullFocus

            GridView.OptionsSelection.MultiSelect = True
            GridView.OptionsSelection.MultiSelectMode = GridMultiSelectMode.CheckBoxRowSelect

        Catch ex As Exception
            MessageBox.Show(ex.Message, "System Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub GridView_RowCellStyle(sender As Object, e As RowCellStyleEventArgs) Handles GridView.RowCellStyle
        Dim view As GridView = TryCast(sender, GridView)
        If (e.Column.FieldName = "SampleID") Or (e.Column.FieldName = "PatientID") Then
            If view.GetRowCellValue(e.RowHandle, "Status") = "1" Then
                e.Appearance.BackColor = Color.ForestGreen
                e.Appearance.BackColor2 = Color.ForestGreen
                e.Appearance.ForeColor = Color.White
            End If
        End If
    End Sub

    Private Sub frmEmailPush_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        LoadRecords()
        AutoLoadTestName()
    End Sub

    Private Sub btnSendResult_ItemClick(sender As Object, e As DevExpress.XtraBars.ItemClickEventArgs) Handles btnSendResult.ItemClick
        sendmail(GridView.GetFocusedRowCellValue("EmailAddress"),
                 GridView.GetFocusedRowCellValue("PatientName"),
                 GridView.GetFocusedRowCellValue("PDFLocation") & GridView.GetFocusedRowCellValue("SampleID") & "_" & GridView.GetFocusedRowCellValue("PatientName") & ".PDF")

        '------------------Save Email Details------------------------------
        rs.Parameters.Clear()
        rs.Parameters.AddWithValue("@SampleID", GridView.GetFocusedRowCellValue("SampleID"))
        rs.Parameters.AddWithValue("@Section", GridView.GetFocusedRowCellValue("Section"))
        rs.Parameters.AddWithValue("@SubSection", GridView.GetFocusedRowCellValue("SubSection"))
        Connect()
        rs.Connection = conn
        rs.CommandType = CommandType.Text
        rs.CommandText = ("UPDATE `email_details` SET `status` = 1 WHERE `sample_id` = @SampleID AND `section` = @Section AND `sub_section` = @SubSection")
        rs.ExecuteNonQuery()
        Disconnect()
        '------------------Save Email Details------------------------------

        LoadRecords()
    End Sub

    Public Sub AutoLoadTestName()
        cboSection.Properties.Items.Clear()
        Connect()
        rs.Connection = conn
        rs.CommandType = CommandType.Text
        rs.CommandText = "SELECT `test_name` FROM `testtype` WHERE `test_name` NOT LIKE 'All' ORDER BY `test_name`"
        reader = rs.ExecuteReader
        While reader.Read
            cboSection.Properties.Items.Add(reader(0).ToString)
        End While
        Disconnect()

        cboSection.SelectedIndex = 0
    End Sub

    Private Sub btnSearch_Click(sender As Object, e As EventArgs) Handles btnSearch.Click
        LoadRecords()
    End Sub
End Class