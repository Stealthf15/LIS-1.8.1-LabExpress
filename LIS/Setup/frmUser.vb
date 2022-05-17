Imports DevExpress.XtraGrid.Views.Grid

Public Class frmUser

    Public Sub LoadRecords()
        Try
            GridView.Columns.Clear()
            GridView.Appearance.HeaderPanel.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center
            GridView.Appearance.HeaderPanel.FontStyleDelta = FontStyle.Bold

            Dim SQL As String = "SELECT
                        `users`.`id` AS SequenceNo,
                        `users`.`fname` AS FirstName,
                        `users`.`mname` AS MiddleInitial,
                        `users`.`lname` AS LastName,
                        `users`.`license` AS LicenseNo, 
                        `users`.`designation` AS Designation FROM `users`"

            Dim command As New MySql.Data.MySqlClient.MySqlCommand(SQL, conn)

            Dim adapter As New MySql.Data.MySqlClient.MySqlDataAdapter(command)

            Dim myTable As DataTable = New DataTable
            adapter.Fill(myTable)

            dtList.DataSource = myTable

            ' Make the grid read-only. 
            GridView.OptionsBehavior.Editable = False
            ' Prevent the focused cell from being highlighted. 
            GridView.OptionsSelection.EnableAppearanceFocusedCell = False
            ' Draw a dotted focus rectangle around the entire row. 
            GridView.FocusRectStyle = DrawFocusRectStyle.RowFullFocus

        Catch ex As Exception
            MessageBox.Show(ex.Message, "System Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub frmMachineList_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        LoadRecords()
    End Sub

    Private Sub SimpleButton2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnClose.ItemClick
        MainFOrm.lblTitle.Text = ""
        Me.Close()
        Me.Dispose()
    End Sub

    Private Sub btnDelete_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnDelete.ItemClick
        Try
            DeleteRecord("users", "id", GridView.GetFocusedRowCellValue(GridView.Columns("SequenceNo")))
            LoadRecords()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "System Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub btnEdit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnEdit.ItemClick

        frmUserAE.ID = GridView.GetFocusedRowCellValue(GridView.Columns("SequenceNo"))
        frmUserAE.txtFName.Text = GridView.GetFocusedRowCellValue(GridView.Columns("FirstName"))
        frmUserAE.txtMName.Text = GridView.GetFocusedRowCellValue(GridView.Columns("MiddleInitial"))
        frmUserAE.txtLName.Text = GridView.GetFocusedRowCellValue(GridView.Columns("LastName"))
        frmUserAE.txtLicense.Text = GridView.GetFocusedRowCellValue(GridView.Columns("LicenseNo"))
        frmUserAE.txtDesign.Text = GridView.GetFocusedRowCellValue(GridView.Columns("Designation"))

        frmUserAE.btnSave.Text = "&Update"
        frmUserAE.ShowDialog()
    End Sub

    Private Sub btnRefresh_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnRefresh.ItemClick
        LoadRecords()
    End Sub

    Private Sub btnAdd_ItemClick(ByVal sender As System.Object, ByVal e As DevExpress.XtraBars.ItemClickEventArgs) Handles btnAdd.ItemClick
        frmUserAE.btnSave.Text = "&Save"
        frmUserAE.ShowDialog()
    End Sub

End Class