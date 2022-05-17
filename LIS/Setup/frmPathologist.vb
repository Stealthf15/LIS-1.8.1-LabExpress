Imports DevExpress.XtraGrid.Views.Grid

Public Class frmPathologist

    Public Sub LoadRecords()
        Try
            GridView.Columns.Clear()
            GridView.Appearance.HeaderPanel.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center
            GridView.Appearance.HeaderPanel.FontStyleDelta = FontStyle.Bold

            Dim SQL As String = "SELECT
                        `pathologist`.`id` AS SequenceNo,
                        `pathologist`.`fname` AS FirstName,
                        `pathologist`.`mname` AS MiddleInitial,
                        `pathologist`.`lname` AS LastName,
                        `pathologist`.`license` AS LicenseNo, 
                        `pathologist`.`designation` AS Designation,
                        `pathologist`.`img` AS Signature FROM `pathologist`"

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
        Me.Close()
        Me.Dispose()
    End Sub

    Private Sub btnDelete_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnDelete.ItemClick
        Try
            Try
                DeleteRecord("pathologist", "id", GridView.GetFocusedRowCellValue(GridView.Columns("SequenceNo")))
                LoadRecords()
            Catch ex As Exception
                MessageBox.Show(ex.Message, "System Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End Try
            LoadRecords()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "System Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub btnEdit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnEdit.ItemClick
        frmPathologistAE.ID = GridView.GetFocusedRowCellValue(GridView.Columns("SequenceNo"))
        frmPathologistAE.txtFName.Text = GridView.GetFocusedRowCellValue(GridView.Columns("FirstName"))
        frmPathologistAE.txtMName.Text = GridView.GetFocusedRowCellValue(GridView.Columns("MiddleInitial"))
        frmPathologistAE.txtLName.Text = GridView.GetFocusedRowCellValue(GridView.Columns("LastName"))
        frmPathologistAE.txtLicense.Text = GridView.GetFocusedRowCellValue(GridView.Columns("LicenseNo"))
        frmPathologistAE.txtDesign.Text = GridView.GetFocusedRowCellValue(GridView.Columns("Designation"))

        frmPathologistAE.btnSave.Text = "&Update"
        frmPathologistAE.ShowDialog()
    End Sub

    Private Sub btnRefresh_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnRefresh.ItemClick
        LoadRecords()
    End Sub

    Private Sub btnAdd_ItemClick(ByVal sender As Object, ByVal e As DevExpress.XtraBars.ItemClickEventArgs) Handles btnAdd.ItemClick
        frmPathologistAE.btnSave.Text = "&Save"
        frmPathologistAE.ShowDialog()
    End Sub
End Class