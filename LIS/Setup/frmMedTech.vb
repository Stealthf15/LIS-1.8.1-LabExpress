Imports DevExpress.XtraGrid.Views.Grid

Public Class frmMedTech

    Public Sub LoadRecords()
        Try
            GridView.Columns.Clear()
            GridView.Appearance.HeaderPanel.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center
            GridView.Appearance.HeaderPanel.FontStyleDelta = FontStyle.Bold

            Dim SQL As String = "SELECT
                        `medtech`.`id` AS SequenceNo,
                        `medtech`.`fname` AS FirstName,
                        `medtech`.`mname` AS MiddleInitial,
                        `medtech`.`lname` AS LastName,
                        `medtech`.`license` AS LicenseNo, 
                        `medtech`.`designation` AS Designation,
                        `medtech`.`signature` AS Signature,
                        `medtech`.`verificator` AS Verificator FROM `medtech`"

            Dim command As New MySql.Data.MySqlClient.MySqlCommand(SQL, conn)

            Dim adapter As New MySql.Data.MySqlClient.MySqlDataAdapter(command)

            Dim myTable As DataTable = New DataTable
            adapter.Fill(myTable)

            dtList.DataSource = myTable

            GridView.Columns("Verificator").Visible = False

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

    Private Sub GridView_RowCellStyle(sender As Object, e As RowCellStyleEventArgs) Handles GridView.RowCellStyle
        Dim view As GridView = TryCast(sender, GridView)
        If (e.Column.FieldName = "SequenceNo") Then
            If view.GetRowCellValue(e.RowHandle, "Verificator") = 1 Then
                e.Appearance.BackColor = Color.Gold
                e.Appearance.BackColor2 = Color.Gold
                e.Appearance.ForeColor = Color.Black
            End If
        End If
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
            DeleteRecord("medtech", "id", GridView.GetFocusedRowCellValue(GridView.Columns("SequenceNo")))
            LoadRecords()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "System Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub btnEdit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnEdit.ItemClick

        frmMedTechAE.ID = GridView.GetFocusedRowCellValue(GridView.Columns("SequenceNo"))
        frmMedTechAE.txtFName.Text = GridView.GetFocusedRowCellValue(GridView.Columns("FirstName"))
        frmMedTechAE.txtMName.Text = GridView.GetFocusedRowCellValue(GridView.Columns("MiddleInitial"))
        frmMedTechAE.txtLName.Text = GridView.GetFocusedRowCellValue(GridView.Columns("LastName"))
        frmMedTechAE.txtLicense.Text = GridView.GetFocusedRowCellValue(GridView.Columns("LicenseNo"))
        frmMedTechAE.txtDesign.Text = GridView.GetFocusedRowCellValue(GridView.Columns("Designation"))

        If GridView.GetFocusedRowCellValue(GridView.Columns("Verificator")) Then
            frmMedTechAE.chkVerify.Checked = True
        Else
            frmMedTechAE.chkVerify.Checked = False
        End If

        frmMedTechAE.btnSave.Text = "&Update"
        frmMedTechAE.ShowDialog()
    End Sub

    Private Sub btnRefresh_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnRefresh.ItemClick
        LoadRecords()
    End Sub

    Private Sub btnAdd_ItemClick(ByVal sender As System.Object, ByVal e As DevExpress.XtraBars.ItemClickEventArgs) Handles btnAdd.ItemClick
        frmMedTechAE.btnSave.Text = "&Save"
        frmMedTechAE.ShowDialog()
    End Sub

End Class