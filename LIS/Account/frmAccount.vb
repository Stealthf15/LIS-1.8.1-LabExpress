Imports DevExpress.XtraGrid.Views.Grid

Public Class frmAccount

    Public Sub LoadRecords()
        Try
            GridView.Columns.Clear()
            GridView.Appearance.HeaderPanel.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center
            GridView.Appearance.HeaderPanel.FontStyleDelta = FontStyle.Bold

            Dim SQL As String = "SELECT `user_id` AS SequenceNo, `email` AS UserID, `name` AS User, `username` AS Username, `password` AS Password, `usertype` AS UserType FROM `user_account` ORDER BY `user_id`"
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
            'DeleteRecord("user_account", "user_id", lvList.FocusedItem.SubItems(0).Text)
            LoadRecords()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "System Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub btnRefresh_ItemClick(ByVal sender As Object, ByVal e As DevExpress.XtraBars.ItemClickEventArgs) Handles btnRefresh.ItemClick
        LoadRecords()
    End Sub

    Private Sub btnAdd_ItemClick(ByVal sender As Object, ByVal e As DevExpress.XtraBars.ItemClickEventArgs) Handles btnAdd.ItemClick
        frmAccountA.ShowDialog()
    End Sub

    Private Sub BarLargeButtonItem1_ItemClick(sender As Object, e As DevExpress.XtraBars.ItemClickEventArgs) Handles BarLargeButtonItem1.ItemClick
        Dim bytes() As Byte = System.Text.Encoding.UTF8.GetBytes(GridView.GetFocusedRowCellValue(GridView.Columns("Password")))
        Dim hashOfBytes() As Byte = New System.Security.Cryptography.SHA1Managed().ComputeHash(bytes)
        Dim strHash As String = Convert.ToBase64String(hashOfBytes)

        frmAccountEdit.ID_NO = GridView.GetFocusedRowCellValue(GridView.Columns("UserID"))

        frmAccountEdit.cboMedTech.Text = GridView.GetFocusedRowCellValue(GridView.Columns("User"))
        frmAccountEdit.txtUsername.Text = GridView.GetFocusedRowCellValue(GridView.Columns("Username"))
        frmAccountEdit.cboUserType.Text = GridView.GetFocusedRowCellValue(GridView.Columns("UserType"))
        frmAccountEdit.Password = GridView.GetFocusedRowCellValue(GridView.Columns("Password"))
        frmAccountEdit.cboLocation.Text = "Laboratory"
        'frmAccountEdit.cboEmailAccess.Text = Me.lvList.FocusedItem.SubItems(1).Text

        frmAccountEdit.ShowDialog()
    End Sub
End Class