Public Class frmOrder

    Public Sub LoadRecords()
        Try
            LoadRecordsOnLVSQL(lvList, "SELECT `id`, `specimen`, `test_code` FROM `default_order` WHERE `section` = '" & cboLimit.Text & "'", 2)
        Catch ex As Exception
            MessageBox.Show(ex.Message, "System Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub frmTestType_FormClosed(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
        Me.Dispose()
    End Sub

    Private Sub frmMachineList_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        AutoLoadTestName()
    End Sub

    Private Sub SimpleButton2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnClose.ItemClick
        Me.Close()
    End Sub

    Private Sub txtSearch_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtSearch.ButtonClick
        Try
            LoadRecordsOnLVSQL(lvList, "SELECT `id`, `specimen`, `test_code` FROM `default_order` WHERE `section` = '" & cboLimit.Text & "' AND `specimen` = '" & txtSearch.Text & "%'", 2)
        Catch ex As Exception
            MessageBox.Show(ex.Message, "System Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub btnDelete_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnDelete.ItemClick
        Try
            DeleteRecord("default_order", "id", lvList.FocusedItem.SubItems(0).Text)
            LoadRecords()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "System Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub btnEdit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnEdit.ItemClick
        If lvList.SelectedItems.Count > 0 Then
            frmOrderAE.AutoLoadCombo()
            frmOrderAE.name = lvList.FocusedItem.SubItems(0).Text
            frmOrderAE.cboTestName.Text = cboLimit.Text
            frmOrderAE.cboTest.Text = lvList.FocusedItem.SubItems(1).Text
            frmOrderAE.txtTestCode.Text = lvList.FocusedItem.SubItems(2).Text
            frmOrderAE.btnSave.Text = "&Update"
            frmOrderAE.ShowDialog()
        End If
    End Sub

    Private Sub btnRefresh_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnRefresh.ItemClick
        LoadRecords()
    End Sub

    Private Sub btnAdd_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnAdd.ItemClick
        frmOrderAE.AutoLoadCombo()
        frmOrderAE.cboTestName.Text = cboLimit.Text

        frmOrderAE.btnSave.Text = "&Save"
        frmOrderAE.ShowDialog()
    End Sub

    Public Sub AutoLoadTestName()
        cboLimit.Properties.Items.Clear()
        Connect()
        rs.Connection = conn
        rs.CommandText = "SELECT * FROM `testtype` WHERE `test_name` NOT LIKE 'All' ORDER BY `test_name`"
        reader = rs.ExecuteReader
        While reader.Read
            cboLimit.Properties.Items.Add(reader(1).ToString)
        End While
        Disconnect()
    End Sub

    Private Sub lvList_DrawColumnHeader(ByVal sender As Object, ByVal e As System.Windows.Forms.DrawListViewColumnHeaderEventArgs) Handles lvList.DrawColumnHeader
        Using sf As StringFormat =
            New StringFormat()

            ' Store the column text alignment, letting it default
            ' to Left if it has not been set to Center or Right.
            Select Case (e.Header.TextAlign)

                Case HorizontalAlignment.Center
                    sf.Alignment = StringAlignment.Center

                Case HorizontalAlignment.Right
                    sf.Alignment = StringAlignment.Far

            End Select

            ' Draw the standard header background.
            e.DrawBackground()

            ' Draw the header text.
            Using headerFont As Font =
                New Font("Helvetica", 10, FontStyle.Bold)

                e.Graphics.DrawString(e.Header.Text, headerFont,
                    Brushes.Black, e.Bounds, sf)
            End Using
        End Using
        Return
    End Sub

    Private Sub cboLimit_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cboLimit.SelectedIndexChanged
        LoadRecords()
    End Sub
End Class