Public Class frmMachineList

    Public Sub LoadRecords()
        Try
            LoadRecordsOnLVSQL(lvList, "SELECT `id`, `name`, `serial_no`, `com_port`, `baud_rate`, `test_name`, `status` FROM `machines`", 6)
        Catch ex As Exception
            MessageBox.Show(ex.Message, "System Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub frmMachineList_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load
        LoadRecords()
    End Sub

    Private Sub btnBack_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnClose.ItemClick
        Me.Close()
        Me.Dispose()
    End Sub

    Private Sub SimpleButton1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSelect.ItemClick, lvList.DoubleClick
        If lvList.SelectedItems.Count > 0 Then
            Connect()
            rs.Connection = conn
            rs.CommandText = "SELECT `location` FROM `machines` WHERE `id` = '" & lvList.FocusedItem.Text & "'"
            reader = rs.ExecuteReader
            reader.Read()
            If reader.HasRows Then
                Process.Start(reader(0).ToString)
            End If
            Disconnect()

            Connect()
            rs.Connection = conn
            rs.CommandText = "UPDATE `machines` SET " _
                & "`status` = 'Connected'" _
                & " WHERE `id` = '" & lvList.FocusedItem.SubItems(0).Text & "'"
            rs.ExecuteNonQuery()
            Disconnect()

            LoadRecords()
        End If
    End Sub

    'Private Sub lvList_ItemSelectionChanged(ByVal sender As Object, ByVal e As System.Windows.Forms.ListViewItemSelectionChangedEventArgs) Handles lvList.ItemSelectionChanged
    '    btnSelect.
    'End Sub
End Class