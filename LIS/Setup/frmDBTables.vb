Public Class frmDBTables

    Public Sub LoadRecords()
        Try
            LoadRecordsOnLVSQL(lvList, "SHOW TABLES", 0)
        Catch ex As Exception
            MessageBox.Show(ex.Message, "System Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub frmDBTables_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        LoadRecords()
    End Sub

    Private Sub SimpleButton2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnClose.ItemClick
        Me.Close()
        Me.Dispose()
    End Sub

    Private Sub btnRefresh_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnRefresh.ItemClick
        LoadRecords()
    End Sub

    Private Sub btnTruncate_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnAdd.ItemClick
        If MessageBox.Show("Are you sure you want to Truncate the selected tables?" & vbNewLine & "This process will delete all existing records.", "Delete Table Data", MessageBoxButtons.YesNo, MessageBoxIcon.Question) = vbYes Then
            truncateTables()
        Else
            Exit Sub
        End If
    End Sub

    Public Sub truncateTables()
        For x As Integer = 0 To lvList.CheckedItems.Count - 1 Step 1
            Dim table As String = lvList.CheckedItems(x).Text
            Connect()
            rs.Connection = conn
            rs.CommandText = "TRUNCATE TABLE `" & table & "`"
            rs.ExecuteNonQuery()
            Disconnect()
        Next
    End Sub
End Class