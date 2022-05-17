Public Class frmPatientList

    Public Type As String = ""

    Public Sub LoadRecords()
        Try
            LoadRecordsOnLV(lvList, "`patient_info`", 8, "patient_id")
        Catch ex As Exception
            MessageBox.Show(ex.Message, "System Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub frmPatientList_FormClosed(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
        Me.Dispose()
    End Sub

    Private Sub frmNewOrder_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        LoadRecords()
    End Sub

    Private Sub SimpleButton2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSelect.ItemClick

        If Me.lvList.SelectedItems.Count > 0 Then
            frmPatientOrderAE.txtPatientID.Text = lvList.FocusedItem.SubItems(0).Text
            frmPatientOrderAE.txtName.Text = lvList.FocusedItem.SubItems(1).Text
            frmPatientOrderAE.cboSex.Text = lvList.FocusedItem.SubItems(2).Text
            frmPatientOrderAE.dtBDate.Text = lvList.FocusedItem.SubItems(3).Text
            frmPatientOrderAE.txtAge.Text = lvList.FocusedItem.SubItems(4).Text
            frmPatientOrderAE.txtAddress.Text = lvList.FocusedItem.SubItems(5).Text
            frmPatientOrderAE.txtContact.Text = lvList.FocusedItem.SubItems(6).Text
            frmPatientOrderAE.GetBDate()
            Me.Close()
            Me.Dispose()

        End If
    End Sub

    Private Sub btnClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnClose.ItemClick
        Me.Close()
    End Sub

    Private Sub btnRefresh_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnRefresh.ItemClick
        LoadRecords()
    End Sub

    Private Sub txtSearch_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtSearch.TextChanged
        'SearchRecordsOnLV(lvList, "`patient_info`", 8, "`name`", Me.txtSearch.Text, "patient_id")
        If rgSelect.SelectedIndex = 0 Then
            lvList.Items.Clear()
            Connect()
            rs.Connection = conn
            rs.CommandType = CommandType.Text
            rs.CommandText = "SELECT * FROM `patient_info` WHERE patient_id LIKE '" & txtSearch.Text & "%'"
            reader = rs.ExecuteReader
            While reader.Read
                iItem = New ListViewItem(reader(0).ToString, 0)
                iItem.Checked = False
                For x As Integer = 1 To 8 Step 1
                    iItem.SubItems.Add(reader(x).ToString())
                Next
                lvList.Items.Add(iItem)
            End While
            Disconnect()
        ElseIf rgSelect.SelectedIndex = 1 Then
            lvList.Items.Clear()
            Connect()
            rs.Connection = conn
            rs.CommandType = CommandType.Text
            rs.CommandText = "SELECT * FROM `patient_info` WHERE `name` LIKE '" & txtSearch.Text & "%'"
            reader = rs.ExecuteReader
            While reader.Read
                iItem = New ListViewItem(reader(0).ToString, 0)
                iItem.Checked = False
                For x As Integer = 1 To 8 Step 1
                    iItem.SubItems.Add(reader(x).ToString())
                Next
                lvList.Items.Add(iItem)
            End While
            Disconnect()
        End If

    End Sub

End Class