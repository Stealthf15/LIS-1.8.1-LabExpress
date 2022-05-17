Imports System.Text.RegularExpressions
Imports System
Imports System.Collections.Generic
Imports System.Linq
Imports System.Drawing
Imports DevExpress.XtraEditors
Imports DevExpress.XtraPrinting.BarCode

Public Class frmPatientOrdersample

    'Public ID As String = ""
    'Dim arrImage() As Byte

    'Private Sub frmResultNew_FormClosed(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
    '    Me.Dispose()
    'End Sub

    'Private Sub frmNewOrderAE_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
    '    AutoPatientID()
    '    AutoLoadRoom()
    '    AutoLoadDoctor()

    '    Connect()
    '    rs.Connection = conn
    '    rs.CommandText = "SELECT `specimen`, `test_code` FROM `default_specimen` WHERE `test_group` = 'Hematology'"
    '    reader = rs.ExecuteReader
    '    While reader.Read
    '        If reader.HasRows Then
    '            tvHematology.Nodes.Add(reader(0).ToString & " (" & reader(1).ToString & ")")
    '        End If
    '    End While
    '    Disconnect()

    '    Connect()
    '    rs.Connection = conn
    '    rs.CommandText = "SELECT `specimen`, `test_code` FROM `default_specimen` WHERE `test_group` = 'Chemistry'"
    '    reader = rs.ExecuteReader
    '    While reader.Read
    '        If reader.HasRows Then
    '            tvChemistry.Nodes.Add(reader(0).ToString & " (" & reader(1).ToString & ")")
    '        End If
    '    End While
    '    Disconnect()

    '    Connect()
    '    rs.Connection = conn
    '    rs.CommandText = "SELECT `specimen`, `test_code` FROM `default_specimen` WHERE `test_group` = 'Electrolytes'"
    '    reader = rs.ExecuteReader
    '    While reader.Read
    '        If reader.HasRows Then
    '            tvElectrolytes.Nodes.Add(reader(0).ToString & " (" & reader(1).ToString & ")")
    '        End If
    '    End While
    '    Disconnect()
    'End Sub

    'Private Sub btnClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnClose.ItemClick
    '    Me.Close()
    '    Me.Dispose()
    'End Sub

    'Private Sub txtAge_KeyPress(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtAge.KeyPress
    '    If InStr("1234567890.", e.KeyChar) = 0 And Not Chr(AscW(e.KeyChar)) = vbBack Then
    '        e.KeyChar = ChrW(0)
    '    End If
    'End Sub

    'Private Sub dtBDate_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles dtBDate.EditValueChanged
    '    Dim Birthday As Date = dtBDate.DateTime
    '    Dim endDate As Date = Date.Now

    '    Dim timeSpan As TimeSpan = endDate.Subtract(Birthday)
    '    Dim totalDays As Integer = timeSpan.TotalDays
    '    Dim totalYears As Integer = Math.Truncate(totalDays / 365)
    '    Dim totalMonths As Integer = Math.Truncate((totalDays Mod 365) / 30)
    '    Dim remainingDays As Integer = Math.Truncate((totalDays Mod 365) Mod 30)

    '    If Birthday.Year = endDate.Year Then
    '        If Birthday.Month = endDate.Month Then
    '            If Birthday.Day = endDate.Day Then
    '                txtAge.Text = remainingDays.ToString
    '                txtClass.Text = "Day(s)"
    '            ElseIf Birthday.Day < endDate.Day Then
    '                txtAge.Text = remainingDays.ToString
    '                txtClass.Text = "Day(s)"
    '            Else
    '                txtAge.Text = "0"
    '                txtClass.Text = "Day(s)"
    '            End If
    '        ElseIf Birthday.Month < endDate.Month Then
    '            txtAge.Text = totalMonths.ToString
    '            txtClass.Text = "Month(s)"
    '        Else
    '            txtAge.Text = "0"
    '            txtClass.Text = "Month(s)"
    '        End If
    '    ElseIf Birthday.Year < endDate.Year Then
    '        txtAge.Text = totalYears.ToString
    '        txtClass.Text = "Year(s)"
    '    Else
    '        txtAge.Text = "0"
    '        txtClass.Text = "Year(s)"
    '    End If
    'End Sub

    'Public Sub AutoPatientID()
    '    Try
    '        Connect()
    '        rs.Connection = conn
    '        rs.CommandText = "SELECT (id) as id FROM `patient_info` ORDER BY `id` DESC"
    '        reader = rs.ExecuteReader
    '        reader.Read()
    '        If reader.HasRows Then
    '            If reader(0).ToString > 0 And reader(0).ToString <= 9 Then
    '                txtPatientID.Text = Now.Year & "-" & Now.Day & "-" & "000000" & reader(0).ToString + 1
    '            ElseIf reader(0).ToString > 9 And reader(0).ToString <= 99 Then
    '                txtPatientID.Text = Now.Year & "-" & Now.Day & "-" & "00000" & reader(0).ToString + 1
    '            ElseIf reader(0).ToString > 99 And reader(0).ToString <= 999 Then
    '                txtPatientID.Text = Now.Year & "-" & Now.Day & "-" & "0000" & reader(0).ToString + 1
    '            ElseIf reader(0).ToString > 999 And reader(0).ToString <= 9999 Then
    '                txtPatientID.Text = Now.Year & "-" & Now.Day & "-" & "000" & reader(0).ToString + 1
    '            ElseIf reader(0).ToString > 9999 And reader(0).ToString <= 99999 Then
    '                txtPatientID.Text = Now.Year & "-" & Now.Day & "-" & "00" & reader(0).ToString + 1
    '            ElseIf reader(0).ToString > 99999 And reader(0).ToString <= 999999 Then
    '                txtPatientID.Text = Now.Year & "-" & Now.Day & "-" & "0" & reader(0).ToString + 1
    '            ElseIf reader(0).ToString > 999999 And reader(0).ToString <= 9999999 Then
    '                txtPatientID.Text = reader(0).ToString + 1
    '            End If
    '        Else
    '            txtPatientID.Text = Now.Year & "-" & Now.Day & "-" & "000000" & 1
    '        End If
    '        Disconnect()
    '    Catch ex As Exception
    '        MessageBox.Show(ex.Message, "Auto Patient ID", MessageBoxButtons.OK, MessageBoxIcon.Error)
    '        Disconnect()
    '        Exit Sub
    '    End Try
    'End Sub

    'Private Sub SimpleButton3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles SimpleButton3.Click
    '    frmPatientList.Type = "Order"
    '    frmPatientList.ShowDialog()
    'End Sub

    'Private Sub btnSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSave.ItemClick
    '    If Me.txtName.Text = "" Or Me.txtAge.Text = "" Or txtPatientID.Text = "" Or cboSex.Text = "" Then
    '        MessageBox.Show("Please verify the data carefully.", "Verification Failed", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
    '        Exit Sub
    '    End If

    '    Me.Close()
    '    Me.Dispose()
    'End Sub

    'Public Sub AutoLoadRoom()
    '    Me.cboDept.Properties.Items.Clear()
    '    Connect()
    '    rs.Connection = conn
    '    rs.CommandText = "SELECT * FROM `department` ORDER BY `dept_name`"
    '    reader = rs.ExecuteReader
    '    While reader.Read
    '        Me.cboDept.Properties.Items.Add(reader(1).ToString)
    '    End While
    '    Disconnect()
    'End Sub

    'Public Sub AutoLoadDoctor()
    '    Me.cboDoctor.Properties.Items.Clear()
    '    Connect()
    '    rs.Connection = conn
    '    rs.CommandText = "SELECT CONCAT(fname, ' ', mname, ' ', lname, ', ', designation) AS `name` FROM `doctors` ORDER BY `name`"
    '    reader = rs.ExecuteReader
    '    While reader.Read
    '        Me.cboDoctor.Properties.Items.Add(reader(0))
    '    End While
    '    Disconnect()
    'End Sub

    'Private Sub LoadSampleID()
    '    Connect()
    '    rs.Connection = conn
    '    rs.CommandText = "SELECT `id` FROM `tmpWorklist` ORDER BY `id` DESC LIMIT 1"
    '    reader = rs.ExecuteReader
    '    reader.Read()
    '    If reader.HasRows Then
    '        If reader(0).ToString > 0 And reader(0).ToString <= 9 Then
    '            SampleID = "000000" & reader(0).ToString + 1
    '        ElseIf reader(0).ToString > 9 And reader(0).ToString <= 99 Then
    '            SampleID = "00000" & reader(0).ToString + 1
    '        ElseIf reader(0).ToString > 99 And reader(0).ToString <= 999 Then
    '            SampleID = "0000" & reader(0).ToString + 1
    '        ElseIf reader(0).ToString > 999 And reader(0).ToString <= 9999 Then
    '            SampleID = "000" & reader(0).ToString + 1
    '        ElseIf reader(0).ToString > 9999 And reader(0).ToString <= 99999 Then
    '            SampleID = "00" & reader(0).ToString + 1
    '        ElseIf reader(0).ToString > 99999 And reader(0).ToString <= 999999 Then
    '            SampleID = "0" & reader(0).ToString + 1
    '        ElseIf reader(0).ToString > 999999 And reader(0).ToString <= 9999999 Then
    '            SampleID = reader(0).ToString + 1
    '        End If
    '    Else
    '        Disconnect()
    '        Connect()
    '        rs.Connection = conn
    '        rs.CommandText = "SELECT `id` FROM `order` ORDER BY `id` DESC LIMIT 1"
    '        reader = rs.ExecuteReader
    '        reader.Read()
    '        If reader.HasRows Then
    '            If reader(0).ToString > 0 And reader(0).ToString <= 9 Then
    '                SampleID = "000000" & reader(0).ToString + 1
    '            ElseIf reader(0).ToString > 9 And reader(0).ToString <= 99 Then
    '                SampleID = "00000" & reader(0).ToString + 1
    '            ElseIf reader(0).ToString > 99 And reader(0).ToString <= 999 Then
    '                SampleID = "0000" & reader(0).ToString + 1
    '            ElseIf reader(0).ToString > 999 And reader(0).ToString <= 9999 Then
    '                SampleID = "000" & reader(0).ToString + 1
    '            ElseIf reader(0).ToString > 9999 And reader(0).ToString <= 99999 Then
    '                SampleID = "00" & reader(0).ToString + 1
    '            ElseIf reader(0).ToString > 99999 And reader(0).ToString <= 999999 Then
    '                SampleID = "0" & reader(0).ToString + 1
    '            ElseIf reader(0).ToString > 999999 And reader(0).ToString <= 9999999 Then
    '                SampleID = reader(0).ToString + 1
    '            End If
    '        Else
    '            SampleID = "000000" & 1
    '        End If
    '        Disconnect()
    '    End If
    '    Disconnect()
    'End Sub

    'Private Sub SaveReceivedBy()
    '    rs.Parameters.Clear()
    '    rs.Parameters.AddWithValue("@Name", CurrUser)
    '    rs.Parameters.AddWithValue("@Date", Now)
    '    rs.Parameters.AddWithValue("@Sample_ID", SampleID)

    '    SaveRecordwthoutMSG("INSERT INTO `received_by` (`name`, `date`, `sample_id`) VALUES " _
    '                        & "(" _
    '                        & "@Name," _
    '                        & "@Date," _
    '                        & "@Sample_ID" _
    '                        & ")"
    '                        )
    'End Sub


    ''Private Sub btnRemoveAll_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnRemoveAll.ItemClick
    ''    If Me.lvList.CheckedItems.Count > 0 Then
    ''        For x As Integer = 0 To lvList.CheckedItems.Count - 1 Step 1
    ''            DeleteRecordSQL("DELETE FROM `patient_order` WHERE `sample_id` = '" & txtSampleID.Text & "'")
    ''            DeleteRecordSQL("DELETE FROM `tmpResult` WHERE `sample_id` = '" & txtSampleID.Text & "'")
    ''        Next
    ''        LoadTest()
    ''    End If
    ''End Sub

    ''Private Sub btnRemove_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnRemove.ItemClick
    ''    If Me.lvList.CheckedItems.Count > 0 Then
    ''        For x As Integer = 0 To lvList.CheckedItems.Count - 1 Step 1
    ''            DeleteRecordSQL("DELETE FROM `patient_order` WHERE `id` = '" & lvList.CheckedItems(x).SubItems(0).Text & "'")
    ''            DeleteRecordSQL("DELETE FROM `tmpResult` WHERE `sample_id` = '" & txtSampleID.Text & "'")
    ''        Next
    ''        LoadTest()
    ''    End If
    ''End Sub

    ''Private Sub btnPrint_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
    ''    Using myRDLCPrinter As New RDLCPrinterPWorkSheet(txtSampleID.Text, CurrUser, My.Settings.DefaultPrinter)
    ''        myRDLCPrinter.Print(1)
    ''    End Using
    ''    Me.Close()
    ''End Sub

    'Private Sub lvTest_ItemChecked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.ItemCheckedEventArgs)

    'End Sub

    'Private Sub tvHematology_AfterCheck(ByVal sender As Object, ByVal e As System.Windows.Forms.TreeViewEventArgs)
    '    'Dim listHema As List(Of String)
    '    lvList.Items.Clear()

    '    For Each tvChecked As TreeNode In tvHematology.Nodes
    '        If tvChecked.Checked Then
    '            lvList.Items.Add(tvChecked.Text)
    '        End If
    '    Next
    'End Sub

End Class