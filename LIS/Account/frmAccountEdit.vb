Imports DevExpress.XtraGrid.Views.Grid
Imports MySql.Data.MySqlClient

Public Class frmAccountEdit
    Public ID As String
    Public ID_NO As Integer = 0

    Public Password As String = ""

    Public Sub LoadPrivileges()
        'On Error Resume Next
        Try
            GridView.Columns.Clear()
            GridView.Appearance.HeaderPanel.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center
            GridView.Appearance.HeaderPanel.FontStyleDelta = FontStyle.Bold

            'GridView.Appearance.OddRow.BackColor = Color.Gainsboro
            'GridView.OptionsView.EnableAppearanceOddRow = True
            'GridView.Appearance.EvenRow.BackColor = Color.White
            'GridView.OptionsView.EnableAppearanceEvenRow = True

            Dim SQL As String = "SELECT `access_name` AS Modules, `access` AS Access FROM `user_access` WHERE `user_id` = '" & ID_NO & "'"

            Dim command As New mySqlCommand(SQL, conn)

            Dim adapter As New MySqlDataAdapter(command)

            Dim myTable As DataTable = New DataTable
            adapter.Fill(myTable)

            dtResult.DataSource = myTable

            GridView.Columns("Access").Visible = False

            ' Make the grid read-only. 
            'GridView.OptionsBehavior.Editable = False
            ' Prevent the focused cell from being highlighted. 
            GridView.OptionsSelection.EnableAppearanceFocusedCell = False
            ' Draw a dotted focus rectangle around the entire row. 
            GridView.FocusRectStyle = DrawFocusRectStyle.RowFullFocus

            For x As Integer = 0 To GridView.RowCount - 1 Step 1
                If GridView.GetRowCellValue(x, GridView.Columns("Access")) Then
                    GridView.SelectRow(x)
                Else

                End If
            Next

        Finally
            If conn.State = ConnectionState.Open Then
                conn.Close()
            End If
        End Try

    End Sub

    Public Sub LoadPermissions()
        'On Error Resume Next
        Try
            GridPermissions.Columns.Clear()
            GridPermissions.Appearance.HeaderPanel.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center
            GridPermissions.Appearance.HeaderPanel.FontStyleDelta = FontStyle.Bold

            'GridView.Appearance.OddRow.BackColor = Color.Gainsboro
            'GridView.OptionsView.EnableAppearanceOddRow = True
            'GridView.Appearance.EvenRow.BackColor = Color.White
            'GridView.OptionsView.EnableAppearanceEvenRow = True

            Dim SQL As String = "SELECT `access_name` AS Permissions, `access` AS Access FROM `user_permission` WHERE `user_id` = '" & ID_NO & "'"

            Dim command As New MySqlCommand(SQL, conn)

            Dim adapter As New MySqlDataAdapter(command)

            Dim myTable As DataTable = New DataTable
            adapter.Fill(myTable)

            dtPermissions.DataSource = myTable

            GridPermissions.Columns("Access").Visible = False

            ' Make the grid read-only. 
            'GridView.OptionsBehavior.Editable = False
            ' Prevent the focused cell from being highlighted. 
            GridPermissions.OptionsSelection.EnableAppearanceFocusedCell = False
            ' Draw a dotted focus rectangle around the entire row. 
            GridPermissions.FocusRectStyle = DrawFocusRectStyle.RowFullFocus

            For x As Integer = 0 To GridView.RowCount - 1 Step 1
                If GridPermissions.GetRowCellValue(x, GridPermissions.Columns("Access")) Then
                    GridPermissions.SelectRow(x)
                Else

                End If
            Next

        Finally
            If conn.State = ConnectionState.Open Then
                conn.Close()
            End If
        End Try

    End Sub

    Private Sub btnBack_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Me.Close()
        Me.Dispose()
    End Sub

    Private Sub btnSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSave.Click
        'Got this Encrypting code from StockOverFlow.com
        Dim bytes() As Byte = System.Text.Encoding.UTF8.GetBytes(txtPassword.Text)
        Dim hashOfBytes() As Byte = New System.Security.Cryptography.SHA1Managed().ComputeHash(bytes)
        Dim strHash As String = Convert.ToBase64String(hashOfBytes)

        If Me.txtUsername.Text = "" Or Me.cboMedTech.Text = "" Or Me.cboUserType.Text = "" Then
            MessageBox.Show("Please fill all required information.", "Empty String", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            Exit Sub
        End If

        rs.Parameters.Clear()
        rs.Parameters.AddWithValue("@name", cboMedTech.Text)
        rs.Parameters.AddWithValue("@username", txtUsername.Text)
        rs.Parameters.AddWithValue("@password", strHash)
        rs.Parameters.AddWithValue("@usertype", cboUserType.Text)
        rs.Parameters.AddWithValue("@MedTechID", ID_NO)
        rs.Parameters.AddWithValue("@location", cboLocation.Text)

        If txtPassword.Text = "" Then
            conn.Close()
            conn.ConnectionString = MyConnectionString
            conn.Open()
            rs.Connection = conn
            rs.CommandText = "UPDATE `user_account` SET `name` = @name, `username` = @username, `usertype` = @usertype, `email` = @MedTechID, `dept` = @location WHERE `email` = @MedTechID"
            rs.ExecuteNonQuery()
            Disconnect()
            MessageBox.Show("New record has been successfully save.", "Save", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Disconnect()
        Else
            If Password = strHash Then
                conn.Close()
                conn.ConnectionString = MyConnectionString
                conn.Open()
                rs.Connection = conn
                rs.CommandText = "UPDATE `user_account` SET `name` = @name, `username` = @username, `password` = @password, `usertype` = @usertype, `email` = @MedTechID, `dept` = @location WHERE `email` = @MedTechID"
                rs.ExecuteNonQuery()
                Disconnect()
                MessageBox.Show("New record has been successfully save.", "Save", MessageBoxButtons.OK, MessageBoxIcon.Information)
                Disconnect()
            Else
                Disconnect()
                MessageBox.Show("Password did not match.", "System Message", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
                Exit Sub
            End If
        End If


        For x As Integer = 0 To GridView.RowCount - 1 Step 1
            If (GridView.IsRowSelected(x)) Then
                UpdateRecordwthoutMSG("UPDATE user_access SET user_id = '" & ID_NO & "', access_name = '" & GridView.GetRowCellValue(x, "Modules") & "', access = 1 WHERE user_id = @MedTechID AND access_name = '" & GridView.GetRowCellValue(x, "Modules") & "'")
            Else
                UpdateRecordwthoutMSG("UPDATE user_access SET user_id = '" & ID_NO & "', access_name = '" & GridView.GetRowCellValue(x, "Modules") & "', access = 0 WHERE user_id = @MedTechID AND access_name = '" & GridView.GetRowCellValue(x, "Modules") & "'")
            End If
        Next

        For x As Integer = 0 To GridPermissions.RowCount - 1 Step 1
            If (GridPermissions.IsRowSelected(x)) Then
                UpdateRecordwthoutMSG("UPDATE user_permission SET user_id = '" & ID_NO & "', access_name = '" & GridPermissions.GetRowCellValue(x, "Permissions") & "', access = 1 WHERE user_id = @MedTechID AND access_name = '" & GridPermissions.GetRowCellValue(x, "Permissions") & "'")
            Else
                UpdateRecordwthoutMSG("UPDATE user_permission SET user_id = '" & ID_NO & "', access_name = '" & GridPermissions.GetRowCellValue(x, "Permissions") & "', access = 0 WHERE user_id = @MedTechID AND access_name = '" & GridPermissions.GetRowCellValue(x, "Permissions") & "'")
            End If
        Next

        Me.Close()
        frmAccount.LoadRecords()
    End Sub

    Private Sub cboMedTech_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cboMedTech.SelectedIndexChanged
        Connect()
        rs.Connection = conn
        rs.CommandText = "SELECT * FROM `viewusers` WHERE `name` = '" & Me.cboMedTech.Text & "'"
        reader = rs.ExecuteReader
        reader.Read()
        If reader.HasRows Then
            ID_NO = reader(0).ToString
        End If
        Disconnect()
    End Sub

    Private Sub LoadCombo()
        Me.cboMedTech.Properties.Items.Clear()
        Connect()
        rs.Connection = conn
        rs.CommandText = "SELECT CONCAT(medtech.fname, ' ', medtech.mname, ' ', medtech.lname, ', ', medtech.designation) AS `Name` FROM medtech
                            UNION ALL 
                          SELECT CONCAT(pathologist.fname, ' ', pathologist.mname, ' ', pathologist.lname, ', ', pathologist.designation) AS `Name` FROM pathologist 
                            UNION ALL 
                          SELECT CONCAT(users.fname, ' ', users.mname, ' ', users.lname, ', ', users.designation) AS `Name` FROM users "

        reader = rs.ExecuteReader
        While reader.Read
            Me.cboMedTech.Properties.Items.Add(reader(0))
        End While
        Disconnect()
    End Sub

    Private Sub LoadLocation()
        Me.cboLocation.Properties.Items.Clear()
        Connect()
        rs.Connection = conn
        rs.CommandText = "SELECT `location` FROM `location` ORDER BY `id`"
        reader = rs.ExecuteReader
        While reader.Read
            Me.cboLocation.Properties.Items.Add(reader(0))
        End While
        Disconnect()
    End Sub

    Private Sub frmAccountSemiA_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        LoadPrivileges()
        LoadPermissions()
        LoadCombo()
        LoadLocation()
    End Sub

    Private Sub btnClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnClose.Click
        Me.Close()
        Me.Dispose()
    End Sub
End Class