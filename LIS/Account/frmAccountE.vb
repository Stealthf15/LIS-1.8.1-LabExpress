Public Class frmAccountE

    Public ID, MedTechID As String

    Private Sub btnBack_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Me.Close()
        Me.Dispose()
    End Sub

    Private Sub btnSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSave.Click
        'Got this Encrypting code from StockOverFlow.com
        Dim bytes() As Byte = System.Text.Encoding.UTF8.GetBytes(txtPassword.Text)
        Dim hashOfBytes() As Byte = New System.Security.Cryptography.SHA1Managed().ComputeHash(bytes)
        Dim strHash As String = Convert.ToBase64String(hashOfBytes)

        If Me.txtUsername.Text = "" Or Me.txtPassword.Text = "" Or Me.cboMedTech.Text = "" Or Me.cboUserType.Text = "" Then
            MessageBox.Show("Please fill all required information.", "Empty String", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            Exit Sub
        End If
        rs.Parameters.Clear()
        rs.Parameters.AddWithValue("@name", cboMedTech.Text)
        rs.Parameters.AddWithValue("@username", txtUsername.Text)
        rs.Parameters.AddWithValue("@password", strHash)
        rs.Parameters.AddWithValue("@usertype", cboUserType.Text)
        'rs.Parameters.AddWithValue("@MedTechID", MedTechID)

        conn.Close()
        conn.ConnectionString = MyConnectionString
        conn.Open()
        rs.Connection = conn
        rs.CommandText = "UPDATE `user_account` SET `name` = @name, `username` = @username, `password` = @password, `usertype` = @usertype WHERE `user_id` = " & CurrID & ""
        rs.ExecuteNonQuery()
        Disconnect()
        MessageBox.Show("Your account has been successfully update.", "Update", MessageBoxButtons.OK, MessageBoxIcon.Information)
        Disconnect()

        frmAccount.LoadRecords()
        cboMedTech.Text = ""
        txtUsername.Text = ""
        txtPassword.Text = ""
        cboUserType.Text = ""
        Me.Close()
    End Sub

    'Private Sub cboMedTech_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cboMedTech.SelectedIndexChanged
    '    Connect()
    '    rs.Connection = conn
    '    rs.CommandText = "SELECT * FROM `viewMedTEch` WHERE `name` LIKE '" & Me.cboMedTech.Text & "'"
    '    reader = rs.ExecuteReader
    '    reader.Read()
    '    If reader.HasRows Then
    '        MedTechID = reader(0).ToString
    '    End If
    '    Disconnect()
    'End Sub

    Private Sub LoadCombo()
        Me.cboMedTech.Properties.Items.Clear()
        Connect()
        rs.Connection = conn
        rs.CommandText = "SELECT `fname` + ' ' + `mname` + ' ' + `lname` + ', ' + `designation` AS name FROM medtech ORDER BY name"
        reader = rs.ExecuteReader
        While reader.Read
            Me.cboMedTech.Properties.Items.Add(reader(0))
        End While
        Disconnect()
    End Sub

    Private Sub frmAccountSemiA_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        LoadCombo()
        LoadAccount()
    End Sub

    Private Sub btnClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnClose.Click
        Me.Close()
        Me.Dispose()
    End Sub

    Private Sub LoadAccount()
        Disconnect()
        Connect()
        rs.Connection = conn
        rs.CommandText = "SELECT `name`, `username`, `usertype` FROM `user_account` WHERE `user_id` = '" & CurrID & "'"
        reader = rs.ExecuteReader
        reader.Read()
        If reader.HasRows Then
            cboMedTech.Text = reader(0).ToString
            txtUsername.Text = reader(1).ToString
            cboUserType.Text = reader(2).ToString
        End If
        Disconnect()
    End Sub
End Class