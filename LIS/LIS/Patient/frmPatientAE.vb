Public Class frmPatientAE

    Public ID As String

    Private Sub btnBack_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnClose.Click
        Me.Close()
    End Sub

    Private Sub btnSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSave.Click
        If (txtPatientID.Text = "" Or txtFName.Text = "" Or txtAge.Text = "" Or cboSex.Text = "") Then
            MessageBox.Show("Please fill alle required information mark with (*).", "Empty String", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            Exit Sub
        End If

        If (Val(txtAge.Text) < 0) Then
            MessageBox.Show("Invalid age. PLease check.", "Invalid Entry", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            Exit Sub
        End If

        rs.Parameters.Clear()
        rs.Parameters.AddWithValue("@patient_id", txtPatientID.Text)
        rs.Parameters.AddWithValue("@name", txtFName.Text)
        rs.Parameters.AddWithValue("@sex", cboSex.Text)
        rs.Parameters.AddWithValue("@dbirth", dtBDate.Text)
        rs.Parameters.AddWithValue("@age", txtAge.Text)
        rs.Parameters.AddWithValue("@type", cboAgeType.Text)
        rs.Parameters.AddWithValue("@address", txtAddress.Text)
        rs.Parameters.AddWithValue("@contact", txtContact.Text)
        rs.Parameters.AddWithValue("@date", Now)
        rs.Parameters.AddWithValue("@ID", ID)

        If Me.btnSave.Text = "&Save" Then
            Connect()
            rs.Connection = conn
            rs.CommandText = "SELECT * FROM patient_info WHERE `patient_id` LIKE @patient_id"
            reader = rs.ExecuteReader
            reader.Read()
            If reader.HasRows Then
                Disconnect()
                MessageBox.Show("Patient ID already exist.", "Duplicate Entry", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
                Exit Sub
            Else
                Disconnect()
                SaveRecord("INSERT INTO patient_info (`patient_id`, `name`, `sex`, `date_of_birth`, `age`, `type`, `address`, `contact_no`, `date`) VALUES " _
                           & "(" _
                           & "@patient_id," _
                           & "@name," _
                           & "@sex," _
                           & "@dbirth," _
                           & "@age," _
                           & "@type," _
                           & "@address," _
                           & "@contact," _
                           & "@date" _
                           & ")"
                           )
            End If
        Else
            UpdateRecord("UPDATE patient_info SET " _
                       & "`patient_id` = @patient_id," _
                       & "`name` = @name," _
                       & "`sex` = @sex," _
                       & "`address` = @address," _
                       & "`date_of_birth` = @dbirth," _
                       & "`age` = @age," _
                       & "`type` = @type," _
                       & "`contact_no` = @contact," _
                       & "`date` = @date" _
                       & " WHERE `patient_id` LIKE @ID"
                       )
        End If
        frmPatient.LoadRecords()
        Me.Close()
    End Sub

    Private Sub cboAuto_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkAuto.CheckedChanged
        If chkAuto.CheckState = CheckState.Checked Then
            AutoPatientID()
            Me.txtPatientID.Enabled = False
        Else
            Me.txtPatientID.Enabled = True
            Me.txtPatientID.Text = ""
            Me.txtPatientID.Focus()
        End If
    End Sub

    Private Sub dtDateBirth_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles dtBDate.EditValueChanged
        GetBDate()
    End Sub

    Public Sub AutoPatientID()
        Try
            Connect()
            rs.Connection = conn
            rs.CommandText = "SELECT (id) as id FROM `patient_info` ORDER BY `id` DESC"
            reader = rs.ExecuteReader
            reader.Read()
            If reader.HasRows Then
                Me.txtpatientID.Text = "P" & reader(0).ToString + 1
            Else
                Me.txtpatientID.Text = "P1"
            End If
            Disconnect()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Auto Patient ID", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Disconnect()
            Exit Sub
        End Try
    End Sub

    Private Sub txtAge_KeyPress(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtAge.KeyPress, txtContact.KeyPress
        If InStr("1234567890", e.KeyChar) = 0 And Not Chr(AscW(e.KeyChar)) = vbBack Then
            e.KeyChar = ChrW(0)
        End If
    End Sub

    Private Sub frmPatientAE_FormClosed(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
        Me.Dispose()
    End Sub

    Public Sub GetBDate()
        Try
            If dtBDate.DateTime = Nothing Then
                Exit Sub
            End If

            Dim Birthday As Date = dtBDate.DateTime
            Dim endDate As Date = Date.Now

            Dim timeSpan As TimeSpan = endDate.Subtract(Birthday)
            Dim totalDays As Integer = timeSpan.TotalDays
            Dim totalYears As Integer = Math.Truncate(totalDays / 365)
            Dim totalMonths As Integer = Math.Truncate((totalDays Mod 365) / 30)
            Dim remainingDays As Integer = Math.Truncate((totalDays Mod 365) Mod 30)

            If Birthday.Year = endDate.Year Then
                If Birthday.Month = endDate.Month Then
                    If Birthday.Day = endDate.Day Then
                        txtAge.Text = totalDays.ToString
                        cboAgeType.Text = "Day(s)"
                    ElseIf Birthday.Day < endDate.Day Then
                        txtAge.Text = totalDays.ToString
                        cboAgeType.Text = "Day(s)"
                    Else
                        txtAge.Text = "0"
                        cboAgeType.Text = "Day(s)"
                    End If
                ElseIf Birthday.Month < endDate.Month Then
                    txtAge.Text = totalMonths.ToString
                    cboAgeType.Text = "Month(s)"
                Else
                    txtAge.Text = "0"
                    cboAgeType.Text = "Month(s)"
                End If
            ElseIf Birthday.Year < endDate.Year Then
                txtAge.Text = totalYears.ToString
                cboAgeType.Text = "Year(s)"
            Else
                txtAge.Text = "0"
                cboAgeType.Text = "Year(s)"
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Exception", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
        End Try
    End Sub
End Class