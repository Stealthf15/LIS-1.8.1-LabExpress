Public Class frmSetup

    Public a As Integer = 1

    Private Sub btnSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSave.Click
        Try
            If a = 0 Then
                Panel1.Visible = True
                Panel2.Visible = False
                Panel3.Visible = False
                Panel4.Visible = False
                a = a + 1
            ElseIf a = 1 Then
                Panel1.Visible = False
                Panel2.Visible = True
                Panel3.Visible = False
                Panel4.Visible = False
                a = a + 1
            ElseIf a = 2 Then
                Panel1.Visible = False
                Panel2.Visible = False
                Panel3.Visible = True
                Panel4.Visible = False
                btnSave.Text = "Activate"
                If Me.btnSave.Text = "Activate" Then
                    If Not Me.txtLicense.Text = "" Then
                        Dim key As Long
                        If Not Long.TryParse(txtLicense.Text, key) Then
                            MessageBox.Show("Invalid activation key")
                            a = 2
                        Else
                            Dim sm As New SecurityManager
                            If sm.CheckKey(key) Then
                                My.Settings.ActivationStatus = True
                                My.Settings.Save()
                                Panel1.Visible = False
                                Panel2.Visible = False
                                Panel3.Visible = False
                                Panel4.Visible = True
                                btnSave.Visible = False
                                btnClose.Text = "&Finish"
                            Else
                                MessageBox.Show("Activation was unsuccessful")
                                a = 2
                            End If
                        End If
                    End If
                End If
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message, "System Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Exit Sub
        End Try
    End Sub

    Private Sub btnClose_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnClose.Click
        If Me.btnClose.Text = "&Finish" Then
            'My.Settings.isFirstRun = False
            My.Settings.Save()
            Me.Hide()
        Else
            If (MessageBox.Show("Closing this window will terminate the setup of the application." & vbNewLine & vbNewLine & "Do you want to close it anyway?", "Exit Setup", MessageBoxButtons.YesNo, MessageBoxIcon.Question)) = vbYes Then
                End
            Else
                Exit Sub
            End If
        End If
    End Sub

    Private Sub frmSetup_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Dim serial As Long
        Dim sm As New SecurityManager
        serial = sm.GetSerial
        txtSerial.Text = serial
    End Sub
End Class