Public Class frmAuth

    Public Section As String
    Public Method As String

    Private Sub btnLogin_Click(sender As Object, e As EventArgs) Handles btnLogin.Click
        Try
            'Got this Encrypting code from StockOverFlow.com
            Dim bytes() As Byte = System.Text.Encoding.UTF8.GetBytes(txtPassword.Text)
            Dim hashOfBytes() As Byte = New System.Security.Cryptography.SHA1Managed().ComputeHash(bytes)
            Dim strHash As String = Convert.ToBase64String(hashOfBytes)

            rs.Parameters.Clear()
            rs.Parameters.AddWithValue("@username", txtUsername.Text)
            rs.Parameters.AddWithValue("@password", strHash)

            Connect()
            rs.Connection = conn
            rs.CommandText = "SELECT * FROM `user_account` WHERE `Username` = @username AND `password` = @password"
            reader = rs.ExecuteReader
            reader.Read()
            If (reader.HasRows = True) Then
                If (reader("username") = Me.txtUsername.Text) Then
                    CurrID = reader(0).ToString
                    CurrUser = reader(1).ToString
                    CurrPost = reader(7).ToString
                    CurrDept = reader(6).ToString
                    CurrType = reader(4).ToString
                    CurrEmail = reader(5).ToString
                    UserIDRelease = CurrID
                    UserRelease = CurrUser
                    Disconnect()

                    If Section = "Chemistry" Then
                        If Method = "Release" Then
                            frmChemNew.ReleaseResultNoPrint()
                        ElseIf Method = "PrintRelease" Then
                            frmChemNew.PrintReleaseResult()
                        ElseIf Method = "Verify" Then
                            frmChemNew.VerifyResult()
                        End If
                    ElseIf Section = "Hematology" Then
                        If Method = "Release" Then
                            frmHemaNew.ReleaseResultNoPrint()
                        ElseIf Method = "PrintRelease" Then
                            frmHemaNew.PrintReleaseResult()
                        ElseIf Method = "Verify" Then
                            frmHemaNew.VerifyResult()
                        End If
                    ElseIf Section = "Urinalysis" Then
                        If Method = "Release" Then
                            frmUrinNew.ReleaseResultNoPrint()
                        ElseIf Method = "PrintRelease" Then
                            frmUrinNew.PrintReleaseResult()
                        ElseIf Method = "Verify" Then
                            frmUrinNew.VerifyResult()
                        End If
                    ElseIf Section = "Fecalysis" Then
                        If Method = "Release" Then
                            frmFecalNew.ReleaseResultNoPrint()
                        ElseIf Method = "PrintRelease" Then
                            frmFecalNew.PrintReleaseResult()
                        ElseIf Method = "Verify" Then
                            frmFecalNew.VerifyResult()
                        End If
                    ElseIf Section = "Serology" Then
                        If Method = "Release" Then
                            frmFecalNew.ReleaseResultNoPrint()
                        ElseIf Method = "PrintRelease" Then
                            frmFecalNew.PrintReleaseResult()
                        ElseIf Method = "Verify" Then
                            frmFecalNew.VerifyResult()
                        End If
                    End If

                    Me.Close()
                    Me.Dispose()
                Else
                    Disconnect()
                    MessageBox.Show("Invalid Username or Password.", "Authentication Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    Exit Sub
                End If
            Else
                Disconnect()
                MessageBox.Show("Invalid Username or Password.", "Authentication Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Exit Sub
            End If
            Disconnect()
        Catch ex As Exception
            Exit Sub
        End Try
    End Sub

    Private Sub btnCancel_Click(sender As Object, e As EventArgs) Handles btnCancel.Click
        Me.Close()
        Me.Dispose()
    End Sub
End Class