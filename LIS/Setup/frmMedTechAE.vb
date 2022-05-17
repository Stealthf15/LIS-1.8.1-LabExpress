Public Class frmMedTechAE

    Public ID As String
    Private medtechName As String
    Dim Verificator As Integer = 0

    Private Sub btnBack_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnClose.Click
        Me.Close()
    End Sub

    Private Sub btnSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSave.Click
        If Me.txtFName.Text = "" Or Me.txtLName.Text = "" Then
            MessageBox.Show("Please fill all required information.", "Empty String", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            Exit Sub
        End If

        Dim arrImage() As Byte

        Dim myMS As New IO.MemoryStream
        If Not IsNothing(picLogo.Image) Then
            Me.picLogo.Image.Save(myMS, Drawing.Imaging.ImageFormat.Jpeg)
            arrImage = myMS.GetBuffer
        Else
            arrImage = Nothing
        End If

        rs.Parameters.Clear()
        rs.Parameters.AddWithValue("@id", id)
        rs.Parameters.AddWithValue("@fname", txtFName.Text)
        rs.Parameters.AddWithValue("@mname", txtMName.Text)
        rs.Parameters.AddWithValue("@lname", txtLName.Text)
        rs.Parameters.AddWithValue("@license", txtLicense.Text)
        rs.Parameters.AddWithValue("@designation", txtDesign.Text)
        rs.Parameters.AddWithValue("@img", arrImage)

        If chkVerify.Checked = True Then
            rs.Parameters.AddWithValue("@verificator", Verificator)
        Else
            rs.Parameters.AddWithValue("@verificator", Verificator)
        End If

        If Me.btnSave.Text = "&Save" Then
            SaveRecord("INSERT INTO `medtech` (`fname`, `mname`, `lname`, `license`, `designation`, `signature`, `verificator`) VALUES (@fname, @mname, @lname, @license, @designation, @img, @verificator)")
            Me.Close()
            frmMedTech.LoadRecords()
        Else
            Try
                medtechName = txtFName.Text + " " + txtMName.Text + " " + txtLName.Text + ", " + txtDesign.Text
                UpdateRecord("UPDATE `medtech` SET `fname` = @fname, `mname` = @mname, `lname` = @lname, `license` = @license, `designation` = @designation, `signature` = @img, `verificator`= @verificator WHERE `id` = @id")
                UpdateRecordwthoutMSG("UPDATE `user_account` SET `name` = '" & medtechName & "' WHERE `email` = @id")
                Me.Close()
                frmMedTech.LoadRecords()
            Catch ex As Exception
                MessageBox.Show(ex.ToString, "Error", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            End Try
        End If
    End Sub

    Private Sub frmMedTechAE_FormClosed(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
        Me.Dispose()
    End Sub

    Private Sub btnSelect_Click(sender As Object, e As EventArgs) Handles btnSelect.Click
        OpenFileDialog.Filter = "Images Files(*.PNG; *.jpg; *.bmp)|*.png; *.jpg; *.bmp"
        If OpenFileDialog.ShowDialog = DialogResult.OK Then
            Dim FileName As String = System.IO.Path.GetFullPath(OpenFileDialog.FileName)
            picLogo.Image = System.Drawing.Bitmap.FromFile(FileName)
            Me.picLogo.SizeMode = PictureBoxSizeMode.StretchImage
        End If
    End Sub

    Private Sub btnClear_Click(sender As Object, e As EventArgs) Handles btnClear.Click
        picLogo.Image.Dispose()
    End Sub

    Private Sub chkVerify_CheckedChanged(sender As Object, e As EventArgs) Handles chkVerify.CheckedChanged
        If chkVerify.Checked = True Then
            Verificator = 1
        Else
            Verificator = 0
        End If
    End Sub
End Class