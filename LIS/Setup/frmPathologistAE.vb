Public Class frmPathologistAE

    Public ID As String

    Private Sub btnBack_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnClose.Click
        Me.Close()
        Me.Dispose()
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
        rs.Parameters.AddWithValue("@id", ID)
        rs.Parameters.AddWithValue("@fname", txtFName.Text)
        rs.Parameters.AddWithValue("@mname", txtMName.Text)
        rs.Parameters.AddWithValue("@lname", txtLName.Text)
        rs.Parameters.AddWithValue("@license", txtLicense.Text)
        rs.Parameters.AddWithValue("@designation", txtDesign.Text)
        rs.Parameters.AddWithValue("@img", arrImage)

        If Me.btnSave.Text = "&Save" Then
            SaveRecord("INSERT INTO `pathologist` (`fname`, `mname`, `lname`, `license`, `designation`,`img`) VALUES (@fname, @mname, @lname, @license, @designation, @img)")
            Me.Close()
            frmPathologist.LoadRecords()
        Else
            UpdateRecord("UPDATE `pathologist` SET `fname` = @fname, `mname` = @mname, `lname` = @lname, `license` = @license, `designation` = @designation, img = @img WHERE `id` = @id")
            Me.Close()
            frmPathologist.LoadRecords()
        End If
    End Sub

    Private Sub btnSelect_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSelect.Click
        OpenFileDialog.Filter = "Images Files(*.PNG; *.jpg; *.bmp)|*.png; *.jpg; *.bmp"
        If OpenFileDialog.ShowDialog = DialogResult.OK Then
            Dim FileName As String = System.IO.Path.GetFullPath(OpenFileDialog.FileName)
            picLogo.Image = System.Drawing.Bitmap.FromFile(FileName)
            Me.picLogo.SizeMode = PictureBoxSizeMode.StretchImage
        End If
    End Sub

    Private Sub btnClear_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnClear.Click
        picLogo.Image.Dispose()
    End Sub
End Class