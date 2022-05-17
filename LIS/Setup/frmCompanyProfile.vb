Imports System.IO

Public Class frmCompanyProfile

    Public imgName As String

    Private Sub btnSelect_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSelect.Click
        OpenFileDialog.Filter = "Images Files(*.PNG; *.jpg; *.bmp)|*.png; *.jpg; *.bmp"
        If OpenFileDialog.ShowDialog = DialogResult.OK Then
            Dim FileName As String = System.IO.Path.GetFullPath(OpenFileDialog.FileName)
            picLogo_primary.Image = System.Drawing.Bitmap.FromFile(FileName)
            Me.picLogo_primary.SizeMode = PictureBoxSizeMode.StretchImage
        End If
    End Sub

    Private Sub btnSelect1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSelect1.Click
        OpenFileDialog.Filter = "Images Files(*.PNG; *.jpg; *.bmp)|*.png; *.jpg; *.bmp"
        If OpenFileDialog.ShowDialog = DialogResult.OK Then
            Dim FileName As String = System.IO.Path.GetFullPath(OpenFileDialog.FileName)
            picLogo_secondary.Image = System.Drawing.Bitmap.FromFile(FileName)
            Me.picLogo_secondary.SizeMode = PictureBoxSizeMode.StretchImage
        End If
    End Sub

    Private Sub btnAction_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSave.Click
        If Me.txtAddress.Text = "" Or Me.txtName.Text = "" Then
            MessageBox.Show("Please fill all required information mark with (*).", "Empty String", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            Exit Sub
        End If

        Dim arrImage() As Byte
        Dim arrImage1() As Byte

        Dim myMS As New IO.MemoryStream
        Dim myMS1 As New IO.MemoryStream

        If Not IsNothing(picLogo_primary.Image) Then
            Me.picLogo_primary.Image.Save(myMS, Drawing.Imaging.ImageFormat.Png)
            arrImage = myMS.GetBuffer
        Else
            arrImage = Nothing
        End If

        If Not IsNothing(picLogo_secondary.Image) Then
            Me.picLogo_secondary.Image.Save(myMS1, Drawing.Imaging.ImageFormat.Png)
            arrImage1 = myMS1.GetBuffer
        Else
            arrImage1 = Nothing
        End If

        rs.Parameters.Clear()
        rs.Parameters.AddWithValue("@name", txtName.Text)
        rs.Parameters.AddWithValue("@address", txtAddress.Text)
        rs.Parameters.AddWithValue("@tel", txtTel.Text)
        rs.Parameters.AddWithValue("@fax", txtFax.Text)
        rs.Parameters.AddWithValue("@comp_logo", arrImage)
        rs.Parameters.AddWithValue("@comp_logo1", arrImage1)

        Connect()
        rs.Connection = conn
        rs.CommandType = CommandType.Text
        rs.CommandText = "SELECT * FROM company_profile"
        reader = rs.ExecuteReader
        reader.Read()
        If reader.HasRows Then
            Disconnect()
            UpdateRecordwthoutMSG("UPDATE company_profile SET " _
               & "`comp_name` = @name," _
               & "`comp_address` = @address," _
               & "`comp_tel_no` = @tel," _
               & "`comp_fax` = @fax," _
               & "`comp_logo` = @comp_logo," _
               & "`comp_logo1` = @comp_logo1"
               )
            Me.Close()
        Else
            Disconnect()
            SaveRecordwthoutMSG("INSERT INTO company_profile (`comp_name`, `comp_address`, `comp_tel_no`, `comp_fax`, `comp_logo`, `comp_logo1`) VALUES" _
               & "(" _
               & "@name," _
               & "@address," _
               & "@tel," _
               & "@fax," _
               & "@comp_logo," _
               & "@comp_logo1" _
               & ")"
               )
            Me.Close()
        End If
    End Sub

    Private Sub btnClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnClose.Click
        Me.Close()
    End Sub

    Private Sub frmCompanyProfile_FormClosed(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
        Me.Dispose()
    End Sub

    Private Sub frmCompanyProfile_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        RetriveData()
    End Sub

    Private Sub RetriveData()
        Connect()
        rs.Connection = conn
        rs.CommandType = CommandType.Text
        rs.CommandText = "SELECT * FROM company_profile"
        reader = rs.ExecuteReader
        reader.Read()
        If reader.HasRows Then
            txtName.Text = reader(1).ToString
            txtAddress.Text = reader(2).ToString
            txtTel.Text = reader(3).ToString
            txtFax.Text = reader(4).ToString

            If Not reader(5).ToString = "" Then
                Dim imageData As Byte() = DirectCast(reader(5), Byte())
                If Not imageData Is Nothing Then
                    Using ms As New MemoryStream(imageData, 0, imageData.Length)
                        ms.Write(imageData, 0, imageData.Length)
                        picLogo_primary.Image = Image.FromStream(ms, True)
                    End Using
                End If
            End If

            If Not reader(6).ToString = "" Then
                Dim imageData_sec As Byte() = DirectCast(reader(6), Byte())
                If Not imageData_sec Is Nothing Then
                    Using ms As New MemoryStream(imageData_sec, 0, imageData_sec.Length)
                        ms.Write(imageData_sec, 0, imageData_sec.Length)
                        picLogo_secondary.Image = Image.FromStream(ms, True)
                    End Using
                End If
            End If
        End If
        Disconnect()
    End Sub
End Class