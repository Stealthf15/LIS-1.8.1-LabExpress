Public Class frmEmailContent
    Private Sub btnClose_Click(sender As Object, e As EventArgs) Handles btnClose.Click
        Me.Close()
    End Sub

    Private Sub btnSave_Click(sender As Object, e As EventArgs) Handles btnSave.Click
        Connect()
        rs.Parameters.Clear()
        rs.Parameters.AddWithValue("@Sender", txtSender.Text)
        rs.Parameters.AddWithValue("@Header", txtHeader.Text)
        rs.Parameters.AddWithValue("@Body", txtBody.Text)

        rs.Connection = conn
        rs.CommandType = CommandType.Text
        rs.CommandText = "SELECT * FROM `email_content`"
        reader = rs.ExecuteReader
        reader.Read()
        If reader.HasRows Then
            Disconnect()
            UpdateRecord("UPDATE `email_content` SET `sender` = @Sender, `header` = @Header, `body` = @Body")
        Else
            Disconnect()
            SaveRecord("INSERT INTO `email_content` (`sender`, `header`, `body`) VALUES (@Sender, @Header, @Body)")
        End If
        Disconnect()

    End Sub

    Private Sub frmEmailContent_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Connect()
        rs.Connection = conn
        rs.CommandType = CommandType.Text
        rs.CommandText = "SELECT * FROM `email_content`"
        reader = rs.ExecuteReader
        reader.Read()
        If reader.HasRows Then
            txtSender.Text = reader(1).ToString
            txtHeader.Text = reader(2).ToString
            txtBody.Text = reader(3).ToString
        End If
        Disconnect()
    End Sub
End Class