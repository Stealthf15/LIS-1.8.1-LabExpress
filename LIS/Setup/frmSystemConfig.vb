Imports System.Drawing.Printing

Public Class frmSystemConfig

    Public name As String

    Private Sub btnBack_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnClose.Click
        Me.Close()
    End Sub

    Private Sub btnSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSave.Click
        Try
            With My.Settings
                .MedTech = cboType.SelectedIndex
                .PrintBarcode = chkBarcode.CheckState
                .BCPrinterName = cboPrinteName.Text
                .LISType = cboLISType.Text
                .AuthenticateRelease = cboAuthenticate.Text

                .HL7Destination = txtHL7Destination.Text
                .HL7Location = txtHL7Location.Text
                .HL7Read = chkHL7Read.CheckState
                .HL7Write = chkHL7Write.CheckState

                .PDFLocation = txtPDF.Text
                .SaveAsPDF = chkPDF.CheckState

                'XML File
                .XMLDestination = txtXMLDestination.Text
                .XMLWrite = chkXML.CheckState

                'Page Setup
                .BCWidth = txtBCWidth.Text
                .BCHeight = txtBCHeight.Text
                .PaperWidth = txtPaperWidth.Text
                .PaperHeight = txtPaperHeight.Text

                'Save SQL Settings
                .SQLServer = txtSQLServerName.Text
                .SQLUID = txtSQLUID.Text
                .SQLPWD = txtSQLPWD.Text
                .SQLDBName = txtSQLDBName.Text
                .SQLConnection = chkEnableSQL.CheckState
                .MyConnectionStringSQL = "SERVER = " & txtSQLServerName.Text & ";" & "DATABASE = " & txtSQLDBName.Text & "; UID = " & txtSQLUID.Text & "; PWD = " & txtSQLPWD.Text & ";"

                'Auto Send Email
                .AutoEmail = chkEmail.CheckState

                .Save()
                .Reload()
            End With

            rs.Parameters.Clear()
            rs.Parameters.AddWithValue("@Server", txtServer.Text)
            rs.Parameters.AddWithValue("@Port", txtPort.Text)
            rs.Parameters.AddWithValue("@Email", txtEmail.Text)
            rs.Parameters.AddWithValue("@Username", txtUsername.Text)
            rs.Parameters.AddWithValue("@Password", txtPassword.Text)
            rs.Parameters.AddWithValue("@CC", txtCC.Text)
            rs.Parameters.AddWithValue("@BC", txtBC.Text)
            rs.Parameters.AddWithValue("@Description", txtDescription.Text)
            rs.Parameters.AddWithValue("@Status", txtStatus.Text)
            rs.Parameters.AddWithValue("@Access", txtAccess.Text)
            rs.Parameters.AddWithValue("@DateTime", Now)
            Connect()
            rs.Connection = conn
            rs.CommandType = CommandType.Text
            rs.CommandText = "SELECT * FROM `email_maintenance`"
            reader = rs.ExecuteReader
            reader.Read()
            If reader.HasRows Then
                Disconnect()
                UpdateRecordwthoutMSG("UPDATE `email_maintenance` SET " &
                        "`smtp_server` = @Server," &
                        "`smtp_port` = @Port," &
                        "`email` = @Email," &
                        "`username` = @Username," &
                        "`password` = @Password," &
                        "`cc` = @CC," &
                        "`bc` = @BC," &
                        "`description` = @Description," &
                        "`access` = @Access," &
                        "`status` = @Status," &
                        "`date_added` = @DateTime"
                        )
            Else
                Disconnect()
                SaveRecordwthoutMSG("INSERT INTO `email_maintenance` (`smtp_server`, `smtp_port`, `email`, `username`, `password`, `cc`, `bc`, `description`, `access`, `status`, `date_added`) VALUES " &
                        "(" &
                        "@Server," &
                        "@Port," &
                        "@Email," &
                        "@Username," &
                        "@Password," &
                        "@CC," &
                        "@BC," &
                        "@Description," &
                        "@Access," &
                        "@Status," &
                        "@DateTime" &
                        ")"
                        )
            End If
            Disconnect()

            MessageBox.Show("All settings has been updated." & vbNewLine &
                            "Please restart the application to apply the changes.", "System Message", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Me.Close()
        Catch ex As Exception
            Disconnect()
            MessageBox.Show(ex.Message, "Exception", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
        End Try
    End Sub

    Private Sub frmTestTypeAE_FormClosed(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
        Me.Dispose()
    End Sub

    Private Sub frmSystemConfig_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Try
            For Each prnt As String In PrinterSettings.InstalledPrinters
                cboPrinteName.Properties.Items.Add(prnt)
            Next

            With My.Settings
                cboType.SelectedIndex = .MedTech

                txtHL7Destination.Text = .HL7Destination
                txtHL7Location.Text = .HL7Location
                chkHL7Read.Checked = .HL7Read
                chkHL7Write.Checked = .HL7Write

                'XML File
                txtXMLDestination.Text = .XMLDestination
                chkXML.Checked = .XMLWrite

                chkBarcode.Checked = .PrintBarcode
                cboPrinteName.Text = .BCPrinterName

                txtPDF.Text = .PDFLocation
                chkPDF.Checked = .SaveAsPDF
                cboLISType.Text = .LISType
                cboAuthenticate.Text = .AuthenticateRelease

                'Page Setup
                txtBCWidth.Text = .BCWidth
                txtBCHeight.Text = .BCHeight
                txtPaperWidth.Text = .PaperWidth
                txtPaperHeight.Text = .PaperHeight

                'Save SQL Settings
                txtSQLServerName.Text = .SQLServer
                txtSQLUID.Text = .SQLUID
                txtSQLPWD.Text = .SQLPWD
                txtSQLDBName.Text = .SQLDBName
                chkEnableSQL.Checked = .SQLConnection

                'Auto Send Email
                chkEmail.Checked = .AutoEmail
            End With

            Connect()
            rs.Connection = conn
            rs.CommandType = CommandType.Text
            rs.CommandText = "SELECT `id`, `access`, `email`, `smtp_server`, `smtp_port`, `username`, `password`, `cc`, `bc`, `description`, `status`, `access` FROM `email_maintenance`"
            reader = rs.ExecuteReader
            reader.Read()
            If reader.HasRows Then
                txtServer.Text = reader(3).ToString
                txtPort.Text = reader(4).ToString
                txtEmail.Text = reader(2).ToString
                txtUsername.Text = reader(5).ToString
                txtPassword.Text = reader(6).ToString
                txtCC.Text = reader(7).ToString
                txtBC.Text = reader(8).ToString
                txtDescription.Text = reader(9).ToString
                txtStatus.Text = reader(10).ToString
                txtAccess.Text = reader(11).ToString
            End If
            Disconnect()
        Catch ex As Exception
            Exit Sub
        End Try
    End Sub
End Class