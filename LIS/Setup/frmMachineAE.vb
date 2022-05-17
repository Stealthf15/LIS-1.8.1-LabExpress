Public Class frmMachineAE

    Public ID As String

    Private Sub btnBack_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Me.Close()
        Me.Dispose()
    End Sub

    Private Sub btnSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        If Me.txtName.Text = "" Or Me.cboStatus.Text = "" Or Me.cboBaudRate.Text = "" Or Me.cboTestName.Text = "" Or Me.cboComPort.Text = "" Then
            MessageBox.Show("Please fill all required information.", "Empty String", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            Exit Sub
        End If

        rs.Parameters.Clear()
        rs.Parameters.AddWithValue("@id", ID)
        rs.Parameters.AddWithValue("@name", txtName.Text)
        rs.Parameters.AddWithValue("@comport", cboComPort.Text)
        rs.Parameters.AddWithValue("@baudrate", cboBaudRate.Text)
        rs.Parameters.AddWithValue("@testname", cboTestName.Text)
        rs.Parameters.AddWithValue("@status", cboStatus.Text)

        If Me.btnSave.Text = "&Save" Then
            SaveRecord("INSERT INTO `machines` (`name`, `com_port`, `baud_rate`, `test_name`, `status`) VALUES (@name, @comport, @baudrate, @testname, @status)")
            Me.Close()
            frmMachineList.LoadRecords()
        Else
            UpdateRecord("UPDATE `machines` SET `name` = @name, `com_port` = @comport, `baud_rate` = @baudrate, `test_name` = @testname, `status` = @status WHERE `id` = @id")
            Me.Close()
            frmMachineList.LoadRecords()
        End If
    End Sub

    Private Sub frmMachineAE_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Me.cboComPort.Properties.Items.Clear()
        Me.cboBaudRate.Properties.Items.Clear()
        Me.cboStatus.Properties.Items.Clear()

        For Each COMString As String In My.Computer.Ports.SerialPortNames ' Load all available COM ports.
            cboComPort.Properties.Items.Add(COMString)
        Next
        cboComPort.Properties.Sorted = True
        cboBaudRate.Properties.Items.Add("110")
        cboBaudRate.Properties.Items.Add("300")
        cboBaudRate.Properties.Items.Add("600")
        cboBaudRate.Properties.Items.Add("1200")
        cboBaudRate.Properties.Items.Add("1800")
        cboBaudRate.Properties.Items.Add("2400")
        cboBaudRate.Properties.Items.Add("4800")
        cboBaudRate.Properties.Items.Add("7200")
        cboBaudRate.Properties.Items.Add("9600")
        cboBaudRate.Properties.Items.Add("14400")
        cboBaudRate.Properties.Items.Add("19200")      ' Min. FIFO size 3 Bytes (8030 or 8530)
        cboBaudRate.Properties.Items.Add("38400")
        cboBaudRate.Properties.Items.Add("57600")      ' Min. FIFO size 8 bytes
        cboBaudRate.Properties.Items.Add("115200")     ' Min. FIFO size 16 bytes (16C550)
        cboBaudRate.Properties.Items.Add("230400")     ' Min. FIFO size 32 bytes (16C650)
        cboBaudRate.Properties.Items.Add("460800")     ' Min. FIFO size 64 bytes (16C750)
        cboBaudRate.Properties.Items.Add("921600")     ' Min. FIFO size 128 bytes (16C850 or 16C950)

        Me.cboStatus.Properties.Items.Add("Active")
        Me.cboStatus.Properties.Items.Add("Inactive")

        cboTestName.Properties.Items.Clear()
        Connect()
        rs.Connection = conn
        rs.CommandText = "SELECT * FROM `TESTtype` ORDER BY `TEST_name`"
        reader = rs.ExecuteReader
        While reader.Read
            cboTestName.Properties.Items.Add(reader(1).ToString)
        End While
        Disconnect()
    End Sub
End Class