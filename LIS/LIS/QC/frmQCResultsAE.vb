Public Class frmQCResultsAE
    Public _section As String = ""
    Public _controlName As String = ""
    Public _runCount As String = ""
    Public Sub LoadRecords()
        Try
            lvlist.Items.Clear()
            Connect()
            rs.Connection = conn
            rs.CommandText = "SELECT `id`, `specimen`, `test_code` FROM `default_specimen` WHERE `section` LIKE '" & _section & "'"
            reader = rs.ExecuteReader
            While reader.Read()
                If reader.HasRows Then
                    iItem = New ListViewItem(reader(0).ToString, 0)
                    For x As Integer = 1 To 2 Step 1
                        iItem.SubItems.Add(reader(x).ToString())
                    Next
                    lvlist.Items.Add(iItem)
                End If
            End While
            Disconnect()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "System Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub
    Private Sub fmrQCResultsAE_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        LoadRecords()
    End Sub

    Private Sub btnAddTest_ItemClick(sender As Object, e As DevExpress.XtraBars.ItemClickEventArgs) Handles btnAddTest.ItemClick
        Dim si_units As String
        For x = Me.lvlist.CheckedItems.Count - 1 To 0 Step -1


            rs.Parameters.Clear()
            rs.Parameters.AddWithValue("@UniversalID", lvlist.CheckedItems(x).SubItems(1).Text)
            rs.Parameters.AddWithValue("@Measurement", "")
            rs.Parameters.AddWithValue("@test_code", lvlist.CheckedItems(x).SubItems(2).Text)
            rs.Parameters.AddWithValue("@sample_id", _controlName)
            rs.Parameters.Add("@date", MySql.Data.MySqlClient.MySqlDbType.Date).Value = Format(frmQCResults.dtFrom.Value, "yyyy-MM-dd")
            rs.Parameters.AddWithValue("@test_type", _section)
            rs.Parameters.AddWithValue("@month", frmQCResults.dtFrom.Value.Month)
            rs.Parameters.AddWithValue("@year", frmQCResults.dtFrom.Value.Year)
            rs.Parameters.AddWithValue("@instrument", frmQC.cboMachines.Text)
            rs.Parameters.AddWithValue("@run_count", _runCount)


            Connect()
            rs.Connection = conn
            rs.CommandText = "SELECT `si_unit`FROM `default_specimen` WHERE `specimen` = @UniversalID AND `test_code` = @test_code"
            reader = rs.ExecuteReader
            If reader.Read Then
                si_units = reader(0).ToString
            End If
            Disconnect()



            SaveRecordwthoutMSG("INSERT INTO `control_result` (`universal_id`, `measurement`,`test_code`, `sample_id`, `date`, `test_type`, `month`, `year`, `instrument`, `unit`, `run_count`) VALUES " _
                        & "(" _
                        & "@UniversalID, @Measurement, @test_code, @sample_id, @date, @test_type, @month, @year, @instrument, '" & si_units & "', @run_count" _
                        & ")"
                        )
            frmQCResults.LoadTest()
            Me.Close()
        Next

    End Sub

    Private Sub btnClose_ItemClick(sender As Object, e As DevExpress.XtraBars.ItemClickEventArgs) Handles btnClose.ItemClick
        Me.Dispose()
    End Sub
End Class