Imports System.Text.RegularExpressions
Imports System
Imports System.Collections.Generic
Imports System.Linq
Imports System.Drawing
Imports DevExpress.XtraEditors
Imports DevExpress.XtraPrinting.BarCode

Public Class frmControlSettingAE

    Private Sub btnClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnClose.Click
        Me.Close()
    End Sub

    Private Sub frmControlSetting_FormClosed(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
        Me.Dispose()
    End Sub

    Private Sub SimpleButton5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSave.Click
        rs.Parameters.Clear()
        rs.Parameters.AddWithValue("@ctrl_id", txtSampleID.Text)
        rs.Parameters.AddWithValue("@test_type", cboTest.Text)
        rs.Parameters.AddWithValue("@type", cboType.Text)
        rs.Parameters.AddWithValue("@lot_no", txtLotNo.Text)
        rs.Parameters.AddWithValue("@control_name", txtContrlName.Text)
        rs.Parameters.AddWithValue("@instrument", cboInstrument.Text & "_" & txtSerialNo.Text)
        rs.Parameters.AddWithValue("@entry_date", Now)
        'rs.Parameters.AddWithValue("@expiry_date", Format(dtExpiry.Value, "MM-dd-yyyy"))

        SaveRecord("INSERT INTO `control_setting` (`ctrl_id`, `test_type`, `type`, `lot_no`, `control_name`, `instrument`, `entry_date`) VALUES " _
                   & "(" _
                   & "@ctrl_id," _
                   & "@test_type," _
                   & "@type," _
                   & "@lot_no," _
                   & "@control_name," _
                   & "@instrument," _
                   & "@entry_date" _
                   & ")"
                   )
        Me.Close()
        frmControlSetting.LoadRecords()
    End Sub

    Private Sub LoadCombo()
        AutoLoadComboDev(cboInstrument, "machines", 1, "name")
        AutoLoadComboDev(cboTest, "testtype", 1, "test_name")
    End Sub

    Private Sub frmControlSettingAE_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        LoadCombo()
    End Sub

    Private Sub cboInstrument_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cboInstrument.SelectedIndexChanged
        rs.Parameters.Clear()
        rs.Parameters.AddWithValue("@machine", cboInstrument.Text)

        Connect()
        rs.Connection = conn
        rs.CommandText = "SELECT `serial_no` FROM `machines` WHERE `name` LIKE @machine"
        reader = rs.ExecuteReader()
        reader.Read()
        If reader.HasRows Then
            txtSerialNo.Text = reader(0).ToString
        End If
        Disconnect()
    End Sub
End Class