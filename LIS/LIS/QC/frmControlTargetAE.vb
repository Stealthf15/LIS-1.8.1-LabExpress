Imports System.Text.RegularExpressions
Imports System
Imports System.Collections.Generic
Imports System.Linq
Imports System.Drawing
Imports DevExpress.XtraEditors
Imports DevExpress.XtraPrinting.BarCode

Public Class frmControlTargetAE

    Public ID As String

    Private Sub btnClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnClose.Click
        Me.Close()
    End Sub

    Private Sub frmControlSetting_FormClosed(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
        Me.Dispose()
    End Sub

    Private Sub SimpleButton5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSave.Click
        rs.Parameters.Clear()
        rs.Parameters.AddWithValue("@id", ID)
        rs.Parameters.AddWithValue("@ctrl_id", txtControlID.Text)
        rs.Parameters.AddWithValue("@specimen", cboSpecimen.Text)
        rs.Parameters.AddWithValue("@testcode", txtTestCode.Text)
        rs.Parameters.AddWithValue("@target", txtTarget.Text)
        rs.Parameters.AddWithValue("@tolerance", txtTolerance.Text)
        rs.Parameters.AddWithValue("@ll", txtLSD.Text)
        rs.Parameters.AddWithValue("@ul", txtUSD.Text)
        rs.Parameters.AddWithValue("@plus_one", txtP1.Text)
        rs.Parameters.AddWithValue("@minus_one", txtN1.Text)
        rs.Parameters.AddWithValue("@plus_three", txtP3.Text)
        rs.Parameters.AddWithValue("@minus_three", txtN3.Text)
        rs.Parameters.AddWithValue("@instrument", cboInstrument.Text)
        rs.Parameters.AddWithValue("@entry_date", Now)

        If btnSave.Text = "&Save" Then
            SaveRecord("INSERT INTO `control_target` (`control_id`, `test_name`, `test_code`, `target`, `sd`, `ll`, `ul`, `plus_one`, `minus_one`, `plus_three`, `minus_three`, `instrument`, `entry_date`) VALUES " _
                       & "(" _
                       & "@ctrl_id," _
                       & "@specimen," _
                       & "@testcode," _
                       & "@target," _
                       & "@tolerance," _
                       & "@ll," _
                       & "@ul," _
                       & "@plus_one," _
                       & "@minus_one," _
                       & "@plus_three," _
                       & "@minus_three," _
                       & "@instrument," _
                       & "@entry_date" _
                       & ")"
                       )
        Else
            UpdateRecord("UPDATE `control_target` SET " _
                   & "`control_id` = @ctrl_id," _
                   & "`test_name` = @specimen," _
                   & "`test_code` = @testcode," _
                   & "`target` = @target," _
                   & "`sd` = @tolerance," _
                   & "`ll` = @ll," _
                   & "`ul` = @ul," _
                   & "`plus_one` = @plus_one," _
                   & "`minus_one` = @minus_one," _
                   & "`plus_three` = @plus_three," _
                   & "`minus_three` = @minus_three," _
                   & "`instrument` = @instrument," _
                   & "`entry_date` = @entry_date" _
                   & " WHERE `id` = @id"
                   )
        End If
        Me.Close()
    End Sub

    Private Sub cboSpecimen_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cboSpecimen.SelectedIndexChanged
        rs.Parameters.Clear()
        rs.Parameters.AddWithValue("@Specimen", cboSpecimen.Text)
        rs.Parameters.AddWithValue("@Instrument", cboInstrument.Text)

        Connect()
        rs.Connection = conn
        rs.CommandText = "SELECT `test_code` FROM `default_specimen` WHERE `specimen` = @Specimen"
        reader = rs.ExecuteReader()
        reader.Read()
        If reader.HasRows Then
            txtTestCode.Text = reader(0).ToString
        End If
        Disconnect()
    End Sub

    Private Sub txtTestType_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtTestType.TextChanged, cboInstrument.TextChanged
        rs.Parameters.Clear()
        rs.Parameters.AddWithValue("@TestType", txtTestType.Text)
        rs.Parameters.AddWithValue("@Instrument", cboInstrument.Text)

        cboSpecimen.Properties.Items.Clear()

        Connect()
        rs.Connection = conn
        rs.CommandText = "SELECT `specimen` FROM `default_specimen` WHERE `section` = @TestType"
        reader = rs.ExecuteReader()
        While reader.Read()
            cboSpecimen.Properties.Items.Add(reader(0).ToString)
        End While
        Disconnect()
    End Sub

    Private Sub txtTolerance_EditValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtTolerance.EditValueChanged, txtTarget.EditValueChanged
        Me.txtLSD.Text = Val(txtTarget.Text) - (Val(txtTolerance.Text) * 2)
        Me.txtUSD.Text = Val(txtTarget.Text) + (Val(txtTolerance.Text) * 2)

        Me.txtN1.Text = Val(txtTarget.Text) - (Val(txtTolerance.Text) * 1)
        Me.txtP1.Text = Val(txtTarget.Text) + (Val(txtTolerance.Text) * 1)

        Me.txtN3.Text = Val(txtTarget.Text) - (Val(txtTolerance.Text) * 3)
        Me.txtP3.Text = Val(txtTarget.Text) + (Val(txtTolerance.Text) * 3)
    End Sub
End Class