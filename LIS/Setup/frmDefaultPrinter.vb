Imports MySql.Data.MySqlClient
Imports System.Drawing.Printing

Public Class frmDefaultPrinter
    Public Sub LoadPrinter()
        lvList.Items.Add("")
        For Each prnt As String In PrinterSettings.InstalledPrinters
            lvList.Items.Add(prnt)
        Next
    End Sub

    Private Sub frmDefaultPrinter_FormClosed(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
        Me.Dispose()
    End Sub

    Private Sub frmDefaultPrinter_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        LoadPrinter()
    End Sub

    Private Sub btnRefresh_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        LoadPrinter()
    End Sub

    Private Sub btnClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnClose.Click
        Me.Close()
    End Sub

    Private Sub btnEdit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSave.Click
        My.Settings.DefaultPrinter = Me.lvList.FocusedItem.Text
        My.Settings.Save()

        MessageBox.Show("Default Printer Set.", "Default Printer", MessageBoxButtons.OK, MessageBoxIcon.Information)
        Me.Close()
    End Sub
End Class