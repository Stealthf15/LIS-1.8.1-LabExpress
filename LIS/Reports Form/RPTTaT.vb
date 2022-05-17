Imports Microsoft.Reporting.WinForms

Public Class RPTTaT
    Private Sub RPTTaT_FormClosed(sender As Object, e As FormClosedEventArgs) Handles Me.FormClosed
        Me.Dispose()
    End Sub

    Private Sub RPTTaT_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Me.ReportViewer1.RefreshReport()
    End Sub
End Class