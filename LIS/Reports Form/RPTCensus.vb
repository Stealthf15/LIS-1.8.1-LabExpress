﻿Imports Microsoft.Reporting.WinForms

Public Class RPTCensus

    Private Sub RPTCensus_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        Me.Dispose()
    End Sub

    'Private Sub RPTWorksheet_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
    '    db_sbsi_lis_universalDataSet.EnforceConstraints = False
    'End Sub
End Class