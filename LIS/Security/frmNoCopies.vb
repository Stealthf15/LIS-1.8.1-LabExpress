Imports System.Drawing.Printing

Public Class frmNoCopies
    Public PrintType As String = ""

    Public Sub LoadPrinter()
        Me.cboPrinter.Properties.Items.Clear()
        For Each prnt As String In PrinterSettings.InstalledPrinters
            cboPrinter.Properties.Items.Add(prnt)
        Next
    End Sub

    Private Sub frmDefaultPrinter_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        LoadPrinter()
        Me.cboPrinter.Text = My.Settings.DefaultPrinter
    End Sub

    Private Sub btnClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnClose.Click
        Me.Close()
        Me.Dispose()
    End Sub

    Private Sub btnPrint_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnPrint.Click
        'If PrintType = "Hema" Then
        '    Using myRDLCPrinter As New RDLCPrinterHema(frmResultNew.MainSampleID, "", Me.cboPrinter.Text)
        '        myRDLCPrinter.Print(Me.txtCopies.Text)
        '    End Using
        'ElseIf PrintType = "Hema Reprint" Then
        '    Using myRDLCPrinter As New RDLCPrinterHema(frmResultOrdered.mainID, "REPRINT", Me.cboPrinter.Text)
        '        myRDLCPrinter.Print(Me.txtCopies.Text)
        '    End Using
        'ElseIf PrintType = "Chem" Then
        '    Using myRDLCPrinter As New RDLCPrinterChem(frmResultNewOther.MainSampleID, "", Me.cboPrinter.Text)
        '        myRDLCPrinter.Print(Me.txtCopies.Text)
        '    End Using
        'ElseIf PrintType = "Order" Then
        '    Using myRDLCPrinter As New RDLCPrinterPWorkSheet(frmPatientOrderAE.txtSampleID.Text, CurrUser, Me.cboPrinter.Text)
        '        myRDLCPrinter.Print(Me.txtCopies.Text)
        '    End Using
        'ElseIf PrintType = "Manual" Then
        '    Using myRDLCPrinter As New RDLCPrinterChem(frmResultNewManual.MainSampleID, "", Me.cboPrinter.Text)
        '        myRDLCPrinter.Print(Me.txtCopies.Text)
        '    End Using
        'End If

        Me.Close()
        Me.Dispose()

    End Sub
End Class