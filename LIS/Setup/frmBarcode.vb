Imports System
Imports System.Collections.Generic
Imports System.Linq
Imports System.Drawing
Imports System.Drawing.Printing
Imports DevExpress.XtraEditors
Imports DevExpress.XtraPrinting.BarCode

Public Class frmBarcode
    Dim picCode As New PictureBox

    Private Sub btnSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSave.Click

        'bcCode.Symbology = New  With {.CharacterSet = }
        bcCode.Text = txtCode.Text
        Dim bmp = New Bitmap(bcCode.Width, bcCode.Height)

        bcCode.DrawToBitmap(bmp, bcCode.ClientRectangle)

        Dim arrImage() As Byte

        Dim myMS As New IO.MemoryStream
        If Not IsNothing(picCode.Image) Then
            picCode.Image.Save(myMS, Imaging.ImageFormat.Jpeg)
            arrImage = myMS.GetBuffer
        Else
            arrImage = Nothing
        End If
    End Sub

    Private Sub btnClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnClose.Click
        Me.Close()
    End Sub

    Private Sub frmBarcode_FormClosed(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
        Me.Dispose()
    End Sub

    Private Sub btnPrint_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnPrint.Click
        PrintImage()
    End Sub

    Private Sub PrintImage()

        'Set label size and margins
        Dim ps As New PaperSize("2x1 Inch Label Size", 200, 100) 'Label Size
        pd.DefaultPageSettings.PaperSize = ps

        pd.DefaultPageSettings.Margins = New Margins(10, 10, 10, 10)
        pd.OriginAtMargins = False
        pd.DefaultPageSettings.Landscape = False

        'AddHandler pd.PrintPage, New PrintPageEventHandler(AddressOf pd_PrintPage)
        AddHandler pd.PrintPage, New PrintPageEventHandler(AddressOf pd_PrintPage)
        PrintPreviewDialog1.Document = pd
        PrintPreviewDialog1.ShowDialog()
        'PrintPreviewDialog1.Document.Print()

    End Sub

    Private Sub pd_PrintPage(ByVal sender As Object, ByVal e As PrintPageEventArgs) Handles pd.PrintPage

        Dim imageTopr = New Bitmap(Me.picCode.Width, Me.picCode.Height)

        picCode.DrawToBitmap(imageTopr, picCode.ClientRectangle)

        e.Graphics.InterpolationMode = Drawing2D.InterpolationMode.HighQualityBilinear
        e.Graphics.SmoothingMode = Drawing2D.SmoothingMode.HighQuality
        e.Graphics.DrawImage(imageTopr, New Rectangle(0, 0, Me.bcCode.Width, Me.bcCode.Height))

        e.Graphics.DrawImage(imageTopr, 0, 0, 175, 75)
    End Sub

End Class