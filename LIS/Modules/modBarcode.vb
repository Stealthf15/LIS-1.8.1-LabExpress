Imports System.Drawing.Printing
Imports DevExpress.XtraPrinting.BarCode

Module modBarcode
    Private PrintDocType As String = "Barcode"
    Private StrPrinterName As String = My.Settings.BCPrinterName
    Private BarcodeControl As New DevExpress.XtraEditors.BarCodeControl

    Private Width As Integer = My.Settings.BCWidth
    Private Height As Integer = My.Settings.BCHeight

    Private WithEvents PrintDocument As New PrintDocument
    Private PrintPreviewDialog As New PrintPreviewDialog

    Public Specimen As String, BCsID As String, BCpID As String, BCpName As String, BCBDate As String, BCGender As String, BCTest As String, BCRequest As String, BCAge As String

    Public BCSection, BCSubSection As String

    Public Sub PrintBarcode(bc_Specimen As String, bc_sID As String, bc_pID As String, bc_pName As String, bc_BDate As String, bc_Gender As String, bc_Section As String, bc_SubSection As String)
        If My.Settings.PrintBarcode Then
            Specimen = bc_Specimen
            BCsID = bc_sID
            BCpID = bc_pID
            BCpName = bc_pName.ToUpper
            BCBDate = bc_BDate
            BCGender = bc_Gender.ToUpper
            BCSection = bc_Section
            BCSubSection = bc_SubSection

            BCRequest = retrieveTest(BCsID, BCSection, BCSubSection)
            BCAge = GetAge(BCBDate)

            'Paper Size
            Dim papersize As New PaperSize("2x1 Inch Label Size", Width, Height)
            PrintDocument.PrinterSettings.PrinterName = StrPrinterName
            PrintDocument.PrintController = New StandardPrintController

            PrintDocument.DefaultPageSettings.PaperSize = papersize
            PrintDocument.DefaultPageSettings.Margins = New Margins(0, 0, 0, 0)
            PrintDocument.OriginAtMargins = False
            PrintDocument.DefaultPageSettings.Landscape = False

            'Printing
            PrintPreviewDialog.Document = PrintDocument
            PrintPreviewDialog.Document.Print()
            'PrintPreviewDialog.ShowDialog()
        End If
    End Sub

    Private Sub PrintDocument_PrintPage(ByVal sender As System.Object, ByVal e As System.Drawing.Printing.PrintPageEventArgs) Handles PrintDocument.PrintPage
        BarcodeControl.HorizontalAlignment = DevExpress.Utils.HorzAlignment.Center
        BarcodeControl.HorizontalTextAlignment = DevExpress.Utils.HorzAlignment.Center
        BarcodeControl.BorderStyle = DevExpress.XtraEditors.Controls.BorderStyles.NoBorder
        BarcodeControl.ShowText = False
        BarcodeControl.AutoModule = True

        BarcodeControl.BackColor = Color.White
        BarcodeControl.ForeColor = Color.Black

        BarcodeControl.Symbology = New Code128Generator With {.CharacterSet = Code128Charset.CharsetAuto}
        BarcodeControl.Text = BCsID

        BarcodeControl.Width = 440
        BarcodeControl.Height = 131

        Dim bm As New Bitmap(440, 131)

        BarcodeControl.DrawToBitmap(bm, BarcodeControl.ClientRectangle)

        'Font and Size of Font
        Dim myFont As Font = New Font("Tahoma", 5.8, FontStyle.Bold)
        Dim FontArrow As Font = New Font("Tahoma", 12, FontStyle.Bold)
        Dim myBrush As Brush = Brushes.Black

        'DEDVMH Barcode Format
        e.Graphics.InterpolationMode = Drawing2D.InterpolationMode.HighQualityBilinear
        e.Graphics.SmoothingMode = Drawing2D.SmoothingMode.HighQuality
        e.Graphics.DrawString(BCsID & " " & Specimen & " " & BCGender & " " & Format(CDate(BCBDate), "MM-dd-yyyy"), myFont, myBrush, 10, 15)
        'e.Graphics.DrawString(BCAge, myFont, myBrush, 145, 5)
        'e.Graphics.DrawString(BCGender, myFont, myBrush, 160, 5)
        'e.Graphics.DrawString("→", FontArrow, myBrush, 170, 30)
        e.Graphics.DrawImage(bm, 10, 25, 110, 55)
        'e.Graphics.DrawString(BCpID, myFont, myBrush, 30, 70)
        e.Graphics.DrawString(BCpName, myFont, myBrush, 10, 80)
        e.Graphics.DrawString(Format(Now, "MM-dd-yy"), myFont, myBrush, 10, 90)

        'MAMC Barcode Format
        'If Specimen = "CBC" Then
        '    Specimen = "WB"
        'ElseIf Specimen = "Coagulation" Then
        '    Specimen = "Plasma"
        'ElseIf Specimen = "Clinical Chemistry" Then
        '    Specimen = "Serum"
        'End If

        'e.Graphics.InterpolationMode = Drawing2D.InterpolationMode.HighQualityBilinear
        'e.Graphics.SmoothingMode = Drawing2D.SmoothingMode.HighQuality

        'e.Graphics.DrawString(BCsID, myFont, myBrush, 5, 5)
        'e.Graphics.DrawString(Specimen, myFont, myBrush, 80, 5)
        'e.Graphics.DrawString(Format(Now, "MM-dd-yy"), myFont, myBrush, 110, 5)
        'e.Graphics.DrawImage(bm, 10, 15, 100, 40)
        'e.Graphics.DrawString("", myFont, myBrush, 5, 55)
        'e.Graphics.DrawString(BCGender, myFont, myBrush, 45, 55)
        'e.Graphics.DrawString(BCpID, myFont, myBrush, 5, 64)
        'e.Graphics.DrawString(BCpName, myFont, myBrush, 40, 64)
        ''e.Graphics.DrawString(Test, myFont, myBrush, 5, 63)
    End Sub


    Private Function GetAge(bBDate As String) As String
        Dim bAge As String = ""

        If bBDate = "" Then

        Else
            Dim parseBdate As DateTime

            parseBdate = Convert.ToDateTime(bBDate)
            bAge = parseBdate.Year

            If parseBdate = Nothing Then
                Exit Function
            End If

            Dim Birthday As Date = parseBdate
            Dim endDate As Date = Date.Now

            Dim timeSpan As TimeSpan = endDate.Subtract(Birthday)
            Dim totalDays As Integer = timeSpan.TotalDays
            Dim totalYears As Integer = Math.Truncate(totalDays / 365)
            Dim totalMonths As Integer = Math.Truncate((totalDays Mod 365) / 30)
            Dim remainingDays As Integer = Math.Truncate((totalDays Mod 365) Mod 30)

            If totalDays <= 61 Then
                bAge = totalDays
            ElseIf totalDays >= 62 And totalDays <= 364 Then
                bAge = totalMonths
            ElseIf totalDays >= 365 Then
                bAge = totalYears
            End If

        End If
        Return bAge
    End Function

    Private Function retrieveTest(SampleID As String, Section As String, SubSection As String) As String
        BCTest = ""

        Connect()
        rs.Parameters.Clear()
        rs.Parameters.AddWithValue("@sample_id", SampleID)
        rs.Parameters.AddWithValue("@Section", Section)
        rs.Parameters.AddWithValue("@SubSection", SubSection)

        rs.Connection = conn
        rs.CommandType = CommandType.Text
        rs.CommandText = "SELECT DISTINCT `his_code` FROM `patient_order` WHERE `sample_id` = @sample_id And `testtype` = @Section And `sub_section` = @SubSection"
        reader = rs.ExecuteReader
        If reader.HasRows Then
            While reader.Read
                BCTest &= reader(0).ToString & ", "
                'txtTest.Text &= Test.TrimEnd(", ")
            End While
        End If
        Disconnect()

        Return BCTest

    End Function
End Module
