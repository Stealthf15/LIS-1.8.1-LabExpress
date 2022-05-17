Imports System.IO
Imports System.Data
Imports System.Text
Imports System.Drawing.Imaging
Imports System.Drawing.Printing
Imports System.Collections.Generic
Imports Microsoft.Reporting.WinForms
Imports System.Reflection
Imports MySql.Data.MySqlClient

Public Class RDLCPrinterPrintReleased : Implements IDisposable

    Private _ReportInstance As Microsoft.Reporting.WinForms.LocalReport = Nothing
    Private _TempFileName As String = String.Empty
    Private _DefaultPrinterName As String = String.Empty

    Public Sub New(ByVal SampleID As String, ByVal Section As String, ByVal SubSection As String, ByVal PrintStatus As String, ByVal DefaultPrinterName As String)

        'reportResourceName = "NameofProject.NameofRDLC"
        Dim reportResourceName = Application.StartupPath & "\Reports\Result\" & SubSection & ".rdlc"
        Dim myReport As New Microsoft.Reporting.WinForms.LocalReport

        Dim myReportParams As List(Of ReportParameter) = Nothing

        myReport = New Microsoft.Reporting.WinForms.LocalReport
        myReport.ReportPath = reportResourceName

        'IF REPORT HAS DATASET WE NEED THIS
        '###################----DataSet----#################################################################
        'Order Table
        myReport.DataSources.Clear()
        Dim cn = New MySqlConnection(MyConnectionString)
        Dim SQL = "SELECT * FROM `order` WHERE `main_id` = '" & SampleID & "' AND testtype = '" & Section & "' AND sub_section = '" & SubSection & "'"
        Dim adapter = New MySqlDataAdapter(SQL, cn)
        Dim rdlc_table = New DataTable()
        adapter.Fill(rdlc_table)
        Dim rdlc_rds = New ReportDataSource("DataSet1", rdlc_table)
        myReport.DataSources.Add(rdlc_rds)

        'Result Table
        Connect()
        Dim result_SQL = "SELECT * FROM `result` WHERE `sample_id` = '" & SampleID & "' AND `print_status` = '1' AND section = '" & Section & "' AND sub_section = '" & SubSection & "' ORDER BY `order_no`"
        Dim result_adapter = New MySqlDataAdapter(result_SQL, conn)
        Dim result_rdlc_table = New DataTable()
        result_adapter.Fill(result_rdlc_table)
        Dim result_rdlc_rds = New ReportDataSource("DataSet2", result_rdlc_table)
        myReport.DataSources.Add(result_rdlc_rds)
        Disconnect()

        'Patient Remarks Table
        Connect()
        SQL = "SELECT * FROM `patient_remarks` WHERE `sample_id` = '" & SampleID & "'"
        adapter = New MySqlDataAdapter(SQL, conn)
        rdlc_table = New DataTable()
        adapter.Fill(rdlc_table)
        rdlc_rds = New ReportDataSource("DataSet3", rdlc_table)
        myReport.DataSources.Add(rdlc_rds)
        Disconnect()

        'ViewOrderMedtech Table
        Connect()
        SQL = "SELECT * FROM `viewordermedtech` WHERE `sample_id` = '" & SampleID & "'"
        adapter = New MySqlDataAdapter(SQL, conn)
        rdlc_table = New DataTable()
        adapter.Fill(rdlc_table)
        rdlc_rds = New ReportDataSource("DataSet4", rdlc_table)
        myReport.DataSources.Add(rdlc_rds)
        Disconnect()

        'ViewOrderPathologist Table
        Connect()
        SQL = "SELECT * FROM `vieworderpathologist` WHERE `sample_id` = '" & SampleID & "'"
        adapter = New MySqlDataAdapter(SQL, conn)
        rdlc_table = New DataTable()
        adapter.Fill(rdlc_table)
        rdlc_rds = New ReportDataSource("DataSet5", rdlc_table)
        myReport.DataSources.Add(rdlc_rds)
        Disconnect()

        'Company Profile Table
        Connect()
        SQL = "SELECT * FROM `company_profile`"
        adapter = New MySqlDataAdapter(SQL, conn)
        rdlc_table = New DataTable()
        adapter.Fill(rdlc_table)
        rdlc_rds = New ReportDataSource("DataSet6", rdlc_table)
        myReport.DataSources.Add(rdlc_rds)
        Disconnect()

        'ViewOrderVerified Table
        Connect()
        SQL = "SELECT * FROM `vieworderverified` WHERE `sample_id` = '" & SampleID & "'"
        adapter = New MySqlDataAdapter(SQL, conn)
        rdlc_table = New DataTable()
        adapter.Fill(rdlc_table)
        rdlc_rds = New ReportDataSource("DataSet8", rdlc_table)
        myReport.DataSources.Add(rdlc_rds)
        Disconnect()

        'Additional(Information)
        Connect()
        SQL = "SELECT * FROM `additional_info` WHERE `sample_id` = '" & SampleID & "'"
        adapter = New MySqlDataAdapter(SQL, conn)
        rdlc_table = New DataTable()
        adapter.Fill(rdlc_table)
        rdlc_rds = New ReportDataSource("DataSet7", rdlc_table)
        myReport.DataSources.Add(rdlc_rds)
        Disconnect()

        'Patient Information
        Connect()
        SQL = "SELECT * FROM `patient_info` WHERE `sample_id` = '" & SampleID & "'"
        adapter = New MySqlDataAdapter(SQL, conn)
        rdlc_table = New DataTable()
        adapter.Fill(rdlc_table)
        rdlc_rds = New ReportDataSource("DataSet9", rdlc_table)
        myReport.DataSources.Add(rdlc_rds)
        Disconnect()

        'Specimen Tracking
        Connect()
        SQL = "SELECT * FROM `specimen_tracking` WHERE `sample_id` LIKE '" & SampleID & "'"
        adapter = New MySqlDataAdapter(SQL, conn)
        rdlc_table = New DataTable()
        adapter.Fill(rdlc_table)
        rdlc_rds = New ReportDataSource("DataSet10", rdlc_table)
        myReport.DataSources.Add(rdlc_rds)
        Disconnect()

        '###################----End of DataSet----############################################################

        'IF REPORT HAS PARAMETERS WE NEED THIS
        '###################----Parameters----################################################################
        'Parameterized data to pass in report parameters

        Dim Print_Status As New ReportParameter("PrintStatus", PrintStatus)
        myReport.SetParameters(Print_Status)
        '######################----End of Parameters----###############################################################

        myReport.Refresh()
        myReport.ReportEmbeddedResource = reportResourceName

        _ReportInstance = myReport
        _DefaultPrinterName = DefaultPrinterName
    End Sub

    Public ReadOnly Property LocalReport() As LocalReport
        Get
            Return _ReportInstance
        End Get
    End Property

    Public Sub Print(ByVal NumberOfCopies As Integer)
        Export(_ReportInstance)
        m_currentPageIndex = 0
        PrintInteral(NumberOfCopies)
    End Sub

    Private m_currentPageIndex As Integer
    Private m_streams As IList(Of Stream)

    Private Sub Export(ByVal report As LocalReport)
        Dim deviceInfo As String =
          "<DeviceInfo>" +
          "  <OutputFormat>EMF</OutputFormat>" +
          "  <PageWidth>8.5in</PageWidth>" +
          "  <PageHeight>11in</PageHeight>" +
          "  <MarginTop>0.1in</MarginTop>" +
          "  <MarginLeft>0.2in</MarginLeft>" +
          "  <MarginRight>0.2in</MarginRight>" +
          "  <MarginBottom>0.1in</MarginBottom>" +
          "</DeviceInfo>"

        Dim warnings() As Warning = Nothing
        m_streams = New List(Of Stream)()

        Try
            report.Render("Image", deviceInfo, AddressOf CreateStream,
                       warnings)
        Catch ex As Exception
            MessageBox.Show(ex.ToString)
        End Try

        Dim stream As Stream
        For Each stream In m_streams
            stream.Position = 0
        Next
    End Sub

    Private Function CreateStream(ByVal name As String,
      ByVal fileNameExtension As String,
      ByVal encoding As Encoding, ByVal mimeType As String,
      ByVal willSeek As Boolean) As Stream
        _TempFileName = IO.Path.GetTempPath & "\" & name + "." + fileNameExtension
        Dim stream As Stream = New FileStream(_TempFileName, FileMode.Create)
        m_streams.Add(stream)
        Return stream
    End Function

    Private Sub PrintPage(ByVal sender As Object,
    ByVal ev As PrintPageEventArgs)

        m_streams(m_currentPageIndex).Position = 0

        Dim pageImage As New Metafile(m_streams(m_currentPageIndex))
        ev.Graphics.DrawImage(pageImage, ev.PageBounds)

        m_currentPageIndex += 1
        ev.HasMorePages = (m_currentPageIndex < m_streams.Count)

    End Sub

    Private Sub PrintInteral(ByVal NumberOfCopies As Integer)

        'IF THERE ARE NO STREAMS, WE CANT PRINT
        If m_streams Is Nothing Or m_streams.Count = 0 Then
            Return
        End If

        'CREATE PRINT DOCUMENT
        Dim printDoc As New PrintDocument()

        'IF THERE IS NO DEFAULT PRINTER NAME IN SETTINGS, PROMPT FOR ONE
        If _DefaultPrinterName = String.Empty Then
            Dim myPrintDialog As New PrintDialog
            myPrintDialog.PrinterSettings = printDoc.PrinterSettings
            If myPrintDialog.ShowDialog() = DialogResult.Cancel Then Return
        Else
            'WE HAVE A DEFAULT PRINTER NAME, ASSIGN IT
            printDoc.PrinterSettings.PrinterName = _DefaultPrinterName
        End If

        '    printDoc.PrinterSettings.Copies = CShort(NumberOfCopies)

        'IF PRINTER IS NOT VALID SHOW MESSAGE INDICATING ERROR
        If Not printDoc.PrinterSettings.IsValid Then
            MessageBox.Show("Error printing")
            Return
        End If

        'PRINT OUT THE PAGE TO THE SPECIFIED PRINTER
        AddHandler printDoc.PrintPage, AddressOf PrintPage

        If NumberOfCopies < 1 Then NumberOfCopies = 1
        For i As Integer = 1 To NumberOfCopies
            m_currentPageIndex = 0
            printDoc.Print()
        Next
    End Sub

    Private disposedValue As Boolean = False        ' To detect redundant calls

    ' IDisposable
    Protected Overridable Sub Dispose(ByVal disposing As Boolean)
        If Not Me.disposedValue Then
            If disposing Then
                If Not (m_streams Is Nothing) Then
                    Dim stream As Stream
                    For Each stream In m_streams
                        stream.Close()
                        stream.Dispose()
                    Next
                    m_streams = Nothing
                    If IO.File.Exists(_TempFileName) Then
                        Try
                            IO.File.Delete(_TempFileName)
                        Catch ex As Exception
                        End Try
                    End If
                End If
            End If
        End If
        Me.disposedValue = True
    End Sub

#Region " IDisposable Support "
    ' This code added by Visual Basic to correctly implement the disposable pattern.
    Public Sub Dispose() Implements IDisposable.Dispose
        ' Do not change this code.  Put cleanup code in Dispose(ByVal disposing As Boolean) above.
        Dispose(True)
        GC.SuppressFinalize(Me)
    End Sub
#End Region

End Class