Imports System.IO
Imports System.IO.Ports
Imports System.Threading
Imports System.Threading.Thread
Imports MySql.Data.MySqlClient
Imports Microsoft.Reporting.WinForms
Imports System.Data.SqlClient

Module modVariables
#Region "ASTM String Readline for COMMPOrt Communication"
    '--------------------------Global Declaration of Variables----------------------
    Public RXString As String 'Used as string to handle COMPort Readline
    Public RXString1 As String 'Used as string to handle COMPort Readline

    'Public COMPort As New List(Of SerialPort) 'Comport Declaration

    Public ReceivedFromSerial As String = ""
    Public ReceivedFromSerial1 As String = ""
    Public DataReceived As Char

    Public wh As EventWaitHandle
    Public orderFlag As Boolean = False

#End Region

#Region "Segregated Data Handler ASTM"
    '--------------------Data Handler(instrument to host file transfer)-------------
#Region "Header Data"
    '----------------------------------Header Data----------------------------------
    Public HeaderData1 As String = ""
    Public HeaderData2 As String = ""
    Public HeaderData3 As String = ""
    Public HeaderData4 As String = ""
#End Region
#Region "Result Data"
    '----------------------------------Result Data----------------------------------
    'Public Test As String = ""
    Public Data1 As String = ""
    Public Data2 As String = ""
    Public Data3 As String = ""
    Public Data4 As String = ""
    Public Data5 As String = ""
    Public Data6 As String = ""
    Public Data7 As String = ""
    Public Data8 As String = ""
    Public Data9 As String = ""
    Public Data10 As String = ""
    Public Data11 As String = ""
    Public Data12 As String = ""
    Public Data13 As String = ""
    Public Data14 As String = ""

    Public FinalData2 As String = ""
    Public FinalData5 As String = ""
#End Region
#Region "Patient Data"
    '----------------------------------Patient Data---------------------------------
    Public PatientData1 As String = ""
    Public PatientData2 As String = ""
    Public PatientData3 As String = ""
    Public PatientData4 As String = ""
    Public PatientData5 As String = ""
    Public PatientData6 As String = ""

    Public PatientLastName As String = ""
    Public PatientFirstName As String = ""

    Public TypeSelection As String = ""
#End Region
#Region "Patient Age"
    '---------------------------------Patient Age Computer---------------------------
    Public FinalbDate As String
    Public ConvertedbDate As Date
    Public FinalAge As String
#End Region
#Region "Order Data"
    '-----------------------------------Order Data----------------------------------
    Public OrderData1 As String = ""
    Public OrderData2 As String = ""
    Public OrderData3 As String = ""
    Public OrderData4 As String = ""
    Public OrderData5 As String = ""
    Public OrderData6 As String = ""
    Public OrderData7 As String = ""

    Public FinalOrderData3 As String = ""
    Public FinalOrderID As String = ""
    Public OrderSampleID As String = ""
#End Region
#Region "Comment Data"
    '---------------------------------------Comment Data-------------------------------------
    Public CommentData1 As String = ""
    Public CommentData2 As String = ""
    Public CommentData3 As String = ""
#End Region
#Region "Universal ID"
    '---------------------------------------Universal ID--------------------------------------
    Public PatientID As String = ""
    '------------------------------------End of Data Handler----------------------------------
#End Region
#Region "Data Hanlder Bi-directional"
    '--------------------Data Handler(host to instrument file transfer)-----------------------
    '-----------------------------------Patient Data Out--------------------------------------
    Public PatientID_Out As String
    Public PatientName_Out As String
    Public PatientBDate_Out As String
    Public PatientSex_Out As String
    Public PatientDept_Out As String
    Public PatientDoctor_Out As String

    '-------------------------------------Order Data Out--------------------------------------
    Public OrderSampleID_out As String
    Public OrderTest_Out As String
    Public OrderPriority_Out As String
    Public OrderDate_Out As String
    Public OrderAction_Out As String
    Public OrderDescriptor_Out As String
    Public OrderChannel_out As String

    '------------------------------------End of Data Handler----------------------------------
#End Region
#End Region

#Region "Database Connection Variables"
    '----------------------------Database Connection String Variables-------------------------
    Public ServerName = My.Settings.Server
    Public UID = My.Settings.UID
    Public PWD = My.Settings.PWD = ""
    Public DBName = My.Settings.DBName

    '----------------------------Database Connection String-------------------------
    Public MyConnectionString = My.Settings.MyConnectionString

    Public MyConnectionStringSQL = My.Settings.MyConnectionStringSQL
#End Region

#Region "MySQL.Data.MySQLClient Variables"
    '---------------------------------MySQL Variables-------------------------------
    Public conn As New MySqlConnection
    Public rs As New MySqlCommand
    Public reader As MySqlDataReader
    Public adapter As New MySqlDataAdapter
    Public dt As New DataTable

    Public conn1 As New MySqlConnection
    Public rs1 As New MySqlCommand
    Public reader1 As MySqlDataReader

    Public conn2 As New MySqlConnection
    Public rs2 As New MySqlCommand
    Public reader2 As MySqlDataReader

    Public conn3 As New MySqlConnection
    Public rs3 As New MySqlCommand
    Public reader3 As MySqlDataReader

    '----------------------------------OLEDB Variables---------------------------------
    Public connSQL As New SqlConnection
    Public rsSQL As New SqlCommand
    Public readerSQL As SqlDataReader

    'Public connOLEDB2 As New OleDbConnection
    'Public rsOLEDB2 As New OleDbCommand
    'Public readerOLEDB2 As OleDbDataReader

    '------------------------------HIS MySQL Variables-------------------------------
    Public HISconn As New MySqlConnection
    Public HISrs As New MySqlCommand
    Public HISreader As MySqlDataReader
    Public HISadapter As New MySqlDataAdapter
    Public HISdt As New DataTable
#End Region

#Region "Global Declaration"
    ''------------------------------------RDLC Adapters-------------------------
    'Public rdlc_table As DataTable
    'Public rdlc_rds As ReportDataSource
    'Public SQL As String = ""

    '------------------------------ListView Subitem Handler-------------------------
    Public iTemSubItems As New ListViewItem
    '------------------------------------User Details-------------------------------
    Public CurrID As String
    Public CurrUser As String
    Public CurrPost As String
    Public CurrDept As String
    Public CurrType As String
    Public CurrEmail As String
    Public UserRelease As String
    Public UserIDRelease As String

    Public imgBlob() As Byte

    Public iItem As New ListViewItem

    Public OrderNo As Integer = 0

    '((( VARIABLES )))
    ''Public _Shadow As DropShadow
    Public NewPoint As New Point
    Public New_X, New_Y As Integer

    Public TestTypes As String = ""

    Public unit, range, unit_conv, range_conv, age_group, patient_id, test_name, test_code, test_group, sub_section As String
    Public order_no As String

    Public Reference_No As String = ""
    Public SampleID As String = ""
    Public Reference_No_P As String = ""
#End Region

End Module
