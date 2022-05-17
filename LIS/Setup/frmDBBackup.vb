Imports System
Imports System.IO
Imports System.Text
Imports System.Collections.Generic
Imports System.Linq
Imports System.Drawing
Imports DevExpress.XtraEditors
Imports DevExpress.XtraPrinting.BarCode

Public Class frmDBBackup

    Dim AppPath As String = Application.StartupPath
    Dim FolderPath As String = AppPath & "\Backup"

    Private Sub btnSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSave.Click
        CreateBackup()
    End Sub

    Private Sub btnClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnClose.Click
        Me.Close()
    End Sub

    Private Sub CreateBackup()
        If Not Directory.Exists(FolderPath) Then
            Directory.CreateDirectory(FolderPath)
        End If
        Process.Start("C:\xampp\mysql\bin\mysqldump.exe", " -u " & My.Settings.UID & " -p" & My.Settings.PWD & " " & My.Settings.DBName & " -r """ & FolderPath & "\" & txtCode.Text & """")
    End Sub


    'Dim OutputStream As System.IO.StreamWriter

    'Sub OnDataReceived1(ByVal Sender As Object, ByVal e As System.Diagnostics.DataReceivedEventArgs)
    '    If e.Data IsNot Nothing Then
    '        Dim text As String = e.Data
    '        Dim bytes As Byte() = Encoding.Default.GetBytes(text)
    '        text = Encoding.UTF8.GetString(bytes)
    '        OutputStream.WriteLine(text)
    '    End If
    'End Sub

    'Sub CreateBackup()
    '    'If Not Directory.Exists(FolderPath) Then
    '    '    Directory.CreateDirectory(FolderPath)
    '    'End If

    '    'Dim mysqldumpPath As String = "C:\xampp\mysql\bin\mysqldump.exe"
    '    'Dim host As String = My.Settings.Server
    '    'Dim user As String = My.Settings.UID
    '    'Dim pswd As String = My.Settings.PWD
    '    'Dim dbnm As String = My.Settings.DBName
    '    'Dim cmd As String = String.Format("-h{0} -u{1} -p{2} {3}", host, user, pswd, dbnm)
    '    'Dim filePath As String = FolderPath & "\" & txtCode.Text
    '    'OutputStream = New System.IO.StreamWriter(filePath, False, System.Text.Encoding.UTF8)

    '    'Dim startInfo As System.Diagnostics.ProcessStartInfo = New System.Diagnostics.ProcessStartInfo()
    '    'startInfo.FileName = mysqldumpPath
    '    'startInfo.Arguments = cmd

    '    'startInfo.RedirectStandardError = True
    '    'startInfo.RedirectStandardInput = False
    '    'startInfo.RedirectStandardOutput = True
    '    'startInfo.UseShellExecute = False
    '    'startInfo.CreateNoWindow = False
    '    'startInfo.ErrorDialog = False

    '    'Dim proc As System.Diagnostics.Process = New System.Diagnostics.Process()
    '    'proc.StartInfo = startInfo
    '    'AddHandler proc.OutputDataReceived, AddressOf OnDataReceived1
    '    'proc.Start()
    '    'proc.BeginOutputReadLine()
    '    'proc.WaitForExit()

    '    'OutputStream.Flush()
    '    'OutputStream.Close()
    '    'proc.Close()
    'End Sub

    'Private Sub RestoreDatase()
    '    Dim myProcess As New Process()
    '    myProcess.StartInfo.FileName = "cmd.exe"
    '    myProcess.StartInfo.UseShellExecute = False
    '    myProcess.StartInfo.WorkingDirectory = "C:\xampp\mysql\bin\mysqldump.exe"
    '    myProcess.StartInfo.RedirectStandardInput = True
    '    myProcess.StartInfo.RedirectStandardOutput = True
    '    myProcess.Start()
    '    Dim myStreamWriter As StreamWriter = myProcess.StandardInput
    '    Dim mystreamreader As StreamReader = myProcess.StandardOutput
    '    myStreamWriter.WriteLine("mysql -u username -p***** databasename < C:\backup.sql ")
    '    myStreamWriter.Close()
    '    myProcess.WaitForExit()
    '    myProcess.Close()
    'End Sub

    Private Sub frmBackup_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Me.txtCode.Text = "Backup_" + Replace(Now.ToShortDateString, "/", "") + ".sql"
    End Sub
End Class