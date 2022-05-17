Imports System
Imports System.IO
Imports System.Text
Imports System.Collections.Generic
Imports System.Linq
Imports System.Drawing
Imports DevExpress.XtraEditors
Imports DevExpress.XtraPrinting.BarCode

Public Class frmDBRestore

    Dim AppPath As String = Application.StartupPath
    Dim FolderPath As String = AppPath & "\Backup"

    Private Sub btnSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSave.Click
        RestoreDatase()
    End Sub

    Private Sub btnClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnClose.Click
        Me.Close()
    End Sub

    Private Sub RestoreDatase()
        Dim myProcess As New Process()
        myProcess.StartInfo.FileName = "cmd.exe"
        myProcess.StartInfo.UseShellExecute = False
        myProcess.StartInfo.WorkingDirectory = "C:\xampp\mysql\bin\"
        myProcess.StartInfo.RedirectStandardInput = True
        myProcess.StartInfo.RedirectStandardOutput = True
        myProcess.Start()
        Dim myStreamWriter As StreamWriter = myProcess.StandardInput
        Dim mystreamreader As StreamReader = myProcess.StandardOutput
        myStreamWriter.WriteLine("mysql -u " & My.Settings.UID & " -p" & My.Settings.PWD & " " & My.Settings.DBName & " < " & FolderPath & "\" & txtCode.Text & "")
        myStreamWriter.Close()
        myProcess.WaitForExit()
        myProcess.Close()

    End Sub

    Private Sub frmBackup_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Me.txtCode.Text = "Backup_" + Replace(Now.ToShortDateString, "/", "") + ".sql"
    End Sub
End Class