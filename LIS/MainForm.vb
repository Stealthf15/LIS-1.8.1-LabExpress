Imports System.Environment
Imports System.IO
Imports DevExpress.XtraBars.Navigation

Imports Bluegrams.Application

Public Class MainFOrm

    Public Shared Sub Main()
        ' Get the path to the Application Data folder
        Dim FolderPath As String = GetFolderPath(SpecialFolder.LocalApplicationData) & "\AppSupport\"

        'Create Folder
        Directory.CreateDirectory(FolderPath)

        PortableSettingsProvider.SettingsFileName = "LIS.cs"
        PortableSettingsProvider.SettingsDirectory = FolderPath
    End Sub

    Public Sub GetHospitalName()
        Try
            Connect()
            rs.Connection = conn
            rs.CommandType = CommandType.Text
            rs.CommandText = "SELECT `comp_name` FROM `company_profile`"
            reader = rs.ExecuteReader
            reader.Read()
            If reader.HasRows Then
                lblHospitalName.Text = reader(0).ToString
            End If
            Disconnect()
        Catch ex As Exception
            Disconnect()
            MessageBox.Show(ex.Message, "System Error", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            Exit Sub
        End Try
    End Sub
    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Main()
        frmLogin.ShowDialog()
        GetHospitalName()
    End Sub

    'Private Sub AccordionControl1_CustomDrawElement(sender As Object, e As DevExpress.XtraBars.Navigation.CustomDrawElementEventArgs) Handles AccordionControl.CustomDrawElement

    '    Dim AControl As AccordionControl = sender
    '    Dim vertLineWidth As Integer = 3

    '    If (e.ObjectInfo.State = DevExpress.Utils.Drawing.ObjectState.Pressed) Then

    '        e.Handled = True
    '        e.Cache.FillRectangle(Color.Blue, New Rectangle(e.ObjectInfo.HeaderBounds.Location, New Size(vertLineWidth, e.ObjectInfo.HeaderBounds.Height)))
    '        Dim pt = New Point(e.ObjectInfo.HeaderBounds.Location.X + vertLineWidth, e.ObjectInfo.HeaderBounds.Location.Y)
    '        e.Cache.FillRectangle(Color.FromArgb(90, Color.Green), New Rectangle(pt, New Size(e.ObjectInfo.HeaderBounds.Width - vertLineWidth, e.ObjectInfo.HeaderBounds.Height)))
    '    End If
    '    e.Appearance.ForeColor = Color.White
    '    e.DrawText()

    'End Sub

#Region "CheckConnection"
    Private Sub tmCheckCon_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tmCheckCon.Tick

        'If CurrType = "Phlebotomy" Or CurrType = "Receptionist" Then
        '    mnuSetting.Visible = False
        '    mnuHelp.Visible = False
        '    mnuSysSettings.Visible = False

        '    aceSection.Visible = False
        '    aceReports.Visible = False
        '    btnRealtime.Visible = False
        '    btnDelta.Visible = False

        '    btnPatientOrder.Visible = True
        '    btnPatient.Visible = True
        '    btnAccounts.Visibility = False
        'ElseIf CurrType = "Laboratory" Or CurrType = "MedTech" Then
        '    mnuSetting.Visible = False
        '    mnuHelp.Visible = False
        '    mnuSysSettings.Visible = False

        '    aceSection.Visible = True
        '    aceReports.Visible = True
        '    btnRealtime.Visible = False
        '    btnDelta.Visible = False
        '    btnPatientOrder.Visible = True 'Temporary True
        '    btnAccounts.Enabled = False 'Prevent creation of user when MedTech Account
        '    btnPatient.Visible = False
        '    btnAccounts.Visibility = False
        'ElseIf CurrType = "Administrator" Then
        '    mnuSetting.Visible = True
        '    mnuHelp.Visible = True
        '    mnuSysSettings.Visible = True
        '    aceSection.Visible = True
        '    aceReports.Visible = True
        '    btnRealtime.Visible = True
        '    btnDelta.Visible = True
        '    btnAccounts.Enabled = True 'Prevent creation of user when MedTech Account
        '    btnPatientOrder.Visible = True
        '    btnPatient.Visible = True
        '    btnAccounts.Visibility = True
        'ElseIf CurrType = "Super Admin" Then
        '    mnuSetting.Visible = True
        '    mnuHelp.Visible = True
        '    mnuSysSettings.Visible = True
        '    aceSection.Visible = True
        '    aceReports.Visible = True
        '    btnRealtime.Visible = True
        '    btnDelta.Visible = True
        '    btnAccounts.Enabled = True 'Prevent creation of user when MedTech Account
        '    btnPatientOrder.Visible = True
        '    btnPatient.Visible = True
        '    btnAccounts.Visibility = True
        'Else
        '    mnuSetting.Visible = False
        '    mnuHelp.Visible = False
        '    mnuSysSettings.Visible = False
        '    aceSection.Visible = False
        '    aceReports.Visible = False
        '    btnDelta.Visible = False
        '    btnPatientOrder.Visible = False
        '    btnPatient.Visible = False
        '    btnAccounts.Visibility = False
        'End If

        Me.mnuUser.Text = CurrUser
        lblDateTime.Text = "Today: " & Now.ToLongDateString & " " & Now.ToShortTimeString
    End Sub
#End Region

#Region "Accordion Menu Click Events"
    Private Sub accEmailPush_Click(sender As Object, e As EventArgs) Handles accEmailPush.Click
        For Each f As Form In Me.MdiChildren
            f.Close()
        Next
        lblTitle.Text = "EMAIL PUSH RESULT"
        frmEmailPush.WindowState = FormWindowState.Maximized
        frmEmailPush.MdiParent = Me
        frmEmailPush.Show()
    End Sub

    Private Sub aceWHema_Click(sender As Object, e As EventArgs) Handles aceHema.Click
        For Each f As Form In Me.MdiChildren
            f.Close()
        Next
        lblTitle.Text = "HEMATOLOGY WORKLIST"
        frmHemaWorklist.WindowState = FormWindowState.Maximized
        frmHemaWorklist.MdiParent = Me
        frmHemaWorklist.Show()
    End Sub

    Private Sub acChem_Click(sender As Object, e As EventArgs) Handles aceChem.Click
        For Each f As Form In Me.MdiChildren
            f.Close()
        Next
        lblTitle.Text = "CHEMISTRY WORKLIST"
        frmChemWorklist.WindowState = FormWindowState.Maximized
        frmChemWorklist.MdiParent = Me
        frmChemWorklist.Show()
    End Sub

    Private Sub aceUrinWorklist_Click(sender As Object, e As EventArgs) Handles aceUrin.Click
        For Each f As Form In Me.MdiChildren
            f.Close()
        Next
        lblTitle.Text = "URINALYSIS WORKLIST"
        frmUrinWorklist.WindowState = FormWindowState.Maximized
        frmUrinWorklist.MdiParent = Me
        frmUrinWorklist.Show()
    End Sub

    Private Sub aceFecalWorklist_Click(sender As Object, e As EventArgs) Handles aceFecal.Click
        For Each f As Form In Me.MdiChildren
            f.Close()
        Next
        lblTitle.Text = "FECALYSIS WORKLIST"
        frmFecalWorklist.WindowState = FormWindowState.Maximized
        frmFecalWorklist.MdiParent = Me
        frmFecalWorklist.Show()
    End Sub

    Private Sub aceSerology_Click(sender As Object, e As EventArgs) Handles aceSerology.Click
        For Each f As Form In Me.MdiChildren
            f.Close()
        Next
        lblTitle.Text = "SEROLOGY WORKLIST"
        frmSeroWorklist.WindowState = FormWindowState.Maximized
        frmSeroWorklist.MdiParent = Me
        frmSeroWorklist.Show()
    End Sub

    Private Sub btnSettings_ItemClick(ByVal sender As System.Object, ByVal e As DevExpress.XtraBars.ItemClickEventArgs) Handles btnSettings.ItemClick
        For Each f As Form In Me.MdiChildren
            f.Close()
        Next
        lblTitle.Text = "QC SETTINGS"
        frmControlSetting.WindowState = FormWindowState.Maximized
        frmControlSetting.MdiParent = Me
        frmControlSetting.Show()
    End Sub

    Private Sub btnMedtech_ItemClick(ByVal sender As System.Object, ByVal e As DevExpress.XtraBars.ItemClickEventArgs) Handles btnMedTech.ItemClick
        For Each f As Form In Me.MdiChildren
            f.Close()
        Next
        lblTitle.Text = "SIGNATORIES"
        frmMedTech.WindowState = FormWindowState.Maximized
        frmMedTech.MdiParent = Me
        frmMedTech.Show()
    End Sub

    Private Sub bntTestName_ItemClick(ByVal sender As System.Object, ByVal e As DevExpress.XtraBars.ItemClickEventArgs) Handles btnTestName.ItemClick
        frmTest.WindowState = FormWindowState.Maximized
        frmTest.MdiParent = Me
        frmTest.Show()
    End Sub

    Private Sub btnPathologist_ItemClick(ByVal sender As System.Object, ByVal e As DevExpress.XtraBars.ItemClickEventArgs) Handles btnPathologist.ItemClick
        For Each f As Form In Me.MdiChildren
            f.Close()
        Next
        lblTitle.Text = "PATHOLOGIST"
        frmPathologist.WindowState = FormWindowState.Maximized
        frmPathologist.MdiParent = Me
        frmPathologist.Show()
    End Sub

    Private Sub btnDepartment_ItemClick(ByVal sender As System.Object, ByVal e As DevExpress.XtraBars.ItemClickEventArgs) Handles btnDepartment.ItemClick
        For Each f As Form In Me.MdiChildren
            f.Close()
        Next
        lblTitle.Text = "ROOM/WARD"
        frmDeparment.WindowState = FormWindowState.Maximized
        frmDeparment.MdiParent = Me
        frmDeparment.Show()
    End Sub

    Private Sub btnCompany_ItemClick(ByVal sender As System.Object, ByVal e As DevExpress.XtraBars.ItemClickEventArgs) Handles btnCompany.ItemClick
        frmCompanyProfile.ShowDialog()
    End Sub

    Private Sub btnDefaultPrinter_ItemClick(ByVal sender As System.Object, ByVal e As DevExpress.XtraBars.ItemClickEventArgs) Handles btnDefaultPrinter.ItemClick
        frmDefaultPrinter.ShowDialog()
    End Sub

    Private Sub btnChemRange_ItemClick(ByVal sender As System.Object, ByVal e As DevExpress.XtraBars.ItemClickEventArgs) Handles btnChemRange.ItemClick
        For Each f As Form In Me.MdiChildren
            f.Close()
        Next
        lblTitle.Text = "REFERENCE RANGE"
        frmRangeList.WindowState = FormWindowState.Maximized
        frmRangeList.MdiParent = Me
        frmRangeList.Show()
    End Sub

    Private Sub btnBackup_ItemClick(ByVal sender As System.Object, ByVal e As DevExpress.XtraBars.ItemClickEventArgs) Handles btnBackup.ItemClick
        frmDBBackup.ShowDialog()
    End Sub

    Private Sub btnUnit_ItemClick(ByVal sender As System.Object, ByVal e As DevExpress.XtraBars.ItemClickEventArgs) Handles btnUnit.ItemClick
        For Each f As Form In Me.MdiChildren
            f.Close()
        Next
        lblTitle.Text = "UNIT"

        frmUnit.WindowState = FormWindowState.Maximized
        frmUnit.MdiParent = Me
        frmUnit.Show()
    End Sub

    Private Sub btnTestName_ItemClick(ByVal sender As System.Object, ByVal e As DevExpress.XtraBars.ItemClickEventArgs)
        frmTestName.WindowState = FormWindowState.Maximized
        frmTestName.MdiParent = Me
        frmTestName.Show()
    End Sub

    Private Sub btnSpecimen_ItemClick(ByVal sender As System.Object, ByVal e As DevExpress.XtraBars.ItemClickEventArgs) Handles btnSpecimen.ItemClick
        For Each f As Form In Me.MdiChildren
            f.Close()
        Next
        lblTitle.Text = "SPECIMEN MAPPING"

        frmSpecimenMapping.WindowState = FormWindowState.Maximized
        frmSpecimenMapping.MdiParent = Me
        frmSpecimenMapping.Show()
    End Sub

    Private Sub btnChange_ItemClick(ByVal sender As System.Object, ByVal e As DevExpress.XtraBars.ItemClickEventArgs) Handles btnChange.ItemClick
        frmChangeUser.ShowDialog()
    End Sub

    Private Sub btnPatient_ItemClick(ByVal sender As System.Object, ByVal e As EventArgs) Handles btnPatient.Click
        For Each f As Form In Me.MdiChildren
            f.Close()
        Next
        lblTitle.Text = "FRONTDESK WINDOW"
        frmFrontDesk.WindowState = FormWindowState.Maximized
        frmFrontDesk.MdiParent = Me
        frmFrontDesk.Show()
    End Sub

    Private Sub btnQC_ItemClick(ByVal sender As System.Object, ByVal e As EventArgs) Handles btnQC.Click
        For Each f As Form In Me.MdiChildren
            f.Close()
        Next
        lblTitle.Text = "QUALITY CONTROL"
        frmQC.WindowState = FormWindowState.Maximized
        frmQC.MdiParent = Me
        frmQC.Show()
    End Sub

    Private Sub btnDoctor_ItemClick(ByVal sender As System.Object, ByVal e As DevExpress.XtraBars.ItemClickEventArgs) Handles btnDoctors.ItemClick
        For Each f As Form In Me.MdiChildren
            f.Close()
        Next
        lblTitle.Text = "REQUESTOR LIST"

        frmDoctor.WindowState = FormWindowState.Maximized
        frmDoctor.MdiParent = Me
        frmDoctor.Show()
    End Sub

    Private Sub btnDefaultSpecimen_ItemClick(ByVal sender As System.Object, ByVal e As DevExpress.XtraBars.ItemClickEventArgs) Handles btnDefaultSpecimen.ItemClick
        For Each f As Form In Me.MdiChildren
            f.Close()
        Next
        lblTitle.Text = "SPECIMEN LIST"

        frmSpecimen.WindowState = FormWindowState.Maximized
        frmSpecimen.MdiParent = Me
        frmSpecimen.Show()
    End Sub

    Private Sub btnCencus_ItemClick(ByVal sender As System.Object, ByVal e As EventArgs) Handles btnCensus.Click
        For Each f As Form In Me.MdiChildren
            f.Close()
        Next
        lblTitle.Text = "CENSUS REPORT"
        frmCensus.WindowState = FormWindowState.Maximized
        frmCensus.MdiParent = Me
        frmCensus.Show()
    End Sub

    Private Sub btnUsers_ItemClick(sender As Object, e As DevExpress.XtraBars.ItemClickEventArgs) Handles btnUsers.ItemClick
        For Each f As Form In Me.MdiChildren
            f.Close()
        Next
        lblTitle.Text = "USER WINDOW"
        frmUser.WindowState = FormWindowState.Maximized
        frmUser.MdiParent = Me
        frmUser.Show()
    End Sub

    Private Sub btnTest_ItemClick(ByVal sender As System.Object, ByVal e As DevExpress.XtraBars.ItemClickEventArgs) Handles btnTest.ItemClick
        frmSystemConfig.ShowDialog()
    End Sub

    Private Sub btnMachines_ItemClick(ByVal sender As System.Object, ByVal e As DevExpress.XtraBars.ItemClickEventArgs) Handles btnMachines.ItemClick
        frmDBUtils.ShowDialog()
    End Sub

    Private Sub btnRestore_ItemClick(ByVal sender As System.Object, ByVal e As DevExpress.XtraBars.ItemClickEventArgs) Handles btnRestore.ItemClick
        frmDBRestore.ShowDialog()
    End Sub

    Private Sub btnTruncate_ItemClick(ByVal sender As System.Object, ByVal e As DevExpress.XtraBars.ItemClickEventArgs) Handles btnTruncate.ItemClick
        frmDBTables.WindowState = FormWindowState.Maximized
        frmDBTables.MdiParent = Me
        frmDBTables.Show()
    End Sub

    Private Sub btnEditAccount_ItemClick(ByVal sender As System.Object, ByVal e As DevExpress.XtraBars.ItemClickEventArgs) Handles btnEditAccount.ItemClick
        frmAccountE.ShowDialog()
    End Sub

    Private Sub btnShift_ItemClick(ByVal sender As System.Object, ByVal e As DevExpress.XtraBars.ItemClickEventArgs) Handles btnShift.ItemClick
        frmShift.WindowState = FormWindowState.Maximized
        frmShift.MdiParent = Me
        frmShift.Show()
    End Sub

    Private Sub btnRealTime_ItemClick(sender As Object, e As EventArgs) Handles btnRealtime.Click
        frmTaT.WindowState = FormWindowState.Maximized
        frmTaT.MdiParent = Me
        frmTaT.Show()
    End Sub

    Private Sub btnSpecimenLogs_ItemClick(sender As Object, e As EventArgs) Handles btnSpecimenLog.Click
        frmTaTRealtime.WindowState = FormWindowState.Maximized
        frmTaTRealtime.MdiParent = Me
        frmTaTRealtime.Show()
    End Sub

    Private Sub btnCompleteTaT_ItemClick(sender As Object, e As EventArgs) Handles btnCompleteTat.Click
        For Each f As Form In Me.MdiChildren
            f.Close()
        Next
        lblTitle.Text = "TURNAROUND TIME REPORT"

        frmTatComplete.WindowState = FormWindowState.Maximized
        frmTatComplete.MdiParent = Me
        frmTatComplete.Show()
    End Sub

    Private Sub aceMicrobiology_Click(sender As Object, e As EventArgs) Handles aceMicrobiology.Click
        For Each f As Form In Me.MdiChildren
            f.Close()
        Next
        lblTitle.Text = "MICROBIOLOGY WORKLIST"

        frmMicroWorklist.WindowState = FormWindowState.Maximized
        frmMicroWorklist.MdiParent = Me
        frmMicroWorklist.Show()
    End Sub

    Private Sub btnPatientOrder_Click(sender As Object, e As EventArgs) Handles btnPatientOrder.Click
        For Each f As Form In Me.MdiChildren
            f.Close()
        Next
        lblTitle.Text = "ORDER MANAGEMENT"

        frmPatientOrder.WindowState = FormWindowState.Maximized
        frmPatientOrder.MdiParent = Me
        frmPatientOrder.Show()
    End Sub

    Private Sub btnWorkSheet_Click(sender As Object, e As EventArgs) Handles btnWorkSheet.Click
        For Each f As Form In Me.MdiChildren
            f.Close()
        Next
        lblTitle.Text = "WORKSHEET"

        frmWorkSheet.WindowState = FormWindowState.Maximized
        frmWorkSheet.MdiParent = Me
        frmWorkSheet.Show()
    End Sub

    Private Sub btnMngUser_ItemClick(sender As Object, e As DevExpress.XtraBars.ItemClickEventArgs) Handles btnMngUser.ItemClick
        For Each f As Form In Me.MdiChildren
            f.Close()
        Next
        lblTitle.Text = "USER ACCOUNTS"
        frmAccount.WindowState = FormWindowState.Maximized
        frmAccount.MdiParent = Me
        frmAccount.Show()
    End Sub

    Private Sub accPhlebotomy_Click(sender As Object, e As EventArgs) Handles accPhlebotomy.Click
        For Each f As Form In Me.MdiChildren
            f.Close()
        Next
        lblTitle.Text = "PHLEBOTOMY WINDOW"
        frmPhlebotomy.WindowState = FormWindowState.Maximized
        frmPhlebotomy.MdiParent = Me
        frmPhlebotomy.Show()
    End Sub

    Private Sub btnSPecimenLogs_ItemClick(sender As Object, e As DevExpress.XtraBars.ItemClickEventArgs) Handles btnSpecimenLogs.ItemClick
        For Each f As Form In Me.MdiChildren
            f.Close()
        Next
        lblTitle.Text = "SPECIMEN LOGS"
        frmSpecimenLog.WindowState = FormWindowState.Maximized
        frmSpecimenLog.MdiParent = Me
        frmSpecimenLog.Show()
    End Sub

    Private Sub btnActivityLogs_ItemClick(sender As Object, e As DevExpress.XtraBars.ItemClickEventArgs) Handles btnActivityLog.ItemClick
        For Each f As Form In Me.MdiChildren
            f.Close()
        Next
        lblTitle.Text = "ACTIVITY LOGS"
        frmActivityLog.WindowState = FormWindowState.Maximized
        frmActivityLog.MdiParent = Me
        frmActivityLog.Show()
    End Sub

    Private Sub acciHOMIS_ItemClick(ByVal sender As System.Object, ByVal e As EventArgs) Handles acciHOMIS.Click
        For Each f As Form In Me.MdiChildren
            f.Close()
        Next
        lblTitle.Text = "iHOMIS INTEGRATION"
        frmiHOMIS.WindowState = FormWindowState.Maximized
        frmiHOMIS.MdiParent = Me
        frmiHOMIS.Show()
    End Sub

    Private Sub aceMolecular_Click(sender As Object, e As EventArgs) Handles aceMolecular.Click
        For Each f As Form In Me.MdiChildren
            f.Close()
        Next
        lblTitle.Text = "MOLECULAR BIOLOGY WORKLIST"
        frmMolWorklist.WindowState = FormWindowState.Maximized
        frmMolWorklist.MdiParent = Me
        frmMolWorklist.Show()
    End Sub

    Private Sub btnLogout_Click(sender As Object, e As EventArgs) Handles btnLogout.Click
        frmChangeUser.ShowDialog()
    End Sub
#End Region

#Region "Get New and Pending Orders"
    Private Sub Timer_Tick_New(sender As Object, e As EventArgs) Handles TimerNew.Tick
        Try
            Dim Date1 As String = ""

            Connect()
            rs.Connection = conn
            rs.CommandType = CommandType.Text
            rs.CommandText = "SELECT COUNT(*) AS `No` FROM tmpworklist WHERE (status = 'Processing' OR status = 'Result Received' OR status = 'Checked-In') AND `date` = DATE_FORMAT(NOW(), '%Y-%m-%d')"
            reader = rs.ExecuteReader
            reader.Read()
            If reader.HasRows Then
                lblNew.Text = "New Order (" & reader(0).ToString & ")"
                Disconnect()
            End If
            Disconnect()

            Connect()
            rs.Connection = conn
            rs.CommandType = CommandType.Text
            rs.CommandText = "SELECT `date` FROM tmpworklist WHERE status = 'Ordered' ORDER BY `date` desc Limit 1"
            reader = rs.ExecuteReader
            reader.Read()
            If reader.HasRows Then
                Date1 = Format(CDate(reader(0).ToString), "yyyy-MM-dd")
                Disconnect()

                If Date1.Contains(Now.ToString("yyyy-MM-dd")) = True Then
                    If lblNew.Visible = False Then
                        lblNew.Visible = True
                    ElseIf lblNew.Visible = True Then
                        lblNew.Visible = False
                    End If
                Else
                    lblNew.Visible = False
                End If
            End If
            Disconnect()
        Catch ex As Exception
            TimerNew.Stop()
            Disconnect()
            MessageBox.Show(ex.Message, "System Error", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            Exit Sub
        End Try
    End Sub

    Private Sub Timer_Tick_Pending(sender As Object, e As EventArgs) Handles TimerPending.Tick
        Try
            Dim countPending As Integer = 0

            Connect()
            rs.Connection = conn
            rs.CommandType = CommandType.Text
            rs.CommandText = "SELECT COUNT(*) AS `No` FROM tmpworklist WHERE (status = 'Processing' OR status = 'Result Received' OR status = 'Checked-In') AND `date` < DATE_FORMAT(NOW(), '%Y-%m-%d')"
            'Try
            reader = rs.ExecuteReader
            reader.Read()
            If reader.HasRows Then
                lblPending.Text = "Pending Order (" & reader(0).ToString & ")"
                Disconnect()
                countPending = countPending + 1
            Else
                Disconnect()
            End If

            If countPending > 0 Then
                If lblPending.Visible = False Then
                    lblPending.Visible = True
                ElseIf lblPending.Visible = True Then
                    lblPending.Visible = False
                End If
            Else
                lblPending.Visible = False
            End If
            Disconnect()
        Catch ex As Exception
            TimerPending.Stop()
            Disconnect()
            MessageBox.Show(ex.Message, "System Error", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            Exit Sub
        End Try
    End Sub
#End Region

#Region "Disable Controls"
    Public Sub DisableControls()
        For x = 0 To Me.AccordionControl.Elements.Count - 1
            Me.AccordionControl.Elements.Item(x).Expanded = False
        Next

        Connect()
        rs.Connection = conn
        rs.CommandText = "SELECT * FROM `user_access` WHERE `user_id` = '" & CurrEmail & "'"
        reader = rs.ExecuteReader
        While reader.Read
            If reader(3).ToString = 0 Then
                For Each ctrl In Me.AccordionControl.GetElements
                    If ctrl.Text = reader(2).ToString Then
                        ctrl.Enabled = False
                    End If

                    If reader(2).ToString = "System Settings" Then
                        Me.mnuSetting.Visible = False
                        Me.mnuSysSettings.Visible = False
                        Me.mnuHelp.Visible = False
                    End If
                Next
            ElseIf reader(3).ToString = 1 Then
                For Each ctrl In Me.AccordionControl.GetElements
                    If ctrl.Text = reader(2).ToString Then
                        ctrl.Enabled = True
                    End If

                    If reader(2).ToString = "System Settings" Then
                        Me.mnuSetting.Visible = True
                        Me.mnuSysSettings.Visible = True
                        Me.mnuHelp.Visible = True
                    End If
                Next
            End If
        End While

        If My.Settings.SQLConnection = True Then
            acciHOMIS.Visible = True
        Else
            acciHOMIS.Visible = False
        End If
        Disconnect()

    End Sub

    Private Sub lblNew_DoubleClick(sender As Object, e As EventArgs) Handles lblNew.DoubleClick, lblPending.DoubleClick
        For Each f As Form In Me.MdiChildren
            f.Close()
        Next
        lblTitle.Text = "PHLEBOTOMY WINDOW"
        frmPhlebotomy.WindowState = FormWindowState.Maximized
        frmPhlebotomy.MdiParent = Me
        frmPhlebotomy.Show()
    End Sub

    Private Sub btnEmail_ItemClick(sender As Object, e As DevExpress.XtraBars.ItemClickEventArgs) Handles btnEmail.ItemClick
        frmEmailContent.ShowDialog()
    End Sub

    Private Sub btnDelta_Click(sender As Object, e As EventArgs) Handles btnDelta.Click
        For Each f As Form In Me.MdiChildren
            f.Close()
        Next
        lblTitle.Text = "DELTA CHECK"
        frmDeltaCheck.WindowState = FormWindowState.Maximized
        frmDeltaCheck.MdiParent = Me
        frmDeltaCheck.Show()
    End Sub

#End Region

End Class
