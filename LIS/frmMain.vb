Imports System.IO
Imports System.IO.Ports
Imports System.Threading
Imports System.Threading.Thread
Imports System.Globalization
Imports System.Text.RegularExpressions

Public Class frmMain

    Private Sub frmMainSemi_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        If (MessageBox.Show("Are you sure you want to close the application?", "Exit", MessageBoxButtons.YesNo, MessageBoxIcon.Question) = DialogResult.Yes) Then
            End
        Else
            e.Cancel = True
        End If
    End Sub

    Private Sub frmMain_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        frmLogin.ShowDialog()
        RetriveData()
    End Sub

#Region "CheckConnection"
    Private Sub tmCheckCon_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tmCheckCon.Tick

        If CurrType = "Phlebotomy" Or CurrType = "Receptionist" Then
            RBPFrontDesk.Visible = False
            RBPLaboratory.Visible = False
            RBPReport.Visible = False
            RibbonPage2.Visible = True
            RibbonPage3.Visible = False
            RBPPhlembo.Visible = True
            btnAccounts.Enabled = False
        ElseIf CurrType = "Laboratory" Or CurrType = "MedTech" Then
            RBPFrontDesk.Visible = False
            RBPReport.Visible = True
            RBPLaboratory.Visible = True
            RibbonPage2.Visible = False
            RibbonPage3.Visible = False
            RBPPhlembo.Visible = False
            btnAccounts.Enabled = False
        ElseIf CurrType = "Administrator" Then
            RBPFrontDesk.Visible = False
            RBPReport.Visible = True
            RBPLaboratory.Visible = True
            RibbonPage2.Visible = True
            RibbonPage3.Visible = False
            RBPPhlembo.Visible = True
            btnAccounts.Enabled = True
        ElseIf CurrType = "Super Admin" Then
            RBPFrontDesk.Visible = False
            RBPReport.Visible = True
            RBPLaboratory.Visible = True
            RibbonPage2.Visible = True
            RibbonPage3.Visible = True
            RBPPhlembo.Visible = True
            btnAccounts.Enabled = True
        Else
            RBPFrontDesk.Visible = True
            RBPReport.Visible = False
            RibbonPage2.Visible = False
            RibbonPage3.Visible = False
            RBPPhlembo.Visible = False
            btnAccounts.Enabled = False
            btnAccounts.Enabled = False
        End If
    End Sub
#End Region

#Region "RibbonBar Menu Click Events"
    Private Sub btnSettings_ItemClick(ByVal sender As System.Object, ByVal e As DevExpress.XtraBars.ItemClickEventArgs) Handles btnSettings.ItemClick
        frmControlSetting.WindowState = FormWindowState.Maximized
        frmControlSetting.MdiParent = Me
        frmControlSetting.Show()
    End Sub

    Private Sub btnMedtech_ItemClick(ByVal sender As System.Object, ByVal e As DevExpress.XtraBars.ItemClickEventArgs) Handles btnMedTech.ItemClick
        frmMedTech.WindowState = FormWindowState.Maximized
        frmMedTech.MdiParent = Me
        frmMedTech.Show()
    End Sub

    Private Sub bvClose_ItemClick(ByVal sender As System.Object, ByVal e As DevExpress.XtraBars.Ribbon.BackstageViewItemEventArgs) Handles bvClose.ItemClick
        Me.Close()
    End Sub

    Private Sub bntTestName_ItemClick(ByVal sender As System.Object, ByVal e As DevExpress.XtraBars.ItemClickEventArgs) Handles bntTestName.ItemClick
        frmTest.WindowState = FormWindowState.Maximized
        frmTest.MdiParent = Me
        frmTest.Show()
    End Sub

    Private Sub btnPatientInfo_ItemClick(ByVal sender As System.Object, ByVal e As DevExpress.XtraBars.ItemClickEventArgs) Handles btnPatientInformation.ItemClick, btnPatientInfos.ItemClick
        frmPatient.WindowState = FormWindowState.Maximized
        frmPatient.MdiParent = Me
        frmPatient.Show()
    End Sub

    Private Sub btnPatientOrder_ItemClick(ByVal sender As System.Object, ByVal e As DevExpress.XtraBars.ItemClickEventArgs) Handles btnPatientNewOrder.ItemClick
        frmPatientOrder.WindowState = FormWindowState.Maximized
        frmPatientOrder.MdiParent = Me
        frmPatientOrder.Show()
    End Sub

    Private Sub btnNewOrder_ItemClick(ByVal sender As System.Object, ByVal e As DevExpress.XtraBars.ItemClickEventArgs) Handles btnNewOrder.ItemClick
        frmHemaWorklist.WindowState = FormWindowState.Maximized
        frmHemaWorklist.MdiParent = Me
        frmHemaWorklist.Show()
    End Sub

    Private Sub btnTestOrder_ItemClick(ByVal sender As System.Object, ByVal e As DevExpress.XtraBars.ItemClickEventArgs) Handles btnTestOrder.ItemClick
        frmChemArchive.WindowState = FormWindowState.Maximized
        frmChemArchive.MdiParent = Me
        frmChemArchive.Show()
    End Sub

    Private Sub btnWorkSheet_ItemClick(ByVal sender As System.Object, ByVal e As DevExpress.XtraBars.ItemClickEventArgs) Handles btnWorkSheet.ItemClick
        frmWorkSheet.WindowState = FormWindowState.Maximized
        frmWorkSheet.MdiParent = Me
        frmWorkSheet.Show()
    End Sub

    Private Sub btnDoctors_ItemClick(ByVal sender As System.Object, ByVal e As DevExpress.XtraBars.ItemClickEventArgs) Handles btnDoctors.ItemClick
        frmDoctor.WindowState = FormWindowState.Maximized
        frmDoctor.MdiParent = Me
        frmDoctor.Show()
    End Sub

    Private Sub btnAccounts_ItemClick(ByVal sender As System.Object, ByVal e As DevExpress.XtraBars.ItemClickEventArgs) Handles btnAccounts.ItemClick
        frmAccount.WindowState = FormWindowState.Maximized
        frmAccount.MdiParent = Me
        frmAccount.Show()
    End Sub

    Private Sub btnPathologist_ItemClick(ByVal sender As System.Object, ByVal e As DevExpress.XtraBars.ItemClickEventArgs) Handles btnPathologist.ItemClick
        frmPathologist.WindowState = FormWindowState.Maximized
        frmPathologist.MdiParent = Me
        frmPathologist.Show()
    End Sub

    Private Sub btnDepartment_ItemClick(ByVal sender As System.Object, ByVal e As DevExpress.XtraBars.ItemClickEventArgs) Handles btnDepartment.ItemClick
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

    Private Sub btnChart_ItemClick(ByVal sender As System.Object, ByVal e As DevExpress.XtraBars.ItemClickEventArgs) Handles btnChart.ItemClick
        'If TestTypes = "Hematology" Or TestTypes = "HEMATOLOGY" Then
        'RPTChartHema.ShowDialog()
        'ElseIf TestTypes = "Chemistry" Or TestTypes = "CHEMISTRY" Then
        'RPTChartChem.ShowDialog()
        'End If
    End Sub

    Private Sub btnChemRange_ItemClick(ByVal sender As System.Object, ByVal e As DevExpress.XtraBars.ItemClickEventArgs) Handles btnChemRange.ItemClick
        frmRangeList.WindowState = FormWindowState.Maximized
        frmRangeList.MdiParent = Me
        frmRangeList.Show()
    End Sub

    Private Sub btnBarcode_ItemClick(ByVal sender As System.Object, ByVal e As DevExpress.XtraBars.ItemClickEventArgs) Handles btnBarcode.ItemClick
        frmBarcode.ShowDialog()
    End Sub

    Private Sub btnBackup_ItemClick(ByVal sender As System.Object, ByVal e As DevExpress.XtraBars.ItemClickEventArgs) Handles btnBackup.ItemClick
        frmDBBackup.ShowDialog()
    End Sub

    Private Sub btnUnit_ItemClick(ByVal sender As System.Object, ByVal e As DevExpress.XtraBars.ItemClickEventArgs) Handles btnUnit.ItemClick
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
        frmSpecimenMapping.WindowState = FormWindowState.Maximized
        frmSpecimenMapping.MdiParent = Me
        frmSpecimenMapping.Show()
    End Sub

    Private Sub btnSPrinterSetting_ItemClick(ByVal sender As System.Object, ByVal e As DevExpress.XtraBars.ItemClickEventArgs) Handles btnSPrinterSetting.ItemClick
        frmDefaultPrinter.ShowDialog()
    End Sub

    Private Sub btnSBackup_ItemClick(ByVal sender As System.Object, ByVal e As DevExpress.XtraBars.ItemClickEventArgs) Handles btnSBackup.ItemClick
        frmDBBackup.ShowDialog()
    End Sub

    Private Sub btnWebsite_ItemClick(ByVal sender As System.Object, ByVal e As DevExpress.XtraBars.ItemClickEventArgs) Handles btnWebsite.ItemClick
        frmWebsite.WindowState = FormWindowState.Maximized
        frmWebsite.MdiParent = Me
        frmWebsite.Show()
    End Sub

    Private Sub btnChange_ItemClick(ByVal sender As System.Object, ByVal e As DevExpress.XtraBars.ItemClickEventArgs) Handles btnChange.ItemClick
        frmChangeUser.ShowDialog()
    End Sub

    Private Sub btnPatient_ItemClick(ByVal sender As System.Object, ByVal e As DevExpress.XtraBars.ItemClickEventArgs) Handles btnPatient.ItemClick
        frmPatient.MdiParent = Me
        frmPatient.WindowState = FormWindowState.Maximized
        frmPatient.Show()
    End Sub

    Private Sub btnQC_ItemClick(ByVal sender As System.Object, ByVal e As DevExpress.XtraBars.ItemClickEventArgs) Handles btnQC.ItemClick
        frmQC.WindowState = FormWindowState.Maximized
        frmQC.MdiParent = Me
        frmQC.Show()
    End Sub

    Private Sub btnDoctor_ItemClick(ByVal sender As System.Object, ByVal e As DevExpress.XtraBars.ItemClickEventArgs) Handles btnDoctor.ItemClick
        frmDoctor.WindowState = FormWindowState.Maximized
        frmDoctor.MdiParent = Me
        frmDoctor.Show()
    End Sub

    Private Sub btnDefaultSpecimen_ItemClick(ByVal sender As System.Object, ByVal e As DevExpress.XtraBars.ItemClickEventArgs) Handles btnDefaultSpecimen.ItemClick
        frmSpecimen.WindowState = FormWindowState.Maximized
        frmSpecimen.MdiParent = Me
        frmSpecimen.Show()
    End Sub

    Private Sub btnDefaultOrder_ItemClick(ByVal sender As System.Object, ByVal e As DevExpress.XtraBars.ItemClickEventArgs) Handles btnDefaultOrder.ItemClick
        frmOrder.WindowState = FormWindowState.Maximized
        frmOrder.MdiParent = Me
        frmOrder.Show()
    End Sub

    Private Sub btnDelta_ItemClick(ByVal sender As System.Object, ByVal e As DevExpress.XtraBars.ItemClickEventArgs) Handles btnDelta.ItemClick
        frmDeltaCheck.WindowState = FormWindowState.Maximized
        frmDeltaCheck.MdiParent = Me
        frmDeltaCheck.Show()
    End Sub

    Private Sub btnCencus_ItemClick(ByVal sender As System.Object, ByVal e As DevExpress.XtraBars.ItemClickEventArgs) Handles btnCencus.ItemClick
        frmCensus.WindowState = FormWindowState.Maximized
        frmCensus.MdiParent = Me
        frmCensus.Show()
    End Sub

    'Private Sub btnManual_ItemClick(ByVal sender As System.Object, ByVal e As DevExpress.XtraBars.ItemClickEventArgs) Handles btnManual.ItemClick
    '    frmResultsNewOrder.Text = "Hematology Manual Result"
    '    frmResultsNewOrder.Section = "Hematology"
    '    frmResultsNewOrder.ShowDialog()
    'End Sub

    Private Sub btnManualHematology_ItemClick(ByVal sender As System.Object, ByVal e As DevExpress.XtraBars.ItemClickEventArgs) Handles btnManualHematology.ItemClick
        frmResultsNewOrder.Text = "Hematology Manual Result"
        frmResultsNewOrder.Section = "Hematology"
        frmResultsNewOrder.ShowDialog()
    End Sub

    Private Sub btnManualChemistry_ItemClick(ByVal sender As System.Object, ByVal e As DevExpress.XtraBars.ItemClickEventArgs) Handles btnManualChemistry.ItemClick
        frmResultsNewOrder.Text = "Chemistry Manual Result"
        frmResultsNewOrder.Section = "Chemistry"
        frmResultsNewOrder.ShowDialog()
    End Sub

    Private Sub btnUrin_ItemClick(ByVal sender As System.Object, ByVal e As DevExpress.XtraBars.ItemClickEventArgs) Handles btnUrin.ItemClick
        frmResultsNewOrder.Text = "Urinalysis Manual Result"
        frmResultsNewOrder.Section = "Urinalysis"
        frmResultsNewOrder.ShowDialog()
    End Sub

    Private Sub btnStool_ItemClick(ByVal sender As System.Object, ByVal e As DevExpress.XtraBars.ItemClickEventArgs) Handles btnStool.ItemClick
        frmResultsNewOrder.Text = "Stool Manual Result"
        frmResultsNewOrder.Section = "Fecalysis"
        frmResultsNewOrder.ShowDialog()
    End Sub

    Private Sub btnOther_ItemClick(ByVal sender As System.Object, ByVal e As DevExpress.XtraBars.ItemClickEventArgs) Handles btnOthers.ItemClick
        frmResultsNewOrder.Text = "Others Manual Result"
        frmResultsNewOrder.Section = "Others"
        frmResultsNewOrder.ShowDialog()
    End Sub

    Private Sub btnTest_ItemClick(ByVal sender As System.Object, ByVal e As DevExpress.XtraBars.ItemClickEventArgs) Handles btnTest.ItemClick
        frmSystemConfig.ShowDialog()
    End Sub

    Private Sub btnMachines_ItemClick(ByVal sender As System.Object, ByVal e As DevExpress.XtraBars.ItemClickEventArgs) Handles btnMachines.ItemClick
        frmMachineList.WindowState = FormWindowState.Maximized
        frmMachineList.MdiParent = Me
        frmMachineList.Show()
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

    Private Sub btnFrontDesk_ItemClick(ByVal sender As System.Object, ByVal e As DevExpress.XtraBars.ItemClickEventArgs) Handles btnFrontDesk.ItemClick
        frmOrders.WindowState = FormWindowState.Maximized
        frmOrders.MdiParent = Me
        frmOrders.Show()
    End Sub

    Private Sub btnShift_ItemClick(ByVal sender As System.Object, ByVal e As DevExpress.XtraBars.ItemClickEventArgs) Handles btnShift.ItemClick
        frmShift.WindowState = FormWindowState.Maximized
        frmShift.MdiParent = Me
        frmShift.Show()
    End Sub

    Private Sub btnRealTime_ItemClick(sender As Object, e As DevExpress.XtraBars.ItemClickEventArgs) Handles btnRealTime.ItemClick
        frmTaT.WindowState = FormWindowState.Maximized
        frmTaT.MdiParent = Me
        frmTaT.Show()
    End Sub

    Private Sub btnSpecimenLogs_ItemClick(sender As Object, e As DevExpress.XtraBars.ItemClickEventArgs) Handles btnSpecimenLogs.ItemClick
        frmTaTRealtime.WindowState = FormWindowState.Maximized
        frmTaTRealtime.MdiParent = Me
        frmTaTRealtime.Show()
    End Sub

    Private Sub btnCompleteTaT_ItemClick(sender As Object, e As DevExpress.XtraBars.ItemClickEventArgs) Handles btnCompleteTaT.ItemClick
        frmTatComplete.WindowState = FormWindowState.Maximized
        frmTatComplete.MdiParent = Me
        frmTatComplete.Show()
    End Sub
#End Region


    Private Sub RetriveData()
        Connect()
        rs.Connection = conn
        rs.CommandText = "SELECT * FROM company_profile"
        reader = rs.ExecuteReader
        reader.Read()
        If reader.HasRows Then
            txtName.Text = reader(1).ToString
            txtAddress.Text = reader(2).ToString
            txtTel.Text = reader(3).ToString
            txtFax.Text = reader(4).ToString

            If Not reader(5).ToString = "" Then
                Dim imageData As Byte() = DirectCast(reader(5), Byte())
                If Not imageData Is Nothing Then
                    Using ms As New MemoryStream(imageData, 0, imageData.Length)
                        ms.Write(imageData, 0, imageData.Length)
                        picLogo_primary.Image = Image.FromStream(ms, True)
                    End Using
                End If
            End If

            If Not reader(6).ToString = "" Then
                Dim imageData_sec As Byte() = DirectCast(reader(6), Byte())
                If Not imageData_sec Is Nothing Then
                    Using ms As New MemoryStream(imageData_sec, 0, imageData_sec.Length)
                        ms.Write(imageData_sec, 0, imageData_sec.Length)
                        picLogo_secondary.Image = Image.FromStream(ms, True)
                    End Using
                End If
            End If
        End If
        Disconnect()
    End Sub

    Private Sub acChem_Click(sender As Object, e As EventArgs) Handles acChem.Click
        frmChemArchive.WindowState = FormWindowState.Maximized
        frmChemArchive.MdiParent = Me
        frmChemArchive.Show()
    End Sub

    Private Sub AccordionControlElement14_Click(sender As Object, e As EventArgs) Handles AccordionControlElement14.Click
        frmChemArchive.WindowState = FormWindowState.Maximized
        frmChemArchive.MdiParent = Me
        frmChemArchive.Show()
    End Sub
End Class