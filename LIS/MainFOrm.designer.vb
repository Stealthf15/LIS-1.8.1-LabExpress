<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class MainFOrm
    Inherits DevExpress.XtraEditors.XtraForm

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()>
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()>
    Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(MainFOrm))
        Me.SidePanel1 = New DevExpress.XtraEditors.SidePanel()
        Me.AccordionControl = New DevExpress.XtraBars.Navigation.AccordionControl()
        Me.AccordionControlSeparator6 = New DevExpress.XtraBars.Navigation.AccordionControlSeparator()
        Me.acciHOMIS = New DevExpress.XtraBars.Navigation.AccordionControlElement()
        Me.btnPatient = New DevExpress.XtraBars.Navigation.AccordionControlElement()
        Me.accPhlebotomy = New DevExpress.XtraBars.Navigation.AccordionControlElement()
        Me.AccordionControlSeparator1 = New DevExpress.XtraBars.Navigation.AccordionControlSeparator()
        Me.btnPatientOrder = New DevExpress.XtraBars.Navigation.AccordionControlElement()
        Me.AccordionControlSeparator2 = New DevExpress.XtraBars.Navigation.AccordionControlSeparator()
        Me.aceSection = New DevExpress.XtraBars.Navigation.AccordionControlElement()
        Me.aceHema = New DevExpress.XtraBars.Navigation.AccordionControlElement()
        Me.aceChem = New DevExpress.XtraBars.Navigation.AccordionControlElement()
        Me.aceSerology = New DevExpress.XtraBars.Navigation.AccordionControlElement()
        Me.aceImmuno = New DevExpress.XtraBars.Navigation.AccordionControlElement()
        Me.aceMicrobiology = New DevExpress.XtraBars.Navigation.AccordionControlElement()
        Me.aceHistopath = New DevExpress.XtraBars.Navigation.AccordionControlElement()
        Me.aceUrin = New DevExpress.XtraBars.Navigation.AccordionControlElement()
        Me.aceFecal = New DevExpress.XtraBars.Navigation.AccordionControlElement()
        Me.aceMolecular = New DevExpress.XtraBars.Navigation.AccordionControlElement()
        Me.AccordionControlSeparator3 = New DevExpress.XtraBars.Navigation.AccordionControlSeparator()
        Me.btnDelta = New DevExpress.XtraBars.Navigation.AccordionControlElement()
        Me.AccordionControlSeparator5 = New DevExpress.XtraBars.Navigation.AccordionControlSeparator()
        Me.aceReports = New DevExpress.XtraBars.Navigation.AccordionControlElement()
        Me.btnQC = New DevExpress.XtraBars.Navigation.AccordionControlElement()
        Me.btnWorkSheet = New DevExpress.XtraBars.Navigation.AccordionControlElement()
        Me.btnCensus = New DevExpress.XtraBars.Navigation.AccordionControlElement()
        Me.btnCompleteTat = New DevExpress.XtraBars.Navigation.AccordionControlElement()
        Me.btnSpecimenLog = New DevExpress.XtraBars.Navigation.AccordionControlElement()
        Me.AccordionControlSeparator7 = New DevExpress.XtraBars.Navigation.AccordionControlSeparator()
        Me.btnRealtime = New DevExpress.XtraBars.Navigation.AccordionControlElement()
        Me.AccordionControlSeparator9 = New DevExpress.XtraBars.Navigation.AccordionControlSeparator()
        Me.accEmailPush = New DevExpress.XtraBars.Navigation.AccordionControlElement()
        Me.btnLogout = New DevExpress.XtraEditors.SimpleButton()
        Me.PanelControl2 = New DevExpress.XtraEditors.PanelControl()
        Me.PictureBox1 = New System.Windows.Forms.PictureBox()
        Me.SeparatorControl1 = New DevExpress.XtraEditors.SeparatorControl()
        Me.LabelControl1 = New DevExpress.XtraEditors.LabelControl()
        Me.PanelControl1 = New DevExpress.XtraEditors.PanelControl()
        Me.lblPending = New DevExpress.XtraEditors.LabelControl()
        Me.lblNew = New DevExpress.XtraEditors.LabelControl()
        Me.LabelControl4 = New DevExpress.XtraEditors.LabelControl()
        Me.LabelControl3 = New DevExpress.XtraEditors.LabelControl()
        Me.SeparatorControl2 = New DevExpress.XtraEditors.SeparatorControl()
        Me.lblTitle = New DevExpress.XtraEditors.LabelControl()
        Me.mnuHelp = New DevExpress.XtraEditors.DropDownButton()
        Me.pmLogs = New DevExpress.XtraBars.PopupMenu(Me.components)
        Me.btnSpecimenLogs = New DevExpress.XtraBars.BarLargeButtonItem()
        Me.btnActivityLog = New DevExpress.XtraBars.BarLargeButtonItem()
        Me.BarManager = New DevExpress.XtraBars.BarManager(Me.components)
        Me.barDockControlTop = New DevExpress.XtraBars.BarDockControl()
        Me.barDockControlBottom = New DevExpress.XtraBars.BarDockControl()
        Me.barDockControlLeft = New DevExpress.XtraBars.BarDockControl()
        Me.barDockControlRight = New DevExpress.XtraBars.BarDockControl()
        Me.btnSpecimen = New DevExpress.XtraBars.BarButtonItem()
        Me.btnDefaultSpecimen = New DevExpress.XtraBars.BarButtonItem()
        Me.btnTestGroup = New DevExpress.XtraBars.BarButtonItem()
        Me.btnChemRange = New DevExpress.XtraBars.BarButtonItem()
        Me.btnMedTech = New DevExpress.XtraBars.BarButtonItem()
        Me.btnPathologist = New DevExpress.XtraBars.BarButtonItem()
        Me.btnDoctors = New DevExpress.XtraBars.BarButtonItem()
        Me.btnDepartment = New DevExpress.XtraBars.BarButtonItem()
        Me.btnCompany = New DevExpress.XtraBars.BarButtonItem()
        Me.btnShift = New DevExpress.XtraBars.BarButtonItem()
        Me.btnTest = New DevExpress.XtraBars.BarButtonItem()
        Me.btnDefaultPrinter = New DevExpress.XtraBars.BarButtonItem()
        Me.btnPrinterSetting = New DevExpress.XtraBars.BarButtonItem()
        Me.btnBackup = New DevExpress.XtraBars.BarButtonItem()
        Me.btnRestore = New DevExpress.XtraBars.BarButtonItem()
        Me.btnTruncate = New DevExpress.XtraBars.BarButtonItem()
        Me.btnMachines = New DevExpress.XtraBars.BarButtonItem()
        Me.btnAccounts = New DevExpress.XtraBars.BarButtonItem()
        Me.btnEditAccount = New DevExpress.XtraBars.BarButtonItem()
        Me.btnChange = New DevExpress.XtraBars.BarButtonItem()
        Me.btnTestName = New DevExpress.XtraBars.BarButtonItem()
        Me.btnUnit = New DevExpress.XtraBars.BarButtonItem()
        Me.btnSettings = New DevExpress.XtraBars.BarButtonItem()
        Me.btnMngUser = New DevExpress.XtraBars.BarLargeButtonItem()
        Me.btnUsers = New DevExpress.XtraBars.BarLargeButtonItem()
        Me.btnEmail = New DevExpress.XtraBars.BarButtonItem()
        Me.lblTitleHome = New DevExpress.XtraEditors.LabelControl()
        Me.mnuSetting = New DevExpress.XtraEditors.DropDownButton()
        Me.pmSettings = New DevExpress.XtraBars.PopupMenu(Me.components)
        Me.mnuSysSettings = New DevExpress.XtraEditors.DropDownButton()
        Me.pmTools = New DevExpress.XtraBars.PopupMenu(Me.components)
        Me.mnuUser = New DevExpress.XtraEditors.DropDownButton()
        Me.pmUser = New DevExpress.XtraBars.PopupMenu(Me.components)
        Me.tmCheckCon = New System.Windows.Forms.Timer(Me.components)
        Me.PanelControl3 = New DevExpress.XtraEditors.PanelControl()
        Me.lblDateTime = New DevExpress.XtraEditors.LabelControl()
        Me.lblHospitalName = New DevExpress.XtraEditors.LabelControl()
        Me.TimerNew = New System.Windows.Forms.Timer(Me.components)
        Me.TimerPending = New System.Windows.Forms.Timer(Me.components)
        Me.SidePanel1.SuspendLayout()
        CType(Me.AccordionControl, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.PanelControl2, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.PanelControl2.SuspendLayout()
        CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.SeparatorControl1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.PanelControl1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.PanelControl1.SuspendLayout()
        CType(Me.SeparatorControl2, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.pmLogs, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.BarManager, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.pmSettings, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.pmTools, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.pmUser, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.PanelControl3, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.PanelControl3.SuspendLayout()
        Me.SuspendLayout()
        '
        'SidePanel1
        '
        Me.SidePanel1.BorderThickness = 0
        Me.SidePanel1.Controls.Add(Me.AccordionControl)
        Me.SidePanel1.Controls.Add(Me.btnLogout)
        Me.SidePanel1.Controls.Add(Me.PanelControl2)
        Me.SidePanel1.Dock = System.Windows.Forms.DockStyle.Left
        Me.SidePanel1.Location = New System.Drawing.Point(0, 0)
        Me.SidePanel1.Name = "SidePanel1"
        Me.SidePanel1.Size = New System.Drawing.Size(250, 625)
        Me.SidePanel1.TabIndex = 0
        Me.SidePanel1.Text = "SidePanel"
        '
        'AccordionControl
        '
        Me.AccordionControl.AllowGlyphSkinning = DevExpress.Utils.DefaultBoolean.[True]
        Me.AccordionControl.Appearance.AccordionControl.BackColor = System.Drawing.SystemColors.Control
        Me.AccordionControl.Appearance.AccordionControl.BackColor2 = System.Drawing.Color.FromArgb(CType(CType(27, Byte), Integer), CType(CType(41, Byte), Integer), CType(CType(62, Byte), Integer))
        Me.AccordionControl.Appearance.AccordionControl.BorderColor = System.Drawing.Color.FromArgb(CType(CType(27, Byte), Integer), CType(CType(41, Byte), Integer), CType(CType(62, Byte), Integer))
        Me.AccordionControl.Appearance.AccordionControl.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.AccordionControl.Appearance.AccordionControl.Options.UseBackColor = True
        Me.AccordionControl.Appearance.AccordionControl.Options.UseBorderColor = True
        Me.AccordionControl.Appearance.AccordionControl.Options.UseFont = True
        Me.AccordionControl.Appearance.Group.Hovered.BackColor = System.Drawing.Color.FromArgb(CType(CType(16, Byte), Integer), CType(CType(110, Byte), Integer), CType(CType(190, Byte), Integer))
        Me.AccordionControl.Appearance.Group.Hovered.BackColor2 = System.Drawing.Color.FromArgb(CType(CType(16, Byte), Integer), CType(CType(110, Byte), Integer), CType(CType(190, Byte), Integer))
        Me.AccordionControl.Appearance.Group.Hovered.BorderColor = System.Drawing.Color.FromArgb(CType(CType(16, Byte), Integer), CType(CType(110, Byte), Integer), CType(CType(190, Byte), Integer))
        Me.AccordionControl.Appearance.Group.Hovered.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.AccordionControl.Appearance.Group.Hovered.ForeColor = System.Drawing.Color.White
        Me.AccordionControl.Appearance.Group.Hovered.Options.UseBackColor = True
        Me.AccordionControl.Appearance.Group.Hovered.Options.UseBorderColor = True
        Me.AccordionControl.Appearance.Group.Hovered.Options.UseFont = True
        Me.AccordionControl.Appearance.Group.Hovered.Options.UseForeColor = True
        Me.AccordionControl.Appearance.Group.Normal.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.AccordionControl.Appearance.Group.Normal.ForeColor = System.Drawing.Color.FromArgb(CType(CType(27, Byte), Integer), CType(CType(41, Byte), Integer), CType(CType(62, Byte), Integer))
        Me.AccordionControl.Appearance.Group.Normal.Options.UseFont = True
        Me.AccordionControl.Appearance.Group.Normal.Options.UseForeColor = True
        Me.AccordionControl.Appearance.Item.Hovered.BackColor = System.Drawing.Color.FromArgb(CType(CType(16, Byte), Integer), CType(CType(110, Byte), Integer), CType(CType(190, Byte), Integer))
        Me.AccordionControl.Appearance.Item.Hovered.BackColor2 = System.Drawing.Color.FromArgb(CType(CType(16, Byte), Integer), CType(CType(110, Byte), Integer), CType(CType(190, Byte), Integer))
        Me.AccordionControl.Appearance.Item.Hovered.BorderColor = System.Drawing.Color.FromArgb(CType(CType(16, Byte), Integer), CType(CType(110, Byte), Integer), CType(CType(190, Byte), Integer))
        Me.AccordionControl.Appearance.Item.Hovered.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.AccordionControl.Appearance.Item.Hovered.ForeColor = System.Drawing.Color.White
        Me.AccordionControl.Appearance.Item.Hovered.Options.UseBackColor = True
        Me.AccordionControl.Appearance.Item.Hovered.Options.UseBorderColor = True
        Me.AccordionControl.Appearance.Item.Hovered.Options.UseFont = True
        Me.AccordionControl.Appearance.Item.Hovered.Options.UseForeColor = True
        Me.AccordionControl.Appearance.Item.Normal.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.AccordionControl.Appearance.Item.Normal.ForeColor = System.Drawing.Color.FromArgb(CType(CType(27, Byte), Integer), CType(CType(41, Byte), Integer), CType(CType(62, Byte), Integer))
        Me.AccordionControl.Appearance.Item.Normal.Options.UseFont = True
        Me.AccordionControl.Appearance.Item.Normal.Options.UseForeColor = True
        Me.AccordionControl.Dock = System.Windows.Forms.DockStyle.Fill
        Me.AccordionControl.Elements.AddRange(New DevExpress.XtraBars.Navigation.AccordionControlElement() {Me.AccordionControlSeparator6, Me.acciHOMIS, Me.btnPatient, Me.accPhlebotomy, Me.AccordionControlSeparator1, Me.btnPatientOrder, Me.AccordionControlSeparator2, Me.aceSection, Me.AccordionControlSeparator3, Me.btnDelta, Me.AccordionControlSeparator5, Me.aceReports, Me.AccordionControlSeparator7, Me.btnRealtime, Me.AccordionControlSeparator9, Me.accEmailPush})
        Me.AccordionControl.ExpandElementMode = DevExpress.XtraBars.Navigation.ExpandElementMode.[Single]
        Me.AccordionControl.Location = New System.Drawing.Point(0, 62)
        Me.AccordionControl.Name = "AccordionControl"
        Me.AccordionControl.ScrollBarMode = DevExpress.XtraBars.Navigation.ScrollBarMode.AutoCollapse
        Me.AccordionControl.Size = New System.Drawing.Size(250, 530)
        Me.AccordionControl.TabIndex = 2
        Me.AccordionControl.ViewType = DevExpress.XtraBars.Navigation.AccordionControlViewType.HamburgerMenu
        '
        'AccordionControlSeparator6
        '
        Me.AccordionControlSeparator6.Name = "AccordionControlSeparator6"
        '
        'acciHOMIS
        '
        Me.acciHOMIS.ImageOptions.SvgImage = CType(resources.GetObject("acciHOMIS.ImageOptions.SvgImage"), DevExpress.Utils.Svg.SvgImage)
        Me.acciHOMIS.Name = "acciHOMIS"
        Me.acciHOMIS.Text = "iHomis Integration"
        Me.acciHOMIS.Visible = False
        '
        'btnPatient
        '
        Me.btnPatient.ImageOptions.SvgImage = CType(resources.GetObject("btnPatient.ImageOptions.SvgImage"), DevExpress.Utils.Svg.SvgImage)
        Me.btnPatient.Name = "btnPatient"
        Me.btnPatient.Style = DevExpress.XtraBars.Navigation.ElementStyle.Item
        Me.btnPatient.Text = "Front Desk"
        '
        'accPhlebotomy
        '
        Me.accPhlebotomy.ImageOptions.SvgImage = CType(resources.GetObject("accPhlebotomy.ImageOptions.SvgImage"), DevExpress.Utils.Svg.SvgImage)
        Me.accPhlebotomy.Name = "accPhlebotomy"
        Me.accPhlebotomy.Style = DevExpress.XtraBars.Navigation.ElementStyle.Item
        Me.accPhlebotomy.Text = "Phlebotomy"
        '
        'AccordionControlSeparator1
        '
        Me.AccordionControlSeparator1.Name = "AccordionControlSeparator1"
        '
        'btnPatientOrder
        '
        Me.btnPatientOrder.Expanded = True
        Me.btnPatientOrder.ImageOptions.SvgImage = CType(resources.GetObject("btnPatientOrder.ImageOptions.SvgImage"), DevExpress.Utils.Svg.SvgImage)
        Me.btnPatientOrder.Name = "btnPatientOrder"
        Me.btnPatientOrder.Style = DevExpress.XtraBars.Navigation.ElementStyle.Item
        Me.btnPatientOrder.Text = "Order Management"
        '
        'AccordionControlSeparator2
        '
        Me.AccordionControlSeparator2.Name = "AccordionControlSeparator2"
        '
        'aceSection
        '
        Me.aceSection.Elements.AddRange(New DevExpress.XtraBars.Navigation.AccordionControlElement() {Me.aceHema, Me.aceChem, Me.aceSerology, Me.aceImmuno, Me.aceMicrobiology, Me.aceHistopath, Me.aceUrin, Me.aceFecal, Me.aceMolecular})
        Me.aceSection.Expanded = True
        Me.aceSection.ImageOptions.SvgImage = CType(resources.GetObject("aceSection.ImageOptions.SvgImage"), DevExpress.Utils.Svg.SvgImage)
        Me.aceSection.Name = "aceSection"
        Me.aceSection.Text = "Result Management"
        '
        'aceHema
        '
        Me.aceHema.Expanded = True
        Me.aceHema.ImageOptions.SvgImage = CType(resources.GetObject("aceHema.ImageOptions.SvgImage"), DevExpress.Utils.Svg.SvgImage)
        Me.aceHema.Name = "aceHema"
        Me.aceHema.Style = DevExpress.XtraBars.Navigation.ElementStyle.Item
        Me.aceHema.Text = "Hematology"
        '
        'aceChem
        '
        Me.aceChem.Expanded = True
        Me.aceChem.ImageOptions.SvgImage = CType(resources.GetObject("aceChem.ImageOptions.SvgImage"), DevExpress.Utils.Svg.SvgImage)
        Me.aceChem.Name = "aceChem"
        Me.aceChem.Style = DevExpress.XtraBars.Navigation.ElementStyle.Item
        Me.aceChem.Text = "Chemistry"
        '
        'aceSerology
        '
        Me.aceSerology.ImageOptions.SvgImage = CType(resources.GetObject("aceSerology.ImageOptions.SvgImage"), DevExpress.Utils.Svg.SvgImage)
        Me.aceSerology.Name = "aceSerology"
        Me.aceSerology.Text = "ImmunoSero"
        '
        'aceImmuno
        '
        Me.aceImmuno.ImageOptions.SvgImage = CType(resources.GetObject("aceImmuno.ImageOptions.SvgImage"), DevExpress.Utils.Svg.SvgImage)
        Me.aceImmuno.Name = "aceImmuno"
        Me.aceImmuno.Text = "Immunology"
        Me.aceImmuno.Visible = False
        '
        'aceMicrobiology
        '
        Me.aceMicrobiology.ImageOptions.SvgImage = CType(resources.GetObject("aceMicrobiology.ImageOptions.SvgImage"), DevExpress.Utils.Svg.SvgImage)
        Me.aceMicrobiology.Name = "aceMicrobiology"
        Me.aceMicrobiology.Text = "Microbiology"
        '
        'aceHistopath
        '
        Me.aceHistopath.ImageOptions.SvgImage = CType(resources.GetObject("aceHistopath.ImageOptions.SvgImage"), DevExpress.Utils.Svg.SvgImage)
        Me.aceHistopath.Name = "aceHistopath"
        Me.aceHistopath.Text = "Histopathology"
        Me.aceHistopath.Visible = False
        '
        'aceUrin
        '
        Me.aceUrin.ImageOptions.SvgImage = CType(resources.GetObject("aceUrin.ImageOptions.SvgImage"), DevExpress.Utils.Svg.SvgImage)
        Me.aceUrin.Name = "aceUrin"
        Me.aceUrin.Text = "Urinalysis"
        '
        'aceFecal
        '
        Me.aceFecal.ImageOptions.SvgImage = CType(resources.GetObject("aceFecal.ImageOptions.SvgImage"), DevExpress.Utils.Svg.SvgImage)
        Me.aceFecal.Name = "aceFecal"
        Me.aceFecal.Text = "Fecalysis"
        '
        'aceMolecular
        '
        Me.aceMolecular.ImageOptions.SvgImage = CType(resources.GetObject("aceMolecular.ImageOptions.SvgImage"), DevExpress.Utils.Svg.SvgImage)
        Me.aceMolecular.Name = "aceMolecular"
        Me.aceMolecular.Text = "Molecular Biology"
        '
        'AccordionControlSeparator3
        '
        Me.AccordionControlSeparator3.Name = "AccordionControlSeparator3"
        '
        'btnDelta
        '
        Me.btnDelta.ImageOptions.SvgImage = CType(resources.GetObject("btnDelta.ImageOptions.SvgImage"), DevExpress.Utils.Svg.SvgImage)
        Me.btnDelta.Name = "btnDelta"
        Me.btnDelta.Text = "Delta Check"
        '
        'AccordionControlSeparator5
        '
        Me.AccordionControlSeparator5.Name = "AccordionControlSeparator5"
        '
        'aceReports
        '
        Me.aceReports.Elements.AddRange(New DevExpress.XtraBars.Navigation.AccordionControlElement() {Me.btnQC, Me.btnWorkSheet, Me.btnCensus, Me.btnCompleteTat, Me.btnSpecimenLog})
        Me.aceReports.ImageOptions.SvgImage = CType(resources.GetObject("aceReports.ImageOptions.SvgImage"), DevExpress.Utils.Svg.SvgImage)
        Me.aceReports.Name = "aceReports"
        Me.aceReports.Text = "Management Reports"
        '
        'btnQC
        '
        Me.btnQC.ImageOptions.SvgImage = CType(resources.GetObject("btnQC.ImageOptions.SvgImage"), DevExpress.Utils.Svg.SvgImage)
        Me.btnQC.Name = "btnQC"
        Me.btnQC.Text = "Quality Control"
        '
        'btnWorkSheet
        '
        Me.btnWorkSheet.ImageOptions.SvgImage = CType(resources.GetObject("btnWorkSheet.ImageOptions.SvgImage"), DevExpress.Utils.Svg.SvgImage)
        Me.btnWorkSheet.Name = "btnWorkSheet"
        Me.btnWorkSheet.Text = "Worksheet"
        '
        'btnCensus
        '
        Me.btnCensus.ImageOptions.SvgImage = CType(resources.GetObject("btnCensus.ImageOptions.SvgImage"), DevExpress.Utils.Svg.SvgImage)
        Me.btnCensus.Name = "btnCensus"
        Me.btnCensus.Text = "Census"
        '
        'btnCompleteTat
        '
        Me.btnCompleteTat.ImageOptions.SvgImage = CType(resources.GetObject("btnCompleteTat.ImageOptions.SvgImage"), DevExpress.Utils.Svg.SvgImage)
        Me.btnCompleteTat.Name = "btnCompleteTat"
        Me.btnCompleteTat.Text = "Turnaround Time"
        '
        'btnSpecimenLog
        '
        Me.btnSpecimenLog.ImageOptions.SvgImage = CType(resources.GetObject("btnSpecimenLog.ImageOptions.SvgImage"), DevExpress.Utils.Svg.SvgImage)
        Me.btnSpecimenLog.Name = "btnSpecimenLog"
        Me.btnSpecimenLog.Text = "Specimen Log"
        '
        'AccordionControlSeparator7
        '
        Me.AccordionControlSeparator7.Name = "AccordionControlSeparator7"
        '
        'btnRealtime
        '
        Me.btnRealtime.ImageOptions.SvgImage = CType(resources.GetObject("btnRealtime.ImageOptions.SvgImage"), DevExpress.Utils.Svg.SvgImage)
        Me.btnRealtime.Name = "btnRealtime"
        Me.btnRealtime.Style = DevExpress.XtraBars.Navigation.ElementStyle.Item
        Me.btnRealtime.Text = "Realtime"
        '
        'AccordionControlSeparator9
        '
        Me.AccordionControlSeparator9.Name = "AccordionControlSeparator9"
        '
        'accEmailPush
        '
        Me.accEmailPush.ImageOptions.SvgImage = CType(resources.GetObject("accEmailPush.ImageOptions.SvgImage"), DevExpress.Utils.Svg.SvgImage)
        Me.accEmailPush.Name = "accEmailPush"
        Me.accEmailPush.Style = DevExpress.XtraBars.Navigation.ElementStyle.Item
        Me.accEmailPush.Text = "Email Push Result"
        '
        'btnLogout
        '
        Me.btnLogout.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.btnLogout.ImageOptions.ImageToTextAlignment = DevExpress.XtraEditors.ImageAlignToText.LeftCenter
        Me.btnLogout.ImageOptions.SvgImage = CType(resources.GetObject("btnLogout.ImageOptions.SvgImage"), DevExpress.Utils.Svg.SvgImage)
        Me.btnLogout.Location = New System.Drawing.Point(0, 592)
        Me.btnLogout.Name = "btnLogout"
        Me.btnLogout.Size = New System.Drawing.Size(250, 33)
        Me.btnLogout.TabIndex = 13
        Me.btnLogout.Text = "Logout"
        '
        'PanelControl2
        '
        Me.PanelControl2.Appearance.BackColor = System.Drawing.Color.FromArgb(CType(CType(16, Byte), Integer), CType(CType(110, Byte), Integer), CType(CType(190, Byte), Integer))
        Me.PanelControl2.Appearance.Options.UseBackColor = True
        Me.PanelControl2.BorderStyle = DevExpress.XtraEditors.Controls.BorderStyles.NoBorder
        Me.PanelControl2.Controls.Add(Me.PictureBox1)
        Me.PanelControl2.Controls.Add(Me.SeparatorControl1)
        Me.PanelControl2.Controls.Add(Me.LabelControl1)
        Me.PanelControl2.Dock = System.Windows.Forms.DockStyle.Top
        Me.PanelControl2.Location = New System.Drawing.Point(0, 0)
        Me.PanelControl2.LookAndFeel.Style = DevExpress.LookAndFeel.LookAndFeelStyle.UltraFlat
        Me.PanelControl2.LookAndFeel.UseDefaultLookAndFeel = False
        Me.PanelControl2.Name = "PanelControl2"
        Me.PanelControl2.Size = New System.Drawing.Size(250, 62)
        Me.PanelControl2.TabIndex = 8
        '
        'PictureBox1
        '
        Me.PictureBox1.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.PictureBox1.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch
        Me.PictureBox1.Image = Global.LIS.My.Resources.Resources.SBSI_Logo
        Me.PictureBox1.Location = New System.Drawing.Point(12, 5)
        Me.PictureBox1.Name = "PictureBox1"
        Me.PictureBox1.Size = New System.Drawing.Size(131, 53)
        Me.PictureBox1.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        Me.PictureBox1.TabIndex = 0
        Me.PictureBox1.TabStop = False
        '
        'SeparatorControl1
        '
        Me.SeparatorControl1.Anchor = System.Windows.Forms.AnchorStyles.Left
        Me.SeparatorControl1.LineColor = System.Drawing.Color.Gainsboro
        Me.SeparatorControl1.LineOrientation = System.Windows.Forms.Orientation.Vertical
        Me.SeparatorControl1.LineThickness = 1
        Me.SeparatorControl1.Location = New System.Drawing.Point(142, 7)
        Me.SeparatorControl1.Name = "SeparatorControl1"
        Me.SeparatorControl1.Size = New System.Drawing.Size(20, 51)
        Me.SeparatorControl1.TabIndex = 11
        '
        'LabelControl1
        '
        Me.LabelControl1.AllowGlyphSkinning = DevExpress.Utils.DefaultBoolean.[True]
        Me.LabelControl1.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.LabelControl1.Appearance.Font = New System.Drawing.Font("Segoe UI Semibold", 15.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.LabelControl1.Appearance.ForeColor = System.Drawing.Color.White
        Me.LabelControl1.Appearance.Options.UseFont = True
        Me.LabelControl1.Appearance.Options.UseForeColor = True
        Me.LabelControl1.Appearance.Options.UseTextOptions = True
        Me.LabelControl1.Appearance.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center
        Me.LabelControl1.AutoSizeMode = DevExpress.XtraEditors.LabelAutoSizeMode.None
        Me.LabelControl1.ImageOptions.Alignment = System.Drawing.ContentAlignment.MiddleLeft
        Me.LabelControl1.Location = New System.Drawing.Point(163, 17)
        Me.LabelControl1.Name = "LabelControl1"
        Me.LabelControl1.Size = New System.Drawing.Size(68, 32)
        Me.LabelControl1.TabIndex = 12
        Me.LabelControl1.Text = "LIS Pro"
        '
        'PanelControl1
        '
        Me.PanelControl1.Appearance.BackColor = System.Drawing.Color.FromArgb(CType(CType(16, Byte), Integer), CType(CType(110, Byte), Integer), CType(CType(190, Byte), Integer))
        Me.PanelControl1.Appearance.Options.UseBackColor = True
        Me.PanelControl1.BorderStyle = DevExpress.XtraEditors.Controls.BorderStyles.NoBorder
        Me.PanelControl1.Controls.Add(Me.lblPending)
        Me.PanelControl1.Controls.Add(Me.lblNew)
        Me.PanelControl1.Controls.Add(Me.LabelControl4)
        Me.PanelControl1.Controls.Add(Me.LabelControl3)
        Me.PanelControl1.Controls.Add(Me.SeparatorControl2)
        Me.PanelControl1.Controls.Add(Me.lblTitle)
        Me.PanelControl1.Controls.Add(Me.mnuHelp)
        Me.PanelControl1.Controls.Add(Me.lblTitleHome)
        Me.PanelControl1.Controls.Add(Me.mnuSetting)
        Me.PanelControl1.Controls.Add(Me.mnuSysSettings)
        Me.PanelControl1.Controls.Add(Me.mnuUser)
        Me.PanelControl1.Dock = System.Windows.Forms.DockStyle.Top
        Me.PanelControl1.Location = New System.Drawing.Point(250, 0)
        Me.PanelControl1.LookAndFeel.Style = DevExpress.LookAndFeel.LookAndFeelStyle.UltraFlat
        Me.PanelControl1.LookAndFeel.UseDefaultLookAndFeel = False
        Me.PanelControl1.Name = "PanelControl1"
        Me.PanelControl1.Size = New System.Drawing.Size(938, 97)
        Me.PanelControl1.TabIndex = 1
        '
        'lblPending
        '
        Me.lblPending.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.lblPending.Appearance.Font = New System.Drawing.Font("Segoe UI Semibold", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblPending.Appearance.ForeColor = System.Drawing.Color.Orange
        Me.lblPending.Appearance.Options.UseFont = True
        Me.lblPending.Appearance.Options.UseForeColor = True
        Me.lblPending.Appearance.Options.UseTextOptions = True
        Me.lblPending.Appearance.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Far
        Me.lblPending.AutoSizeMode = DevExpress.XtraEditors.LabelAutoSizeMode.Vertical
        Me.lblPending.Location = New System.Drawing.Point(743, 70)
        Me.lblPending.Name = "lblPending"
        Me.lblPending.Size = New System.Drawing.Size(186, 21)
        Me.lblPending.TabIndex = 17
        Me.lblPending.Text = "Pending Order (123123)"
        '
        'lblNew
        '
        Me.lblNew.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.lblNew.Appearance.Font = New System.Drawing.Font("Segoe UI Semibold", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblNew.Appearance.ForeColor = System.Drawing.Color.Lime
        Me.lblNew.Appearance.Options.UseFont = True
        Me.lblNew.Appearance.Options.UseForeColor = True
        Me.lblNew.Appearance.Options.UseTextOptions = True
        Me.lblNew.Appearance.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Far
        Me.lblNew.AutoSizeMode = DevExpress.XtraEditors.LabelAutoSizeMode.Vertical
        Me.lblNew.Location = New System.Drawing.Point(558, 69)
        Me.lblNew.Name = "lblNew"
        Me.lblNew.Size = New System.Drawing.Size(181, 21)
        Me.lblNew.TabIndex = 16
        Me.lblNew.Text = "New Order (1231231)"
        '
        'LabelControl4
        '
        Me.LabelControl4.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.LabelControl4.Appearance.Font = New System.Drawing.Font("Segoe UI Semibold", 13.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.LabelControl4.Appearance.ForeColor = System.Drawing.Color.White
        Me.LabelControl4.Appearance.Options.UseFont = True
        Me.LabelControl4.Appearance.Options.UseForeColor = True
        Me.LabelControl4.Location = New System.Drawing.Point(11, 19)
        Me.LabelControl4.Name = "LabelControl4"
        Me.LabelControl4.Size = New System.Drawing.Size(106, 23)
        Me.LabelControl4.TabIndex = 15
        Me.LabelControl4.Text = "Version 1.8.1.0"
        '
        'LabelControl3
        '
        Me.LabelControl3.AllowGlyphSkinning = DevExpress.Utils.DefaultBoolean.[True]
        Me.LabelControl3.Appearance.Font = New System.Drawing.Font("Segoe UI Semibold", 13.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.LabelControl3.Appearance.ForeColor = System.Drawing.Color.White
        Me.LabelControl3.Appearance.Options.UseFont = True
        Me.LabelControl3.Appearance.Options.UseForeColor = True
        Me.LabelControl3.ImageOptions.SvgImage = CType(resources.GetObject("LabelControl3.ImageOptions.SvgImage"), DevExpress.Utils.Svg.SvgImage)
        Me.LabelControl3.ImageOptions.SvgImageSize = New System.Drawing.Size(24, 24)
        Me.LabelControl3.Location = New System.Drawing.Point(86, 66)
        Me.LabelControl3.Name = "LabelControl3"
        Me.LabelControl3.Size = New System.Drawing.Size(24, 24)
        Me.LabelControl3.TabIndex = 14
        '
        'SeparatorControl2
        '
        Me.SeparatorControl2.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.SeparatorControl2.LineColor = System.Drawing.Color.Gainsboro
        Me.SeparatorControl2.LineThickness = 1
        Me.SeparatorControl2.Location = New System.Drawing.Point(3, 45)
        Me.SeparatorControl2.Name = "SeparatorControl2"
        Me.SeparatorControl2.Size = New System.Drawing.Size(932, 18)
        Me.SeparatorControl2.TabIndex = 13
        '
        'lblTitle
        '
        Me.lblTitle.Appearance.Font = New System.Drawing.Font("Segoe UI Semibold", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblTitle.Appearance.ForeColor = System.Drawing.Color.White
        Me.lblTitle.Appearance.Options.UseFont = True
        Me.lblTitle.Appearance.Options.UseForeColor = True
        Me.lblTitle.Location = New System.Drawing.Point(122, 70)
        Me.lblTitle.Name = "lblTitle"
        Me.lblTitle.Size = New System.Drawing.Size(47, 21)
        Me.lblTitle.TabIndex = 11
        Me.lblTitle.Text = "HOME"
        '
        'mnuHelp
        '
        Me.mnuHelp.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.mnuHelp.Appearance.BackColor = System.Drawing.Color.FromArgb(CType(CType(16, Byte), Integer), CType(CType(110, Byte), Integer), CType(CType(190, Byte), Integer))
        Me.mnuHelp.Appearance.BackColor2 = System.Drawing.Color.FromArgb(CType(CType(16, Byte), Integer), CType(CType(110, Byte), Integer), CType(CType(190, Byte), Integer))
        Me.mnuHelp.Appearance.BorderColor = System.Drawing.Color.FromArgb(CType(CType(16, Byte), Integer), CType(CType(110, Byte), Integer), CType(CType(190, Byte), Integer))
        Me.mnuHelp.Appearance.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.mnuHelp.Appearance.ForeColor = System.Drawing.Color.White
        Me.mnuHelp.Appearance.Options.UseBackColor = True
        Me.mnuHelp.Appearance.Options.UseBorderColor = True
        Me.mnuHelp.Appearance.Options.UseFont = True
        Me.mnuHelp.Appearance.Options.UseForeColor = True
        Me.mnuHelp.AppearanceDropDown.BackColor = System.Drawing.Color.FromArgb(CType(CType(6, Byte), Integer), CType(CType(31, Byte), Integer), CType(CType(71, Byte), Integer))
        Me.mnuHelp.AppearanceDropDown.BackColor2 = System.Drawing.Color.FromArgb(CType(CType(6, Byte), Integer), CType(CType(31, Byte), Integer), CType(CType(71, Byte), Integer))
        Me.mnuHelp.AppearanceDropDown.BorderColor = System.Drawing.Color.FromArgb(CType(CType(6, Byte), Integer), CType(CType(31, Byte), Integer), CType(CType(71, Byte), Integer))
        Me.mnuHelp.AppearanceDropDown.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.mnuHelp.AppearanceDropDown.ForeColor = System.Drawing.Color.White
        Me.mnuHelp.AppearanceDropDown.Options.UseBackColor = True
        Me.mnuHelp.AppearanceDropDown.Options.UseBorderColor = True
        Me.mnuHelp.AppearanceDropDown.Options.UseFont = True
        Me.mnuHelp.AppearanceDropDown.Options.UseForeColor = True
        Me.mnuHelp.AppearanceDropDownHovered.BackColor = System.Drawing.Color.FromArgb(CType(CType(16, Byte), Integer), CType(CType(110, Byte), Integer), CType(CType(190, Byte), Integer))
        Me.mnuHelp.AppearanceDropDownHovered.BackColor2 = System.Drawing.Color.FromArgb(CType(CType(16, Byte), Integer), CType(CType(110, Byte), Integer), CType(CType(190, Byte), Integer))
        Me.mnuHelp.AppearanceDropDownHovered.BorderColor = System.Drawing.Color.FromArgb(CType(CType(16, Byte), Integer), CType(CType(110, Byte), Integer), CType(CType(190, Byte), Integer))
        Me.mnuHelp.AppearanceDropDownHovered.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.mnuHelp.AppearanceDropDownHovered.ForeColor = System.Drawing.Color.White
        Me.mnuHelp.AppearanceDropDownHovered.Options.UseBackColor = True
        Me.mnuHelp.AppearanceDropDownHovered.Options.UseBorderColor = True
        Me.mnuHelp.AppearanceDropDownHovered.Options.UseFont = True
        Me.mnuHelp.AppearanceDropDownHovered.Options.UseForeColor = True
        Me.mnuHelp.DropDownArrowStyle = DevExpress.XtraEditors.DropDownArrowStyle.Show
        Me.mnuHelp.DropDownControl = Me.pmLogs
        Me.mnuHelp.ImageOptions.AllowGlyphSkinning = DevExpress.Utils.DefaultBoolean.[True]
        Me.mnuHelp.ImageOptions.SvgImage = CType(resources.GetObject("mnuHelp.ImageOptions.SvgImage"), DevExpress.Utils.Svg.SvgImage)
        Me.mnuHelp.Location = New System.Drawing.Point(572, 10)
        Me.mnuHelp.Name = "mnuHelp"
        Me.mnuHelp.Size = New System.Drawing.Size(46, 32)
        Me.mnuHelp.TabIndex = 12
        '
        'pmLogs
        '
        Me.pmLogs.LinksPersistInfo.AddRange(New DevExpress.XtraBars.LinkPersistInfo() {New DevExpress.XtraBars.LinkPersistInfo(Me.btnSpecimenLogs), New DevExpress.XtraBars.LinkPersistInfo(Me.btnActivityLog)})
        Me.pmLogs.Manager = Me.BarManager
        Me.pmLogs.Name = "pmLogs"
        '
        'btnSpecimenLogs
        '
        Me.btnSpecimenLogs.Caption = "Specimen Logs"
        Me.btnSpecimenLogs.Id = 25
        Me.btnSpecimenLogs.ImageOptions.SvgImage = CType(resources.GetObject("btnSpecimenLogs.ImageOptions.SvgImage"), DevExpress.Utils.Svg.SvgImage)
        Me.btnSpecimenLogs.Name = "btnSpecimenLogs"
        '
        'btnActivityLog
        '
        Me.btnActivityLog.Caption = "Activity Logs"
        Me.btnActivityLog.Id = 26
        Me.btnActivityLog.ImageOptions.SvgImage = CType(resources.GetObject("btnActivityLog.ImageOptions.SvgImage"), DevExpress.Utils.Svg.SvgImage)
        Me.btnActivityLog.Name = "btnActivityLog"
        '
        'BarManager
        '
        Me.BarManager.DockControls.Add(Me.barDockControlTop)
        Me.BarManager.DockControls.Add(Me.barDockControlBottom)
        Me.BarManager.DockControls.Add(Me.barDockControlLeft)
        Me.BarManager.DockControls.Add(Me.barDockControlRight)
        Me.BarManager.Form = Me
        Me.BarManager.Items.AddRange(New DevExpress.XtraBars.BarItem() {Me.btnSpecimen, Me.btnDefaultSpecimen, Me.btnTestGroup, Me.btnChemRange, Me.btnMedTech, Me.btnPathologist, Me.btnDoctors, Me.btnDepartment, Me.btnCompany, Me.btnShift, Me.btnTest, Me.btnDefaultPrinter, Me.btnPrinterSetting, Me.btnBackup, Me.btnRestore, Me.btnTruncate, Me.btnMachines, Me.btnAccounts, Me.btnEditAccount, Me.btnChange, Me.btnTestName, Me.btnUnit, Me.btnSettings, Me.btnMngUser, Me.btnUsers, Me.btnSpecimenLogs, Me.btnActivityLog, Me.btnEmail})
        Me.BarManager.MaxItemId = 28
        Me.BarManager.MdiMenuMergeStyle = DevExpress.XtraBars.BarMdiMenuMergeStyle.Never
        '
        'barDockControlTop
        '
        Me.barDockControlTop.CausesValidation = False
        Me.barDockControlTop.Dock = System.Windows.Forms.DockStyle.Top
        Me.barDockControlTop.Location = New System.Drawing.Point(0, 0)
        Me.barDockControlTop.Manager = Me.BarManager
        Me.barDockControlTop.Size = New System.Drawing.Size(1188, 0)
        '
        'barDockControlBottom
        '
        Me.barDockControlBottom.CausesValidation = False
        Me.barDockControlBottom.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.barDockControlBottom.Location = New System.Drawing.Point(0, 650)
        Me.barDockControlBottom.Manager = Me.BarManager
        Me.barDockControlBottom.Size = New System.Drawing.Size(1188, 0)
        '
        'barDockControlLeft
        '
        Me.barDockControlLeft.CausesValidation = False
        Me.barDockControlLeft.Dock = System.Windows.Forms.DockStyle.Left
        Me.barDockControlLeft.Location = New System.Drawing.Point(0, 0)
        Me.barDockControlLeft.Manager = Me.BarManager
        Me.barDockControlLeft.Size = New System.Drawing.Size(0, 650)
        '
        'barDockControlRight
        '
        Me.barDockControlRight.CausesValidation = False
        Me.barDockControlRight.Dock = System.Windows.Forms.DockStyle.Right
        Me.barDockControlRight.Location = New System.Drawing.Point(1188, 0)
        Me.barDockControlRight.Manager = Me.BarManager
        Me.barDockControlRight.Size = New System.Drawing.Size(0, 650)
        '
        'btnSpecimen
        '
        Me.btnSpecimen.Caption = "Machine Test Mapping"
        Me.btnSpecimen.Id = 0
        Me.btnSpecimen.ImageOptions.SvgImage = CType(resources.GetObject("btnSpecimen.ImageOptions.SvgImage"), DevExpress.Utils.Svg.SvgImage)
        Me.btnSpecimen.Name = "btnSpecimen"
        '
        'btnDefaultSpecimen
        '
        Me.btnDefaultSpecimen.Caption = "HIS Test Mapping"
        Me.btnDefaultSpecimen.Id = 1
        Me.btnDefaultSpecimen.ImageOptions.SvgImage = CType(resources.GetObject("btnDefaultSpecimen.ImageOptions.SvgImage"), DevExpress.Utils.Svg.SvgImage)
        Me.btnDefaultSpecimen.Name = "btnDefaultSpecimen"
        '
        'btnTestGroup
        '
        Me.btnTestGroup.Caption = "Test Groups"
        Me.btnTestGroup.Id = 2
        Me.btnTestGroup.Name = "btnTestGroup"
        '
        'btnChemRange
        '
        Me.btnChemRange.Caption = "Referance Range"
        Me.btnChemRange.Id = 3
        Me.btnChemRange.ImageOptions.SvgImage = CType(resources.GetObject("btnChemRange.ImageOptions.SvgImage"), DevExpress.Utils.Svg.SvgImage)
        Me.btnChemRange.Name = "btnChemRange"
        '
        'btnMedTech
        '
        Me.btnMedTech.Caption = "Signatories"
        Me.btnMedTech.Id = 4
        Me.btnMedTech.ImageOptions.SvgImage = CType(resources.GetObject("btnMedTech.ImageOptions.SvgImage"), DevExpress.Utils.Svg.SvgImage)
        Me.btnMedTech.Name = "btnMedTech"
        '
        'btnPathologist
        '
        Me.btnPathologist.Caption = "Pathologist"
        Me.btnPathologist.Id = 5
        Me.btnPathologist.ImageOptions.SvgImage = CType(resources.GetObject("btnPathologist.ImageOptions.SvgImage"), DevExpress.Utils.Svg.SvgImage)
        Me.btnPathologist.Name = "btnPathologist"
        '
        'btnDoctors
        '
        Me.btnDoctors.Caption = "Requestor"
        Me.btnDoctors.Id = 6
        Me.btnDoctors.ImageOptions.SvgImage = CType(resources.GetObject("btnDoctors.ImageOptions.SvgImage"), DevExpress.Utils.Svg.SvgImage)
        Me.btnDoctors.Name = "btnDoctors"
        '
        'btnDepartment
        '
        Me.btnDepartment.Border = DevExpress.XtraEditors.Controls.BorderStyles.[Default]
        Me.btnDepartment.Caption = "Room/Ward"
        Me.btnDepartment.Id = 7
        Me.btnDepartment.ImageOptions.SvgImage = CType(resources.GetObject("btnDepartment.ImageOptions.SvgImage"), DevExpress.Utils.Svg.SvgImage)
        Me.btnDepartment.Name = "btnDepartment"
        '
        'btnCompany
        '
        Me.btnCompany.Caption = "Hospital Profile"
        Me.btnCompany.Id = 8
        Me.btnCompany.ImageOptions.SvgImage = CType(resources.GetObject("btnCompany.ImageOptions.SvgImage"), DevExpress.Utils.Svg.SvgImage)
        Me.btnCompany.Name = "btnCompany"
        '
        'btnShift
        '
        Me.btnShift.Caption = "Shift Assignment"
        Me.btnShift.Id = 9
        Me.btnShift.Name = "btnShift"
        '
        'btnTest
        '
        Me.btnTest.Caption = "System Configuration"
        Me.btnTest.Id = 10
        Me.btnTest.ImageOptions.SvgImage = CType(resources.GetObject("btnTest.ImageOptions.SvgImage"), DevExpress.Utils.Svg.SvgImage)
        Me.btnTest.Name = "btnTest"
        '
        'btnDefaultPrinter
        '
        Me.btnDefaultPrinter.Caption = "Default Printer"
        Me.btnDefaultPrinter.Id = 11
        Me.btnDefaultPrinter.ImageOptions.SvgImage = CType(resources.GetObject("btnDefaultPrinter.ImageOptions.SvgImage"), DevExpress.Utils.Svg.SvgImage)
        Me.btnDefaultPrinter.Name = "btnDefaultPrinter"
        '
        'btnPrinterSetting
        '
        Me.btnPrinterSetting.Caption = "Printer Settings"
        Me.btnPrinterSetting.Id = 12
        Me.btnPrinterSetting.ImageOptions.SvgImage = CType(resources.GetObject("btnPrinterSetting.ImageOptions.SvgImage"), DevExpress.Utils.Svg.SvgImage)
        Me.btnPrinterSetting.Name = "btnPrinterSetting"
        '
        'btnBackup
        '
        Me.btnBackup.Caption = "Backup Database"
        Me.btnBackup.Id = 13
        Me.btnBackup.ImageOptions.SvgImage = CType(resources.GetObject("btnBackup.ImageOptions.SvgImage"), DevExpress.Utils.Svg.SvgImage)
        Me.btnBackup.Name = "btnBackup"
        '
        'btnRestore
        '
        Me.btnRestore.Caption = "Restore Database"
        Me.btnRestore.Id = 14
        Me.btnRestore.ImageOptions.SvgImage = CType(resources.GetObject("btnRestore.ImageOptions.SvgImage"), DevExpress.Utils.Svg.SvgImage)
        Me.btnRestore.Name = "btnRestore"
        '
        'btnTruncate
        '
        Me.btnTruncate.Caption = "Truncate Tables"
        Me.btnTruncate.Id = 15
        Me.btnTruncate.ImageOptions.SvgImage = CType(resources.GetObject("btnTruncate.ImageOptions.SvgImage"), DevExpress.Utils.Svg.SvgImage)
        Me.btnTruncate.Name = "btnTruncate"
        '
        'btnMachines
        '
        Me.btnMachines.Caption = "Connection Settings"
        Me.btnMachines.Id = 16
        Me.btnMachines.ImageOptions.SvgImage = CType(resources.GetObject("btnMachines.ImageOptions.SvgImage"), DevExpress.Utils.Svg.SvgImage)
        Me.btnMachines.Name = "btnMachines"
        '
        'btnAccounts
        '
        Me.btnAccounts.Caption = "Manage User"
        Me.btnAccounts.Id = 17
        Me.btnAccounts.ImageOptions.SvgImage = CType(resources.GetObject("btnAccounts.ImageOptions.SvgImage"), DevExpress.Utils.Svg.SvgImage)
        Me.btnAccounts.Name = "btnAccounts"
        '
        'btnEditAccount
        '
        Me.btnEditAccount.Caption = "Edit Account"
        Me.btnEditAccount.Id = 18
        Me.btnEditAccount.ImageOptions.SvgImage = CType(resources.GetObject("btnEditAccount.ImageOptions.SvgImage"), DevExpress.Utils.Svg.SvgImage)
        Me.btnEditAccount.Name = "btnEditAccount"
        '
        'btnChange
        '
        Me.btnChange.Caption = "Change User"
        Me.btnChange.Id = 19
        Me.btnChange.ImageOptions.SvgImage = CType(resources.GetObject("btnChange.ImageOptions.SvgImage"), DevExpress.Utils.Svg.SvgImage)
        Me.btnChange.Name = "btnChange"
        '
        'btnTestName
        '
        Me.btnTestName.Caption = "Section"
        Me.btnTestName.Id = 20
        Me.btnTestName.ImageOptions.SvgImage = CType(resources.GetObject("btnTestName.ImageOptions.SvgImage"), DevExpress.Utils.Svg.SvgImage)
        Me.btnTestName.Name = "btnTestName"
        '
        'btnUnit
        '
        Me.btnUnit.Caption = "Units"
        Me.btnUnit.Id = 21
        Me.btnUnit.Name = "btnUnit"
        '
        'btnSettings
        '
        Me.btnSettings.Caption = "QC Settings"
        Me.btnSettings.Id = 22
        Me.btnSettings.ImageOptions.SvgImage = CType(resources.GetObject("btnSettings.ImageOptions.SvgImage"), DevExpress.Utils.Svg.SvgImage)
        Me.btnSettings.Name = "btnSettings"
        '
        'btnMngUser
        '
        Me.btnMngUser.Caption = "Manage Account"
        Me.btnMngUser.Id = 23
        Me.btnMngUser.ImageOptions.SvgImage = CType(resources.GetObject("btnMngUser.ImageOptions.SvgImage"), DevExpress.Utils.Svg.SvgImage)
        Me.btnMngUser.Name = "btnMngUser"
        '
        'btnUsers
        '
        Me.btnUsers.Caption = "User"
        Me.btnUsers.Id = 24
        Me.btnUsers.ImageOptions.SvgImage = CType(resources.GetObject("btnUsers.ImageOptions.SvgImage"), DevExpress.Utils.Svg.SvgImage)
        Me.btnUsers.Name = "btnUsers"
        '
        'btnEmail
        '
        Me.btnEmail.Caption = "Email Content"
        Me.btnEmail.Id = 27
        Me.btnEmail.ImageOptions.SvgImage = CType(resources.GetObject("btnEmail.ImageOptions.SvgImage"), DevExpress.Utils.Svg.SvgImage)
        Me.btnEmail.Name = "btnEmail"
        '
        'lblTitleHome
        '
        Me.lblTitleHome.AllowGlyphSkinning = DevExpress.Utils.DefaultBoolean.[True]
        Me.lblTitleHome.Appearance.Font = New System.Drawing.Font("Segoe UI Semibold", 15.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblTitleHome.Appearance.ForeColor = System.Drawing.Color.White
        Me.lblTitleHome.Appearance.Options.UseFont = True
        Me.lblTitleHome.Appearance.Options.UseForeColor = True
        Me.lblTitleHome.AutoSizeMode = DevExpress.XtraEditors.LabelAutoSizeMode.None
        Me.lblTitleHome.ImageOptions.Alignment = System.Drawing.ContentAlignment.MiddleLeft
        Me.lblTitleHome.Location = New System.Drawing.Point(11, 62)
        Me.lblTitleHome.Name = "lblTitleHome"
        Me.lblTitleHome.Size = New System.Drawing.Size(132, 32)
        Me.lblTitleHome.TabIndex = 6
        Me.lblTitleHome.Text = "HOME"
        '
        'mnuSetting
        '
        Me.mnuSetting.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.mnuSetting.Appearance.BackColor = System.Drawing.Color.FromArgb(CType(CType(16, Byte), Integer), CType(CType(110, Byte), Integer), CType(CType(190, Byte), Integer))
        Me.mnuSetting.Appearance.BackColor2 = System.Drawing.Color.FromArgb(CType(CType(16, Byte), Integer), CType(CType(110, Byte), Integer), CType(CType(190, Byte), Integer))
        Me.mnuSetting.Appearance.BorderColor = System.Drawing.Color.FromArgb(CType(CType(16, Byte), Integer), CType(CType(110, Byte), Integer), CType(CType(190, Byte), Integer))
        Me.mnuSetting.Appearance.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.mnuSetting.Appearance.ForeColor = System.Drawing.Color.White
        Me.mnuSetting.Appearance.Options.UseBackColor = True
        Me.mnuSetting.Appearance.Options.UseBorderColor = True
        Me.mnuSetting.Appearance.Options.UseFont = True
        Me.mnuSetting.Appearance.Options.UseForeColor = True
        Me.mnuSetting.AppearanceDropDown.BackColor = System.Drawing.Color.FromArgb(CType(CType(6, Byte), Integer), CType(CType(31, Byte), Integer), CType(CType(71, Byte), Integer))
        Me.mnuSetting.AppearanceDropDown.BackColor2 = System.Drawing.Color.FromArgb(CType(CType(6, Byte), Integer), CType(CType(31, Byte), Integer), CType(CType(71, Byte), Integer))
        Me.mnuSetting.AppearanceDropDown.BorderColor = System.Drawing.Color.FromArgb(CType(CType(6, Byte), Integer), CType(CType(31, Byte), Integer), CType(CType(71, Byte), Integer))
        Me.mnuSetting.AppearanceDropDown.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.mnuSetting.AppearanceDropDown.ForeColor = System.Drawing.Color.White
        Me.mnuSetting.AppearanceDropDown.Options.UseBackColor = True
        Me.mnuSetting.AppearanceDropDown.Options.UseBorderColor = True
        Me.mnuSetting.AppearanceDropDown.Options.UseFont = True
        Me.mnuSetting.AppearanceDropDown.Options.UseForeColor = True
        Me.mnuSetting.AppearanceDropDownHovered.BackColor = System.Drawing.Color.FromArgb(CType(CType(16, Byte), Integer), CType(CType(110, Byte), Integer), CType(CType(190, Byte), Integer))
        Me.mnuSetting.AppearanceDropDownHovered.BackColor2 = System.Drawing.Color.FromArgb(CType(CType(16, Byte), Integer), CType(CType(110, Byte), Integer), CType(CType(190, Byte), Integer))
        Me.mnuSetting.AppearanceDropDownHovered.BorderColor = System.Drawing.Color.FromArgb(CType(CType(16, Byte), Integer), CType(CType(110, Byte), Integer), CType(CType(190, Byte), Integer))
        Me.mnuSetting.AppearanceDropDownHovered.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.mnuSetting.AppearanceDropDownHovered.ForeColor = System.Drawing.Color.White
        Me.mnuSetting.AppearanceDropDownHovered.Options.UseBackColor = True
        Me.mnuSetting.AppearanceDropDownHovered.Options.UseBorderColor = True
        Me.mnuSetting.AppearanceDropDownHovered.Options.UseFont = True
        Me.mnuSetting.AppearanceDropDownHovered.Options.UseForeColor = True
        Me.mnuSetting.DropDownArrowStyle = DevExpress.XtraEditors.DropDownArrowStyle.Show
        Me.mnuSetting.DropDownControl = Me.pmSettings
        Me.mnuSetting.ImageOptions.AllowGlyphSkinning = DevExpress.Utils.DefaultBoolean.[True]
        Me.mnuSetting.ImageOptions.SvgImage = CType(resources.GetObject("mnuSetting.ImageOptions.SvgImage"), DevExpress.Utils.Svg.SvgImage)
        Me.mnuSetting.Location = New System.Drawing.Point(624, 10)
        Me.mnuSetting.Name = "mnuSetting"
        Me.mnuSetting.Size = New System.Drawing.Size(47, 32)
        Me.mnuSetting.TabIndex = 5
        '
        'pmSettings
        '
        Me.pmSettings.LinksPersistInfo.AddRange(New DevExpress.XtraBars.LinkPersistInfo() {New DevExpress.XtraBars.LinkPersistInfo(Me.btnDefaultSpecimen), New DevExpress.XtraBars.LinkPersistInfo(Me.btnSpecimen), New DevExpress.XtraBars.LinkPersistInfo(Me.btnChemRange, True), New DevExpress.XtraBars.LinkPersistInfo(Me.btnSettings), New DevExpress.XtraBars.LinkPersistInfo(Me.btnTestName), New DevExpress.XtraBars.LinkPersistInfo(Me.btnMedTech, True), New DevExpress.XtraBars.LinkPersistInfo(Me.btnPathologist), New DevExpress.XtraBars.LinkPersistInfo(Me.btnDoctors), New DevExpress.XtraBars.LinkPersistInfo(Me.btnUsers), New DevExpress.XtraBars.LinkPersistInfo(Me.btnDepartment, True), New DevExpress.XtraBars.LinkPersistInfo(Me.btnCompany), New DevExpress.XtraBars.LinkPersistInfo(Me.btnEmail), New DevExpress.XtraBars.LinkPersistInfo(Me.btnMngUser, True)})
        Me.pmSettings.Manager = Me.BarManager
        Me.pmSettings.Name = "pmSettings"
        '
        'mnuSysSettings
        '
        Me.mnuSysSettings.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.mnuSysSettings.Appearance.BackColor = System.Drawing.Color.FromArgb(CType(CType(16, Byte), Integer), CType(CType(110, Byte), Integer), CType(CType(190, Byte), Integer))
        Me.mnuSysSettings.Appearance.BackColor2 = System.Drawing.Color.FromArgb(CType(CType(16, Byte), Integer), CType(CType(110, Byte), Integer), CType(CType(190, Byte), Integer))
        Me.mnuSysSettings.Appearance.BorderColor = System.Drawing.Color.FromArgb(CType(CType(16, Byte), Integer), CType(CType(110, Byte), Integer), CType(CType(190, Byte), Integer))
        Me.mnuSysSettings.Appearance.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.mnuSysSettings.Appearance.ForeColor = System.Drawing.Color.White
        Me.mnuSysSettings.Appearance.Options.UseBackColor = True
        Me.mnuSysSettings.Appearance.Options.UseBorderColor = True
        Me.mnuSysSettings.Appearance.Options.UseFont = True
        Me.mnuSysSettings.Appearance.Options.UseForeColor = True
        Me.mnuSysSettings.AppearanceDropDown.BackColor = System.Drawing.Color.FromArgb(CType(CType(6, Byte), Integer), CType(CType(31, Byte), Integer), CType(CType(71, Byte), Integer))
        Me.mnuSysSettings.AppearanceDropDown.BackColor2 = System.Drawing.Color.FromArgb(CType(CType(6, Byte), Integer), CType(CType(31, Byte), Integer), CType(CType(71, Byte), Integer))
        Me.mnuSysSettings.AppearanceDropDown.BorderColor = System.Drawing.Color.FromArgb(CType(CType(6, Byte), Integer), CType(CType(31, Byte), Integer), CType(CType(71, Byte), Integer))
        Me.mnuSysSettings.AppearanceDropDown.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.mnuSysSettings.AppearanceDropDown.ForeColor = System.Drawing.Color.White
        Me.mnuSysSettings.AppearanceDropDown.Options.UseBackColor = True
        Me.mnuSysSettings.AppearanceDropDown.Options.UseBorderColor = True
        Me.mnuSysSettings.AppearanceDropDown.Options.UseFont = True
        Me.mnuSysSettings.AppearanceDropDown.Options.UseForeColor = True
        Me.mnuSysSettings.AppearanceDropDownHovered.BackColor = System.Drawing.Color.FromArgb(CType(CType(16, Byte), Integer), CType(CType(110, Byte), Integer), CType(CType(190, Byte), Integer))
        Me.mnuSysSettings.AppearanceDropDownHovered.BackColor2 = System.Drawing.Color.FromArgb(CType(CType(16, Byte), Integer), CType(CType(110, Byte), Integer), CType(CType(190, Byte), Integer))
        Me.mnuSysSettings.AppearanceDropDownHovered.BorderColor = System.Drawing.Color.FromArgb(CType(CType(16, Byte), Integer), CType(CType(110, Byte), Integer), CType(CType(190, Byte), Integer))
        Me.mnuSysSettings.AppearanceDropDownHovered.Options.UseBackColor = True
        Me.mnuSysSettings.AppearanceDropDownHovered.Options.UseBorderColor = True
        Me.mnuSysSettings.DropDownArrowStyle = DevExpress.XtraEditors.DropDownArrowStyle.Show
        Me.mnuSysSettings.DropDownControl = Me.pmTools
        Me.mnuSysSettings.ImageOptions.AllowGlyphSkinning = DevExpress.Utils.DefaultBoolean.[True]
        Me.mnuSysSettings.ImageOptions.SvgImage = CType(resources.GetObject("mnuSysSettings.ImageOptions.SvgImage"), DevExpress.Utils.Svg.SvgImage)
        Me.mnuSysSettings.Location = New System.Drawing.Point(677, 10)
        Me.mnuSysSettings.Name = "mnuSysSettings"
        Me.mnuSysSettings.Size = New System.Drawing.Size(47, 32)
        Me.mnuSysSettings.TabIndex = 4
        '
        'pmTools
        '
        Me.pmTools.LinksPersistInfo.AddRange(New DevExpress.XtraBars.LinkPersistInfo() {New DevExpress.XtraBars.LinkPersistInfo(Me.btnTest), New DevExpress.XtraBars.LinkPersistInfo(Me.btnDefaultPrinter), New DevExpress.XtraBars.LinkPersistInfo(Me.btnBackup, True), New DevExpress.XtraBars.LinkPersistInfo(Me.btnRestore), New DevExpress.XtraBars.LinkPersistInfo(Me.btnTruncate), New DevExpress.XtraBars.LinkPersistInfo(Me.btnMachines, True)})
        Me.pmTools.Manager = Me.BarManager
        Me.pmTools.Name = "pmTools"
        '
        'mnuUser
        '
        Me.mnuUser.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.mnuUser.Appearance.BackColor = System.Drawing.Color.FromArgb(CType(CType(16, Byte), Integer), CType(CType(110, Byte), Integer), CType(CType(190, Byte), Integer))
        Me.mnuUser.Appearance.BackColor2 = System.Drawing.Color.FromArgb(CType(CType(16, Byte), Integer), CType(CType(110, Byte), Integer), CType(CType(190, Byte), Integer))
        Me.mnuUser.Appearance.BorderColor = System.Drawing.Color.FromArgb(CType(CType(16, Byte), Integer), CType(CType(110, Byte), Integer), CType(CType(190, Byte), Integer))
        Me.mnuUser.Appearance.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.mnuUser.Appearance.ForeColor = System.Drawing.Color.White
        Me.mnuUser.Appearance.Options.UseBackColor = True
        Me.mnuUser.Appearance.Options.UseBorderColor = True
        Me.mnuUser.Appearance.Options.UseFont = True
        Me.mnuUser.Appearance.Options.UseForeColor = True
        Me.mnuUser.Appearance.Options.UseTextOptions = True
        Me.mnuUser.Appearance.TextOptions.Trimming = DevExpress.Utils.Trimming.EllipsisCharacter
        Me.mnuUser.AppearanceDropDown.BackColor = System.Drawing.Color.FromArgb(CType(CType(6, Byte), Integer), CType(CType(31, Byte), Integer), CType(CType(71, Byte), Integer))
        Me.mnuUser.AppearanceDropDown.BackColor2 = System.Drawing.Color.FromArgb(CType(CType(6, Byte), Integer), CType(CType(31, Byte), Integer), CType(CType(71, Byte), Integer))
        Me.mnuUser.AppearanceDropDown.BorderColor = System.Drawing.Color.FromArgb(CType(CType(6, Byte), Integer), CType(CType(31, Byte), Integer), CType(CType(71, Byte), Integer))
        Me.mnuUser.AppearanceDropDown.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.mnuUser.AppearanceDropDown.ForeColor = System.Drawing.Color.White
        Me.mnuUser.AppearanceDropDown.Options.UseBackColor = True
        Me.mnuUser.AppearanceDropDown.Options.UseBorderColor = True
        Me.mnuUser.AppearanceDropDown.Options.UseFont = True
        Me.mnuUser.AppearanceDropDown.Options.UseForeColor = True
        Me.mnuUser.AppearanceDropDownHovered.BackColor = System.Drawing.Color.FromArgb(CType(CType(16, Byte), Integer), CType(CType(110, Byte), Integer), CType(CType(190, Byte), Integer))
        Me.mnuUser.AppearanceDropDownHovered.BackColor2 = System.Drawing.Color.FromArgb(CType(CType(16, Byte), Integer), CType(CType(110, Byte), Integer), CType(CType(190, Byte), Integer))
        Me.mnuUser.AppearanceDropDownHovered.BorderColor = System.Drawing.Color.FromArgb(CType(CType(16, Byte), Integer), CType(CType(110, Byte), Integer), CType(CType(190, Byte), Integer))
        Me.mnuUser.AppearanceDropDownHovered.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.mnuUser.AppearanceDropDownHovered.ForeColor = System.Drawing.Color.White
        Me.mnuUser.AppearanceDropDownHovered.Options.UseBackColor = True
        Me.mnuUser.AppearanceDropDownHovered.Options.UseBorderColor = True
        Me.mnuUser.AppearanceDropDownHovered.Options.UseFont = True
        Me.mnuUser.AppearanceDropDownHovered.Options.UseForeColor = True
        Me.mnuUser.DropDownArrowStyle = DevExpress.XtraEditors.DropDownArrowStyle.Show
        Me.mnuUser.DropDownControl = Me.pmUser
        Me.mnuUser.ImageOptions.AllowGlyphSkinning = DevExpress.Utils.DefaultBoolean.[True]
        Me.mnuUser.ImageOptions.SvgImage = CType(resources.GetObject("mnuUser.ImageOptions.SvgImage"), DevExpress.Utils.Svg.SvgImage)
        Me.mnuUser.Location = New System.Drawing.Point(730, 10)
        Me.mnuUser.Name = "mnuUser"
        Me.mnuUser.Size = New System.Drawing.Size(197, 32)
        Me.mnuUser.TabIndex = 3
        Me.mnuUser.Text = "Ardison F. Pagulayan, RMT"
        '
        'pmUser
        '
        Me.pmUser.LinksPersistInfo.AddRange(New DevExpress.XtraBars.LinkPersistInfo() {New DevExpress.XtraBars.LinkPersistInfo(Me.btnEditAccount)})
        Me.pmUser.Manager = Me.BarManager
        Me.pmUser.Name = "pmUser"
        '
        'tmCheckCon
        '
        Me.tmCheckCon.Enabled = True
        '
        'PanelControl3
        '
        Me.PanelControl3.Appearance.BackColor = System.Drawing.Color.FromArgb(CType(CType(16, Byte), Integer), CType(CType(110, Byte), Integer), CType(CType(190, Byte), Integer))
        Me.PanelControl3.Appearance.Options.UseBackColor = True
        Me.PanelControl3.BorderStyle = DevExpress.XtraEditors.Controls.BorderStyles.NoBorder
        Me.PanelControl3.Controls.Add(Me.lblDateTime)
        Me.PanelControl3.Controls.Add(Me.lblHospitalName)
        Me.PanelControl3.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.PanelControl3.Location = New System.Drawing.Point(0, 625)
        Me.PanelControl3.LookAndFeel.Style = DevExpress.LookAndFeel.LookAndFeelStyle.UltraFlat
        Me.PanelControl3.LookAndFeel.UseDefaultLookAndFeel = False
        Me.PanelControl3.Name = "PanelControl3"
        Me.PanelControl3.Size = New System.Drawing.Size(1188, 25)
        Me.PanelControl3.TabIndex = 7
        '
        'lblDateTime
        '
        Me.lblDateTime.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.lblDateTime.Appearance.Font = New System.Drawing.Font("Segoe UI Semibold", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblDateTime.Appearance.ForeColor = System.Drawing.Color.White
        Me.lblDateTime.Appearance.Options.UseFont = True
        Me.lblDateTime.Appearance.Options.UseForeColor = True
        Me.lblDateTime.Appearance.Options.UseTextOptions = True
        Me.lblDateTime.Appearance.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Far
        Me.lblDateTime.AutoSizeMode = DevExpress.XtraEditors.LabelAutoSizeMode.None
        Me.lblDateTime.Location = New System.Drawing.Point(840, 2)
        Me.lblDateTime.Name = "lblDateTime"
        Me.lblDateTime.Size = New System.Drawing.Size(340, 21)
        Me.lblDateTime.TabIndex = 1
        '
        'lblHospitalName
        '
        Me.lblHospitalName.Appearance.Font = New System.Drawing.Font("Segoe UI Semibold", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblHospitalName.Appearance.ForeColor = System.Drawing.Color.White
        Me.lblHospitalName.Appearance.Options.UseFont = True
        Me.lblHospitalName.Appearance.Options.UseForeColor = True
        Me.lblHospitalName.Location = New System.Drawing.Point(8, 4)
        Me.lblHospitalName.Name = "lblHospitalName"
        Me.lblHospitalName.Size = New System.Drawing.Size(66, 17)
        Me.lblHospitalName.TabIndex = 0
        Me.lblHospitalName.Text = "Laboratory"
        '
        'TimerNew
        '
        Me.TimerNew.Interval = 1000
        '
        'TimerPending
        '
        Me.TimerPending.Interval = 1000
        '
        'MainFOrm
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackgroundImageLayoutStore = System.Windows.Forms.ImageLayout.Stretch
        Me.BackgroundImageStore = CType(resources.GetObject("$this.BackgroundImageStore"), System.Drawing.Image)
        Me.ClientSize = New System.Drawing.Size(1188, 650)
        Me.Controls.Add(Me.PanelControl1)
        Me.Controls.Add(Me.SidePanel1)
        Me.Controls.Add(Me.PanelControl3)
        Me.Controls.Add(Me.barDockControlLeft)
        Me.Controls.Add(Me.barDockControlRight)
        Me.Controls.Add(Me.barDockControlBottom)
        Me.Controls.Add(Me.barDockControlTop)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.IsMdiContainer = True
        Me.LookAndFeel.SkinName = "The Bezier"
        Me.LookAndFeel.UseDefaultLookAndFeel = False
        Me.Name = "MainFOrm"
        Me.Text = "Laboratory Information System"
        Me.WindowState = System.Windows.Forms.FormWindowState.Maximized
        Me.SidePanel1.ResumeLayout(False)
        CType(Me.AccordionControl, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.PanelControl2, System.ComponentModel.ISupportInitialize).EndInit()
        Me.PanelControl2.ResumeLayout(False)
        CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.SeparatorControl1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.PanelControl1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.PanelControl1.ResumeLayout(False)
        Me.PanelControl1.PerformLayout()
        CType(Me.SeparatorControl2, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.pmLogs, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.BarManager, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.pmSettings, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.pmTools, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.pmUser, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.PanelControl3, System.ComponentModel.ISupportInitialize).EndInit()
        Me.PanelControl3.ResumeLayout(False)
        Me.PanelControl3.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents SidePanel1 As DevExpress.XtraEditors.SidePanel
    Friend WithEvents PanelControl1 As DevExpress.XtraEditors.PanelControl
    Friend WithEvents AccordionControl As DevExpress.XtraBars.Navigation.AccordionControl
    Friend WithEvents btnPatient As DevExpress.XtraBars.Navigation.AccordionControlElement
    Friend WithEvents btnPatientOrder As DevExpress.XtraBars.Navigation.AccordionControlElement
    Friend WithEvents aceSection As DevExpress.XtraBars.Navigation.AccordionControlElement
    Private WithEvents AccordionControlSeparator1 As DevExpress.XtraBars.Navigation.AccordionControlSeparator
    Private WithEvents AccordionControlSeparator2 As DevExpress.XtraBars.Navigation.AccordionControlSeparator
    Friend WithEvents aceHema As DevExpress.XtraBars.Navigation.AccordionControlElement
    Friend WithEvents aceChem As DevExpress.XtraBars.Navigation.AccordionControlElement
    Friend WithEvents aceSerology As DevExpress.XtraBars.Navigation.AccordionControlElement
    Friend WithEvents aceImmuno As DevExpress.XtraBars.Navigation.AccordionControlElement
    Friend WithEvents aceMicrobiology As DevExpress.XtraBars.Navigation.AccordionControlElement
    Friend WithEvents aceHistopath As DevExpress.XtraBars.Navigation.AccordionControlElement
    Friend WithEvents aceUrin As DevExpress.XtraBars.Navigation.AccordionControlElement
    Friend WithEvents aceFecal As DevExpress.XtraBars.Navigation.AccordionControlElement
    Private WithEvents AccordionControlSeparator3 As DevExpress.XtraBars.Navigation.AccordionControlSeparator
    Friend WithEvents btnDelta As DevExpress.XtraBars.Navigation.AccordionControlElement
    Private WithEvents AccordionControlSeparator5 As DevExpress.XtraBars.Navigation.AccordionControlSeparator
    Private WithEvents AccordionControlSeparator6 As DevExpress.XtraBars.Navigation.AccordionControlSeparator
    Friend WithEvents aceReports As DevExpress.XtraBars.Navigation.AccordionControlElement
    Friend WithEvents btnQC As DevExpress.XtraBars.Navigation.AccordionControlElement
    Friend WithEvents btnWorkSheet As DevExpress.XtraBars.Navigation.AccordionControlElement
    Friend WithEvents btnCensus As DevExpress.XtraBars.Navigation.AccordionControlElement
    Friend WithEvents btnCompleteTat As DevExpress.XtraBars.Navigation.AccordionControlElement
    Friend WithEvents mnuUser As DevExpress.XtraEditors.DropDownButton
    Friend WithEvents mnuSysSettings As DevExpress.XtraEditors.DropDownButton
    Friend WithEvents pmSettings As DevExpress.XtraBars.PopupMenu
    Friend WithEvents pmTools As DevExpress.XtraBars.PopupMenu
    Friend WithEvents pmUser As DevExpress.XtraBars.PopupMenu
    Friend WithEvents lblTitleHome As DevExpress.XtraEditors.LabelControl
    Friend WithEvents PanelControl2 As DevExpress.XtraEditors.PanelControl
    Friend WithEvents PictureBox1 As PictureBox
    Friend WithEvents BarManager As DevExpress.XtraBars.BarManager
    Friend WithEvents barDockControlTop As DevExpress.XtraBars.BarDockControl
    Friend WithEvents barDockControlBottom As DevExpress.XtraBars.BarDockControl
    Friend WithEvents barDockControlLeft As DevExpress.XtraBars.BarDockControl
    Friend WithEvents barDockControlRight As DevExpress.XtraBars.BarDockControl
    Friend WithEvents btnSpecimen As DevExpress.XtraBars.BarButtonItem
    Friend WithEvents btnDefaultSpecimen As DevExpress.XtraBars.BarButtonItem
    Friend WithEvents btnTestGroup As DevExpress.XtraBars.BarButtonItem
    Friend WithEvents btnChemRange As DevExpress.XtraBars.BarButtonItem
    Friend WithEvents btnMedTech As DevExpress.XtraBars.BarButtonItem
    Friend WithEvents btnPathologist As DevExpress.XtraBars.BarButtonItem
    Friend WithEvents btnDoctors As DevExpress.XtraBars.BarButtonItem
    Friend WithEvents btnDepartment As DevExpress.XtraBars.BarButtonItem
    Friend WithEvents btnCompany As DevExpress.XtraBars.BarButtonItem
    Friend WithEvents btnShift As DevExpress.XtraBars.BarButtonItem
    Friend WithEvents btnTest As DevExpress.XtraBars.BarButtonItem
    Friend WithEvents btnDefaultPrinter As DevExpress.XtraBars.BarButtonItem
    Friend WithEvents btnPrinterSetting As DevExpress.XtraBars.BarButtonItem
    Friend WithEvents btnBackup As DevExpress.XtraBars.BarButtonItem
    Friend WithEvents btnRestore As DevExpress.XtraBars.BarButtonItem
    Friend WithEvents btnTruncate As DevExpress.XtraBars.BarButtonItem
    Friend WithEvents btnMachines As DevExpress.XtraBars.BarButtonItem
    Friend WithEvents btnAccounts As DevExpress.XtraBars.BarButtonItem
    Friend WithEvents btnEditAccount As DevExpress.XtraBars.BarButtonItem
    Friend WithEvents btnChange As DevExpress.XtraBars.BarButtonItem
    Friend WithEvents btnTestName As DevExpress.XtraBars.BarButtonItem
    Friend WithEvents btnUnit As DevExpress.XtraBars.BarButtonItem
    Friend WithEvents btnSettings As DevExpress.XtraBars.BarButtonItem
    Private WithEvents AccordionControlSeparator7 As DevExpress.XtraBars.Navigation.AccordionControlSeparator
    Friend WithEvents btnRealtime As DevExpress.XtraBars.Navigation.AccordionControlElement
    Friend WithEvents btnSpecimenLog As DevExpress.XtraBars.Navigation.AccordionControlElement
    Friend WithEvents lblTitle As DevExpress.XtraEditors.LabelControl
    Friend WithEvents mnuHelp As DevExpress.XtraEditors.DropDownButton
    Friend WithEvents mnuSetting As DevExpress.XtraEditors.DropDownButton
    Friend WithEvents LabelControl1 As DevExpress.XtraEditors.LabelControl
    Friend WithEvents SeparatorControl1 As DevExpress.XtraEditors.SeparatorControl
    Private WithEvents AccordionControlSeparator9 As DevExpress.XtraBars.Navigation.AccordionControlSeparator
    Friend WithEvents LabelControl3 As DevExpress.XtraEditors.LabelControl
    Friend WithEvents SeparatorControl2 As DevExpress.XtraEditors.SeparatorControl
    Friend WithEvents tmCheckCon As Timer
    Friend WithEvents LabelControl4 As DevExpress.XtraEditors.LabelControl
    Friend WithEvents btnMngUser As DevExpress.XtraBars.BarLargeButtonItem
    Friend WithEvents btnUsers As DevExpress.XtraBars.BarLargeButtonItem
    Friend WithEvents accPhlebotomy As DevExpress.XtraBars.Navigation.AccordionControlElement
    Friend WithEvents PanelControl3 As DevExpress.XtraEditors.PanelControl
    Friend WithEvents lblHospitalName As DevExpress.XtraEditors.LabelControl
    Friend WithEvents pmLogs As DevExpress.XtraBars.PopupMenu
    Friend WithEvents btnSpecimenLogs As DevExpress.XtraBars.BarLargeButtonItem
    Friend WithEvents acciHOMIS As DevExpress.XtraBars.Navigation.AccordionControlElement
    Friend WithEvents accEmailPush As DevExpress.XtraBars.Navigation.AccordionControlElement
    Friend WithEvents btnLogout As DevExpress.XtraEditors.SimpleButton
    Friend WithEvents lblDateTime As DevExpress.XtraEditors.LabelControl
    Friend WithEvents btnActivityLog As DevExpress.XtraBars.BarLargeButtonItem
    Friend WithEvents aceMolecular As DevExpress.XtraBars.Navigation.AccordionControlElement
    Friend WithEvents lblNew As DevExpress.XtraEditors.LabelControl
    Friend WithEvents btnEmail As DevExpress.XtraBars.BarButtonItem
    Friend WithEvents TimerNew As Timer
    Friend WithEvents TimerPending As Timer
    Friend WithEvents lblPending As DevExpress.XtraEditors.LabelControl
End Class
