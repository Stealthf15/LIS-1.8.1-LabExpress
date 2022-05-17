<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class frmFecalNew
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmFecalNew))
        Me.cboVerify = New DevExpress.XtraEditors.ComboBoxEdit()
        Me.cboPathologist = New DevExpress.XtraEditors.ComboBoxEdit()
        Me.cboMedTech = New DevExpress.XtraEditors.ComboBoxEdit()
        Me.ImageList = New System.Windows.Forms.ImageList(Me.components)
        Me.txtAge = New DevExpress.XtraEditors.TextEdit()
        Me.txtSampleID = New DevExpress.XtraEditors.TextEdit()
        Me.txtName = New DevExpress.XtraEditors.TextEdit()
        Me.txtPatientID = New DevExpress.XtraEditors.TextEdit()
        Me.cboSex = New DevExpress.XtraEditors.ComboBoxEdit()
        Me.BarManager = New DevExpress.XtraBars.BarManager(Me.components)
        Me.Bar2 = New DevExpress.XtraBars.Bar()
        Me.btnViewPrint = New DevExpress.XtraBars.BarLargeButtonItem()
        Me.btnPrint = New DevExpress.XtraBars.BarLargeButtonItem()
        Me.btnValidate = New DevExpress.XtraBars.BarLargeButtonItem()
        Me.btnPrintNow = New DevExpress.XtraBars.BarLargeButtonItem()
        Me.btnClose = New DevExpress.XtraBars.BarLargeButtonItem()
        Me.btnRelease = New DevExpress.XtraBars.BarLargeButtonItem()
        Me.btnVerification = New DevExpress.XtraBars.BarLargeButtonItem()
        Me.btnEmailPush = New DevExpress.XtraBars.BarLargeButtonItem()
        Me.BarAndDockingController = New DevExpress.XtraBars.BarAndDockingController(Me.components)
        Me.barDockControlTop = New DevExpress.XtraBars.BarDockControl()
        Me.barDockControlBottom = New DevExpress.XtraBars.BarDockControl()
        Me.barDockControlLeft = New DevExpress.XtraBars.BarDockControl()
        Me.barDockControlRight = New DevExpress.XtraBars.BarDockControl()
        Me.BarLargeButtonItem3 = New DevExpress.XtraBars.BarLargeButtonItem()
        Me.dtBDate = New DevExpress.XtraEditors.DateEdit()
        Me.txtClass = New DevExpress.XtraEditors.ComboBoxEdit()
        Me.txtComment = New DevExpress.XtraEditors.MemoEdit()
        Me.txtRemarks = New DevExpress.XtraEditors.MemoEdit()
        Me.cboPatientType = New DevExpress.XtraEditors.ComboBoxEdit()
        Me.cboRoom = New DevExpress.XtraEditors.ComboBoxEdit()
        Me.cboPhysician = New DevExpress.XtraEditors.ComboBoxEdit()
        Me.cboRequest = New DevExpress.XtraEditors.ComboBoxEdit()
        Me.txtORNo = New DevExpress.XtraEditors.TextEdit()
        Me.txtAccession = New DevExpress.XtraEditors.TextEdit()
        Me.txtContact = New DevExpress.XtraEditors.TextEdit()
        Me.cboCS = New DevExpress.XtraEditors.ComboBoxEdit()
        Me.btnSearch = New DevExpress.XtraEditors.SimpleButton()
        Me.tmTimeReceived = New System.Windows.Forms.DateTimePicker()
        Me.dtReceived = New System.Windows.Forms.DateTimePicker()
        Me.GroupControl3 = New DevExpress.XtraEditors.GroupControl()
        Me.btnRetrive = New DevExpress.XtraEditors.SimpleButton()
        Me.btnPreview = New DevExpress.XtraEditors.SimpleButton()
        Me.btnAddTest = New DevExpress.XtraEditors.SimpleButton()
        Me.dtResult = New DevExpress.XtraGrid.GridControl()
        Me.GridView = New DevExpress.XtraGrid.Views.Grid.GridView()
        Me.LabelControl26 = New DevExpress.XtraEditors.LabelControl()
        Me.LabelControl27 = New DevExpress.XtraEditors.LabelControl()
        Me.LabelControl29 = New DevExpress.XtraEditors.LabelControl()
        Me.GroupControl2 = New DevExpress.XtraEditors.GroupControl()
        Me.tmTimeReleased = New System.Windows.Forms.DateTimePicker()
        Me.LabelControl2 = New DevExpress.XtraEditors.LabelControl()
        Me.txtChargeSlip = New DevExpress.XtraEditors.TextEdit()
        Me.LabelControl1 = New DevExpress.XtraEditors.LabelControl()
        Me.LabelControl32 = New DevExpress.XtraEditors.LabelControl()
        Me.LabelControl33 = New DevExpress.XtraEditors.LabelControl()
        Me.LabelControl34 = New DevExpress.XtraEditors.LabelControl()
        Me.LabelControl35 = New DevExpress.XtraEditors.LabelControl()
        Me.LabelControl36 = New DevExpress.XtraEditors.LabelControl()
        Me.LabelControl37 = New DevExpress.XtraEditors.LabelControl()
        Me.LabelControl38 = New DevExpress.XtraEditors.LabelControl()
        Me.LabelControl41 = New DevExpress.XtraEditors.LabelControl()
        Me.GroupControl1 = New DevExpress.XtraEditors.GroupControl()
        Me.txtEmail = New DevExpress.XtraEditors.TextEdit()
        Me.LabelControl42 = New DevExpress.XtraEditors.LabelControl()
        Me.LabelControl43 = New DevExpress.XtraEditors.LabelControl()
        Me.LabelControl44 = New DevExpress.XtraEditors.LabelControl()
        Me.LabelControl45 = New DevExpress.XtraEditors.LabelControl()
        Me.LabelControl46 = New DevExpress.XtraEditors.LabelControl()
        Me.LabelControl47 = New DevExpress.XtraEditors.LabelControl()
        Me.LabelControl48 = New DevExpress.XtraEditors.LabelControl()
        Me.LabelControl49 = New DevExpress.XtraEditors.LabelControl()
        Me.LabelControl50 = New DevExpress.XtraEditors.LabelControl()
        Me.txtAddress = New DevExpress.XtraEditors.TextEdit()
        Me.GroupControl4 = New DevExpress.XtraEditors.GroupControl()
        Me.GroupControl5 = New DevExpress.XtraEditors.GroupControl()
        Me.GroupControl6 = New DevExpress.XtraEditors.GroupControl()
        CType(Me.cboVerify.Properties, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.cboPathologist.Properties, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.cboMedTech.Properties, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.txtAge.Properties, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.txtSampleID.Properties, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.txtName.Properties, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.txtPatientID.Properties, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.cboSex.Properties, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.BarManager, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.BarAndDockingController, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.dtBDate.Properties.CalendarTimeProperties, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.dtBDate.Properties, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.txtClass.Properties, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.txtComment.Properties, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.txtRemarks.Properties, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.cboPatientType.Properties, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.cboRoom.Properties, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.cboPhysician.Properties, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.cboRequest.Properties, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.txtORNo.Properties, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.txtAccession.Properties, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.txtContact.Properties, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.cboCS.Properties, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.GroupControl3, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.GroupControl3.SuspendLayout()
        CType(Me.dtResult, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.GridView, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.GroupControl2, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.GroupControl2.SuspendLayout()
        CType(Me.txtChargeSlip.Properties, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.GroupControl1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.GroupControl1.SuspendLayout()
        CType(Me.txtEmail.Properties, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.txtAddress.Properties, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.GroupControl4, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.GroupControl4.SuspendLayout()
        CType(Me.GroupControl5, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.GroupControl5.SuspendLayout()
        CType(Me.GroupControl6, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.GroupControl6.SuspendLayout()
        Me.SuspendLayout()
        '
        'cboVerify
        '
        Me.cboVerify.Location = New System.Drawing.Point(90, 56)
        Me.cboVerify.Name = "cboVerify"
        Me.cboVerify.Properties.Appearance.ForeColor = System.Drawing.Color.FromArgb(CType(CType(27, Byte), Integer), CType(CType(41, Byte), Integer), CType(CType(62, Byte), Integer))
        Me.cboVerify.Properties.Appearance.Options.UseForeColor = True
        Me.cboVerify.Properties.Buttons.AddRange(New DevExpress.XtraEditors.Controls.EditorButton() {New DevExpress.XtraEditors.Controls.EditorButton(DevExpress.XtraEditors.Controls.ButtonPredefines.Combo)})
        Me.cboVerify.Properties.Items.AddRange(New Object() {"Day", "Week", "Month", "Year"})
        Me.cboVerify.Properties.LookAndFeel.SkinName = "Visual Studio 2013 Blue"
        Me.cboVerify.Properties.LookAndFeel.UseDefaultLookAndFeel = False
        Me.cboVerify.Size = New System.Drawing.Size(213, 20)
        Me.cboVerify.TabIndex = 19
        '
        'cboPathologist
        '
        Me.cboPathologist.Location = New System.Drawing.Point(90, 82)
        Me.cboPathologist.Name = "cboPathologist"
        Me.cboPathologist.Properties.Appearance.ForeColor = System.Drawing.Color.FromArgb(CType(CType(27, Byte), Integer), CType(CType(41, Byte), Integer), CType(CType(62, Byte), Integer))
        Me.cboPathologist.Properties.Appearance.Options.UseForeColor = True
        Me.cboPathologist.Properties.Buttons.AddRange(New DevExpress.XtraEditors.Controls.EditorButton() {New DevExpress.XtraEditors.Controls.EditorButton(DevExpress.XtraEditors.Controls.ButtonPredefines.Combo)})
        Me.cboPathologist.Properties.Items.AddRange(New Object() {"Day", "Week", "Month", "Year"})
        Me.cboPathologist.Properties.LookAndFeel.SkinName = "Visual Studio 2013 Blue"
        Me.cboPathologist.Properties.LookAndFeel.UseDefaultLookAndFeel = False
        Me.cboPathologist.Size = New System.Drawing.Size(213, 20)
        Me.cboPathologist.TabIndex = 20
        '
        'cboMedTech
        '
        Me.cboMedTech.Location = New System.Drawing.Point(90, 30)
        Me.cboMedTech.Name = "cboMedTech"
        Me.cboMedTech.Properties.Appearance.ForeColor = System.Drawing.Color.FromArgb(CType(CType(27, Byte), Integer), CType(CType(41, Byte), Integer), CType(CType(62, Byte), Integer))
        Me.cboMedTech.Properties.Appearance.Options.UseForeColor = True
        Me.cboMedTech.Properties.Buttons.AddRange(New DevExpress.XtraEditors.Controls.EditorButton() {New DevExpress.XtraEditors.Controls.EditorButton(DevExpress.XtraEditors.Controls.ButtonPredefines.Combo)})
        Me.cboMedTech.Properties.Items.AddRange(New Object() {"Day", "Week", "Month", "Year"})
        Me.cboMedTech.Properties.LookAndFeel.SkinName = "Visual Studio 2013 Blue"
        Me.cboMedTech.Properties.LookAndFeel.UseDefaultLookAndFeel = False
        Me.cboMedTech.Size = New System.Drawing.Size(213, 20)
        Me.cboMedTech.TabIndex = 18
        '
        'ImageList
        '
        Me.ImageList.ImageStream = CType(resources.GetObject("ImageList.ImageStream"), System.Windows.Forms.ImageListStreamer)
        Me.ImageList.TransparentColor = System.Drawing.Color.Transparent
        Me.ImageList.Images.SetKeyName(0, "icon.png")
        '
        'txtAge
        '
        Me.txtAge.Location = New System.Drawing.Point(90, 147)
        Me.txtAge.Name = "txtAge"
        Me.txtAge.Properties.Appearance.ForeColor = System.Drawing.Color.FromArgb(CType(CType(27, Byte), Integer), CType(CType(41, Byte), Integer), CType(CType(62, Byte), Integer))
        Me.txtAge.Properties.Appearance.Options.UseForeColor = True
        Me.txtAge.Properties.AppearanceReadOnly.BackColor = System.Drawing.Color.White
        Me.txtAge.Properties.AppearanceReadOnly.BackColor2 = System.Drawing.Color.White
        Me.txtAge.Properties.AppearanceReadOnly.BorderColor = System.Drawing.Color.White
        Me.txtAge.Properties.AppearanceReadOnly.Options.UseBackColor = True
        Me.txtAge.Properties.AppearanceReadOnly.Options.UseBorderColor = True
        Me.txtAge.Properties.LookAndFeel.SkinName = "Visual Studio 2013 Blue"
        Me.txtAge.Properties.LookAndFeel.UseDefaultLookAndFeel = False
        Me.txtAge.Size = New System.Drawing.Size(116, 20)
        Me.txtAge.TabIndex = 5
        '
        'txtSampleID
        '
        Me.txtSampleID.Location = New System.Drawing.Point(90, 32)
        Me.txtSampleID.Name = "txtSampleID"
        Me.txtSampleID.Properties.Appearance.ForeColor = System.Drawing.Color.FromArgb(CType(CType(27, Byte), Integer), CType(CType(41, Byte), Integer), CType(CType(62, Byte), Integer))
        Me.txtSampleID.Properties.Appearance.Options.UseForeColor = True
        Me.txtSampleID.Properties.AppearanceReadOnly.BackColor = System.Drawing.Color.White
        Me.txtSampleID.Properties.AppearanceReadOnly.BackColor2 = System.Drawing.Color.White
        Me.txtSampleID.Properties.AppearanceReadOnly.BorderColor = System.Drawing.Color.White
        Me.txtSampleID.Properties.AppearanceReadOnly.Options.UseBackColor = True
        Me.txtSampleID.Properties.AppearanceReadOnly.Options.UseBorderColor = True
        Me.txtSampleID.Properties.LookAndFeel.SkinName = "Visual Studio 2013 Blue"
        Me.txtSampleID.Properties.LookAndFeel.UseDefaultLookAndFeel = False
        Me.txtSampleID.Size = New System.Drawing.Size(213, 20)
        Me.txtSampleID.TabIndex = 0
        '
        'txtName
        '
        Me.txtName.Location = New System.Drawing.Point(90, 78)
        Me.txtName.Name = "txtName"
        Me.txtName.Properties.Appearance.ForeColor = System.Drawing.Color.FromArgb(CType(CType(27, Byte), Integer), CType(CType(41, Byte), Integer), CType(CType(62, Byte), Integer))
        Me.txtName.Properties.Appearance.Options.UseForeColor = True
        Me.txtName.Properties.AppearanceReadOnly.BackColor = System.Drawing.Color.White
        Me.txtName.Properties.AppearanceReadOnly.BackColor2 = System.Drawing.Color.White
        Me.txtName.Properties.AppearanceReadOnly.BorderColor = System.Drawing.Color.White
        Me.txtName.Properties.AppearanceReadOnly.Options.UseBackColor = True
        Me.txtName.Properties.AppearanceReadOnly.Options.UseBorderColor = True
        Me.txtName.Properties.LookAndFeel.SkinName = "Visual Studio 2013 Blue"
        Me.txtName.Properties.LookAndFeel.UseDefaultLookAndFeel = False
        Me.txtName.Size = New System.Drawing.Size(213, 20)
        Me.txtName.TabIndex = 2
        '
        'txtPatientID
        '
        Me.txtPatientID.Location = New System.Drawing.Point(90, 55)
        Me.txtPatientID.Name = "txtPatientID"
        Me.txtPatientID.Properties.Appearance.ForeColor = System.Drawing.Color.FromArgb(CType(CType(27, Byte), Integer), CType(CType(41, Byte), Integer), CType(CType(62, Byte), Integer))
        Me.txtPatientID.Properties.Appearance.Options.UseForeColor = True
        Me.txtPatientID.Properties.AppearanceReadOnly.BackColor = System.Drawing.Color.White
        Me.txtPatientID.Properties.AppearanceReadOnly.BackColor2 = System.Drawing.Color.White
        Me.txtPatientID.Properties.AppearanceReadOnly.BorderColor = System.Drawing.Color.White
        Me.txtPatientID.Properties.AppearanceReadOnly.Options.UseBackColor = True
        Me.txtPatientID.Properties.AppearanceReadOnly.Options.UseBorderColor = True
        Me.txtPatientID.Properties.LookAndFeel.SkinName = "Visual Studio 2013 Blue"
        Me.txtPatientID.Properties.LookAndFeel.UseDefaultLookAndFeel = False
        Me.txtPatientID.Size = New System.Drawing.Size(156, 20)
        Me.txtPatientID.TabIndex = 1
        '
        'cboSex
        '
        Me.cboSex.Location = New System.Drawing.Point(90, 101)
        Me.cboSex.Name = "cboSex"
        Me.cboSex.Properties.Appearance.ForeColor = System.Drawing.Color.FromArgb(CType(CType(27, Byte), Integer), CType(CType(41, Byte), Integer), CType(CType(62, Byte), Integer))
        Me.cboSex.Properties.Appearance.Options.UseForeColor = True
        Me.cboSex.Properties.AppearanceReadOnly.BackColor = System.Drawing.Color.White
        Me.cboSex.Properties.AppearanceReadOnly.BackColor2 = System.Drawing.Color.White
        Me.cboSex.Properties.AppearanceReadOnly.BorderColor = System.Drawing.Color.White
        Me.cboSex.Properties.AppearanceReadOnly.Options.UseBackColor = True
        Me.cboSex.Properties.AppearanceReadOnly.Options.UseBorderColor = True
        Me.cboSex.Properties.Buttons.AddRange(New DevExpress.XtraEditors.Controls.EditorButton() {New DevExpress.XtraEditors.Controls.EditorButton(DevExpress.XtraEditors.Controls.ButtonPredefines.Combo)})
        Me.cboSex.Properties.Items.AddRange(New Object() {"Male", "Female"})
        Me.cboSex.Properties.LookAndFeel.SkinName = "Visual Studio 2013 Blue"
        Me.cboSex.Properties.LookAndFeel.UseDefaultLookAndFeel = False
        Me.cboSex.Size = New System.Drawing.Size(213, 20)
        Me.cboSex.TabIndex = 3
        '
        'BarManager
        '
        Me.BarManager.Bars.AddRange(New DevExpress.XtraBars.Bar() {Me.Bar2})
        Me.BarManager.Controller = Me.BarAndDockingController
        Me.BarManager.DockControls.Add(Me.barDockControlTop)
        Me.BarManager.DockControls.Add(Me.barDockControlBottom)
        Me.BarManager.DockControls.Add(Me.barDockControlLeft)
        Me.BarManager.DockControls.Add(Me.barDockControlRight)
        Me.BarManager.Form = Me
        Me.BarManager.Items.AddRange(New DevExpress.XtraBars.BarItem() {Me.btnPrint, Me.BarLargeButtonItem3, Me.btnClose, Me.btnValidate, Me.btnPrintNow, Me.btnViewPrint, Me.btnRelease, Me.btnVerification, Me.btnEmailPush})
        Me.BarManager.MainMenu = Me.Bar2
        Me.BarManager.MaxItemId = 17
        '
        'Bar2
        '
        Me.Bar2.BarName = "Main menu"
        Me.Bar2.DockCol = 0
        Me.Bar2.DockRow = 0
        Me.Bar2.DockStyle = DevExpress.XtraBars.BarDockStyle.Top
        Me.Bar2.FloatLocation = New System.Drawing.Point(213, 125)
        Me.Bar2.LinksPersistInfo.AddRange(New DevExpress.XtraBars.LinkPersistInfo() {New DevExpress.XtraBars.LinkPersistInfo(Me.btnViewPrint, True), New DevExpress.XtraBars.LinkPersistInfo(Me.btnPrint, True), New DevExpress.XtraBars.LinkPersistInfo(Me.btnValidate, True), New DevExpress.XtraBars.LinkPersistInfo(Me.btnPrintNow, True), New DevExpress.XtraBars.LinkPersistInfo(Me.btnClose, True), New DevExpress.XtraBars.LinkPersistInfo(Me.btnRelease, True), New DevExpress.XtraBars.LinkPersistInfo(Me.btnVerification, True), New DevExpress.XtraBars.LinkPersistInfo(Me.btnEmailPush, True)})
        Me.Bar2.OptionsBar.MultiLine = True
        Me.Bar2.OptionsBar.UseWholeRow = True
        Me.Bar2.Text = "Main menu"
        '
        'btnViewPrint
        '
        Me.btnViewPrint.Caption = "Preview"
        Me.btnViewPrint.CaptionAlignment = DevExpress.XtraBars.BarItemCaptionAlignment.Right
        Me.btnViewPrint.Id = 13
        Me.btnViewPrint.ImageOptions.AllowGlyphSkinning = DevExpress.Utils.DefaultBoolean.[True]
        Me.btnViewPrint.ImageOptions.SvgImage = CType(resources.GetObject("btnViewPrint.ImageOptions.SvgImage"), DevExpress.Utils.Svg.SvgImage)
        Me.btnViewPrint.ImageOptions.SvgImageSize = New System.Drawing.Size(24, 24)
        Me.btnViewPrint.ItemShortcut = New DevExpress.XtraBars.BarShortcut(System.Windows.Forms.Keys.F5)
        Me.btnViewPrint.Name = "btnViewPrint"
        '
        'btnPrint
        '
        Me.btnPrint.Caption = "Print && Release"
        Me.btnPrint.CaptionAlignment = DevExpress.XtraBars.BarItemCaptionAlignment.Right
        Me.btnPrint.Id = 1
        Me.btnPrint.ImageOptions.AllowGlyphSkinning = DevExpress.Utils.DefaultBoolean.[True]
        Me.btnPrint.ImageOptions.SvgImage = CType(resources.GetObject("btnPrint.ImageOptions.SvgImage"), DevExpress.Utils.Svg.SvgImage)
        Me.btnPrint.ImageOptions.SvgImageSize = New System.Drawing.Size(24, 24)
        Me.btnPrint.ItemShortcut = New DevExpress.XtraBars.BarShortcut(System.Windows.Forms.Keys.F3)
        Me.btnPrint.Name = "btnPrint"
        '
        'btnValidate
        '
        Me.btnValidate.Caption = "Save"
        Me.btnValidate.CaptionAlignment = DevExpress.XtraBars.BarItemCaptionAlignment.Right
        Me.btnValidate.Id = 11
        Me.btnValidate.ImageOptions.AllowGlyphSkinning = DevExpress.Utils.DefaultBoolean.[True]
        Me.btnValidate.ImageOptions.SvgImage = CType(resources.GetObject("btnValidate.ImageOptions.SvgImage"), DevExpress.Utils.Svg.SvgImage)
        Me.btnValidate.ImageOptions.SvgImageSize = New System.Drawing.Size(24, 24)
        Me.btnValidate.ItemShortcut = New DevExpress.XtraBars.BarShortcut(System.Windows.Forms.Keys.F8)
        Me.btnValidate.Name = "btnValidate"
        '
        'btnPrintNow
        '
        Me.btnPrintNow.Caption = "Print"
        Me.btnPrintNow.CaptionAlignment = DevExpress.XtraBars.BarItemCaptionAlignment.Right
        Me.btnPrintNow.Id = 12
        Me.btnPrintNow.ImageOptions.AllowGlyphSkinning = DevExpress.Utils.DefaultBoolean.[True]
        Me.btnPrintNow.ImageOptions.SvgImage = CType(resources.GetObject("btnPrintNow.ImageOptions.SvgImage"), DevExpress.Utils.Svg.SvgImage)
        Me.btnPrintNow.ImageOptions.SvgImageSize = New System.Drawing.Size(24, 24)
        Me.btnPrintNow.ItemShortcut = New DevExpress.XtraBars.BarShortcut(System.Windows.Forms.Keys.F4)
        Me.btnPrintNow.Name = "btnPrintNow"
        '
        'btnClose
        '
        Me.btnClose.Alignment = DevExpress.XtraBars.BarItemLinkAlignment.Right
        Me.btnClose.Caption = "&Close"
        Me.btnClose.CaptionAlignment = DevExpress.XtraBars.BarItemCaptionAlignment.Right
        Me.btnClose.Id = 7
        Me.btnClose.ImageOptions.AllowGlyphSkinning = DevExpress.Utils.DefaultBoolean.[True]
        Me.btnClose.ImageOptions.SvgImage = CType(resources.GetObject("btnClose.ImageOptions.SvgImage"), DevExpress.Utils.Svg.SvgImage)
        Me.btnClose.ImageOptions.SvgImageSize = New System.Drawing.Size(24, 24)
        Me.btnClose.Name = "btnClose"
        '
        'btnRelease
        '
        Me.btnRelease.Caption = "Release"
        Me.btnRelease.CaptionAlignment = DevExpress.XtraBars.BarItemCaptionAlignment.Right
        Me.btnRelease.Id = 14
        Me.btnRelease.ImageOptions.SvgImage = CType(resources.GetObject("btnRelease.ImageOptions.SvgImage"), DevExpress.Utils.Svg.SvgImage)
        Me.btnRelease.ImageOptions.SvgImageSize = New System.Drawing.Size(24, 24)
        Me.btnRelease.Name = "btnRelease"
        '
        'btnVerification
        '
        Me.btnVerification.Caption = "For Verification"
        Me.btnVerification.CaptionAlignment = DevExpress.XtraBars.BarItemCaptionAlignment.Right
        Me.btnVerification.Id = 15
        Me.btnVerification.ImageOptions.SvgImage = CType(resources.GetObject("btnVerification.ImageOptions.SvgImage"), DevExpress.Utils.Svg.SvgImage)
        Me.btnVerification.ImageOptions.SvgImageSize = New System.Drawing.Size(24, 24)
        Me.btnVerification.Name = "btnVerification"
        '
        'btnEmailPush
        '
        Me.btnEmailPush.Caption = "Email Push"
        Me.btnEmailPush.CaptionAlignment = DevExpress.XtraBars.BarItemCaptionAlignment.Right
        Me.btnEmailPush.Id = 16
        Me.btnEmailPush.ImageOptions.SvgImage = CType(resources.GetObject("btnEmailPush.ImageOptions.SvgImage"), DevExpress.Utils.Svg.SvgImage)
        Me.btnEmailPush.ImageOptions.SvgImageSize = New System.Drawing.Size(24, 24)
        Me.btnEmailPush.Name = "btnEmailPush"
        '
        'BarAndDockingController
        '
        Me.BarAndDockingController.AppearancesBar.MainMenuAppearance.Normal.ForeColor = System.Drawing.Color.White
        Me.BarAndDockingController.AppearancesBar.MainMenuAppearance.Normal.Options.UseForeColor = True
        Me.BarAndDockingController.PropertiesBar.AllowLinkLighting = False
        '
        'barDockControlTop
        '
        Me.barDockControlTop.CausesValidation = False
        Me.barDockControlTop.Dock = System.Windows.Forms.DockStyle.Top
        Me.barDockControlTop.Location = New System.Drawing.Point(0, 0)
        Me.barDockControlTop.Manager = Me.BarManager
        Me.barDockControlTop.Size = New System.Drawing.Size(1305, 36)
        '
        'barDockControlBottom
        '
        Me.barDockControlBottom.CausesValidation = False
        Me.barDockControlBottom.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.barDockControlBottom.Location = New System.Drawing.Point(0, 705)
        Me.barDockControlBottom.Manager = Me.BarManager
        Me.barDockControlBottom.Size = New System.Drawing.Size(1305, 0)
        '
        'barDockControlLeft
        '
        Me.barDockControlLeft.CausesValidation = False
        Me.barDockControlLeft.Dock = System.Windows.Forms.DockStyle.Left
        Me.barDockControlLeft.Location = New System.Drawing.Point(0, 36)
        Me.barDockControlLeft.Manager = Me.BarManager
        Me.barDockControlLeft.Size = New System.Drawing.Size(0, 669)
        '
        'barDockControlRight
        '
        Me.barDockControlRight.CausesValidation = False
        Me.barDockControlRight.Dock = System.Windows.Forms.DockStyle.Right
        Me.barDockControlRight.Location = New System.Drawing.Point(1305, 36)
        Me.barDockControlRight.Manager = Me.BarManager
        Me.barDockControlRight.Size = New System.Drawing.Size(0, 669)
        '
        'BarLargeButtonItem3
        '
        Me.BarLargeButtonItem3.Caption = "&"
        Me.BarLargeButtonItem3.Id = 2
        Me.BarLargeButtonItem3.Name = "BarLargeButtonItem3"
        '
        'dtBDate
        '
        Me.dtBDate.EditValue = Nothing
        Me.dtBDate.Location = New System.Drawing.Point(90, 124)
        Me.dtBDate.MenuManager = Me.BarManager
        Me.dtBDate.Name = "dtBDate"
        Me.dtBDate.Properties.AppearanceReadOnly.BackColor = System.Drawing.Color.White
        Me.dtBDate.Properties.AppearanceReadOnly.BackColor2 = System.Drawing.Color.White
        Me.dtBDate.Properties.AppearanceReadOnly.Options.UseBackColor = True
        Me.dtBDate.Properties.Buttons.AddRange(New DevExpress.XtraEditors.Controls.EditorButton() {New DevExpress.XtraEditors.Controls.EditorButton(DevExpress.XtraEditors.Controls.ButtonPredefines.Combo)})
        Me.dtBDate.Properties.CalendarTimeProperties.Buttons.AddRange(New DevExpress.XtraEditors.Controls.EditorButton() {New DevExpress.XtraEditors.Controls.EditorButton(DevExpress.XtraEditors.Controls.ButtonPredefines.Combo)})
        Me.dtBDate.Size = New System.Drawing.Size(213, 20)
        Me.dtBDate.TabIndex = 4
        '
        'txtClass
        '
        Me.txtClass.EditValue = "Year(s)"
        Me.txtClass.Location = New System.Drawing.Point(212, 147)
        Me.txtClass.Name = "txtClass"
        Me.txtClass.Properties.Appearance.ForeColor = System.Drawing.Color.FromArgb(CType(CType(27, Byte), Integer), CType(CType(41, Byte), Integer), CType(CType(62, Byte), Integer))
        Me.txtClass.Properties.Appearance.Options.UseForeColor = True
        Me.txtClass.Properties.AppearanceReadOnly.BackColor = System.Drawing.Color.White
        Me.txtClass.Properties.AppearanceReadOnly.BackColor2 = System.Drawing.Color.White
        Me.txtClass.Properties.AppearanceReadOnly.BorderColor = System.Drawing.Color.White
        Me.txtClass.Properties.AppearanceReadOnly.Options.UseBackColor = True
        Me.txtClass.Properties.AppearanceReadOnly.Options.UseBorderColor = True
        Me.txtClass.Properties.Buttons.AddRange(New DevExpress.XtraEditors.Controls.EditorButton() {New DevExpress.XtraEditors.Controls.EditorButton(DevExpress.XtraEditors.Controls.ButtonPredefines.Combo)})
        Me.txtClass.Properties.Items.AddRange(New Object() {"Day(s)", "Week(s)", "Month(s)", "Year(s)", "NB"})
        Me.txtClass.Properties.LookAndFeel.SkinName = "Visual Studio 2013 Blue"
        Me.txtClass.Properties.LookAndFeel.UseDefaultLookAndFeel = False
        Me.txtClass.Properties.TextEditStyle = DevExpress.XtraEditors.Controls.TextEditStyles.DisableTextEditor
        Me.txtClass.Size = New System.Drawing.Size(91, 20)
        Me.txtClass.TabIndex = 6
        '
        'txtComment
        '
        Me.txtComment.Dock = System.Windows.Forms.DockStyle.Fill
        Me.txtComment.EditValue = ""
        Me.txtComment.Location = New System.Drawing.Point(2, 27)
        Me.txtComment.Name = "txtComment"
        Me.txtComment.Properties.Appearance.ForeColor = System.Drawing.Color.FromArgb(CType(CType(27, Byte), Integer), CType(CType(41, Byte), Integer), CType(CType(62, Byte), Integer))
        Me.txtComment.Properties.Appearance.Options.UseForeColor = True
        Me.txtComment.Properties.BorderStyle = DevExpress.XtraEditors.Controls.BorderStyles.NoBorder
        Me.txtComment.Size = New System.Drawing.Size(462, 79)
        Me.txtComment.TabIndex = 22
        '
        'txtRemarks
        '
        Me.txtRemarks.Dock = System.Windows.Forms.DockStyle.Fill
        Me.txtRemarks.Location = New System.Drawing.Point(2, 27)
        Me.txtRemarks.Name = "txtRemarks"
        Me.txtRemarks.Properties.Appearance.ForeColor = System.Drawing.Color.FromArgb(CType(CType(27, Byte), Integer), CType(CType(41, Byte), Integer), CType(CType(62, Byte), Integer))
        Me.txtRemarks.Properties.Appearance.Options.UseForeColor = True
        Me.txtRemarks.Properties.BorderStyle = DevExpress.XtraEditors.Controls.BorderStyles.NoBorder
        Me.txtRemarks.Size = New System.Drawing.Size(472, 79)
        Me.txtRemarks.TabIndex = 21
        '
        'cboPatientType
        '
        Me.cboPatientType.Location = New System.Drawing.Point(90, 243)
        Me.cboPatientType.Name = "cboPatientType"
        Me.cboPatientType.Properties.Appearance.ForeColor = System.Drawing.Color.FromArgb(CType(CType(27, Byte), Integer), CType(CType(41, Byte), Integer), CType(CType(62, Byte), Integer))
        Me.cboPatientType.Properties.Appearance.Options.UseForeColor = True
        Me.cboPatientType.Properties.Buttons.AddRange(New DevExpress.XtraEditors.Controls.EditorButton() {New DevExpress.XtraEditors.Controls.EditorButton(DevExpress.XtraEditors.Controls.ButtonPredefines.Combo)})
        Me.cboPatientType.Properties.Items.AddRange(New Object() {"In Patient", "Out Patient"})
        Me.cboPatientType.Properties.LookAndFeel.SkinName = "Visual Studio 2013 Blue"
        Me.cboPatientType.Properties.LookAndFeel.UseDefaultLookAndFeel = False
        Me.cboPatientType.Size = New System.Drawing.Size(213, 20)
        Me.cboPatientType.TabIndex = 17
        '
        'cboRoom
        '
        Me.cboRoom.Location = New System.Drawing.Point(90, 220)
        Me.cboRoom.Name = "cboRoom"
        Me.cboRoom.Properties.Appearance.ForeColor = System.Drawing.Color.FromArgb(CType(CType(27, Byte), Integer), CType(CType(41, Byte), Integer), CType(CType(62, Byte), Integer))
        Me.cboRoom.Properties.Appearance.Options.UseForeColor = True
        Me.cboRoom.Properties.Buttons.AddRange(New DevExpress.XtraEditors.Controls.EditorButton() {New DevExpress.XtraEditors.Controls.EditorButton(DevExpress.XtraEditors.Controls.ButtonPredefines.Combo)})
        Me.cboRoom.Properties.LookAndFeel.SkinName = "Visual Studio 2013 Blue"
        Me.cboRoom.Properties.LookAndFeel.UseDefaultLookAndFeel = False
        Me.cboRoom.Size = New System.Drawing.Size(213, 20)
        Me.cboRoom.TabIndex = 16
        '
        'cboPhysician
        '
        Me.cboPhysician.Location = New System.Drawing.Point(90, 197)
        Me.cboPhysician.Name = "cboPhysician"
        Me.cboPhysician.Properties.Appearance.ForeColor = System.Drawing.Color.FromArgb(CType(CType(27, Byte), Integer), CType(CType(41, Byte), Integer), CType(CType(62, Byte), Integer))
        Me.cboPhysician.Properties.Appearance.Options.UseForeColor = True
        Me.cboPhysician.Properties.Buttons.AddRange(New DevExpress.XtraEditors.Controls.EditorButton() {New DevExpress.XtraEditors.Controls.EditorButton(DevExpress.XtraEditors.Controls.ButtonPredefines.Combo)})
        Me.cboPhysician.Properties.LookAndFeel.SkinName = "Visual Studio 2013 Blue"
        Me.cboPhysician.Properties.LookAndFeel.UseDefaultLookAndFeel = False
        Me.cboPhysician.Size = New System.Drawing.Size(213, 20)
        Me.cboPhysician.TabIndex = 15
        '
        'cboRequest
        '
        Me.cboRequest.Location = New System.Drawing.Point(90, 174)
        Me.cboRequest.Name = "cboRequest"
        Me.cboRequest.Properties.Appearance.ForeColor = System.Drawing.Color.FromArgb(CType(CType(27, Byte), Integer), CType(CType(41, Byte), Integer), CType(CType(62, Byte), Integer))
        Me.cboRequest.Properties.Appearance.Options.UseForeColor = True
        Me.cboRequest.Properties.Buttons.AddRange(New DevExpress.XtraEditors.Controls.EditorButton() {New DevExpress.XtraEditors.Controls.EditorButton(DevExpress.XtraEditors.Controls.ButtonPredefines.Combo)})
        Me.cboRequest.Properties.LookAndFeel.SkinName = "Visual Studio 2013 Blue"
        Me.cboRequest.Properties.LookAndFeel.UseDefaultLookAndFeel = False
        Me.cboRequest.Size = New System.Drawing.Size(213, 20)
        Me.cboRequest.TabIndex = 14
        '
        'txtORNo
        '
        Me.txtORNo.Location = New System.Drawing.Point(90, 76)
        Me.txtORNo.Name = "txtORNo"
        Me.txtORNo.Properties.Appearance.ForeColor = System.Drawing.Color.FromArgb(CType(CType(27, Byte), Integer), CType(CType(41, Byte), Integer), CType(CType(62, Byte), Integer))
        Me.txtORNo.Properties.Appearance.Options.UseForeColor = True
        Me.txtORNo.Properties.LookAndFeel.SkinName = "Visual Studio 2013 Blue"
        Me.txtORNo.Properties.LookAndFeel.UseDefaultLookAndFeel = False
        Me.txtORNo.Size = New System.Drawing.Size(213, 20)
        Me.txtORNo.TabIndex = 11
        '
        'txtAccession
        '
        Me.txtAccession.Location = New System.Drawing.Point(90, 30)
        Me.txtAccession.Name = "txtAccession"
        Me.txtAccession.Properties.Appearance.ForeColor = System.Drawing.Color.FromArgb(CType(CType(27, Byte), Integer), CType(CType(41, Byte), Integer), CType(CType(62, Byte), Integer))
        Me.txtAccession.Properties.Appearance.Options.UseForeColor = True
        Me.txtAccession.Properties.LookAndFeel.SkinName = "Visual Studio 2013 Blue"
        Me.txtAccession.Properties.LookAndFeel.UseDefaultLookAndFeel = False
        Me.txtAccession.Size = New System.Drawing.Size(213, 20)
        Me.txtAccession.TabIndex = 10
        '
        'txtContact
        '
        Me.txtContact.Location = New System.Drawing.Point(90, 216)
        Me.txtContact.Name = "txtContact"
        Me.txtContact.Properties.Appearance.ForeColor = System.Drawing.Color.FromArgb(CType(CType(27, Byte), Integer), CType(CType(41, Byte), Integer), CType(CType(62, Byte), Integer))
        Me.txtContact.Properties.Appearance.Options.UseForeColor = True
        Me.txtContact.Properties.AppearanceReadOnly.BackColor = System.Drawing.Color.White
        Me.txtContact.Properties.AppearanceReadOnly.BackColor2 = System.Drawing.Color.White
        Me.txtContact.Properties.AppearanceReadOnly.BorderColor = System.Drawing.Color.White
        Me.txtContact.Properties.AppearanceReadOnly.Options.UseBackColor = True
        Me.txtContact.Properties.AppearanceReadOnly.Options.UseBorderColor = True
        Me.txtContact.Properties.LookAndFeel.SkinName = "Visual Studio 2013 Blue"
        Me.txtContact.Properties.LookAndFeel.UseDefaultLookAndFeel = False
        Me.txtContact.Size = New System.Drawing.Size(213, 20)
        Me.txtContact.TabIndex = 9
        '
        'cboCS
        '
        Me.cboCS.Location = New System.Drawing.Point(90, 170)
        Me.cboCS.Name = "cboCS"
        Me.cboCS.Properties.Appearance.ForeColor = System.Drawing.Color.FromArgb(CType(CType(27, Byte), Integer), CType(CType(41, Byte), Integer), CType(CType(62, Byte), Integer))
        Me.cboCS.Properties.Appearance.Options.UseForeColor = True
        Me.cboCS.Properties.AppearanceReadOnly.BackColor = System.Drawing.Color.White
        Me.cboCS.Properties.AppearanceReadOnly.BackColor2 = System.Drawing.Color.White
        Me.cboCS.Properties.AppearanceReadOnly.BorderColor = System.Drawing.Color.White
        Me.cboCS.Properties.AppearanceReadOnly.Options.UseBackColor = True
        Me.cboCS.Properties.AppearanceReadOnly.Options.UseBorderColor = True
        Me.cboCS.Properties.Buttons.AddRange(New DevExpress.XtraEditors.Controls.EditorButton() {New DevExpress.XtraEditors.Controls.EditorButton(DevExpress.XtraEditors.Controls.ButtonPredefines.Combo)})
        Me.cboCS.Properties.Items.AddRange(New Object() {"Single", "Married", "Widow", "Widower"})
        Me.cboCS.Properties.LookAndFeel.SkinName = "Visual Studio 2013 Blue"
        Me.cboCS.Properties.LookAndFeel.UseDefaultLookAndFeel = False
        Me.cboCS.Size = New System.Drawing.Size(213, 20)
        Me.cboCS.TabIndex = 7
        '
        'btnSearch
        '
        Me.btnSearch.Appearance.Font = New System.Drawing.Font("Segoe UI", 9.75!)
        Me.btnSearch.Appearance.Options.UseFont = True
        Me.btnSearch.Cursor = System.Windows.Forms.Cursors.Hand
        Me.btnSearch.ImageOptions.Location = DevExpress.XtraEditors.ImageLocation.MiddleCenter
        Me.btnSearch.ImageOptions.SvgImage = CType(resources.GetObject("btnSearch.ImageOptions.SvgImage"), DevExpress.Utils.Svg.SvgImage)
        Me.btnSearch.ImageOptions.SvgImageSize = New System.Drawing.Size(16, 16)
        Me.btnSearch.Location = New System.Drawing.Point(251, 55)
        Me.btnSearch.Margin = New System.Windows.Forms.Padding(2, 3, 2, 3)
        Me.btnSearch.Name = "btnSearch"
        Me.btnSearch.Size = New System.Drawing.Size(52, 20)
        Me.btnSearch.TabIndex = 160
        '
        'tmTimeReceived
        '
        Me.tmTimeReceived.CalendarFont = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.tmTimeReceived.Format = System.Windows.Forms.DateTimePickerFormat.Time
        Me.tmTimeReceived.Location = New System.Drawing.Point(90, 124)
        Me.tmTimeReceived.Name = "tmTimeReceived"
        Me.tmTimeReceived.Size = New System.Drawing.Size(213, 22)
        Me.tmTimeReceived.TabIndex = 13
        '
        'dtReceived
        '
        Me.dtReceived.CalendarFont = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.dtReceived.Format = System.Windows.Forms.DateTimePickerFormat.[Short]
        Me.dtReceived.Location = New System.Drawing.Point(90, 99)
        Me.dtReceived.Name = "dtReceived"
        Me.dtReceived.Size = New System.Drawing.Size(213, 22)
        Me.dtReceived.TabIndex = 12
        '
        'GroupControl3
        '
        Me.GroupControl3.AppearanceCaption.BackColor = System.Drawing.Color.FromArgb(CType(CType(16, Byte), Integer), CType(CType(110, Byte), Integer), CType(CType(190, Byte), Integer))
        Me.GroupControl3.AppearanceCaption.BackColor2 = System.Drawing.Color.FromArgb(CType(CType(16, Byte), Integer), CType(CType(110, Byte), Integer), CType(CType(190, Byte), Integer))
        Me.GroupControl3.AppearanceCaption.BorderColor = System.Drawing.Color.FromArgb(CType(CType(16, Byte), Integer), CType(CType(110, Byte), Integer), CType(CType(190, Byte), Integer))
        Me.GroupControl3.AppearanceCaption.ForeColor = System.Drawing.Color.White
        Me.GroupControl3.AppearanceCaption.Options.UseBackColor = True
        Me.GroupControl3.AppearanceCaption.Options.UseBorderColor = True
        Me.GroupControl3.AppearanceCaption.Options.UseForeColor = True
        Me.GroupControl3.CaptionImageOptions.AllowGlyphSkinning = True
        Me.GroupControl3.CaptionImageOptions.SvgImage = CType(resources.GetObject("GroupControl3.CaptionImageOptions.SvgImage"), DevExpress.Utils.Svg.SvgImage)
        Me.GroupControl3.CaptionImageOptions.SvgImageSize = New System.Drawing.Size(16, 16)
        Me.GroupControl3.Controls.Add(Me.btnRetrive)
        Me.GroupControl3.Controls.Add(Me.btnPreview)
        Me.GroupControl3.Controls.Add(Me.btnAddTest)
        Me.GroupControl3.Controls.Add(Me.dtResult)
        Me.GroupControl3.Location = New System.Drawing.Point(328, 42)
        Me.GroupControl3.LookAndFeel.SkinName = "The Bezier"
        Me.GroupControl3.LookAndFeel.UseDefaultLookAndFeel = False
        Me.GroupControl3.Name = "GroupControl3"
        Me.GroupControl3.Size = New System.Drawing.Size(948, 519)
        Me.GroupControl3.TabIndex = 175
        Me.GroupControl3.Text = "Test Result"
        '
        'btnRetrive
        '
        Me.btnRetrive.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnRetrive.ImageOptions.SvgImage = CType(resources.GetObject("btnRetrive.ImageOptions.SvgImage"), DevExpress.Utils.Svg.SvgImage)
        Me.btnRetrive.ImageOptions.SvgImageSize = New System.Drawing.Size(20, 20)
        Me.btnRetrive.Location = New System.Drawing.Point(556, 1)
        Me.btnRetrive.Name = "btnRetrive"
        Me.btnRetrive.Size = New System.Drawing.Size(135, 23)
        Me.btnRetrive.TabIndex = 62
        Me.btnRetrive.Text = "Retrieve Re-run"
        '
        'btnPreview
        '
        Me.btnPreview.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnPreview.ImageOptions.SvgImage = CType(resources.GetObject("btnPreview.ImageOptions.SvgImage"), DevExpress.Utils.Svg.SvgImage)
        Me.btnPreview.ImageOptions.SvgImageSize = New System.Drawing.Size(20, 20)
        Me.btnPreview.Location = New System.Drawing.Point(697, 1)
        Me.btnPreview.Name = "btnPreview"
        Me.btnPreview.Size = New System.Drawing.Size(135, 23)
        Me.btnPreview.TabIndex = 61
        Me.btnPreview.Text = "Show Delta Check"
        '
        'btnAddTest
        '
        Me.btnAddTest.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnAddTest.ImageOptions.SvgImage = CType(resources.GetObject("btnAddTest.ImageOptions.SvgImage"), DevExpress.Utils.Svg.SvgImage)
        Me.btnAddTest.ImageOptions.SvgImageSize = New System.Drawing.Size(20, 20)
        Me.btnAddTest.Location = New System.Drawing.Point(838, 1)
        Me.btnAddTest.Name = "btnAddTest"
        Me.btnAddTest.Size = New System.Drawing.Size(108, 23)
        Me.btnAddTest.TabIndex = 60
        Me.btnAddTest.Text = "Add Test"
        '
        'dtResult
        '
        Me.dtResult.Dock = System.Windows.Forms.DockStyle.Fill
        Me.dtResult.Location = New System.Drawing.Point(2, 27)
        Me.dtResult.MainView = Me.GridView
        Me.dtResult.MenuManager = Me.BarManager
        Me.dtResult.Name = "dtResult"
        Me.dtResult.Size = New System.Drawing.Size(944, 490)
        Me.dtResult.TabIndex = 58
        Me.dtResult.ViewCollection.AddRange(New DevExpress.XtraGrid.Views.Base.BaseView() {Me.GridView})
        '
        'GridView
        '
        Me.GridView.Appearance.SelectedRow.BackColor = System.Drawing.Color.WhiteSmoke
        Me.GridView.Appearance.SelectedRow.BackColor2 = System.Drawing.Color.WhiteSmoke
        Me.GridView.Appearance.SelectedRow.BorderColor = System.Drawing.Color.WhiteSmoke
        Me.GridView.Appearance.SelectedRow.ForeColor = System.Drawing.Color.FromArgb(CType(CType(38, Byte), Integer), CType(CType(38, Byte), Integer), CType(CType(38, Byte), Integer))
        Me.GridView.Appearance.SelectedRow.Options.UseBackColor = True
        Me.GridView.Appearance.SelectedRow.Options.UseBorderColor = True
        Me.GridView.Appearance.SelectedRow.Options.UseForeColor = True
        Me.GridView.GridControl = Me.dtResult
        Me.GridView.Name = "GridView"
        Me.GridView.OptionsBehavior.AutoExpandAllGroups = True
        Me.GridView.OptionsCustomization.AllowColumnMoving = False
        Me.GridView.OptionsCustomization.AllowColumnResizing = False
        Me.GridView.OptionsCustomization.AllowFilter = False
        Me.GridView.OptionsCustomization.AllowGroup = False
        Me.GridView.OptionsCustomization.AllowQuickHideColumns = False
        Me.GridView.OptionsCustomization.AllowSort = False
        Me.GridView.OptionsSelection.CheckBoxSelectorColumnWidth = 30
        Me.GridView.OptionsSelection.MultiSelect = True
        Me.GridView.OptionsSelection.MultiSelectMode = DevExpress.XtraGrid.Views.Grid.GridMultiSelectMode.CheckBoxRowSelect
        Me.GridView.OptionsSelection.ShowCheckBoxSelectorInColumnHeader = DevExpress.Utils.DefaultBoolean.[True]
        Me.GridView.OptionsView.ShowGroupPanel = False
        Me.GridView.OptionsView.ShowVerticalLines = DevExpress.Utils.DefaultBoolean.[False]
        '
        'LabelControl26
        '
        Me.LabelControl26.Appearance.ForeColor = System.Drawing.Color.FromArgb(CType(CType(27, Byte), Integer), CType(CType(41, Byte), Integer), CType(CType(62, Byte), Integer))
        Me.LabelControl26.Appearance.Options.UseForeColor = True
        Me.LabelControl26.Location = New System.Drawing.Point(8, 85)
        Me.LabelControl26.LookAndFeel.SkinName = "Visual Studio 2013 Blue"
        Me.LabelControl26.LookAndFeel.UseDefaultLookAndFeel = False
        Me.LabelControl26.Name = "LabelControl26"
        Me.LabelControl26.Size = New System.Drawing.Size(62, 13)
        Me.LabelControl26.TabIndex = 50
        Me.LabelControl26.Text = "Pathologist:"
        '
        'LabelControl27
        '
        Me.LabelControl27.Appearance.ForeColor = System.Drawing.Color.FromArgb(CType(CType(27, Byte), Integer), CType(CType(41, Byte), Integer), CType(CType(62, Byte), Integer))
        Me.LabelControl27.Appearance.Options.UseForeColor = True
        Me.LabelControl27.Location = New System.Drawing.Point(8, 59)
        Me.LabelControl27.LookAndFeel.SkinName = "Visual Studio 2013 Blue"
        Me.LabelControl27.LookAndFeel.UseDefaultLookAndFeel = False
        Me.LabelControl27.Name = "LabelControl27"
        Me.LabelControl27.Size = New System.Drawing.Size(57, 13)
        Me.LabelControl27.TabIndex = 87
        Me.LabelControl27.Text = "Verified By:"
        '
        'LabelControl29
        '
        Me.LabelControl29.Appearance.ForeColor = System.Drawing.Color.FromArgb(CType(CType(27, Byte), Integer), CType(CType(41, Byte), Integer), CType(CType(62, Byte), Integer))
        Me.LabelControl29.Appearance.Options.UseForeColor = True
        Me.LabelControl29.Location = New System.Drawing.Point(8, 33)
        Me.LabelControl29.LookAndFeel.SkinName = "Visual Studio 2013 Blue"
        Me.LabelControl29.LookAndFeel.UseDefaultLookAndFeel = False
        Me.LabelControl29.Name = "LabelControl29"
        Me.LabelControl29.Size = New System.Drawing.Size(70, 13)
        Me.LabelControl29.TabIndex = 48
        Me.LabelControl29.Text = "Performed By:"
        '
        'GroupControl2
        '
        Me.GroupControl2.AppearanceCaption.BackColor = System.Drawing.Color.FromArgb(CType(CType(16, Byte), Integer), CType(CType(110, Byte), Integer), CType(CType(190, Byte), Integer))
        Me.GroupControl2.AppearanceCaption.BackColor2 = System.Drawing.Color.FromArgb(CType(CType(16, Byte), Integer), CType(CType(110, Byte), Integer), CType(CType(190, Byte), Integer))
        Me.GroupControl2.AppearanceCaption.BorderColor = System.Drawing.Color.FromArgb(CType(CType(16, Byte), Integer), CType(CType(110, Byte), Integer), CType(CType(190, Byte), Integer))
        Me.GroupControl2.AppearanceCaption.ForeColor = System.Drawing.Color.White
        Me.GroupControl2.AppearanceCaption.Options.UseBackColor = True
        Me.GroupControl2.AppearanceCaption.Options.UseBorderColor = True
        Me.GroupControl2.AppearanceCaption.Options.UseForeColor = True
        Me.GroupControl2.CaptionImageOptions.AllowGlyphSkinning = True
        Me.GroupControl2.CaptionImageOptions.SvgImage = CType(resources.GetObject("GroupControl2.CaptionImageOptions.SvgImage"), DevExpress.Utils.Svg.SvgImage)
        Me.GroupControl2.CaptionImageOptions.SvgImageSize = New System.Drawing.Size(16, 16)
        Me.GroupControl2.Controls.Add(Me.tmTimeReleased)
        Me.GroupControl2.Controls.Add(Me.LabelControl2)
        Me.GroupControl2.Controls.Add(Me.txtChargeSlip)
        Me.GroupControl2.Controls.Add(Me.LabelControl1)
        Me.GroupControl2.Controls.Add(Me.tmTimeReceived)
        Me.GroupControl2.Controls.Add(Me.LabelControl32)
        Me.GroupControl2.Controls.Add(Me.dtReceived)
        Me.GroupControl2.Controls.Add(Me.LabelControl33)
        Me.GroupControl2.Controls.Add(Me.txtAccession)
        Me.GroupControl2.Controls.Add(Me.LabelControl34)
        Me.GroupControl2.Controls.Add(Me.cboPatientType)
        Me.GroupControl2.Controls.Add(Me.LabelControl35)
        Me.GroupControl2.Controls.Add(Me.txtORNo)
        Me.GroupControl2.Controls.Add(Me.LabelControl36)
        Me.GroupControl2.Controls.Add(Me.LabelControl37)
        Me.GroupControl2.Controls.Add(Me.LabelControl38)
        Me.GroupControl2.Controls.Add(Me.cboRoom)
        Me.GroupControl2.Controls.Add(Me.LabelControl41)
        Me.GroupControl2.Controls.Add(Me.cboPhysician)
        Me.GroupControl2.Controls.Add(Me.cboRequest)
        Me.GroupControl2.Location = New System.Drawing.Point(10, 291)
        Me.GroupControl2.LookAndFeel.SkinName = "The Bezier"
        Me.GroupControl2.LookAndFeel.UseDefaultLookAndFeel = False
        Me.GroupControl2.Name = "GroupControl2"
        Me.GroupControl2.Size = New System.Drawing.Size(312, 270)
        Me.GroupControl2.TabIndex = 174
        Me.GroupControl2.Text = "Additional Details"
        '
        'tmTimeReleased
        '
        Me.tmTimeReleased.CalendarFont = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.tmTimeReleased.CustomFormat = "MM/dd/yyyy hh:mm:ss tt"
        Me.tmTimeReleased.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.tmTimeReleased.Location = New System.Drawing.Point(90, 149)
        Me.tmTimeReleased.Name = "tmTimeReleased"
        Me.tmTimeReleased.Size = New System.Drawing.Size(213, 22)
        Me.tmTimeReleased.TabIndex = 169
        '
        'LabelControl2
        '
        Me.LabelControl2.Appearance.ForeColor = System.Drawing.Color.FromArgb(CType(CType(27, Byte), Integer), CType(CType(41, Byte), Integer), CType(CType(62, Byte), Integer))
        Me.LabelControl2.Appearance.Options.UseForeColor = True
        Me.LabelControl2.Location = New System.Drawing.Point(8, 156)
        Me.LabelControl2.LookAndFeel.SkinName = "Visual Studio 2013 Blue"
        Me.LabelControl2.LookAndFeel.UseDefaultLookAndFeel = False
        Me.LabelControl2.Name = "LabelControl2"
        Me.LabelControl2.Size = New System.Drawing.Size(76, 13)
        Me.LabelControl2.TabIndex = 170
        Me.LabelControl2.Text = "Date Released:"
        '
        'txtChargeSlip
        '
        Me.txtChargeSlip.Location = New System.Drawing.Point(90, 53)
        Me.txtChargeSlip.Name = "txtChargeSlip"
        Me.txtChargeSlip.Properties.Appearance.ForeColor = System.Drawing.Color.FromArgb(CType(CType(27, Byte), Integer), CType(CType(41, Byte), Integer), CType(CType(62, Byte), Integer))
        Me.txtChargeSlip.Properties.Appearance.Options.UseForeColor = True
        Me.txtChargeSlip.Properties.LookAndFeel.SkinName = "Visual Studio 2013 Blue"
        Me.txtChargeSlip.Properties.LookAndFeel.UseDefaultLookAndFeel = False
        Me.txtChargeSlip.Size = New System.Drawing.Size(213, 20)
        Me.txtChargeSlip.TabIndex = 161
        '
        'LabelControl1
        '
        Me.LabelControl1.Appearance.ForeColor = System.Drawing.Color.FromArgb(CType(CType(27, Byte), Integer), CType(CType(41, Byte), Integer), CType(CType(62, Byte), Integer))
        Me.LabelControl1.Appearance.Options.UseForeColor = True
        Me.LabelControl1.Location = New System.Drawing.Point(8, 56)
        Me.LabelControl1.LookAndFeel.SkinName = "Visual Studio 2013 Blue"
        Me.LabelControl1.LookAndFeel.UseDefaultLookAndFeel = False
        Me.LabelControl1.Name = "LabelControl1"
        Me.LabelControl1.Size = New System.Drawing.Size(72, 13)
        Me.LabelControl1.TabIndex = 162
        Me.LabelControl1.Text = "Charge Slip #:"
        '
        'LabelControl32
        '
        Me.LabelControl32.Appearance.ForeColor = System.Drawing.Color.FromArgb(CType(CType(27, Byte), Integer), CType(CType(41, Byte), Integer), CType(CType(62, Byte), Integer))
        Me.LabelControl32.Appearance.Options.UseForeColor = True
        Me.LabelControl32.Location = New System.Drawing.Point(8, 246)
        Me.LabelControl32.LookAndFeel.SkinName = "Visual Studio 2013 Blue"
        Me.LabelControl32.LookAndFeel.UseDefaultLookAndFeel = False
        Me.LabelControl32.Name = "LabelControl32"
        Me.LabelControl32.Size = New System.Drawing.Size(66, 13)
        Me.LabelControl32.TabIndex = 143
        Me.LabelControl32.Text = "Patient Type:"
        '
        'LabelControl33
        '
        Me.LabelControl33.Appearance.ForeColor = System.Drawing.Color.FromArgb(CType(CType(27, Byte), Integer), CType(CType(41, Byte), Integer), CType(CType(62, Byte), Integer))
        Me.LabelControl33.Appearance.Options.UseForeColor = True
        Me.LabelControl33.Location = New System.Drawing.Point(8, 223)
        Me.LabelControl33.LookAndFeel.SkinName = "Visual Studio 2013 Blue"
        Me.LabelControl33.LookAndFeel.UseDefaultLookAndFeel = False
        Me.LabelControl33.Name = "LabelControl33"
        Me.LabelControl33.Size = New System.Drawing.Size(65, 13)
        Me.LabelControl33.TabIndex = 28
        Me.LabelControl33.Text = "Room/Ward:"
        '
        'LabelControl34
        '
        Me.LabelControl34.Appearance.ForeColor = System.Drawing.Color.FromArgb(CType(CType(27, Byte), Integer), CType(CType(41, Byte), Integer), CType(CType(62, Byte), Integer))
        Me.LabelControl34.Appearance.Options.UseForeColor = True
        Me.LabelControl34.Location = New System.Drawing.Point(8, 79)
        Me.LabelControl34.LookAndFeel.SkinName = "Visual Studio 2013 Blue"
        Me.LabelControl34.LookAndFeel.UseDefaultLookAndFeel = False
        Me.LabelControl34.Name = "LabelControl34"
        Me.LabelControl34.Size = New System.Drawing.Size(40, 13)
        Me.LabelControl34.TabIndex = 160
        Me.LabelControl34.Text = "OR No.:"
        '
        'LabelControl35
        '
        Me.LabelControl35.Appearance.ForeColor = System.Drawing.Color.FromArgb(CType(CType(27, Byte), Integer), CType(CType(41, Byte), Integer), CType(CType(62, Byte), Integer))
        Me.LabelControl35.Appearance.Options.UseForeColor = True
        Me.LabelControl35.Location = New System.Drawing.Point(8, 200)
        Me.LabelControl35.LookAndFeel.SkinName = "Visual Studio 2013 Blue"
        Me.LabelControl35.LookAndFeel.UseDefaultLookAndFeel = False
        Me.LabelControl35.Name = "LabelControl35"
        Me.LabelControl35.Size = New System.Drawing.Size(72, 13)
        Me.LabelControl35.TabIndex = 30
        Me.LabelControl35.Text = "Requested By:"
        '
        'LabelControl36
        '
        Me.LabelControl36.Appearance.ForeColor = System.Drawing.Color.FromArgb(CType(CType(27, Byte), Integer), CType(CType(41, Byte), Integer), CType(CType(62, Byte), Integer))
        Me.LabelControl36.Appearance.Options.UseForeColor = True
        Me.LabelControl36.Location = New System.Drawing.Point(8, 177)
        Me.LabelControl36.LookAndFeel.SkinName = "Visual Studio 2013 Blue"
        Me.LabelControl36.LookAndFeel.UseDefaultLookAndFeel = False
        Me.LabelControl36.Name = "LabelControl36"
        Me.LabelControl36.Size = New System.Drawing.Size(45, 13)
        Me.LabelControl36.TabIndex = 139
        Me.LabelControl36.Text = "Request:"
        '
        'LabelControl37
        '
        Me.LabelControl37.Appearance.ForeColor = System.Drawing.Color.FromArgb(CType(CType(27, Byte), Integer), CType(CType(41, Byte), Integer), CType(CType(62, Byte), Integer))
        Me.LabelControl37.Appearance.Options.UseForeColor = True
        Me.LabelControl37.Location = New System.Drawing.Point(8, 131)
        Me.LabelControl37.LookAndFeel.SkinName = "Visual Studio 2013 Blue"
        Me.LabelControl37.LookAndFeel.UseDefaultLookAndFeel = False
        Me.LabelControl37.Name = "LabelControl37"
        Me.LabelControl37.Size = New System.Drawing.Size(75, 13)
        Me.LabelControl37.TabIndex = 152
        Me.LabelControl37.Text = "Time Received:"
        '
        'LabelControl38
        '
        Me.LabelControl38.Appearance.ForeColor = System.Drawing.Color.FromArgb(CType(CType(27, Byte), Integer), CType(CType(41, Byte), Integer), CType(CType(62, Byte), Integer))
        Me.LabelControl38.Appearance.Options.UseForeColor = True
        Me.LabelControl38.Location = New System.Drawing.Point(8, 33)
        Me.LabelControl38.LookAndFeel.SkinName = "Visual Studio 2013 Blue"
        Me.LabelControl38.LookAndFeel.UseDefaultLookAndFeel = False
        Me.LabelControl38.Name = "LabelControl38"
        Me.LabelControl38.Size = New System.Drawing.Size(77, 13)
        Me.LabelControl38.TabIndex = 158
        Me.LabelControl38.Text = "HIS Tracking #:"
        '
        'LabelControl41
        '
        Me.LabelControl41.Appearance.ForeColor = System.Drawing.Color.FromArgb(CType(CType(27, Byte), Integer), CType(CType(41, Byte), Integer), CType(CType(62, Byte), Integer))
        Me.LabelControl41.Appearance.Options.UseForeColor = True
        Me.LabelControl41.Location = New System.Drawing.Point(8, 106)
        Me.LabelControl41.LookAndFeel.SkinName = "Visual Studio 2013 Blue"
        Me.LabelControl41.LookAndFeel.UseDefaultLookAndFeel = False
        Me.LabelControl41.Name = "LabelControl41"
        Me.LabelControl41.Size = New System.Drawing.Size(75, 13)
        Me.LabelControl41.TabIndex = 136
        Me.LabelControl41.Text = "Date Received:"
        '
        'GroupControl1
        '
        Me.GroupControl1.AppearanceCaption.BackColor = System.Drawing.Color.FromArgb(CType(CType(16, Byte), Integer), CType(CType(110, Byte), Integer), CType(CType(190, Byte), Integer))
        Me.GroupControl1.AppearanceCaption.BackColor2 = System.Drawing.Color.FromArgb(CType(CType(16, Byte), Integer), CType(CType(110, Byte), Integer), CType(CType(190, Byte), Integer))
        Me.GroupControl1.AppearanceCaption.BorderColor = System.Drawing.Color.FromArgb(CType(CType(16, Byte), Integer), CType(CType(110, Byte), Integer), CType(CType(190, Byte), Integer))
        Me.GroupControl1.AppearanceCaption.ForeColor = System.Drawing.Color.White
        Me.GroupControl1.AppearanceCaption.Options.UseBackColor = True
        Me.GroupControl1.AppearanceCaption.Options.UseBorderColor = True
        Me.GroupControl1.AppearanceCaption.Options.UseForeColor = True
        Me.GroupControl1.CaptionImageOptions.AllowGlyphSkinning = True
        Me.GroupControl1.CaptionImageOptions.SvgImage = CType(resources.GetObject("GroupControl1.CaptionImageOptions.SvgImage"), DevExpress.Utils.Svg.SvgImage)
        Me.GroupControl1.CaptionImageOptions.SvgImageSize = New System.Drawing.Size(16, 16)
        Me.GroupControl1.Controls.Add(Me.txtEmail)
        Me.GroupControl1.Controls.Add(Me.LabelControl42)
        Me.GroupControl1.Controls.Add(Me.cboCS)
        Me.GroupControl1.Controls.Add(Me.LabelControl43)
        Me.GroupControl1.Controls.Add(Me.btnSearch)
        Me.GroupControl1.Controls.Add(Me.LabelControl44)
        Me.GroupControl1.Controls.Add(Me.txtSampleID)
        Me.GroupControl1.Controls.Add(Me.txtContact)
        Me.GroupControl1.Controls.Add(Me.LabelControl45)
        Me.GroupControl1.Controls.Add(Me.txtClass)
        Me.GroupControl1.Controls.Add(Me.LabelControl46)
        Me.GroupControl1.Controls.Add(Me.LabelControl47)
        Me.GroupControl1.Controls.Add(Me.LabelControl48)
        Me.GroupControl1.Controls.Add(Me.LabelControl49)
        Me.GroupControl1.Controls.Add(Me.txtAge)
        Me.GroupControl1.Controls.Add(Me.LabelControl50)
        Me.GroupControl1.Controls.Add(Me.txtPatientID)
        Me.GroupControl1.Controls.Add(Me.cboSex)
        Me.GroupControl1.Controls.Add(Me.txtName)
        Me.GroupControl1.Controls.Add(Me.dtBDate)
        Me.GroupControl1.Controls.Add(Me.txtAddress)
        Me.GroupControl1.Location = New System.Drawing.Point(10, 42)
        Me.GroupControl1.Name = "GroupControl1"
        Me.GroupControl1.Size = New System.Drawing.Size(312, 243)
        Me.GroupControl1.TabIndex = 173
        Me.GroupControl1.Text = "Patient Details"
        '
        'txtEmail
        '
        Me.txtEmail.Location = New System.Drawing.Point(90, 243)
        Me.txtEmail.Name = "txtEmail"
        Me.txtEmail.Properties.Appearance.ForeColor = System.Drawing.Color.FromArgb(CType(CType(27, Byte), Integer), CType(CType(41, Byte), Integer), CType(CType(62, Byte), Integer))
        Me.txtEmail.Properties.Appearance.Options.UseForeColor = True
        Me.txtEmail.Properties.AppearanceReadOnly.BackColor = System.Drawing.Color.White
        Me.txtEmail.Properties.AppearanceReadOnly.BackColor2 = System.Drawing.Color.White
        Me.txtEmail.Properties.AppearanceReadOnly.BorderColor = System.Drawing.Color.White
        Me.txtEmail.Properties.AppearanceReadOnly.Options.UseBackColor = True
        Me.txtEmail.Properties.AppearanceReadOnly.Options.UseBorderColor = True
        Me.txtEmail.Properties.LookAndFeel.SkinName = "Visual Studio 2013 Blue"
        Me.txtEmail.Properties.LookAndFeel.UseDefaultLookAndFeel = False
        Me.txtEmail.Size = New System.Drawing.Size(213, 20)
        Me.txtEmail.TabIndex = 170
        '
        'LabelControl42
        '
        Me.LabelControl42.Appearance.ForeColor = System.Drawing.Color.FromArgb(CType(CType(27, Byte), Integer), CType(CType(41, Byte), Integer), CType(CType(62, Byte), Integer))
        Me.LabelControl42.Appearance.Options.UseForeColor = True
        Me.LabelControl42.Location = New System.Drawing.Point(8, 173)
        Me.LabelControl42.LookAndFeel.SkinName = "Visual Studio 2013 Blue"
        Me.LabelControl42.LookAndFeel.UseDefaultLookAndFeel = False
        Me.LabelControl42.Name = "LabelControl42"
        Me.LabelControl42.Size = New System.Drawing.Size(59, 13)
        Me.LabelControl42.TabIndex = 161
        Me.LabelControl42.Text = "Civil Status:"
        '
        'LabelControl43
        '
        Me.LabelControl43.Appearance.ForeColor = System.Drawing.Color.FromArgb(CType(CType(27, Byte), Integer), CType(CType(41, Byte), Integer), CType(CType(62, Byte), Integer))
        Me.LabelControl43.Appearance.Options.UseForeColor = True
        Me.LabelControl43.Location = New System.Drawing.Point(8, 35)
        Me.LabelControl43.LookAndFeel.SkinName = "Visual Studio 2013 Blue"
        Me.LabelControl43.LookAndFeel.UseDefaultLookAndFeel = False
        Me.LabelControl43.Name = "LabelControl43"
        Me.LabelControl43.Size = New System.Drawing.Size(54, 13)
        Me.LabelControl43.TabIndex = 22
        Me.LabelControl43.Text = "Sample ID:"
        '
        'LabelControl44
        '
        Me.LabelControl44.Appearance.ForeColor = System.Drawing.Color.FromArgb(CType(CType(27, Byte), Integer), CType(CType(41, Byte), Integer), CType(CType(62, Byte), Integer))
        Me.LabelControl44.Appearance.Options.UseForeColor = True
        Me.LabelControl44.Location = New System.Drawing.Point(8, 127)
        Me.LabelControl44.LookAndFeel.SkinName = "Visual Studio 2013 Blue"
        Me.LabelControl44.LookAndFeel.UseDefaultLookAndFeel = False
        Me.LabelControl44.Name = "LabelControl44"
        Me.LabelControl44.Size = New System.Drawing.Size(68, 13)
        Me.LabelControl44.TabIndex = 85
        Me.LabelControl44.Text = "Date of Birth:"
        '
        'LabelControl45
        '
        Me.LabelControl45.Appearance.ForeColor = System.Drawing.Color.FromArgb(CType(CType(27, Byte), Integer), CType(CType(41, Byte), Integer), CType(CType(62, Byte), Integer))
        Me.LabelControl45.Appearance.Options.UseForeColor = True
        Me.LabelControl45.Location = New System.Drawing.Point(8, 58)
        Me.LabelControl45.LookAndFeel.SkinName = "Visual Studio 2013 Blue"
        Me.LabelControl45.LookAndFeel.UseDefaultLookAndFeel = False
        Me.LabelControl45.Name = "LabelControl45"
        Me.LabelControl45.Size = New System.Drawing.Size(53, 13)
        Me.LabelControl45.TabIndex = 12
        Me.LabelControl45.Text = "Patient ID:"
        '
        'LabelControl46
        '
        Me.LabelControl46.Appearance.ForeColor = System.Drawing.Color.FromArgb(CType(CType(27, Byte), Integer), CType(CType(41, Byte), Integer), CType(CType(62, Byte), Integer))
        Me.LabelControl46.Appearance.Options.UseForeColor = True
        Me.LabelControl46.Location = New System.Drawing.Point(8, 219)
        Me.LabelControl46.LookAndFeel.SkinName = "Visual Studio 2013 Blue"
        Me.LabelControl46.LookAndFeel.UseDefaultLookAndFeel = False
        Me.LabelControl46.Name = "LabelControl46"
        Me.LabelControl46.Size = New System.Drawing.Size(64, 13)
        Me.LabelControl46.TabIndex = 159
        Me.LabelControl46.Text = "Contact No.:"
        '
        'LabelControl47
        '
        Me.LabelControl47.Appearance.ForeColor = System.Drawing.Color.FromArgb(CType(CType(27, Byte), Integer), CType(CType(41, Byte), Integer), CType(CType(62, Byte), Integer))
        Me.LabelControl47.Appearance.Options.UseForeColor = True
        Me.LabelControl47.Location = New System.Drawing.Point(8, 81)
        Me.LabelControl47.LookAndFeel.SkinName = "Visual Studio 2013 Blue"
        Me.LabelControl47.LookAndFeel.UseDefaultLookAndFeel = False
        Me.LabelControl47.Name = "LabelControl47"
        Me.LabelControl47.Size = New System.Drawing.Size(71, 13)
        Me.LabelControl47.TabIndex = 14
        Me.LabelControl47.Text = "Patient Name:"
        '
        'LabelControl48
        '
        Me.LabelControl48.Appearance.ForeColor = System.Drawing.Color.FromArgb(CType(CType(27, Byte), Integer), CType(CType(41, Byte), Integer), CType(CType(62, Byte), Integer))
        Me.LabelControl48.Appearance.Options.UseForeColor = True
        Me.LabelControl48.Location = New System.Drawing.Point(8, 150)
        Me.LabelControl48.LookAndFeel.SkinName = "Visual Studio 2013 Blue"
        Me.LabelControl48.LookAndFeel.UseDefaultLookAndFeel = False
        Me.LabelControl48.Name = "LabelControl48"
        Me.LabelControl48.Size = New System.Drawing.Size(23, 13)
        Me.LabelControl48.TabIndex = 36
        Me.LabelControl48.Text = "Age:"
        '
        'LabelControl49
        '
        Me.LabelControl49.Appearance.ForeColor = System.Drawing.Color.FromArgb(CType(CType(27, Byte), Integer), CType(CType(41, Byte), Integer), CType(CType(62, Byte), Integer))
        Me.LabelControl49.Appearance.Options.UseForeColor = True
        Me.LabelControl49.Location = New System.Drawing.Point(8, 195)
        Me.LabelControl49.LookAndFeel.SkinName = "Visual Studio 2013 Blue"
        Me.LabelControl49.LookAndFeel.UseDefaultLookAndFeel = False
        Me.LabelControl49.Name = "LabelControl49"
        Me.LabelControl49.Size = New System.Drawing.Size(44, 13)
        Me.LabelControl49.TabIndex = 157
        Me.LabelControl49.Text = "Address:"
        '
        'LabelControl50
        '
        Me.LabelControl50.Appearance.ForeColor = System.Drawing.Color.FromArgb(CType(CType(27, Byte), Integer), CType(CType(41, Byte), Integer), CType(CType(62, Byte), Integer))
        Me.LabelControl50.Appearance.Options.UseForeColor = True
        Me.LabelControl50.Location = New System.Drawing.Point(8, 104)
        Me.LabelControl50.LookAndFeel.SkinName = "Visual Studio 2013 Blue"
        Me.LabelControl50.LookAndFeel.UseDefaultLookAndFeel = False
        Me.LabelControl50.Name = "LabelControl50"
        Me.LabelControl50.Size = New System.Drawing.Size(20, 13)
        Me.LabelControl50.TabIndex = 20
        Me.LabelControl50.Text = "Sex:"
        '
        'txtAddress
        '
        Me.txtAddress.Location = New System.Drawing.Point(90, 193)
        Me.txtAddress.Name = "txtAddress"
        Me.txtAddress.Properties.Appearance.ForeColor = System.Drawing.Color.FromArgb(CType(CType(27, Byte), Integer), CType(CType(41, Byte), Integer), CType(CType(62, Byte), Integer))
        Me.txtAddress.Properties.Appearance.Options.UseForeColor = True
        Me.txtAddress.Properties.AppearanceDisabled.ForeColor = System.Drawing.Color.FromArgb(CType(CType(27, Byte), Integer), CType(CType(41, Byte), Integer), CType(CType(62, Byte), Integer))
        Me.txtAddress.Properties.AppearanceDisabled.Options.UseForeColor = True
        Me.txtAddress.Properties.AppearanceReadOnly.BackColor = System.Drawing.Color.White
        Me.txtAddress.Properties.AppearanceReadOnly.BackColor2 = System.Drawing.Color.White
        Me.txtAddress.Properties.AppearanceReadOnly.BorderColor = System.Drawing.Color.White
        Me.txtAddress.Properties.AppearanceReadOnly.Options.UseBackColor = True
        Me.txtAddress.Properties.AppearanceReadOnly.Options.UseBorderColor = True
        Me.txtAddress.Properties.LookAndFeel.SkinName = "Visual Studio 2013 Blue"
        Me.txtAddress.Properties.LookAndFeel.UseDefaultLookAndFeel = False
        Me.txtAddress.Size = New System.Drawing.Size(213, 20)
        Me.txtAddress.TabIndex = 8
        '
        'GroupControl4
        '
        Me.GroupControl4.AppearanceCaption.BackColor = System.Drawing.Color.FromArgb(CType(CType(16, Byte), Integer), CType(CType(110, Byte), Integer), CType(CType(190, Byte), Integer))
        Me.GroupControl4.AppearanceCaption.BackColor2 = System.Drawing.Color.FromArgb(CType(CType(16, Byte), Integer), CType(CType(110, Byte), Integer), CType(CType(190, Byte), Integer))
        Me.GroupControl4.AppearanceCaption.BorderColor = System.Drawing.Color.FromArgb(CType(CType(16, Byte), Integer), CType(CType(110, Byte), Integer), CType(CType(190, Byte), Integer))
        Me.GroupControl4.AppearanceCaption.ForeColor = System.Drawing.Color.White
        Me.GroupControl4.AppearanceCaption.Options.UseBackColor = True
        Me.GroupControl4.AppearanceCaption.Options.UseBorderColor = True
        Me.GroupControl4.AppearanceCaption.Options.UseForeColor = True
        Me.GroupControl4.CaptionImageOptions.AllowGlyphSkinning = True
        Me.GroupControl4.CaptionImageOptions.SvgImage = CType(resources.GetObject("GroupControl4.CaptionImageOptions.SvgImage"), DevExpress.Utils.Svg.SvgImage)
        Me.GroupControl4.CaptionImageOptions.SvgImageSize = New System.Drawing.Size(16, 16)
        Me.GroupControl4.Controls.Add(Me.txtRemarks)
        Me.GroupControl4.Location = New System.Drawing.Point(328, 567)
        Me.GroupControl4.LookAndFeel.SkinName = "The Bezier"
        Me.GroupControl4.LookAndFeel.UseDefaultLookAndFeel = False
        Me.GroupControl4.Name = "GroupControl4"
        Me.GroupControl4.Size = New System.Drawing.Size(476, 108)
        Me.GroupControl4.TabIndex = 180
        Me.GroupControl4.Text = "Remarks"
        '
        'GroupControl5
        '
        Me.GroupControl5.AppearanceCaption.BackColor = System.Drawing.Color.FromArgb(CType(CType(16, Byte), Integer), CType(CType(110, Byte), Integer), CType(CType(190, Byte), Integer))
        Me.GroupControl5.AppearanceCaption.BackColor2 = System.Drawing.Color.FromArgb(CType(CType(16, Byte), Integer), CType(CType(110, Byte), Integer), CType(CType(190, Byte), Integer))
        Me.GroupControl5.AppearanceCaption.BorderColor = System.Drawing.Color.FromArgb(CType(CType(16, Byte), Integer), CType(CType(110, Byte), Integer), CType(CType(190, Byte), Integer))
        Me.GroupControl5.AppearanceCaption.ForeColor = System.Drawing.Color.White
        Me.GroupControl5.AppearanceCaption.Options.UseBackColor = True
        Me.GroupControl5.AppearanceCaption.Options.UseBorderColor = True
        Me.GroupControl5.AppearanceCaption.Options.UseForeColor = True
        Me.GroupControl5.CaptionImageOptions.AllowGlyphSkinning = True
        Me.GroupControl5.CaptionImageOptions.SvgImage = CType(resources.GetObject("GroupControl5.CaptionImageOptions.SvgImage"), DevExpress.Utils.Svg.SvgImage)
        Me.GroupControl5.CaptionImageOptions.SvgImageSize = New System.Drawing.Size(16, 16)
        Me.GroupControl5.Controls.Add(Me.cboMedTech)
        Me.GroupControl5.Controls.Add(Me.cboVerify)
        Me.GroupControl5.Controls.Add(Me.cboPathologist)
        Me.GroupControl5.Controls.Add(Me.LabelControl26)
        Me.GroupControl5.Controls.Add(Me.LabelControl29)
        Me.GroupControl5.Controls.Add(Me.LabelControl27)
        Me.GroupControl5.Location = New System.Drawing.Point(10, 567)
        Me.GroupControl5.LookAndFeel.SkinName = "The Bezier"
        Me.GroupControl5.LookAndFeel.UseDefaultLookAndFeel = False
        Me.GroupControl5.Name = "GroupControl5"
        Me.GroupControl5.Size = New System.Drawing.Size(312, 108)
        Me.GroupControl5.TabIndex = 181
        Me.GroupControl5.Text = "Signatories"
        '
        'GroupControl6
        '
        Me.GroupControl6.AppearanceCaption.BackColor = System.Drawing.Color.FromArgb(CType(CType(16, Byte), Integer), CType(CType(110, Byte), Integer), CType(CType(190, Byte), Integer))
        Me.GroupControl6.AppearanceCaption.BackColor2 = System.Drawing.Color.FromArgb(CType(CType(16, Byte), Integer), CType(CType(110, Byte), Integer), CType(CType(190, Byte), Integer))
        Me.GroupControl6.AppearanceCaption.BorderColor = System.Drawing.Color.FromArgb(CType(CType(16, Byte), Integer), CType(CType(110, Byte), Integer), CType(CType(190, Byte), Integer))
        Me.GroupControl6.AppearanceCaption.ForeColor = System.Drawing.Color.White
        Me.GroupControl6.AppearanceCaption.Options.UseBackColor = True
        Me.GroupControl6.AppearanceCaption.Options.UseBorderColor = True
        Me.GroupControl6.AppearanceCaption.Options.UseForeColor = True
        Me.GroupControl6.CaptionImageOptions.AllowGlyphSkinning = True
        Me.GroupControl6.CaptionImageOptions.SvgImage = CType(resources.GetObject("GroupControl6.CaptionImageOptions.SvgImage"), DevExpress.Utils.Svg.SvgImage)
        Me.GroupControl6.CaptionImageOptions.SvgImageSize = New System.Drawing.Size(16, 16)
        Me.GroupControl6.Controls.Add(Me.txtComment)
        Me.GroupControl6.Location = New System.Drawing.Point(810, 567)
        Me.GroupControl6.LookAndFeel.SkinName = "The Bezier"
        Me.GroupControl6.LookAndFeel.UseDefaultLookAndFeel = False
        Me.GroupControl6.Name = "GroupControl6"
        Me.GroupControl6.Size = New System.Drawing.Size(466, 108)
        Me.GroupControl6.TabIndex = 187
        Me.GroupControl6.Text = "Comment"
        '
        'frmFecalNew
        '
        Me.Appearance.BackColor = System.Drawing.SystemColors.Control
        Me.Appearance.ForeColor = System.Drawing.Color.FromArgb(CType(CType(64, Byte), Integer), CType(CType(64, Byte), Integer), CType(CType(64, Byte), Integer))
        Me.Appearance.Options.UseBackColor = True
        Me.Appearance.Options.UseForeColor = True
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(1305, 705)
        Me.ControlBox = False
        Me.Controls.Add(Me.GroupControl6)
        Me.Controls.Add(Me.GroupControl5)
        Me.Controls.Add(Me.GroupControl4)
        Me.Controls.Add(Me.GroupControl3)
        Me.Controls.Add(Me.GroupControl2)
        Me.Controls.Add(Me.GroupControl1)
        Me.Controls.Add(Me.barDockControlLeft)
        Me.Controls.Add(Me.barDockControlRight)
        Me.Controls.Add(Me.barDockControlBottom)
        Me.Controls.Add(Me.barDockControlTop)
        Me.FormBorderEffect = DevExpress.XtraEditors.FormBorderEffect.Glow
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "frmFecalNew"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
        Me.Text = "Fecalysis Result Window"
        CType(Me.cboVerify.Properties, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.cboPathologist.Properties, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.cboMedTech.Properties, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.txtAge.Properties, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.txtSampleID.Properties, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.txtName.Properties, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.txtPatientID.Properties, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.cboSex.Properties, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.BarManager, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.BarAndDockingController, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.dtBDate.Properties.CalendarTimeProperties, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.dtBDate.Properties, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.txtClass.Properties, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.txtComment.Properties, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.txtRemarks.Properties, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.cboPatientType.Properties, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.cboRoom.Properties, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.cboPhysician.Properties, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.cboRequest.Properties, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.txtORNo.Properties, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.txtAccession.Properties, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.txtContact.Properties, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.cboCS.Properties, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.GroupControl3, System.ComponentModel.ISupportInitialize).EndInit()
        Me.GroupControl3.ResumeLayout(False)
        CType(Me.dtResult, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.GridView, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.GroupControl2, System.ComponentModel.ISupportInitialize).EndInit()
        Me.GroupControl2.ResumeLayout(False)
        Me.GroupControl2.PerformLayout()
        CType(Me.txtChargeSlip.Properties, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.GroupControl1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.GroupControl1.ResumeLayout(False)
        Me.GroupControl1.PerformLayout()
        CType(Me.txtEmail.Properties, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.txtAddress.Properties, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.GroupControl4, System.ComponentModel.ISupportInitialize).EndInit()
        Me.GroupControl4.ResumeLayout(False)
        CType(Me.GroupControl5, System.ComponentModel.ISupportInitialize).EndInit()
        Me.GroupControl5.ResumeLayout(False)
        Me.GroupControl5.PerformLayout()
        CType(Me.GroupControl6, System.ComponentModel.ISupportInitialize).EndInit()
        Me.GroupControl6.ResumeLayout(False)
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents txtPatientID As DevExpress.XtraEditors.TextEdit
    Friend WithEvents txtSampleID As DevExpress.XtraEditors.TextEdit
    Friend WithEvents txtName As DevExpress.XtraEditors.TextEdit
    Friend WithEvents txtAge As DevExpress.XtraEditors.TextEdit
    Friend WithEvents cboSex As DevExpress.XtraEditors.ComboBoxEdit
    Friend WithEvents cboPathologist As DevExpress.XtraEditors.ComboBoxEdit
    Friend WithEvents cboMedTech As DevExpress.XtraEditors.ComboBoxEdit
    Friend WithEvents cboVerify As DevExpress.XtraEditors.ComboBoxEdit
    Friend WithEvents BarManager As DevExpress.XtraBars.BarManager
    Friend WithEvents Bar2 As DevExpress.XtraBars.Bar
    Friend WithEvents barDockControlTop As DevExpress.XtraBars.BarDockControl
    Friend WithEvents barDockControlBottom As DevExpress.XtraBars.BarDockControl
    Friend WithEvents barDockControlLeft As DevExpress.XtraBars.BarDockControl
    Friend WithEvents barDockControlRight As DevExpress.XtraBars.BarDockControl
    Friend WithEvents ImageList As System.Windows.Forms.ImageList
    Friend WithEvents BarAndDockingController As DevExpress.XtraBars.BarAndDockingController
    Friend WithEvents btnPrint As DevExpress.XtraBars.BarLargeButtonItem
    Friend WithEvents BarLargeButtonItem3 As DevExpress.XtraBars.BarLargeButtonItem
    Friend WithEvents dtBDate As DevExpress.XtraEditors.DateEdit
    Friend WithEvents txtClass As DevExpress.XtraEditors.ComboBoxEdit
    Friend WithEvents txtContact As DevExpress.XtraEditors.TextEdit
    Friend WithEvents btnClose As DevExpress.XtraBars.BarLargeButtonItem
    Friend WithEvents cboPatientType As DevExpress.XtraEditors.ComboBoxEdit
    Friend WithEvents cboRoom As DevExpress.XtraEditors.ComboBoxEdit
    Friend WithEvents cboPhysician As DevExpress.XtraEditors.ComboBoxEdit
    Friend WithEvents cboRequest As DevExpress.XtraEditors.ComboBoxEdit
    Friend WithEvents btnSearch As DevExpress.XtraEditors.SimpleButton
    Friend WithEvents txtORNo As DevExpress.XtraEditors.TextEdit
    Friend WithEvents txtAccession As DevExpress.XtraEditors.TextEdit
    Friend WithEvents txtComment As DevExpress.XtraEditors.MemoEdit
    Friend WithEvents txtRemarks As DevExpress.XtraEditors.MemoEdit
    Friend WithEvents btnValidate As DevExpress.XtraBars.BarLargeButtonItem
    Friend WithEvents btnPrintNow As DevExpress.XtraBars.BarLargeButtonItem
    Friend WithEvents btnViewPrint As DevExpress.XtraBars.BarLargeButtonItem
    Friend WithEvents dtReceived As System.Windows.Forms.DateTimePicker
    Friend WithEvents tmTimeReceived As System.Windows.Forms.DateTimePicker
    Friend WithEvents cboCS As DevExpress.XtraEditors.ComboBoxEdit
    Friend WithEvents GroupControl3 As DevExpress.XtraEditors.GroupControl
    Friend WithEvents LabelControl26 As DevExpress.XtraEditors.LabelControl
    Friend WithEvents LabelControl27 As DevExpress.XtraEditors.LabelControl
    Friend WithEvents LabelControl29 As DevExpress.XtraEditors.LabelControl
    Friend WithEvents GroupControl2 As DevExpress.XtraEditors.GroupControl
    Friend WithEvents LabelControl32 As DevExpress.XtraEditors.LabelControl
    Friend WithEvents LabelControl33 As DevExpress.XtraEditors.LabelControl
    Friend WithEvents LabelControl34 As DevExpress.XtraEditors.LabelControl
    Friend WithEvents LabelControl35 As DevExpress.XtraEditors.LabelControl
    Friend WithEvents LabelControl36 As DevExpress.XtraEditors.LabelControl
    Friend WithEvents LabelControl37 As DevExpress.XtraEditors.LabelControl
    Friend WithEvents LabelControl38 As DevExpress.XtraEditors.LabelControl
    Friend WithEvents LabelControl41 As DevExpress.XtraEditors.LabelControl
    Friend WithEvents GroupControl1 As DevExpress.XtraEditors.GroupControl
    Friend WithEvents LabelControl42 As DevExpress.XtraEditors.LabelControl
    Friend WithEvents LabelControl43 As DevExpress.XtraEditors.LabelControl
    Friend WithEvents LabelControl44 As DevExpress.XtraEditors.LabelControl
    Friend WithEvents LabelControl45 As DevExpress.XtraEditors.LabelControl
    Friend WithEvents LabelControl46 As DevExpress.XtraEditors.LabelControl
    Friend WithEvents LabelControl47 As DevExpress.XtraEditors.LabelControl
    Friend WithEvents LabelControl48 As DevExpress.XtraEditors.LabelControl
    Friend WithEvents LabelControl49 As DevExpress.XtraEditors.LabelControl
    Friend WithEvents LabelControl50 As DevExpress.XtraEditors.LabelControl
    Friend WithEvents dtResult As DevExpress.XtraGrid.GridControl
    Friend WithEvents GridView As DevExpress.XtraGrid.Views.Grid.GridView
    Friend WithEvents GroupControl5 As DevExpress.XtraEditors.GroupControl
    Friend WithEvents GroupControl4 As DevExpress.XtraEditors.GroupControl
    Friend WithEvents btnRetrive As DevExpress.XtraEditors.SimpleButton
    Friend WithEvents btnPreview As DevExpress.XtraEditors.SimpleButton
    Friend WithEvents btnAddTest As DevExpress.XtraEditors.SimpleButton
    Friend WithEvents GroupControl6 As DevExpress.XtraEditors.GroupControl
    Friend WithEvents txtChargeSlip As DevExpress.XtraEditors.TextEdit
    Friend WithEvents LabelControl1 As DevExpress.XtraEditors.LabelControl
    Friend WithEvents txtAddress As DevExpress.XtraEditors.TextEdit
    Friend WithEvents btnRelease As DevExpress.XtraBars.BarLargeButtonItem
    Friend WithEvents tmTimeReleased As DateTimePicker
    Friend WithEvents LabelControl2 As DevExpress.XtraEditors.LabelControl
    Friend WithEvents txtEmail As DevExpress.XtraEditors.TextEdit
    Friend WithEvents btnVerification As DevExpress.XtraBars.BarLargeButtonItem
    Friend WithEvents btnEmailPush As DevExpress.XtraBars.BarLargeButtonItem
End Class
