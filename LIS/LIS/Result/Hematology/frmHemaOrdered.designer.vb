<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class frmHemaOrdered
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmHemaOrdered))
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
        Me.btnEdit = New DevExpress.XtraBars.BarLargeButtonItem()
        Me.btnValidate = New DevExpress.XtraBars.BarLargeButtonItem()
        Me.btnClose = New DevExpress.XtraBars.BarLargeButtonItem()
        Me.btnPrintNow = New DevExpress.XtraBars.BarLargeButtonItem()
        Me.btnResend = New DevExpress.XtraBars.BarLargeButtonItem()
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
        Me.gcTest = New DevExpress.XtraEditors.GroupControl()
        Me.dtResult = New DevExpress.XtraGrid.GridControl()
        Me.GridView = New DevExpress.XtraGrid.Views.Grid.GridView()
        Me.lblDiffCount = New DevExpress.XtraEditors.LabelControl()
        Me.btnRetrive = New DevExpress.XtraEditors.SimpleButton()
        Me.btnPreview = New DevExpress.XtraEditors.SimpleButton()
        Me.btnAddTest = New DevExpress.XtraEditors.SimpleButton()
        Me.LabelControl26 = New DevExpress.XtraEditors.LabelControl()
        Me.LabelControl27 = New DevExpress.XtraEditors.LabelControl()
        Me.LabelControl29 = New DevExpress.XtraEditors.LabelControl()
        Me.gcAdditional = New DevExpress.XtraEditors.GroupControl()
        Me.tmTimeReceived = New DevExpress.XtraEditors.DateEdit()
        Me.tmTimeReleased = New System.Windows.Forms.DateTimePicker()
        Me.LabelControl2 = New DevExpress.XtraEditors.LabelControl()
        Me.LabelControl1 = New DevExpress.XtraEditors.LabelControl()
        Me.txtChargeSlip = New DevExpress.XtraEditors.TextEdit()
        Me.dtReceived = New DevExpress.XtraEditors.DateEdit()
        Me.LabelControl32 = New DevExpress.XtraEditors.LabelControl()
        Me.LabelControl33 = New DevExpress.XtraEditors.LabelControl()
        Me.LabelControl34 = New DevExpress.XtraEditors.LabelControl()
        Me.LabelControl35 = New DevExpress.XtraEditors.LabelControl()
        Me.LabelControl36 = New DevExpress.XtraEditors.LabelControl()
        Me.LabelControl37 = New DevExpress.XtraEditors.LabelControl()
        Me.LabelControl38 = New DevExpress.XtraEditors.LabelControl()
        Me.LabelControl41 = New DevExpress.XtraEditors.LabelControl()
        Me.gcPatient = New DevExpress.XtraEditors.GroupControl()
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
        Me.gcRemarks = New DevExpress.XtraEditors.GroupControl()
        Me.gcSignature = New DevExpress.XtraEditors.GroupControl()
        Me.GroupControl1 = New DevExpress.XtraEditors.GroupControl()
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
        CType(Me.gcTest, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.gcTest.SuspendLayout()
        CType(Me.dtResult, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.GridView, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.gcAdditional, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.gcAdditional.SuspendLayout()
        CType(Me.tmTimeReceived.Properties.CalendarTimeProperties, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.tmTimeReceived.Properties, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.txtChargeSlip.Properties, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.dtReceived.Properties.CalendarTimeProperties, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.dtReceived.Properties, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.gcPatient, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.gcPatient.SuspendLayout()
        CType(Me.txtAddress.Properties, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.gcRemarks, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.gcRemarks.SuspendLayout()
        CType(Me.gcSignature, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.gcSignature.SuspendLayout()
        CType(Me.GroupControl1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.GroupControl1.SuspendLayout()
        Me.SuspendLayout()
        '
        'cboVerify
        '
        Me.cboVerify.Location = New System.Drawing.Point(94, 56)
        Me.cboVerify.Name = "cboVerify"
        Me.cboVerify.Properties.Appearance.Font = New System.Drawing.Font("Segoe UI", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cboVerify.Properties.Appearance.ForeColor = System.Drawing.Color.FromArgb(CType(CType(27, Byte), Integer), CType(CType(41, Byte), Integer), CType(CType(62, Byte), Integer))
        Me.cboVerify.Properties.Appearance.Options.UseFont = True
        Me.cboVerify.Properties.Appearance.Options.UseForeColor = True
        Me.cboVerify.Properties.Buttons.AddRange(New DevExpress.XtraEditors.Controls.EditorButton() {New DevExpress.XtraEditors.Controls.EditorButton(DevExpress.XtraEditors.Controls.ButtonPredefines.Combo)})
        Me.cboVerify.Properties.Items.AddRange(New Object() {"Day", "Week", "Month", "Year"})
        Me.cboVerify.Size = New System.Drawing.Size(213, 20)
        Me.cboVerify.TabIndex = 19
        '
        'cboPathologist
        '
        Me.cboPathologist.Location = New System.Drawing.Point(94, 79)
        Me.cboPathologist.Name = "cboPathologist"
        Me.cboPathologist.Properties.Appearance.Font = New System.Drawing.Font("Segoe UI", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cboPathologist.Properties.Appearance.ForeColor = System.Drawing.Color.FromArgb(CType(CType(27, Byte), Integer), CType(CType(41, Byte), Integer), CType(CType(62, Byte), Integer))
        Me.cboPathologist.Properties.Appearance.Options.UseFont = True
        Me.cboPathologist.Properties.Appearance.Options.UseForeColor = True
        Me.cboPathologist.Properties.Buttons.AddRange(New DevExpress.XtraEditors.Controls.EditorButton() {New DevExpress.XtraEditors.Controls.EditorButton(DevExpress.XtraEditors.Controls.ButtonPredefines.Combo)})
        Me.cboPathologist.Properties.Items.AddRange(New Object() {"Day", "Week", "Month", "Year"})
        Me.cboPathologist.Size = New System.Drawing.Size(213, 20)
        Me.cboPathologist.TabIndex = 20
        '
        'cboMedTech
        '
        Me.cboMedTech.Location = New System.Drawing.Point(94, 33)
        Me.cboMedTech.Name = "cboMedTech"
        Me.cboMedTech.Properties.Appearance.Font = New System.Drawing.Font("Segoe UI", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cboMedTech.Properties.Appearance.ForeColor = System.Drawing.Color.FromArgb(CType(CType(27, Byte), Integer), CType(CType(41, Byte), Integer), CType(CType(62, Byte), Integer))
        Me.cboMedTech.Properties.Appearance.Options.UseFont = True
        Me.cboMedTech.Properties.Appearance.Options.UseForeColor = True
        Me.cboMedTech.Properties.Buttons.AddRange(New DevExpress.XtraEditors.Controls.EditorButton() {New DevExpress.XtraEditors.Controls.EditorButton(DevExpress.XtraEditors.Controls.ButtonPredefines.Combo)})
        Me.cboMedTech.Properties.Items.AddRange(New Object() {"Day", "Week", "Month", "Year"})
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
        Me.txtAge.Location = New System.Drawing.Point(94, 147)
        Me.txtAge.Name = "txtAge"
        Me.txtAge.Properties.Appearance.Font = New System.Drawing.Font("Segoe UI", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtAge.Properties.Appearance.ForeColor = System.Drawing.Color.FromArgb(CType(CType(27, Byte), Integer), CType(CType(41, Byte), Integer), CType(CType(62, Byte), Integer))
        Me.txtAge.Properties.Appearance.Options.UseFont = True
        Me.txtAge.Properties.Appearance.Options.UseForeColor = True
        Me.txtAge.Properties.AppearanceReadOnly.BackColor = System.Drawing.Color.White
        Me.txtAge.Properties.AppearanceReadOnly.BackColor2 = System.Drawing.Color.White
        Me.txtAge.Properties.AppearanceReadOnly.BorderColor = System.Drawing.Color.White
        Me.txtAge.Properties.AppearanceReadOnly.Options.UseBackColor = True
        Me.txtAge.Properties.AppearanceReadOnly.Options.UseBorderColor = True
        Me.txtAge.Size = New System.Drawing.Size(106, 20)
        Me.txtAge.TabIndex = 5
        '
        'txtSampleID
        '
        Me.txtSampleID.Location = New System.Drawing.Point(94, 32)
        Me.txtSampleID.Name = "txtSampleID"
        Me.txtSampleID.Properties.Appearance.Font = New System.Drawing.Font("Segoe UI", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtSampleID.Properties.Appearance.ForeColor = System.Drawing.Color.FromArgb(CType(CType(27, Byte), Integer), CType(CType(41, Byte), Integer), CType(CType(62, Byte), Integer))
        Me.txtSampleID.Properties.Appearance.Options.UseFont = True
        Me.txtSampleID.Properties.Appearance.Options.UseForeColor = True
        Me.txtSampleID.Properties.AppearanceReadOnly.BackColor = System.Drawing.Color.White
        Me.txtSampleID.Properties.AppearanceReadOnly.BackColor2 = System.Drawing.Color.White
        Me.txtSampleID.Properties.AppearanceReadOnly.BorderColor = System.Drawing.Color.White
        Me.txtSampleID.Properties.AppearanceReadOnly.Options.UseBackColor = True
        Me.txtSampleID.Properties.AppearanceReadOnly.Options.UseBorderColor = True
        Me.txtSampleID.Size = New System.Drawing.Size(213, 20)
        Me.txtSampleID.TabIndex = 0
        '
        'txtName
        '
        Me.txtName.Location = New System.Drawing.Point(94, 78)
        Me.txtName.Name = "txtName"
        Me.txtName.Properties.Appearance.Font = New System.Drawing.Font("Segoe UI", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtName.Properties.Appearance.ForeColor = System.Drawing.Color.FromArgb(CType(CType(27, Byte), Integer), CType(CType(41, Byte), Integer), CType(CType(62, Byte), Integer))
        Me.txtName.Properties.Appearance.Options.UseFont = True
        Me.txtName.Properties.Appearance.Options.UseForeColor = True
        Me.txtName.Properties.AppearanceReadOnly.BackColor = System.Drawing.Color.White
        Me.txtName.Properties.AppearanceReadOnly.BackColor2 = System.Drawing.Color.White
        Me.txtName.Properties.AppearanceReadOnly.BorderColor = System.Drawing.Color.White
        Me.txtName.Properties.AppearanceReadOnly.Options.UseBackColor = True
        Me.txtName.Properties.AppearanceReadOnly.Options.UseBorderColor = True
        Me.txtName.Size = New System.Drawing.Size(213, 20)
        Me.txtName.TabIndex = 2
        '
        'txtPatientID
        '
        Me.txtPatientID.Location = New System.Drawing.Point(94, 55)
        Me.txtPatientID.Name = "txtPatientID"
        Me.txtPatientID.Properties.Appearance.Font = New System.Drawing.Font("Segoe UI", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtPatientID.Properties.Appearance.ForeColor = System.Drawing.Color.FromArgb(CType(CType(27, Byte), Integer), CType(CType(41, Byte), Integer), CType(CType(62, Byte), Integer))
        Me.txtPatientID.Properties.Appearance.Options.UseFont = True
        Me.txtPatientID.Properties.Appearance.Options.UseForeColor = True
        Me.txtPatientID.Properties.AppearanceReadOnly.BackColor = System.Drawing.Color.White
        Me.txtPatientID.Properties.AppearanceReadOnly.BackColor2 = System.Drawing.Color.White
        Me.txtPatientID.Properties.AppearanceReadOnly.BorderColor = System.Drawing.Color.White
        Me.txtPatientID.Properties.AppearanceReadOnly.Options.UseBackColor = True
        Me.txtPatientID.Properties.AppearanceReadOnly.Options.UseBorderColor = True
        Me.txtPatientID.Size = New System.Drawing.Size(162, 20)
        Me.txtPatientID.TabIndex = 1
        '
        'cboSex
        '
        Me.cboSex.Location = New System.Drawing.Point(94, 101)
        Me.cboSex.Name = "cboSex"
        Me.cboSex.Properties.Appearance.Font = New System.Drawing.Font("Segoe UI", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cboSex.Properties.Appearance.ForeColor = System.Drawing.Color.FromArgb(CType(CType(27, Byte), Integer), CType(CType(41, Byte), Integer), CType(CType(62, Byte), Integer))
        Me.cboSex.Properties.Appearance.Options.UseFont = True
        Me.cboSex.Properties.Appearance.Options.UseForeColor = True
        Me.cboSex.Properties.AppearanceReadOnly.BackColor = System.Drawing.Color.White
        Me.cboSex.Properties.AppearanceReadOnly.BackColor2 = System.Drawing.Color.White
        Me.cboSex.Properties.AppearanceReadOnly.BorderColor = System.Drawing.Color.White
        Me.cboSex.Properties.AppearanceReadOnly.Options.UseBackColor = True
        Me.cboSex.Properties.AppearanceReadOnly.Options.UseBorderColor = True
        Me.cboSex.Properties.Buttons.AddRange(New DevExpress.XtraEditors.Controls.EditorButton() {New DevExpress.XtraEditors.Controls.EditorButton(DevExpress.XtraEditors.Controls.ButtonPredefines.Combo)})
        Me.cboSex.Properties.Items.AddRange(New Object() {"Male", "Female"})
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
        Me.BarManager.Items.AddRange(New DevExpress.XtraBars.BarItem() {Me.BarLargeButtonItem3, Me.btnClose, Me.btnValidate, Me.btnPrintNow, Me.btnViewPrint, Me.btnEdit, Me.btnResend})
        Me.BarManager.MainMenu = Me.Bar2
        Me.BarManager.MaxItemId = 16
        '
        'Bar2
        '
        Me.Bar2.BarName = "Main menu"
        Me.Bar2.DockCol = 0
        Me.Bar2.DockRow = 0
        Me.Bar2.DockStyle = DevExpress.XtraBars.BarDockStyle.Top
        Me.Bar2.FloatLocation = New System.Drawing.Point(213, 125)
        Me.Bar2.LinksPersistInfo.AddRange(New DevExpress.XtraBars.LinkPersistInfo() {New DevExpress.XtraBars.LinkPersistInfo(Me.btnViewPrint, True), New DevExpress.XtraBars.LinkPersistInfo(Me.btnEdit, True), New DevExpress.XtraBars.LinkPersistInfo(Me.btnValidate, True), New DevExpress.XtraBars.LinkPersistInfo(Me.btnClose, True), New DevExpress.XtraBars.LinkPersistInfo(Me.btnPrintNow, True), New DevExpress.XtraBars.LinkPersistInfo(Me.btnResend, True)})
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
        Me.btnViewPrint.ItemShortcut = New DevExpress.XtraBars.BarShortcut(System.Windows.Forms.Keys.F1)
        Me.btnViewPrint.Name = "btnViewPrint"
        '
        'btnEdit
        '
        Me.btnEdit.Caption = "Edit"
        Me.btnEdit.CaptionAlignment = DevExpress.XtraBars.BarItemCaptionAlignment.Right
        Me.btnEdit.Id = 14
        Me.btnEdit.ImageOptions.SvgImage = CType(resources.GetObject("btnEdit.ImageOptions.SvgImage"), DevExpress.Utils.Svg.SvgImage)
        Me.btnEdit.ImageOptions.SvgImageSize = New System.Drawing.Size(24, 24)
        Me.btnEdit.Name = "btnEdit"
        '
        'btnValidate
        '
        Me.btnValidate.Caption = "Validate"
        Me.btnValidate.CaptionAlignment = DevExpress.XtraBars.BarItemCaptionAlignment.Right
        Me.btnValidate.Enabled = False
        Me.btnValidate.Id = 11
        Me.btnValidate.ImageOptions.AllowGlyphSkinning = DevExpress.Utils.DefaultBoolean.[True]
        Me.btnValidate.ImageOptions.SvgImage = CType(resources.GetObject("btnValidate.ImageOptions.SvgImage"), DevExpress.Utils.Svg.SvgImage)
        Me.btnValidate.ImageOptions.SvgImageSize = New System.Drawing.Size(24, 24)
        Me.btnValidate.ItemShortcut = New DevExpress.XtraBars.BarShortcut(System.Windows.Forms.Keys.F4)
        Me.btnValidate.Name = "btnValidate"
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
        'btnPrintNow
        '
        Me.btnPrintNow.Caption = "Print"
        Me.btnPrintNow.CaptionAlignment = DevExpress.XtraBars.BarItemCaptionAlignment.Right
        Me.btnPrintNow.Id = 12
        Me.btnPrintNow.ImageOptions.AllowGlyphSkinning = DevExpress.Utils.DefaultBoolean.[True]
        Me.btnPrintNow.ImageOptions.SvgImage = CType(resources.GetObject("btnPrintNow.ImageOptions.SvgImage"), DevExpress.Utils.Svg.SvgImage)
        Me.btnPrintNow.ImageOptions.SvgImageSize = New System.Drawing.Size(24, 24)
        Me.btnPrintNow.ItemShortcut = New DevExpress.XtraBars.BarShortcut(System.Windows.Forms.Keys.F3)
        Me.btnPrintNow.Name = "btnPrintNow"
        '
        'btnResend
        '
        Me.btnResend.Caption = "Resend Result to HIS"
        Me.btnResend.CaptionAlignment = DevExpress.XtraBars.BarItemCaptionAlignment.Right
        Me.btnResend.Id = 15
        Me.btnResend.ImageOptions.SvgImage = CType(resources.GetObject("btnResend.ImageOptions.SvgImage"), DevExpress.Utils.Svg.SvgImage)
        Me.btnResend.ImageOptions.SvgImageSize = New System.Drawing.Size(24, 24)
        Me.btnResend.Name = "btnResend"
        Me.btnResend.Visibility = DevExpress.XtraBars.BarItemVisibility.Never
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
        Me.barDockControlBottom.Location = New System.Drawing.Point(0, 702)
        Me.barDockControlBottom.Manager = Me.BarManager
        Me.barDockControlBottom.Size = New System.Drawing.Size(1305, 0)
        '
        'barDockControlLeft
        '
        Me.barDockControlLeft.CausesValidation = False
        Me.barDockControlLeft.Dock = System.Windows.Forms.DockStyle.Left
        Me.barDockControlLeft.Location = New System.Drawing.Point(0, 36)
        Me.barDockControlLeft.Manager = Me.BarManager
        Me.barDockControlLeft.Size = New System.Drawing.Size(0, 666)
        '
        'barDockControlRight
        '
        Me.barDockControlRight.CausesValidation = False
        Me.barDockControlRight.Dock = System.Windows.Forms.DockStyle.Right
        Me.barDockControlRight.Location = New System.Drawing.Point(1305, 36)
        Me.barDockControlRight.Manager = Me.BarManager
        Me.barDockControlRight.Size = New System.Drawing.Size(0, 666)
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
        Me.dtBDate.Location = New System.Drawing.Point(94, 124)
        Me.dtBDate.MenuManager = Me.BarManager
        Me.dtBDate.Name = "dtBDate"
        Me.dtBDate.Properties.Appearance.Font = New System.Drawing.Font("Segoe UI", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.dtBDate.Properties.Appearance.Options.UseFont = True
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
        Me.txtClass.Location = New System.Drawing.Point(206, 147)
        Me.txtClass.Name = "txtClass"
        Me.txtClass.Properties.Appearance.Font = New System.Drawing.Font("Segoe UI", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtClass.Properties.Appearance.ForeColor = System.Drawing.Color.FromArgb(CType(CType(27, Byte), Integer), CType(CType(41, Byte), Integer), CType(CType(62, Byte), Integer))
        Me.txtClass.Properties.Appearance.Options.UseFont = True
        Me.txtClass.Properties.Appearance.Options.UseForeColor = True
        Me.txtClass.Properties.AppearanceReadOnly.BackColor = System.Drawing.Color.White
        Me.txtClass.Properties.AppearanceReadOnly.BackColor2 = System.Drawing.Color.White
        Me.txtClass.Properties.AppearanceReadOnly.BorderColor = System.Drawing.Color.White
        Me.txtClass.Properties.AppearanceReadOnly.Options.UseBackColor = True
        Me.txtClass.Properties.AppearanceReadOnly.Options.UseBorderColor = True
        Me.txtClass.Properties.Buttons.AddRange(New DevExpress.XtraEditors.Controls.EditorButton() {New DevExpress.XtraEditors.Controls.EditorButton(DevExpress.XtraEditors.Controls.ButtonPredefines.Combo)})
        Me.txtClass.Properties.Items.AddRange(New Object() {"Day(s)", "Week(s)", "Month(s)", "Year(s)", "NB"})
        Me.txtClass.Properties.TextEditStyle = DevExpress.XtraEditors.Controls.TextEditStyles.DisableTextEditor
        Me.txtClass.Size = New System.Drawing.Size(101, 20)
        Me.txtClass.TabIndex = 6
        '
        'txtComment
        '
        Me.txtComment.Dock = System.Windows.Forms.DockStyle.Fill
        Me.txtComment.Location = New System.Drawing.Point(2, 27)
        Me.txtComment.Name = "txtComment"
        Me.txtComment.Properties.Appearance.ForeColor = System.Drawing.Color.FromArgb(CType(CType(27, Byte), Integer), CType(CType(41, Byte), Integer), CType(CType(62, Byte), Integer))
        Me.txtComment.Properties.Appearance.Options.UseForeColor = True
        Me.txtComment.Properties.BorderStyle = DevExpress.XtraEditors.Controls.BorderStyles.NoBorder
        Me.txtComment.Size = New System.Drawing.Size(489, 79)
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
        Me.txtRemarks.Size = New System.Drawing.Size(443, 78)
        Me.txtRemarks.TabIndex = 21
        '
        'cboPatientType
        '
        Me.cboPatientType.Location = New System.Drawing.Point(94, 244)
        Me.cboPatientType.Name = "cboPatientType"
        Me.cboPatientType.Properties.Appearance.Font = New System.Drawing.Font("Segoe UI", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cboPatientType.Properties.Appearance.ForeColor = System.Drawing.Color.FromArgb(CType(CType(27, Byte), Integer), CType(CType(41, Byte), Integer), CType(CType(62, Byte), Integer))
        Me.cboPatientType.Properties.Appearance.Options.UseFont = True
        Me.cboPatientType.Properties.Appearance.Options.UseForeColor = True
        Me.cboPatientType.Properties.Buttons.AddRange(New DevExpress.XtraEditors.Controls.EditorButton() {New DevExpress.XtraEditors.Controls.EditorButton(DevExpress.XtraEditors.Controls.ButtonPredefines.Combo)})
        Me.cboPatientType.Properties.Items.AddRange(New Object() {"In Patient", "Out Patient"})
        Me.cboPatientType.Size = New System.Drawing.Size(213, 20)
        Me.cboPatientType.TabIndex = 17
        '
        'cboRoom
        '
        Me.cboRoom.Location = New System.Drawing.Point(94, 221)
        Me.cboRoom.Name = "cboRoom"
        Me.cboRoom.Properties.Appearance.Font = New System.Drawing.Font("Segoe UI", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cboRoom.Properties.Appearance.ForeColor = System.Drawing.Color.FromArgb(CType(CType(27, Byte), Integer), CType(CType(41, Byte), Integer), CType(CType(62, Byte), Integer))
        Me.cboRoom.Properties.Appearance.Options.UseFont = True
        Me.cboRoom.Properties.Appearance.Options.UseForeColor = True
        Me.cboRoom.Properties.Buttons.AddRange(New DevExpress.XtraEditors.Controls.EditorButton() {New DevExpress.XtraEditors.Controls.EditorButton(DevExpress.XtraEditors.Controls.ButtonPredefines.Combo)})
        Me.cboRoom.Size = New System.Drawing.Size(213, 20)
        Me.cboRoom.TabIndex = 16
        '
        'cboPhysician
        '
        Me.cboPhysician.Location = New System.Drawing.Point(94, 198)
        Me.cboPhysician.Name = "cboPhysician"
        Me.cboPhysician.Properties.Appearance.Font = New System.Drawing.Font("Segoe UI", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cboPhysician.Properties.Appearance.ForeColor = System.Drawing.Color.FromArgb(CType(CType(27, Byte), Integer), CType(CType(41, Byte), Integer), CType(CType(62, Byte), Integer))
        Me.cboPhysician.Properties.Appearance.Options.UseFont = True
        Me.cboPhysician.Properties.Appearance.Options.UseForeColor = True
        Me.cboPhysician.Properties.Buttons.AddRange(New DevExpress.XtraEditors.Controls.EditorButton() {New DevExpress.XtraEditors.Controls.EditorButton(DevExpress.XtraEditors.Controls.ButtonPredefines.Combo)})
        Me.cboPhysician.Size = New System.Drawing.Size(213, 20)
        Me.cboPhysician.TabIndex = 15
        '
        'cboRequest
        '
        Me.cboRequest.Location = New System.Drawing.Point(94, 175)
        Me.cboRequest.Name = "cboRequest"
        Me.cboRequest.Properties.Appearance.Font = New System.Drawing.Font("Segoe UI", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cboRequest.Properties.Appearance.ForeColor = System.Drawing.Color.FromArgb(CType(CType(27, Byte), Integer), CType(CType(41, Byte), Integer), CType(CType(62, Byte), Integer))
        Me.cboRequest.Properties.Appearance.Options.UseFont = True
        Me.cboRequest.Properties.Appearance.Options.UseForeColor = True
        Me.cboRequest.Properties.Buttons.AddRange(New DevExpress.XtraEditors.Controls.EditorButton() {New DevExpress.XtraEditors.Controls.EditorButton(DevExpress.XtraEditors.Controls.ButtonPredefines.Combo)})
        Me.cboRequest.Size = New System.Drawing.Size(213, 20)
        Me.cboRequest.TabIndex = 14
        '
        'txtORNo
        '
        Me.txtORNo.Location = New System.Drawing.Point(94, 79)
        Me.txtORNo.Name = "txtORNo"
        Me.txtORNo.Properties.Appearance.Font = New System.Drawing.Font("Segoe UI", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtORNo.Properties.Appearance.ForeColor = System.Drawing.Color.FromArgb(CType(CType(27, Byte), Integer), CType(CType(41, Byte), Integer), CType(CType(62, Byte), Integer))
        Me.txtORNo.Properties.Appearance.Options.UseFont = True
        Me.txtORNo.Properties.Appearance.Options.UseForeColor = True
        Me.txtORNo.Size = New System.Drawing.Size(213, 20)
        Me.txtORNo.TabIndex = 11
        '
        'txtAccession
        '
        Me.txtAccession.Location = New System.Drawing.Point(94, 33)
        Me.txtAccession.Name = "txtAccession"
        Me.txtAccession.Properties.Appearance.Font = New System.Drawing.Font("Segoe UI", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtAccession.Properties.Appearance.ForeColor = System.Drawing.Color.FromArgb(CType(CType(27, Byte), Integer), CType(CType(41, Byte), Integer), CType(CType(62, Byte), Integer))
        Me.txtAccession.Properties.Appearance.Options.UseFont = True
        Me.txtAccession.Properties.Appearance.Options.UseForeColor = True
        Me.txtAccession.Size = New System.Drawing.Size(213, 20)
        Me.txtAccession.TabIndex = 10
        '
        'txtContact
        '
        Me.txtContact.Location = New System.Drawing.Point(94, 216)
        Me.txtContact.Name = "txtContact"
        Me.txtContact.Properties.Appearance.Font = New System.Drawing.Font("Segoe UI", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtContact.Properties.Appearance.ForeColor = System.Drawing.Color.FromArgb(CType(CType(27, Byte), Integer), CType(CType(41, Byte), Integer), CType(CType(62, Byte), Integer))
        Me.txtContact.Properties.Appearance.Options.UseFont = True
        Me.txtContact.Properties.Appearance.Options.UseForeColor = True
        Me.txtContact.Properties.AppearanceReadOnly.BackColor = System.Drawing.Color.White
        Me.txtContact.Properties.AppearanceReadOnly.BackColor2 = System.Drawing.Color.White
        Me.txtContact.Properties.AppearanceReadOnly.BorderColor = System.Drawing.Color.White
        Me.txtContact.Properties.AppearanceReadOnly.Options.UseBackColor = True
        Me.txtContact.Properties.AppearanceReadOnly.Options.UseBorderColor = True
        Me.txtContact.Size = New System.Drawing.Size(213, 20)
        Me.txtContact.TabIndex = 9
        '
        'cboCS
        '
        Me.cboCS.Location = New System.Drawing.Point(94, 170)
        Me.cboCS.Name = "cboCS"
        Me.cboCS.Properties.Appearance.Font = New System.Drawing.Font("Segoe UI", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cboCS.Properties.Appearance.ForeColor = System.Drawing.Color.FromArgb(CType(CType(27, Byte), Integer), CType(CType(41, Byte), Integer), CType(CType(62, Byte), Integer))
        Me.cboCS.Properties.Appearance.Options.UseFont = True
        Me.cboCS.Properties.Appearance.Options.UseForeColor = True
        Me.cboCS.Properties.AppearanceReadOnly.BackColor = System.Drawing.Color.White
        Me.cboCS.Properties.AppearanceReadOnly.BackColor2 = System.Drawing.Color.White
        Me.cboCS.Properties.AppearanceReadOnly.BorderColor = System.Drawing.Color.White
        Me.cboCS.Properties.AppearanceReadOnly.Options.UseBackColor = True
        Me.cboCS.Properties.AppearanceReadOnly.Options.UseBorderColor = True
        Me.cboCS.Properties.Buttons.AddRange(New DevExpress.XtraEditors.Controls.EditorButton() {New DevExpress.XtraEditors.Controls.EditorButton(DevExpress.XtraEditors.Controls.ButtonPredefines.Combo)})
        Me.cboCS.Properties.Items.AddRange(New Object() {"Single", "Married", "Widow", "Widower"})
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
        Me.btnSearch.Location = New System.Drawing.Point(261, 55)
        Me.btnSearch.Margin = New System.Windows.Forms.Padding(2, 3, 2, 3)
        Me.btnSearch.Name = "btnSearch"
        Me.btnSearch.Size = New System.Drawing.Size(46, 20)
        Me.btnSearch.TabIndex = 160
        '
        'gcTest
        '
        Me.gcTest.AppearanceCaption.BackColor = System.Drawing.Color.FromArgb(CType(CType(16, Byte), Integer), CType(CType(110, Byte), Integer), CType(CType(190, Byte), Integer))
        Me.gcTest.AppearanceCaption.BackColor2 = System.Drawing.Color.FromArgb(CType(CType(16, Byte), Integer), CType(CType(110, Byte), Integer), CType(CType(190, Byte), Integer))
        Me.gcTest.AppearanceCaption.BorderColor = System.Drawing.Color.FromArgb(CType(CType(16, Byte), Integer), CType(CType(110, Byte), Integer), CType(CType(190, Byte), Integer))
        Me.gcTest.AppearanceCaption.ForeColor = System.Drawing.Color.White
        Me.gcTest.AppearanceCaption.Options.UseBackColor = True
        Me.gcTest.AppearanceCaption.Options.UseBorderColor = True
        Me.gcTest.AppearanceCaption.Options.UseForeColor = True
        Me.gcTest.CaptionImageOptions.AllowGlyphSkinning = True
        Me.gcTest.CaptionImageOptions.SvgImage = CType(resources.GetObject("gcTest.CaptionImageOptions.SvgImage"), DevExpress.Utils.Svg.SvgImage)
        Me.gcTest.CaptionImageOptions.SvgImageSize = New System.Drawing.Size(16, 16)
        Me.gcTest.Controls.Add(Me.dtResult)
        Me.gcTest.Controls.Add(Me.lblDiffCount)
        Me.gcTest.Controls.Add(Me.btnRetrive)
        Me.gcTest.Controls.Add(Me.btnPreview)
        Me.gcTest.Controls.Add(Me.btnAddTest)
        Me.gcTest.Enabled = False
        Me.gcTest.Location = New System.Drawing.Point(330, 44)
        Me.gcTest.LookAndFeel.SkinName = "The Bezier"
        Me.gcTest.LookAndFeel.UseDefaultLookAndFeel = False
        Me.gcTest.Name = "gcTest"
        Me.gcTest.Size = New System.Drawing.Size(946, 518)
        Me.gcTest.TabIndex = 175
        Me.gcTest.Text = "Test Result"
        '
        'dtResult
        '
        Me.dtResult.Dock = System.Windows.Forms.DockStyle.Fill
        Me.dtResult.Location = New System.Drawing.Point(2, 27)
        Me.dtResult.MainView = Me.GridView
        Me.dtResult.MenuManager = Me.BarManager
        Me.dtResult.Name = "dtResult"
        Me.dtResult.Size = New System.Drawing.Size(942, 469)
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
        'lblDiffCount
        '
        Me.lblDiffCount.Appearance.BackColor = System.Drawing.Color.White
        Me.lblDiffCount.Appearance.Font = New System.Drawing.Font("Segoe UI", 11.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblDiffCount.Appearance.ForeColor = System.Drawing.Color.FromArgb(CType(CType(27, Byte), Integer), CType(CType(41, Byte), Integer), CType(CType(62, Byte), Integer))
        Me.lblDiffCount.Appearance.Options.UseBackColor = True
        Me.lblDiffCount.Appearance.Options.UseFont = True
        Me.lblDiffCount.Appearance.Options.UseForeColor = True
        Me.lblDiffCount.AutoSizeMode = DevExpress.XtraEditors.LabelAutoSizeMode.Vertical
        Me.lblDiffCount.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.lblDiffCount.Location = New System.Drawing.Point(2, 496)
        Me.lblDiffCount.LookAndFeel.SkinName = "Visual Studio 2013 Blue"
        Me.lblDiffCount.LookAndFeel.UseDefaultLookAndFeel = False
        Me.lblDiffCount.Name = "lblDiffCount"
        Me.lblDiffCount.Size = New System.Drawing.Size(942, 20)
        Me.lblDiffCount.TabIndex = 145
        Me.lblDiffCount.Text = "Total Dif Count: "
        '
        'btnRetrive
        '
        Me.btnRetrive.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnRetrive.ImageOptions.SvgImage = CType(resources.GetObject("btnRetrive.ImageOptions.SvgImage"), DevExpress.Utils.Svg.SvgImage)
        Me.btnRetrive.ImageOptions.SvgImageSize = New System.Drawing.Size(20, 20)
        Me.btnRetrive.Location = New System.Drawing.Point(552, 1)
        Me.btnRetrive.Name = "btnRetrive"
        Me.btnRetrive.Size = New System.Drawing.Size(135, 23)
        Me.btnRetrive.TabIndex = 62
        Me.btnRetrive.Text = "Retrieve Re-run"
        Me.btnRetrive.Visible = False
        '
        'btnPreview
        '
        Me.btnPreview.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnPreview.ImageOptions.SvgImage = CType(resources.GetObject("btnPreview.ImageOptions.SvgImage"), DevExpress.Utils.Svg.SvgImage)
        Me.btnPreview.ImageOptions.SvgImageSize = New System.Drawing.Size(20, 20)
        Me.btnPreview.Location = New System.Drawing.Point(693, 1)
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
        Me.btnAddTest.Location = New System.Drawing.Point(834, 1)
        Me.btnAddTest.Name = "btnAddTest"
        Me.btnAddTest.Size = New System.Drawing.Size(108, 23)
        Me.btnAddTest.TabIndex = 60
        Me.btnAddTest.Text = "Add Test"
        '
        'LabelControl26
        '
        Me.LabelControl26.Appearance.Font = New System.Drawing.Font("Segoe UI", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.LabelControl26.Appearance.ForeColor = System.Drawing.Color.FromArgb(CType(CType(27, Byte), Integer), CType(CType(41, Byte), Integer), CType(CType(62, Byte), Integer))
        Me.LabelControl26.Appearance.Options.UseFont = True
        Me.LabelControl26.Appearance.Options.UseForeColor = True
        Me.LabelControl26.Location = New System.Drawing.Point(8, 82)
        Me.LabelControl26.LookAndFeel.SkinName = "Visual Studio 2013 Blue"
        Me.LabelControl26.LookAndFeel.UseDefaultLookAndFeel = False
        Me.LabelControl26.Name = "LabelControl26"
        Me.LabelControl26.Size = New System.Drawing.Size(62, 13)
        Me.LabelControl26.TabIndex = 50
        Me.LabelControl26.Text = "Pathologist:"
        '
        'LabelControl27
        '
        Me.LabelControl27.Appearance.Font = New System.Drawing.Font("Segoe UI", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.LabelControl27.Appearance.ForeColor = System.Drawing.Color.FromArgb(CType(CType(27, Byte), Integer), CType(CType(41, Byte), Integer), CType(CType(62, Byte), Integer))
        Me.LabelControl27.Appearance.Options.UseFont = True
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
        Me.LabelControl29.Appearance.Font = New System.Drawing.Font("Segoe UI", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.LabelControl29.Appearance.ForeColor = System.Drawing.Color.FromArgb(CType(CType(27, Byte), Integer), CType(CType(41, Byte), Integer), CType(CType(62, Byte), Integer))
        Me.LabelControl29.Appearance.Options.UseFont = True
        Me.LabelControl29.Appearance.Options.UseForeColor = True
        Me.LabelControl29.Location = New System.Drawing.Point(8, 36)
        Me.LabelControl29.LookAndFeel.SkinName = "Visual Studio 2013 Blue"
        Me.LabelControl29.LookAndFeel.UseDefaultLookAndFeel = False
        Me.LabelControl29.Name = "LabelControl29"
        Me.LabelControl29.Size = New System.Drawing.Size(70, 13)
        Me.LabelControl29.TabIndex = 48
        Me.LabelControl29.Text = "Performed By:"
        '
        'gcAdditional
        '
        Me.gcAdditional.AppearanceCaption.BackColor = System.Drawing.Color.FromArgb(CType(CType(16, Byte), Integer), CType(CType(110, Byte), Integer), CType(CType(190, Byte), Integer))
        Me.gcAdditional.AppearanceCaption.BackColor2 = System.Drawing.Color.FromArgb(CType(CType(16, Byte), Integer), CType(CType(110, Byte), Integer), CType(CType(190, Byte), Integer))
        Me.gcAdditional.AppearanceCaption.BorderColor = System.Drawing.Color.FromArgb(CType(CType(16, Byte), Integer), CType(CType(110, Byte), Integer), CType(CType(190, Byte), Integer))
        Me.gcAdditional.AppearanceCaption.ForeColor = System.Drawing.Color.White
        Me.gcAdditional.AppearanceCaption.Options.UseBackColor = True
        Me.gcAdditional.AppearanceCaption.Options.UseBorderColor = True
        Me.gcAdditional.AppearanceCaption.Options.UseForeColor = True
        Me.gcAdditional.CaptionImageOptions.AllowGlyphSkinning = True
        Me.gcAdditional.CaptionImageOptions.SvgImage = CType(resources.GetObject("gcAdditional.CaptionImageOptions.SvgImage"), DevExpress.Utils.Svg.SvgImage)
        Me.gcAdditional.CaptionImageOptions.SvgImageSize = New System.Drawing.Size(16, 16)
        Me.gcAdditional.Controls.Add(Me.tmTimeReceived)
        Me.gcAdditional.Controls.Add(Me.tmTimeReleased)
        Me.gcAdditional.Controls.Add(Me.LabelControl2)
        Me.gcAdditional.Controls.Add(Me.LabelControl1)
        Me.gcAdditional.Controls.Add(Me.txtChargeSlip)
        Me.gcAdditional.Controls.Add(Me.dtReceived)
        Me.gcAdditional.Controls.Add(Me.LabelControl32)
        Me.gcAdditional.Controls.Add(Me.LabelControl33)
        Me.gcAdditional.Controls.Add(Me.txtAccession)
        Me.gcAdditional.Controls.Add(Me.LabelControl34)
        Me.gcAdditional.Controls.Add(Me.cboPatientType)
        Me.gcAdditional.Controls.Add(Me.LabelControl35)
        Me.gcAdditional.Controls.Add(Me.txtORNo)
        Me.gcAdditional.Controls.Add(Me.LabelControl36)
        Me.gcAdditional.Controls.Add(Me.LabelControl37)
        Me.gcAdditional.Controls.Add(Me.LabelControl38)
        Me.gcAdditional.Controls.Add(Me.cboRoom)
        Me.gcAdditional.Controls.Add(Me.LabelControl41)
        Me.gcAdditional.Controls.Add(Me.cboPhysician)
        Me.gcAdditional.Controls.Add(Me.cboRequest)
        Me.gcAdditional.Enabled = False
        Me.gcAdditional.Location = New System.Drawing.Point(12, 292)
        Me.gcAdditional.LookAndFeel.SkinName = "The Bezier"
        Me.gcAdditional.LookAndFeel.UseDefaultLookAndFeel = False
        Me.gcAdditional.Name = "gcAdditional"
        Me.gcAdditional.Size = New System.Drawing.Size(312, 270)
        Me.gcAdditional.TabIndex = 174
        Me.gcAdditional.Text = "Additional Details"
        '
        'tmTimeReceived
        '
        Me.tmTimeReceived.EditValue = Nothing
        Me.tmTimeReceived.Location = New System.Drawing.Point(94, 126)
        Me.tmTimeReceived.MenuManager = Me.BarManager
        Me.tmTimeReceived.Name = "tmTimeReceived"
        Me.tmTimeReceived.Properties.AppearanceReadOnly.BackColor = System.Drawing.Color.White
        Me.tmTimeReceived.Properties.AppearanceReadOnly.BackColor2 = System.Drawing.Color.White
        Me.tmTimeReceived.Properties.AppearanceReadOnly.Options.UseBackColor = True
        Me.tmTimeReceived.Properties.Buttons.AddRange(New DevExpress.XtraEditors.Controls.EditorButton() {New DevExpress.XtraEditors.Controls.EditorButton(DevExpress.XtraEditors.Controls.ButtonPredefines.Combo)})
        Me.tmTimeReceived.Properties.CalendarTimeProperties.Buttons.AddRange(New DevExpress.XtraEditors.Controls.EditorButton() {New DevExpress.XtraEditors.Controls.EditorButton(DevExpress.XtraEditors.Controls.ButtonPredefines.Combo)})
        Me.tmTimeReceived.Size = New System.Drawing.Size(213, 20)
        Me.tmTimeReceived.TabIndex = 169
        '
        'tmTimeReleased
        '
        Me.tmTimeReleased.CalendarFont = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.tmTimeReleased.CustomFormat = "MM/dd/yyyy hh:mm:ss tt"
        Me.tmTimeReleased.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.tmTimeReleased.Location = New System.Drawing.Point(94, 149)
        Me.tmTimeReleased.Name = "tmTimeReleased"
        Me.tmTimeReleased.Size = New System.Drawing.Size(213, 22)
        Me.tmTimeReleased.TabIndex = 167
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
        Me.LabelControl2.TabIndex = 168
        Me.LabelControl2.Text = "Date Released:"
        '
        'LabelControl1
        '
        Me.LabelControl1.Appearance.Font = New System.Drawing.Font("Segoe UI", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.LabelControl1.Appearance.ForeColor = System.Drawing.Color.FromArgb(CType(CType(27, Byte), Integer), CType(CType(41, Byte), Integer), CType(CType(62, Byte), Integer))
        Me.LabelControl1.Appearance.Options.UseFont = True
        Me.LabelControl1.Appearance.Options.UseForeColor = True
        Me.LabelControl1.Location = New System.Drawing.Point(8, 59)
        Me.LabelControl1.LookAndFeel.SkinName = "Visual Studio 2013 Blue"
        Me.LabelControl1.LookAndFeel.UseDefaultLookAndFeel = False
        Me.LabelControl1.Name = "LabelControl1"
        Me.LabelControl1.Size = New System.Drawing.Size(72, 13)
        Me.LabelControl1.TabIndex = 164
        Me.LabelControl1.Text = "Charge Slip #:"
        '
        'txtChargeSlip
        '
        Me.txtChargeSlip.Location = New System.Drawing.Point(94, 56)
        Me.txtChargeSlip.Name = "txtChargeSlip"
        Me.txtChargeSlip.Properties.Appearance.Font = New System.Drawing.Font("Segoe UI", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtChargeSlip.Properties.Appearance.ForeColor = System.Drawing.Color.FromArgb(CType(CType(27, Byte), Integer), CType(CType(41, Byte), Integer), CType(CType(62, Byte), Integer))
        Me.txtChargeSlip.Properties.Appearance.Options.UseFont = True
        Me.txtChargeSlip.Properties.Appearance.Options.UseForeColor = True
        Me.txtChargeSlip.Size = New System.Drawing.Size(213, 20)
        Me.txtChargeSlip.TabIndex = 163
        '
        'dtReceived
        '
        Me.dtReceived.EditValue = Nothing
        Me.dtReceived.Location = New System.Drawing.Point(94, 102)
        Me.dtReceived.MenuManager = Me.BarManager
        Me.dtReceived.Name = "dtReceived"
        Me.dtReceived.Properties.AppearanceReadOnly.BackColor = System.Drawing.Color.White
        Me.dtReceived.Properties.AppearanceReadOnly.BackColor2 = System.Drawing.Color.White
        Me.dtReceived.Properties.AppearanceReadOnly.Options.UseBackColor = True
        Me.dtReceived.Properties.Buttons.AddRange(New DevExpress.XtraEditors.Controls.EditorButton() {New DevExpress.XtraEditors.Controls.EditorButton(DevExpress.XtraEditors.Controls.ButtonPredefines.Combo)})
        Me.dtReceived.Properties.CalendarTimeProperties.Buttons.AddRange(New DevExpress.XtraEditors.Controls.EditorButton() {New DevExpress.XtraEditors.Controls.EditorButton(DevExpress.XtraEditors.Controls.ButtonPredefines.Combo)})
        Me.dtReceived.Size = New System.Drawing.Size(213, 20)
        Me.dtReceived.TabIndex = 161
        '
        'LabelControl32
        '
        Me.LabelControl32.Appearance.Font = New System.Drawing.Font("Segoe UI", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.LabelControl32.Appearance.ForeColor = System.Drawing.Color.FromArgb(CType(CType(27, Byte), Integer), CType(CType(41, Byte), Integer), CType(CType(62, Byte), Integer))
        Me.LabelControl32.Appearance.Options.UseFont = True
        Me.LabelControl32.Appearance.Options.UseForeColor = True
        Me.LabelControl32.Location = New System.Drawing.Point(8, 247)
        Me.LabelControl32.LookAndFeel.SkinName = "Visual Studio 2013 Blue"
        Me.LabelControl32.LookAndFeel.UseDefaultLookAndFeel = False
        Me.LabelControl32.Name = "LabelControl32"
        Me.LabelControl32.Size = New System.Drawing.Size(66, 13)
        Me.LabelControl32.TabIndex = 143
        Me.LabelControl32.Text = "Patient Type:"
        '
        'LabelControl33
        '
        Me.LabelControl33.Appearance.Font = New System.Drawing.Font("Segoe UI", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.LabelControl33.Appearance.ForeColor = System.Drawing.Color.FromArgb(CType(CType(27, Byte), Integer), CType(CType(41, Byte), Integer), CType(CType(62, Byte), Integer))
        Me.LabelControl33.Appearance.Options.UseFont = True
        Me.LabelControl33.Appearance.Options.UseForeColor = True
        Me.LabelControl33.Location = New System.Drawing.Point(8, 224)
        Me.LabelControl33.LookAndFeel.SkinName = "Visual Studio 2013 Blue"
        Me.LabelControl33.LookAndFeel.UseDefaultLookAndFeel = False
        Me.LabelControl33.Name = "LabelControl33"
        Me.LabelControl33.Size = New System.Drawing.Size(65, 13)
        Me.LabelControl33.TabIndex = 28
        Me.LabelControl33.Text = "Room/Ward:"
        '
        'LabelControl34
        '
        Me.LabelControl34.Appearance.Font = New System.Drawing.Font("Segoe UI", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.LabelControl34.Appearance.ForeColor = System.Drawing.Color.FromArgb(CType(CType(27, Byte), Integer), CType(CType(41, Byte), Integer), CType(CType(62, Byte), Integer))
        Me.LabelControl34.Appearance.Options.UseFont = True
        Me.LabelControl34.Appearance.Options.UseForeColor = True
        Me.LabelControl34.Location = New System.Drawing.Point(8, 82)
        Me.LabelControl34.LookAndFeel.SkinName = "Visual Studio 2013 Blue"
        Me.LabelControl34.LookAndFeel.UseDefaultLookAndFeel = False
        Me.LabelControl34.Name = "LabelControl34"
        Me.LabelControl34.Size = New System.Drawing.Size(40, 13)
        Me.LabelControl34.TabIndex = 160
        Me.LabelControl34.Text = "OR No.:"
        '
        'LabelControl35
        '
        Me.LabelControl35.Appearance.Font = New System.Drawing.Font("Segoe UI", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.LabelControl35.Appearance.ForeColor = System.Drawing.Color.FromArgb(CType(CType(27, Byte), Integer), CType(CType(41, Byte), Integer), CType(CType(62, Byte), Integer))
        Me.LabelControl35.Appearance.Options.UseFont = True
        Me.LabelControl35.Appearance.Options.UseForeColor = True
        Me.LabelControl35.Location = New System.Drawing.Point(8, 201)
        Me.LabelControl35.LookAndFeel.SkinName = "Visual Studio 2013 Blue"
        Me.LabelControl35.LookAndFeel.UseDefaultLookAndFeel = False
        Me.LabelControl35.Name = "LabelControl35"
        Me.LabelControl35.Size = New System.Drawing.Size(72, 13)
        Me.LabelControl35.TabIndex = 30
        Me.LabelControl35.Text = "Requested By:"
        '
        'LabelControl36
        '
        Me.LabelControl36.Appearance.Font = New System.Drawing.Font("Segoe UI", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.LabelControl36.Appearance.ForeColor = System.Drawing.Color.FromArgb(CType(CType(27, Byte), Integer), CType(CType(41, Byte), Integer), CType(CType(62, Byte), Integer))
        Me.LabelControl36.Appearance.Options.UseFont = True
        Me.LabelControl36.Appearance.Options.UseForeColor = True
        Me.LabelControl36.Location = New System.Drawing.Point(8, 178)
        Me.LabelControl36.LookAndFeel.SkinName = "Visual Studio 2013 Blue"
        Me.LabelControl36.LookAndFeel.UseDefaultLookAndFeel = False
        Me.LabelControl36.Name = "LabelControl36"
        Me.LabelControl36.Size = New System.Drawing.Size(45, 13)
        Me.LabelControl36.TabIndex = 139
        Me.LabelControl36.Text = "Request:"
        '
        'LabelControl37
        '
        Me.LabelControl37.Appearance.Font = New System.Drawing.Font("Segoe UI", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.LabelControl37.Appearance.ForeColor = System.Drawing.Color.FromArgb(CType(CType(27, Byte), Integer), CType(CType(41, Byte), Integer), CType(CType(62, Byte), Integer))
        Me.LabelControl37.Appearance.Options.UseFont = True
        Me.LabelControl37.Appearance.Options.UseForeColor = True
        Me.LabelControl37.Location = New System.Drawing.Point(8, 129)
        Me.LabelControl37.LookAndFeel.SkinName = "Visual Studio 2013 Blue"
        Me.LabelControl37.LookAndFeel.UseDefaultLookAndFeel = False
        Me.LabelControl37.Name = "LabelControl37"
        Me.LabelControl37.Size = New System.Drawing.Size(75, 13)
        Me.LabelControl37.TabIndex = 152
        Me.LabelControl37.Text = "Time Received:"
        '
        'LabelControl38
        '
        Me.LabelControl38.Appearance.Font = New System.Drawing.Font("Segoe UI", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.LabelControl38.Appearance.ForeColor = System.Drawing.Color.FromArgb(CType(CType(27, Byte), Integer), CType(CType(41, Byte), Integer), CType(CType(62, Byte), Integer))
        Me.LabelControl38.Appearance.Options.UseFont = True
        Me.LabelControl38.Appearance.Options.UseForeColor = True
        Me.LabelControl38.Location = New System.Drawing.Point(8, 36)
        Me.LabelControl38.LookAndFeel.SkinName = "Visual Studio 2013 Blue"
        Me.LabelControl38.LookAndFeel.UseDefaultLookAndFeel = False
        Me.LabelControl38.Name = "LabelControl38"
        Me.LabelControl38.Size = New System.Drawing.Size(77, 13)
        Me.LabelControl38.TabIndex = 158
        Me.LabelControl38.Text = "HIS Tracking #:"
        '
        'LabelControl41
        '
        Me.LabelControl41.Appearance.Font = New System.Drawing.Font("Segoe UI", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.LabelControl41.Appearance.ForeColor = System.Drawing.Color.FromArgb(CType(CType(27, Byte), Integer), CType(CType(41, Byte), Integer), CType(CType(62, Byte), Integer))
        Me.LabelControl41.Appearance.Options.UseFont = True
        Me.LabelControl41.Appearance.Options.UseForeColor = True
        Me.LabelControl41.Location = New System.Drawing.Point(8, 105)
        Me.LabelControl41.LookAndFeel.SkinName = "Visual Studio 2013 Blue"
        Me.LabelControl41.LookAndFeel.UseDefaultLookAndFeel = False
        Me.LabelControl41.Name = "LabelControl41"
        Me.LabelControl41.Size = New System.Drawing.Size(75, 13)
        Me.LabelControl41.TabIndex = 136
        Me.LabelControl41.Text = "Date Received:"
        '
        'gcPatient
        '
        Me.gcPatient.AppearanceCaption.BackColor = System.Drawing.Color.FromArgb(CType(CType(16, Byte), Integer), CType(CType(110, Byte), Integer), CType(CType(190, Byte), Integer))
        Me.gcPatient.AppearanceCaption.BackColor2 = System.Drawing.Color.FromArgb(CType(CType(16, Byte), Integer), CType(CType(110, Byte), Integer), CType(CType(190, Byte), Integer))
        Me.gcPatient.AppearanceCaption.BorderColor = System.Drawing.Color.FromArgb(CType(CType(16, Byte), Integer), CType(CType(110, Byte), Integer), CType(CType(190, Byte), Integer))
        Me.gcPatient.AppearanceCaption.ForeColor = System.Drawing.Color.White
        Me.gcPatient.AppearanceCaption.Options.UseBackColor = True
        Me.gcPatient.AppearanceCaption.Options.UseBorderColor = True
        Me.gcPatient.AppearanceCaption.Options.UseForeColor = True
        Me.gcPatient.CaptionImageOptions.AllowGlyphSkinning = True
        Me.gcPatient.CaptionImageOptions.SvgImage = CType(resources.GetObject("gcPatient.CaptionImageOptions.SvgImage"), DevExpress.Utils.Svg.SvgImage)
        Me.gcPatient.CaptionImageOptions.SvgImageSize = New System.Drawing.Size(16, 16)
        Me.gcPatient.Controls.Add(Me.LabelControl42)
        Me.gcPatient.Controls.Add(Me.cboCS)
        Me.gcPatient.Controls.Add(Me.LabelControl43)
        Me.gcPatient.Controls.Add(Me.btnSearch)
        Me.gcPatient.Controls.Add(Me.LabelControl44)
        Me.gcPatient.Controls.Add(Me.txtSampleID)
        Me.gcPatient.Controls.Add(Me.txtContact)
        Me.gcPatient.Controls.Add(Me.LabelControl45)
        Me.gcPatient.Controls.Add(Me.txtClass)
        Me.gcPatient.Controls.Add(Me.LabelControl46)
        Me.gcPatient.Controls.Add(Me.LabelControl47)
        Me.gcPatient.Controls.Add(Me.LabelControl48)
        Me.gcPatient.Controls.Add(Me.LabelControl49)
        Me.gcPatient.Controls.Add(Me.txtAge)
        Me.gcPatient.Controls.Add(Me.LabelControl50)
        Me.gcPatient.Controls.Add(Me.txtPatientID)
        Me.gcPatient.Controls.Add(Me.cboSex)
        Me.gcPatient.Controls.Add(Me.txtName)
        Me.gcPatient.Controls.Add(Me.dtBDate)
        Me.gcPatient.Controls.Add(Me.txtAddress)
        Me.gcPatient.Enabled = False
        Me.gcPatient.Location = New System.Drawing.Point(12, 44)
        Me.gcPatient.Name = "gcPatient"
        Me.gcPatient.Size = New System.Drawing.Size(312, 242)
        Me.gcPatient.TabIndex = 173
        Me.gcPatient.Text = "Patient Details"
        '
        'LabelControl42
        '
        Me.LabelControl42.Appearance.Font = New System.Drawing.Font("Segoe UI", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.LabelControl42.Appearance.ForeColor = System.Drawing.Color.FromArgb(CType(CType(27, Byte), Integer), CType(CType(41, Byte), Integer), CType(CType(62, Byte), Integer))
        Me.LabelControl42.Appearance.Options.UseFont = True
        Me.LabelControl42.Appearance.Options.UseForeColor = True
        Me.LabelControl42.Location = New System.Drawing.Point(9, 173)
        Me.LabelControl42.LookAndFeel.SkinName = "Visual Studio 2013 Blue"
        Me.LabelControl42.LookAndFeel.UseDefaultLookAndFeel = False
        Me.LabelControl42.Name = "LabelControl42"
        Me.LabelControl42.Size = New System.Drawing.Size(59, 13)
        Me.LabelControl42.TabIndex = 161
        Me.LabelControl42.Text = "Civil Status:"
        '
        'LabelControl43
        '
        Me.LabelControl43.Appearance.Font = New System.Drawing.Font("Segoe UI", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.LabelControl43.Appearance.ForeColor = System.Drawing.Color.FromArgb(CType(CType(27, Byte), Integer), CType(CType(41, Byte), Integer), CType(CType(62, Byte), Integer))
        Me.LabelControl43.Appearance.Options.UseFont = True
        Me.LabelControl43.Appearance.Options.UseForeColor = True
        Me.LabelControl43.Location = New System.Drawing.Point(9, 35)
        Me.LabelControl43.LookAndFeel.SkinName = "Visual Studio 2013 Blue"
        Me.LabelControl43.LookAndFeel.UseDefaultLookAndFeel = False
        Me.LabelControl43.Name = "LabelControl43"
        Me.LabelControl43.Size = New System.Drawing.Size(54, 13)
        Me.LabelControl43.TabIndex = 22
        Me.LabelControl43.Text = "Sample ID:"
        '
        'LabelControl44
        '
        Me.LabelControl44.Appearance.Font = New System.Drawing.Font("Segoe UI", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.LabelControl44.Appearance.ForeColor = System.Drawing.Color.FromArgb(CType(CType(27, Byte), Integer), CType(CType(41, Byte), Integer), CType(CType(62, Byte), Integer))
        Me.LabelControl44.Appearance.Options.UseFont = True
        Me.LabelControl44.Appearance.Options.UseForeColor = True
        Me.LabelControl44.Location = New System.Drawing.Point(9, 126)
        Me.LabelControl44.LookAndFeel.SkinName = "Visual Studio 2013 Blue"
        Me.LabelControl44.LookAndFeel.UseDefaultLookAndFeel = False
        Me.LabelControl44.Name = "LabelControl44"
        Me.LabelControl44.Size = New System.Drawing.Size(68, 13)
        Me.LabelControl44.TabIndex = 85
        Me.LabelControl44.Text = "Date of Birth:"
        '
        'LabelControl45
        '
        Me.LabelControl45.Appearance.Font = New System.Drawing.Font("Segoe UI", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.LabelControl45.Appearance.ForeColor = System.Drawing.Color.FromArgb(CType(CType(27, Byte), Integer), CType(CType(41, Byte), Integer), CType(CType(62, Byte), Integer))
        Me.LabelControl45.Appearance.Options.UseFont = True
        Me.LabelControl45.Appearance.Options.UseForeColor = True
        Me.LabelControl45.Location = New System.Drawing.Point(9, 58)
        Me.LabelControl45.LookAndFeel.SkinName = "Visual Studio 2013 Blue"
        Me.LabelControl45.LookAndFeel.UseDefaultLookAndFeel = False
        Me.LabelControl45.Name = "LabelControl45"
        Me.LabelControl45.Size = New System.Drawing.Size(53, 13)
        Me.LabelControl45.TabIndex = 12
        Me.LabelControl45.Text = "Patient ID:"
        '
        'LabelControl46
        '
        Me.LabelControl46.Appearance.Font = New System.Drawing.Font("Segoe UI", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.LabelControl46.Appearance.ForeColor = System.Drawing.Color.FromArgb(CType(CType(27, Byte), Integer), CType(CType(41, Byte), Integer), CType(CType(62, Byte), Integer))
        Me.LabelControl46.Appearance.Options.UseFont = True
        Me.LabelControl46.Appearance.Options.UseForeColor = True
        Me.LabelControl46.Location = New System.Drawing.Point(9, 219)
        Me.LabelControl46.LookAndFeel.SkinName = "Visual Studio 2013 Blue"
        Me.LabelControl46.LookAndFeel.UseDefaultLookAndFeel = False
        Me.LabelControl46.Name = "LabelControl46"
        Me.LabelControl46.Size = New System.Drawing.Size(64, 13)
        Me.LabelControl46.TabIndex = 159
        Me.LabelControl46.Text = "Contact No.:"
        '
        'LabelControl47
        '
        Me.LabelControl47.Appearance.Font = New System.Drawing.Font("Segoe UI", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.LabelControl47.Appearance.ForeColor = System.Drawing.Color.FromArgb(CType(CType(27, Byte), Integer), CType(CType(41, Byte), Integer), CType(CType(62, Byte), Integer))
        Me.LabelControl47.Appearance.Options.UseFont = True
        Me.LabelControl47.Appearance.Options.UseForeColor = True
        Me.LabelControl47.Location = New System.Drawing.Point(9, 81)
        Me.LabelControl47.LookAndFeel.SkinName = "Visual Studio 2013 Blue"
        Me.LabelControl47.LookAndFeel.UseDefaultLookAndFeel = False
        Me.LabelControl47.Name = "LabelControl47"
        Me.LabelControl47.Size = New System.Drawing.Size(71, 13)
        Me.LabelControl47.TabIndex = 14
        Me.LabelControl47.Text = "Patient Name:"
        '
        'LabelControl48
        '
        Me.LabelControl48.Appearance.Font = New System.Drawing.Font("Segoe UI", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.LabelControl48.Appearance.ForeColor = System.Drawing.Color.FromArgb(CType(CType(27, Byte), Integer), CType(CType(41, Byte), Integer), CType(CType(62, Byte), Integer))
        Me.LabelControl48.Appearance.Options.UseFont = True
        Me.LabelControl48.Appearance.Options.UseForeColor = True
        Me.LabelControl48.Location = New System.Drawing.Point(9, 150)
        Me.LabelControl48.LookAndFeel.SkinName = "Visual Studio 2013 Blue"
        Me.LabelControl48.LookAndFeel.UseDefaultLookAndFeel = False
        Me.LabelControl48.Name = "LabelControl48"
        Me.LabelControl48.Size = New System.Drawing.Size(23, 13)
        Me.LabelControl48.TabIndex = 36
        Me.LabelControl48.Text = "Age:"
        '
        'LabelControl49
        '
        Me.LabelControl49.Appearance.Font = New System.Drawing.Font("Segoe UI", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.LabelControl49.Appearance.ForeColor = System.Drawing.Color.FromArgb(CType(CType(27, Byte), Integer), CType(CType(41, Byte), Integer), CType(CType(62, Byte), Integer))
        Me.LabelControl49.Appearance.Options.UseFont = True
        Me.LabelControl49.Appearance.Options.UseForeColor = True
        Me.LabelControl49.Location = New System.Drawing.Point(9, 196)
        Me.LabelControl49.LookAndFeel.SkinName = "Visual Studio 2013 Blue"
        Me.LabelControl49.LookAndFeel.UseDefaultLookAndFeel = False
        Me.LabelControl49.Name = "LabelControl49"
        Me.LabelControl49.Size = New System.Drawing.Size(44, 13)
        Me.LabelControl49.TabIndex = 157
        Me.LabelControl49.Text = "Address:"
        '
        'LabelControl50
        '
        Me.LabelControl50.Appearance.Font = New System.Drawing.Font("Segoe UI", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.LabelControl50.Appearance.ForeColor = System.Drawing.Color.FromArgb(CType(CType(27, Byte), Integer), CType(CType(41, Byte), Integer), CType(CType(62, Byte), Integer))
        Me.LabelControl50.Appearance.Options.UseFont = True
        Me.LabelControl50.Appearance.Options.UseForeColor = True
        Me.LabelControl50.Location = New System.Drawing.Point(9, 104)
        Me.LabelControl50.LookAndFeel.SkinName = "Visual Studio 2013 Blue"
        Me.LabelControl50.LookAndFeel.UseDefaultLookAndFeel = False
        Me.LabelControl50.Name = "LabelControl50"
        Me.LabelControl50.Size = New System.Drawing.Size(20, 13)
        Me.LabelControl50.TabIndex = 20
        Me.LabelControl50.Text = "Sex:"
        '
        'txtAddress
        '
        Me.txtAddress.Location = New System.Drawing.Point(94, 193)
        Me.txtAddress.Name = "txtAddress"
        Me.txtAddress.Properties.Appearance.Font = New System.Drawing.Font("Segoe UI", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtAddress.Properties.Appearance.ForeColor = System.Drawing.Color.FromArgb(CType(CType(27, Byte), Integer), CType(CType(41, Byte), Integer), CType(CType(62, Byte), Integer))
        Me.txtAddress.Properties.Appearance.Options.UseFont = True
        Me.txtAddress.Properties.Appearance.Options.UseForeColor = True
        Me.txtAddress.Properties.AppearanceDisabled.ForeColor = System.Drawing.Color.FromArgb(CType(CType(27, Byte), Integer), CType(CType(41, Byte), Integer), CType(CType(62, Byte), Integer))
        Me.txtAddress.Properties.AppearanceDisabled.Options.UseForeColor = True
        Me.txtAddress.Properties.AppearanceReadOnly.BackColor = System.Drawing.Color.White
        Me.txtAddress.Properties.AppearanceReadOnly.BackColor2 = System.Drawing.Color.White
        Me.txtAddress.Properties.AppearanceReadOnly.BorderColor = System.Drawing.Color.White
        Me.txtAddress.Properties.AppearanceReadOnly.Options.UseBackColor = True
        Me.txtAddress.Properties.AppearanceReadOnly.Options.UseBorderColor = True
        Me.txtAddress.Size = New System.Drawing.Size(213, 20)
        Me.txtAddress.TabIndex = 8
        '
        'gcRemarks
        '
        Me.gcRemarks.AppearanceCaption.BackColor = System.Drawing.Color.FromArgb(CType(CType(16, Byte), Integer), CType(CType(110, Byte), Integer), CType(CType(190, Byte), Integer))
        Me.gcRemarks.AppearanceCaption.BackColor2 = System.Drawing.Color.FromArgb(CType(CType(16, Byte), Integer), CType(CType(110, Byte), Integer), CType(CType(190, Byte), Integer))
        Me.gcRemarks.AppearanceCaption.BorderColor = System.Drawing.Color.FromArgb(CType(CType(16, Byte), Integer), CType(CType(110, Byte), Integer), CType(CType(190, Byte), Integer))
        Me.gcRemarks.AppearanceCaption.ForeColor = System.Drawing.Color.White
        Me.gcRemarks.AppearanceCaption.Options.UseBackColor = True
        Me.gcRemarks.AppearanceCaption.Options.UseBorderColor = True
        Me.gcRemarks.AppearanceCaption.Options.UseForeColor = True
        Me.gcRemarks.CaptionImageOptions.AllowGlyphSkinning = True
        Me.gcRemarks.CaptionImageOptions.SvgImage = CType(resources.GetObject("gcRemarks.CaptionImageOptions.SvgImage"), DevExpress.Utils.Svg.SvgImage)
        Me.gcRemarks.CaptionImageOptions.SvgImageSize = New System.Drawing.Size(16, 16)
        Me.gcRemarks.Controls.Add(Me.txtRemarks)
        Me.gcRemarks.Location = New System.Drawing.Point(330, 568)
        Me.gcRemarks.LookAndFeel.SkinName = "The Bezier"
        Me.gcRemarks.LookAndFeel.UseDefaultLookAndFeel = False
        Me.gcRemarks.Name = "gcRemarks"
        Me.gcRemarks.Size = New System.Drawing.Size(447, 107)
        Me.gcRemarks.TabIndex = 180
        Me.gcRemarks.Text = "Remarks"
        '
        'gcSignature
        '
        Me.gcSignature.AppearanceCaption.BackColor = System.Drawing.Color.FromArgb(CType(CType(16, Byte), Integer), CType(CType(110, Byte), Integer), CType(CType(190, Byte), Integer))
        Me.gcSignature.AppearanceCaption.BackColor2 = System.Drawing.Color.FromArgb(CType(CType(16, Byte), Integer), CType(CType(110, Byte), Integer), CType(CType(190, Byte), Integer))
        Me.gcSignature.AppearanceCaption.BorderColor = System.Drawing.Color.FromArgb(CType(CType(16, Byte), Integer), CType(CType(110, Byte), Integer), CType(CType(190, Byte), Integer))
        Me.gcSignature.AppearanceCaption.ForeColor = System.Drawing.Color.White
        Me.gcSignature.AppearanceCaption.Options.UseBackColor = True
        Me.gcSignature.AppearanceCaption.Options.UseBorderColor = True
        Me.gcSignature.AppearanceCaption.Options.UseForeColor = True
        Me.gcSignature.CaptionImageOptions.AllowGlyphSkinning = True
        Me.gcSignature.CaptionImageOptions.SvgImage = CType(resources.GetObject("gcSignature.CaptionImageOptions.SvgImage"), DevExpress.Utils.Svg.SvgImage)
        Me.gcSignature.CaptionImageOptions.SvgImageSize = New System.Drawing.Size(16, 16)
        Me.gcSignature.Controls.Add(Me.cboMedTech)
        Me.gcSignature.Controls.Add(Me.cboVerify)
        Me.gcSignature.Controls.Add(Me.cboPathologist)
        Me.gcSignature.Controls.Add(Me.LabelControl26)
        Me.gcSignature.Controls.Add(Me.LabelControl29)
        Me.gcSignature.Controls.Add(Me.LabelControl27)
        Me.gcSignature.Enabled = False
        Me.gcSignature.Location = New System.Drawing.Point(12, 568)
        Me.gcSignature.LookAndFeel.SkinName = "The Bezier"
        Me.gcSignature.LookAndFeel.UseDefaultLookAndFeel = False
        Me.gcSignature.Name = "gcSignature"
        Me.gcSignature.Size = New System.Drawing.Size(312, 108)
        Me.gcSignature.TabIndex = 181
        Me.gcSignature.Text = "Signatories"
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
        Me.GroupControl1.Controls.Add(Me.txtComment)
        Me.GroupControl1.Location = New System.Drawing.Point(783, 568)
        Me.GroupControl1.LookAndFeel.SkinName = "The Bezier"
        Me.GroupControl1.LookAndFeel.UseDefaultLookAndFeel = False
        Me.GroupControl1.Name = "GroupControl1"
        Me.GroupControl1.Size = New System.Drawing.Size(493, 108)
        Me.GroupControl1.TabIndex = 186
        Me.GroupControl1.Text = "Comment"
        '
        'frmHemaOrdered
        '
        Me.Appearance.BackColor = System.Drawing.SystemColors.Control
        Me.Appearance.ForeColor = System.Drawing.Color.FromArgb(CType(CType(64, Byte), Integer), CType(CType(64, Byte), Integer), CType(CType(64, Byte), Integer))
        Me.Appearance.Options.UseBackColor = True
        Me.Appearance.Options.UseForeColor = True
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(1305, 702)
        Me.ControlBox = False
        Me.Controls.Add(Me.GroupControl1)
        Me.Controls.Add(Me.gcSignature)
        Me.Controls.Add(Me.gcRemarks)
        Me.Controls.Add(Me.gcTest)
        Me.Controls.Add(Me.gcAdditional)
        Me.Controls.Add(Me.gcPatient)
        Me.Controls.Add(Me.barDockControlLeft)
        Me.Controls.Add(Me.barDockControlRight)
        Me.Controls.Add(Me.barDockControlBottom)
        Me.Controls.Add(Me.barDockControlTop)
        Me.FormBorderEffect = DevExpress.XtraEditors.FormBorderEffect.Glow
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "frmHemaOrdered"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
        Me.Text = "Hematology Result Window - Archived"
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
        CType(Me.gcTest, System.ComponentModel.ISupportInitialize).EndInit()
        Me.gcTest.ResumeLayout(False)
        CType(Me.dtResult, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.GridView, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.gcAdditional, System.ComponentModel.ISupportInitialize).EndInit()
        Me.gcAdditional.ResumeLayout(False)
        Me.gcAdditional.PerformLayout()
        CType(Me.tmTimeReceived.Properties.CalendarTimeProperties, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.tmTimeReceived.Properties, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.txtChargeSlip.Properties, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.dtReceived.Properties.CalendarTimeProperties, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.dtReceived.Properties, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.gcPatient, System.ComponentModel.ISupportInitialize).EndInit()
        Me.gcPatient.ResumeLayout(False)
        Me.gcPatient.PerformLayout()
        CType(Me.txtAddress.Properties, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.gcRemarks, System.ComponentModel.ISupportInitialize).EndInit()
        Me.gcRemarks.ResumeLayout(False)
        CType(Me.gcSignature, System.ComponentModel.ISupportInitialize).EndInit()
        Me.gcSignature.ResumeLayout(False)
        Me.gcSignature.PerformLayout()
        CType(Me.GroupControl1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.GroupControl1.ResumeLayout(False)
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
    Friend WithEvents cboCS As DevExpress.XtraEditors.ComboBoxEdit
    Friend WithEvents gcTest As DevExpress.XtraEditors.GroupControl
    Friend WithEvents LabelControl26 As DevExpress.XtraEditors.LabelControl
    Friend WithEvents LabelControl27 As DevExpress.XtraEditors.LabelControl
    Friend WithEvents LabelControl29 As DevExpress.XtraEditors.LabelControl
    Friend WithEvents gcAdditional As DevExpress.XtraEditors.GroupControl
    Friend WithEvents LabelControl32 As DevExpress.XtraEditors.LabelControl
    Friend WithEvents LabelControl33 As DevExpress.XtraEditors.LabelControl
    Friend WithEvents LabelControl34 As DevExpress.XtraEditors.LabelControl
    Friend WithEvents LabelControl35 As DevExpress.XtraEditors.LabelControl
    Friend WithEvents LabelControl36 As DevExpress.XtraEditors.LabelControl
    Friend WithEvents LabelControl37 As DevExpress.XtraEditors.LabelControl
    Friend WithEvents LabelControl38 As DevExpress.XtraEditors.LabelControl
    Friend WithEvents LabelControl41 As DevExpress.XtraEditors.LabelControl
    Friend WithEvents gcPatient As DevExpress.XtraEditors.GroupControl
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
    Friend WithEvents gcSignature As DevExpress.XtraEditors.GroupControl
    Friend WithEvents gcRemarks As DevExpress.XtraEditors.GroupControl
    Friend WithEvents btnRetrive As DevExpress.XtraEditors.SimpleButton
    Friend WithEvents btnPreview As DevExpress.XtraEditors.SimpleButton
    Friend WithEvents btnAddTest As DevExpress.XtraEditors.SimpleButton
    Friend WithEvents btnEdit As DevExpress.XtraBars.BarLargeButtonItem
    Friend WithEvents dtReceived As DevExpress.XtraEditors.DateEdit
    Friend WithEvents GroupControl1 As DevExpress.XtraEditors.GroupControl
    Friend WithEvents LabelControl1 As DevExpress.XtraEditors.LabelControl
    Friend WithEvents txtChargeSlip As DevExpress.XtraEditors.TextEdit
    Friend WithEvents txtAddress As DevExpress.XtraEditors.TextEdit
    Friend WithEvents lblDiffCount As DevExpress.XtraEditors.LabelControl
    Friend WithEvents btnResend As DevExpress.XtraBars.BarLargeButtonItem
    Friend WithEvents tmTimeReleased As DateTimePicker
    Friend WithEvents LabelControl2 As DevExpress.XtraEditors.LabelControl
    Friend WithEvents tmTimeReceived As DevExpress.XtraEditors.DateEdit
End Class
