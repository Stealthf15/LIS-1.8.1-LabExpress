<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmPatientOrderAE
    Inherits DevExpress.XtraEditors.XtraForm

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> _
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
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmPatientOrderAE))
        Dim Code128Generator2 As DevExpress.XtraPrinting.BarCode.Code128Generator = New DevExpress.XtraPrinting.BarCode.Code128Generator()
        Me.lvList = New System.Windows.Forms.ListView()
        Me.ColumnHeader10 = CType(New System.Windows.Forms.ColumnHeader(), System.Windows.Forms.ColumnHeader)
        Me.ColumnHeader1 = CType(New System.Windows.Forms.ColumnHeader(), System.Windows.Forms.ColumnHeader)
        Me.ColumnHeader3 = CType(New System.Windows.Forms.ColumnHeader(), System.Windows.Forms.ColumnHeader)
        Me.ColumnHeader18 = CType(New System.Windows.Forms.ColumnHeader(), System.Windows.Forms.ColumnHeader)
        Me.ColumnHeader11 = CType(New System.Windows.Forms.ColumnHeader(), System.Windows.Forms.ColumnHeader)
        Me.ColumnHeader24 = CType(New System.Windows.Forms.ColumnHeader(), System.Windows.Forms.ColumnHeader)
        Me.ColumnHeader14 = CType(New System.Windows.Forms.ColumnHeader(), System.Windows.Forms.ColumnHeader)
        Me.ImageList = New System.Windows.Forms.ImageList(Me.components)
        Me.bcCode = New DevExpress.XtraEditors.BarCodeControl()
        Me.picCode = New System.Windows.Forms.PictureBox()
        Me.LabelControl7 = New DevExpress.XtraEditors.LabelControl()
        Me.txtSampleID = New DevExpress.XtraEditors.TextEdit()
        Me.txtClass = New DevExpress.XtraEditors.ComboBoxEdit()
        Me.LabelControl13 = New DevExpress.XtraEditors.LabelControl()
        Me.txtAge = New DevExpress.XtraEditors.TextEdit()
        Me.LabelControl6 = New DevExpress.XtraEditors.LabelControl()
        Me.LabelControl2 = New DevExpress.XtraEditors.LabelControl()
        Me.dtBDate = New DevExpress.XtraEditors.DateEdit()
        Me.LabelControl1 = New DevExpress.XtraEditors.LabelControl()
        Me.txtPatientID = New DevExpress.XtraEditors.TextEdit()
        Me.LabelControl39 = New DevExpress.XtraEditors.LabelControl()
        Me.LabelControl15 = New DevExpress.XtraEditors.LabelControl()
        Me.cboSex = New DevExpress.XtraEditors.ComboBoxEdit()
        Me.txtContact = New DevExpress.XtraEditors.TextEdit()
        Me.btnSearch = New DevExpress.XtraEditors.SimpleButton()
        Me.LabelControl14 = New DevExpress.XtraEditors.LabelControl()
        Me.txtAddress = New DevExpress.XtraEditors.MemoEdit()
        Me.txtName = New DevExpress.XtraEditors.MemoEdit()
        Me.cboTestName = New DevExpress.XtraEditors.ComboBoxEdit()
        Me.LabelControl5 = New DevExpress.XtraEditors.LabelControl()
        Me.lvOrder = New System.Windows.Forms.ListView()
        Me.ColumnHeader6 = CType(New System.Windows.Forms.ColumnHeader(), System.Windows.Forms.ColumnHeader)
        Me.ColumnHeader12 = CType(New System.Windows.Forms.ColumnHeader(), System.Windows.Forms.ColumnHeader)
        Me.ColumnHeader13 = CType(New System.Windows.Forms.ColumnHeader(), System.Windows.Forms.ColumnHeader)
        Me.ColumnHeader4 = CType(New System.Windows.Forms.ColumnHeader(), System.Windows.Forms.ColumnHeader)
        Me.ColumnHeader7 = CType(New System.Windows.Forms.ColumnHeader(), System.Windows.Forms.ColumnHeader)
        Me.ColumnHeader19 = CType(New System.Windows.Forms.ColumnHeader(), System.Windows.Forms.ColumnHeader)
        Me.ColumnHeader21 = CType(New System.Windows.Forms.ColumnHeader(), System.Windows.Forms.ColumnHeader)
        Me.ColumnHeader20 = CType(New System.Windows.Forms.ColumnHeader(), System.Windows.Forms.ColumnHeader)
        Me.ColumnHeader26 = CType(New System.Windows.Forms.ColumnHeader(), System.Windows.Forms.ColumnHeader)
        Me.Timer1 = New System.Windows.Forms.Timer(Me.components)
        Me.GroupOrder = New DevExpress.XtraEditors.GroupControl()
        Me.SimpleButton1 = New DevExpress.XtraEditors.SimpleButton()
        Me.btnRemove = New DevExpress.XtraEditors.SimpleButton()
        Me.btnDelete = New DevExpress.XtraEditors.SimpleButton()
        Me.btnAdd = New DevExpress.XtraEditors.SimpleButton()
        Me.lvDup = New System.Windows.Forms.ListView()
        Me.ColumnHeader15 = CType(New System.Windows.Forms.ColumnHeader(), System.Windows.Forms.ColumnHeader)
        Me.ColumnHeader16 = CType(New System.Windows.Forms.ColumnHeader(), System.Windows.Forms.ColumnHeader)
        Me.ColumnHeader17 = CType(New System.Windows.Forms.ColumnHeader(), System.Windows.Forms.ColumnHeader)
        Me.ColumnHeader23 = CType(New System.Windows.Forms.ColumnHeader(), System.Windows.Forms.ColumnHeader)
        Me.lv = New System.Windows.Forms.ListView()
        Me.ColumnHeader8 = CType(New System.Windows.Forms.ColumnHeader(), System.Windows.Forms.ColumnHeader)
        Me.ColumnHeader9 = CType(New System.Windows.Forms.ColumnHeader(), System.Windows.Forms.ColumnHeader)
        Me.ColumnHeader22 = CType(New System.Windows.Forms.ColumnHeader(), System.Windows.Forms.ColumnHeader)
        Me.ColumnHeader25 = CType(New System.Windows.Forms.ColumnHeader(), System.Windows.Forms.ColumnHeader)
        Me.GroupControl1 = New DevExpress.XtraEditors.GroupControl()
        Me.lvTest = New System.Windows.Forms.ListView()
        Me.ColumnHeader2 = CType(New System.Windows.Forms.ColumnHeader(), System.Windows.Forms.ColumnHeader)
        Me.ColumnHeader5 = CType(New System.Windows.Forms.ColumnHeader(), System.Windows.Forms.ColumnHeader)
        Me.GroupTestDetails = New DevExpress.XtraEditors.GroupControl()
        Me.chAll = New System.Windows.Forms.CheckBox()
        Me.GroupPatient = New DevExpress.XtraEditors.GroupControl()
        Me.txtEmail = New DevExpress.XtraEditors.TextEdit()
        Me.LabelControl9 = New DevExpress.XtraEditors.LabelControl()
        Me.LabelControl8 = New DevExpress.XtraEditors.LabelControl()
        Me.cboStatus = New DevExpress.XtraEditors.ComboBoxEdit()
        Me.BarManager = New DevExpress.XtraBars.BarManager(Me.components)
        Me.Bar = New DevExpress.XtraBars.Bar()
        Me.btnSave = New DevExpress.XtraBars.BarLargeButtonItem()
        Me.btnSavePrint = New DevExpress.XtraBars.BarLargeButtonItem()
        Me.btnClear = New DevExpress.XtraBars.BarLargeButtonItem()
        Me.btnClose = New DevExpress.XtraBars.BarLargeButtonItem()
        Me.BarAndDockingController = New DevExpress.XtraBars.BarAndDockingController(Me.components)
        Me.barDockControlTop = New DevExpress.XtraBars.BarDockControl()
        Me.barDockControlBottom = New DevExpress.XtraBars.BarDockControl()
        Me.barDockControlLeft = New DevExpress.XtraBars.BarDockControl()
        Me.barDockControlRight = New DevExpress.XtraBars.BarDockControl()
        Me.BarLargeButtonItem1 = New DevExpress.XtraBars.BarLargeButtonItem()
        Me.BarLargeButtonItem2 = New DevExpress.XtraBars.BarLargeButtonItem()
        Me.PrintPreviewDialog = New System.Windows.Forms.PrintPreviewDialog()
        Me.PrintDocument = New System.Drawing.Printing.PrintDocument()
        Me.GroupAdditional = New DevExpress.XtraEditors.GroupControl()
        Me.LabelControl4 = New DevExpress.XtraEditors.LabelControl()
        Me.cboMode = New DevExpress.XtraEditors.ComboBoxEdit()
        Me.txtChargeSlip = New DevExpress.XtraEditors.TextEdit()
        Me.LabelControl3 = New DevExpress.XtraEditors.LabelControl()
        Me.tmTimeReceived = New System.Windows.Forms.DateTimePicker()
        Me.LabelControl32 = New DevExpress.XtraEditors.LabelControl()
        Me.dtReceived = New System.Windows.Forms.DateTimePicker()
        Me.LabelControl33 = New DevExpress.XtraEditors.LabelControl()
        Me.txtAccession = New DevExpress.XtraEditors.TextEdit()
        Me.LabelControl34 = New DevExpress.XtraEditors.LabelControl()
        Me.cboPType = New DevExpress.XtraEditors.ComboBoxEdit()
        Me.LabelControl35 = New DevExpress.XtraEditors.LabelControl()
        Me.txtORNo = New DevExpress.XtraEditors.TextEdit()
        Me.LabelControl36 = New DevExpress.XtraEditors.LabelControl()
        Me.LabelControl37 = New DevExpress.XtraEditors.LabelControl()
        Me.LabelControl38 = New DevExpress.XtraEditors.LabelControl()
        Me.cboRoom = New DevExpress.XtraEditors.ComboBoxEdit()
        Me.LabelControl41 = New DevExpress.XtraEditors.LabelControl()
        Me.cboPhysician = New DevExpress.XtraEditors.ComboBoxEdit()
        Me.cboRequest = New DevExpress.XtraEditors.ComboBoxEdit()
        CType(Me.picCode, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.txtSampleID.Properties, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.txtClass.Properties, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.txtAge.Properties, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.dtBDate.Properties.CalendarTimeProperties, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.dtBDate.Properties, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.txtPatientID.Properties, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.cboSex.Properties, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.txtContact.Properties, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.txtAddress.Properties, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.txtName.Properties, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.cboTestName.Properties, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.GroupOrder, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.GroupOrder.SuspendLayout()
        CType(Me.GroupControl1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.GroupControl1.SuspendLayout()
        CType(Me.GroupTestDetails, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.GroupTestDetails.SuspendLayout()
        CType(Me.GroupPatient, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.GroupPatient.SuspendLayout()
        CType(Me.txtEmail.Properties, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.cboStatus.Properties, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.BarManager, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.BarAndDockingController, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.GroupAdditional, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.GroupAdditional.SuspendLayout()
        CType(Me.cboMode.Properties, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.txtChargeSlip.Properties, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.txtAccession.Properties, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.cboPType.Properties, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.txtORNo.Properties, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.cboRoom.Properties, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.cboPhysician.Properties, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.cboRequest.Properties, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'lvList
        '
        Me.lvList.BackColor = System.Drawing.Color.FromArgb(CType(CType(240, Byte), Integer), CType(CType(240, Byte), Integer), CType(CType(240, Byte), Integer))
        Me.lvList.BorderStyle = System.Windows.Forms.BorderStyle.None
        Me.lvList.CheckBoxes = True
        Me.lvList.Columns.AddRange(New System.Windows.Forms.ColumnHeader() {Me.ColumnHeader10, Me.ColumnHeader1, Me.ColumnHeader3, Me.ColumnHeader18, Me.ColumnHeader11, Me.ColumnHeader24, Me.ColumnHeader14})
        Me.lvList.Dock = System.Windows.Forms.DockStyle.Fill
        Me.lvList.ForeColor = System.Drawing.Color.FromArgb(CType(CType(27, Byte), Integer), CType(CType(41, Byte), Integer), CType(CType(62, Byte), Integer))
        Me.lvList.FullRowSelect = True
        Me.lvList.HideSelection = False
        Me.lvList.Location = New System.Drawing.Point(2, 27)
        Me.lvList.Margin = New System.Windows.Forms.Padding(2, 3, 2, 3)
        Me.lvList.Name = "lvList"
        Me.lvList.Size = New System.Drawing.Size(429, 275)
        Me.lvList.SmallImageList = Me.ImageList
        Me.lvList.TabIndex = 56
        Me.lvList.UseCompatibleStateImageBehavior = False
        Me.lvList.View = System.Windows.Forms.View.Details
        '
        'ColumnHeader10
        '
        Me.ColumnHeader10.Text = "Sequence No."
        Me.ColumnHeader10.Width = 129
        '
        'ColumnHeader1
        '
        Me.ColumnHeader1.Text = "Test Name"
        Me.ColumnHeader1.Width = 342
        '
        'ColumnHeader3
        '
        Me.ColumnHeader3.Text = "Test Code"
        Me.ColumnHeader3.Width = 150
        '
        'ColumnHeader18
        '
        Me.ColumnHeader18.Text = "Order Code"
        '
        'ColumnHeader11
        '
        Me.ColumnHeader11.Text = "Section"
        Me.ColumnHeader11.Width = 100
        '
        'ColumnHeader24
        '
        Me.ColumnHeader24.Text = "Sub Section"
        Me.ColumnHeader24.Width = 100
        '
        'ColumnHeader14
        '
        Me.ColumnHeader14.Text = "Specimen Type"
        Me.ColumnHeader14.Width = 100
        '
        'ImageList
        '
        Me.ImageList.ImageStream = CType(resources.GetObject("ImageList.ImageStream"), System.Windows.Forms.ImageListStreamer)
        Me.ImageList.TransparentColor = System.Drawing.Color.Transparent
        Me.ImageList.Images.SetKeyName(0, "icon.png")
        '
        'bcCode
        '
        Me.bcCode.AutoModule = True
        Me.bcCode.BackColor = System.Drawing.Color.White
        Me.bcCode.BorderStyle = DevExpress.XtraEditors.Controls.BorderStyles.NoBorder
        Me.bcCode.Font = New System.Drawing.Font("Tahoma", 13.8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.bcCode.ForeColor = System.Drawing.Color.Black
        Me.bcCode.HorizontalAlignment = DevExpress.Utils.HorzAlignment.Center
        Me.bcCode.HorizontalTextAlignment = DevExpress.Utils.HorzAlignment.Center
        Me.bcCode.Location = New System.Drawing.Point(387, 718)
        Me.bcCode.Margin = New System.Windows.Forms.Padding(2)
        Me.bcCode.Name = "bcCode"
        Me.bcCode.Padding = New System.Windows.Forms.Padding(8, 2, 8, 0)
        Me.bcCode.ShowText = False
        Me.bcCode.Size = New System.Drawing.Size(321, 93)
        Me.bcCode.Symbology = Code128Generator2
        Me.bcCode.TabIndex = 142
        '
        'picCode
        '
        Me.picCode.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.picCode.BackColor = System.Drawing.Color.White
        Me.picCode.Location = New System.Drawing.Point(368, 816)
        Me.picCode.Margin = New System.Windows.Forms.Padding(2, 3, 2, 3)
        Me.picCode.Name = "picCode"
        Me.picCode.Size = New System.Drawing.Size(319, 73)
        Me.picCode.TabIndex = 110
        Me.picCode.TabStop = False
        Me.picCode.Visible = False
        '
        'LabelControl7
        '
        Me.LabelControl7.Appearance.ForeColor = System.Drawing.Color.FromArgb(CType(CType(27, Byte), Integer), CType(CType(41, Byte), Integer), CType(CType(62, Byte), Integer))
        Me.LabelControl7.Appearance.Options.UseForeColor = True
        Me.LabelControl7.Location = New System.Drawing.Point(11, 38)
        Me.LabelControl7.LookAndFeel.SkinName = "Visual Studio 2013 Blue"
        Me.LabelControl7.LookAndFeel.UseDefaultLookAndFeel = False
        Me.LabelControl7.Margin = New System.Windows.Forms.Padding(2, 3, 2, 3)
        Me.LabelControl7.Name = "LabelControl7"
        Me.LabelControl7.Size = New System.Drawing.Size(54, 13)
        Me.LabelControl7.TabIndex = 97
        Me.LabelControl7.Text = "Sample ID:"
        '
        'txtSampleID
        '
        Me.txtSampleID.Enabled = False
        Me.txtSampleID.Location = New System.Drawing.Point(102, 35)
        Me.txtSampleID.Margin = New System.Windows.Forms.Padding(2, 3, 2, 3)
        Me.txtSampleID.Name = "txtSampleID"
        Me.txtSampleID.Properties.Appearance.ForeColor = System.Drawing.Color.FromArgb(CType(CType(27, Byte), Integer), CType(CType(41, Byte), Integer), CType(CType(62, Byte), Integer))
        Me.txtSampleID.Properties.Appearance.Options.UseForeColor = True
        Me.txtSampleID.Properties.LookAndFeel.SkinName = "Visual Studio 2013 Blue"
        Me.txtSampleID.Properties.LookAndFeel.UseDefaultLookAndFeel = False
        Me.txtSampleID.Size = New System.Drawing.Size(235, 20)
        Me.txtSampleID.TabIndex = 1
        '
        'txtClass
        '
        Me.txtClass.Location = New System.Drawing.Point(217, 165)
        Me.txtClass.Margin = New System.Windows.Forms.Padding(2, 3, 2, 3)
        Me.txtClass.Name = "txtClass"
        Me.txtClass.Properties.Appearance.ForeColor = System.Drawing.Color.FromArgb(CType(CType(27, Byte), Integer), CType(CType(41, Byte), Integer), CType(CType(62, Byte), Integer))
        Me.txtClass.Properties.Appearance.Options.UseForeColor = True
        Me.txtClass.Properties.Buttons.AddRange(New DevExpress.XtraEditors.Controls.EditorButton() {New DevExpress.XtraEditors.Controls.EditorButton(DevExpress.XtraEditors.Controls.ButtonPredefines.Combo)})
        Me.txtClass.Properties.Items.AddRange(New Object() {"Day(s)", "Month(s)", "Year(s)"})
        Me.txtClass.Properties.LookAndFeel.SkinName = "Visual Studio 2013 Blue"
        Me.txtClass.Properties.LookAndFeel.UseDefaultLookAndFeel = False
        Me.txtClass.Properties.TextEditStyle = DevExpress.XtraEditors.Controls.TextEditStyles.DisableTextEditor
        Me.txtClass.Size = New System.Drawing.Size(120, 20)
        Me.txtClass.TabIndex = 8
        '
        'LabelControl13
        '
        Me.LabelControl13.Appearance.ForeColor = System.Drawing.Color.FromArgb(CType(CType(27, Byte), Integer), CType(CType(41, Byte), Integer), CType(CType(62, Byte), Integer))
        Me.LabelControl13.Appearance.Options.UseForeColor = True
        Me.LabelControl13.Location = New System.Drawing.Point(11, 168)
        Me.LabelControl13.LookAndFeel.SkinName = "Visual Studio 2013 Blue"
        Me.LabelControl13.LookAndFeel.UseDefaultLookAndFeel = False
        Me.LabelControl13.Margin = New System.Windows.Forms.Padding(2, 3, 2, 3)
        Me.LabelControl13.Name = "LabelControl13"
        Me.LabelControl13.Size = New System.Drawing.Size(23, 13)
        Me.LabelControl13.TabIndex = 36
        Me.LabelControl13.Text = "Age:"
        '
        'txtAge
        '
        Me.txtAge.Location = New System.Drawing.Point(102, 165)
        Me.txtAge.Margin = New System.Windows.Forms.Padding(2, 3, 2, 3)
        Me.txtAge.Name = "txtAge"
        Me.txtAge.Properties.Appearance.ForeColor = System.Drawing.Color.FromArgb(CType(CType(27, Byte), Integer), CType(CType(41, Byte), Integer), CType(CType(62, Byte), Integer))
        Me.txtAge.Properties.Appearance.Options.UseForeColor = True
        Me.txtAge.Properties.LookAndFeel.SkinName = "Visual Studio 2013 Blue"
        Me.txtAge.Properties.LookAndFeel.UseDefaultLookAndFeel = False
        Me.txtAge.Size = New System.Drawing.Size(111, 20)
        Me.txtAge.TabIndex = 7
        '
        'LabelControl6
        '
        Me.LabelControl6.Appearance.ForeColor = System.Drawing.Color.FromArgb(CType(CType(27, Byte), Integer), CType(CType(41, Byte), Integer), CType(CType(62, Byte), Integer))
        Me.LabelControl6.Appearance.Options.UseForeColor = True
        Me.LabelControl6.Location = New System.Drawing.Point(11, 122)
        Me.LabelControl6.LookAndFeel.SkinName = "Visual Studio 2013 Blue"
        Me.LabelControl6.LookAndFeel.UseDefaultLookAndFeel = False
        Me.LabelControl6.Margin = New System.Windows.Forms.Padding(2, 3, 2, 3)
        Me.LabelControl6.Name = "LabelControl6"
        Me.LabelControl6.Size = New System.Drawing.Size(20, 13)
        Me.LabelControl6.TabIndex = 20
        Me.LabelControl6.Text = "Sex:"
        '
        'LabelControl2
        '
        Me.LabelControl2.Appearance.ForeColor = System.Drawing.Color.FromArgb(CType(CType(27, Byte), Integer), CType(CType(41, Byte), Integer), CType(CType(62, Byte), Integer))
        Me.LabelControl2.Appearance.Options.UseForeColor = True
        Me.LabelControl2.Location = New System.Drawing.Point(11, 83)
        Me.LabelControl2.LookAndFeel.SkinName = "Visual Studio 2013 Blue"
        Me.LabelControl2.LookAndFeel.UseDefaultLookAndFeel = False
        Me.LabelControl2.Margin = New System.Windows.Forms.Padding(2, 3, 2, 3)
        Me.LabelControl2.Name = "LabelControl2"
        Me.LabelControl2.Size = New System.Drawing.Size(71, 13)
        Me.LabelControl2.TabIndex = 14
        Me.LabelControl2.Text = "Patient Name:"
        '
        'dtBDate
        '
        Me.dtBDate.EditValue = Nothing
        Me.dtBDate.Location = New System.Drawing.Point(102, 142)
        Me.dtBDate.Margin = New System.Windows.Forms.Padding(2, 3, 2, 3)
        Me.dtBDate.Name = "dtBDate"
        Me.dtBDate.Properties.Buttons.AddRange(New DevExpress.XtraEditors.Controls.EditorButton() {New DevExpress.XtraEditors.Controls.EditorButton(DevExpress.XtraEditors.Controls.ButtonPredefines.Combo)})
        Me.dtBDate.Properties.CalendarTimeProperties.Buttons.AddRange(New DevExpress.XtraEditors.Controls.EditorButton() {New DevExpress.XtraEditors.Controls.EditorButton(DevExpress.XtraEditors.Controls.ButtonPredefines.Combo)})
        Me.dtBDate.Size = New System.Drawing.Size(234, 20)
        Me.dtBDate.TabIndex = 6
        '
        'LabelControl1
        '
        Me.LabelControl1.Appearance.ForeColor = System.Drawing.Color.FromArgb(CType(CType(27, Byte), Integer), CType(CType(41, Byte), Integer), CType(CType(62, Byte), Integer))
        Me.LabelControl1.Appearance.Options.UseForeColor = True
        Me.LabelControl1.Location = New System.Drawing.Point(11, 61)
        Me.LabelControl1.LookAndFeel.SkinName = "Visual Studio 2013 Blue"
        Me.LabelControl1.LookAndFeel.UseDefaultLookAndFeel = False
        Me.LabelControl1.Margin = New System.Windows.Forms.Padding(2, 3, 2, 3)
        Me.LabelControl1.Name = "LabelControl1"
        Me.LabelControl1.Size = New System.Drawing.Size(53, 13)
        Me.LabelControl1.TabIndex = 12
        Me.LabelControl1.Text = "Patient ID:"
        '
        'txtPatientID
        '
        Me.txtPatientID.Location = New System.Drawing.Point(102, 58)
        Me.txtPatientID.Margin = New System.Windows.Forms.Padding(2, 3, 2, 3)
        Me.txtPatientID.Name = "txtPatientID"
        Me.txtPatientID.Properties.Appearance.ForeColor = System.Drawing.Color.FromArgb(CType(CType(27, Byte), Integer), CType(CType(41, Byte), Integer), CType(CType(62, Byte), Integer))
        Me.txtPatientID.Properties.Appearance.Options.UseForeColor = True
        Me.txtPatientID.Properties.LookAndFeel.SkinName = "Visual Studio 2013 Blue"
        Me.txtPatientID.Properties.LookAndFeel.UseDefaultLookAndFeel = False
        Me.txtPatientID.Size = New System.Drawing.Size(183, 20)
        Me.txtPatientID.TabIndex = 2
        '
        'LabelControl39
        '
        Me.LabelControl39.Appearance.ForeColor = System.Drawing.Color.FromArgb(CType(CType(27, Byte), Integer), CType(CType(41, Byte), Integer), CType(CType(62, Byte), Integer))
        Me.LabelControl39.Appearance.Options.UseForeColor = True
        Me.LabelControl39.Location = New System.Drawing.Point(11, 145)
        Me.LabelControl39.LookAndFeel.SkinName = "Visual Studio 2013 Blue"
        Me.LabelControl39.LookAndFeel.UseDefaultLookAndFeel = False
        Me.LabelControl39.Margin = New System.Windows.Forms.Padding(2, 3, 2, 3)
        Me.LabelControl39.Name = "LabelControl39"
        Me.LabelControl39.Size = New System.Drawing.Size(68, 13)
        Me.LabelControl39.TabIndex = 83
        Me.LabelControl39.Text = "Date of Birth:"
        '
        'LabelControl15
        '
        Me.LabelControl15.Appearance.ForeColor = System.Drawing.Color.FromArgb(CType(CType(27, Byte), Integer), CType(CType(41, Byte), Integer), CType(CType(62, Byte), Integer))
        Me.LabelControl15.Appearance.Options.UseForeColor = True
        Me.LabelControl15.Location = New System.Drawing.Point(11, 252)
        Me.LabelControl15.LookAndFeel.SkinName = "Visual Studio 2013 Blue"
        Me.LabelControl15.LookAndFeel.UseDefaultLookAndFeel = False
        Me.LabelControl15.Margin = New System.Windows.Forms.Padding(2, 3, 2, 3)
        Me.LabelControl15.Name = "LabelControl15"
        Me.LabelControl15.Size = New System.Drawing.Size(64, 13)
        Me.LabelControl15.TabIndex = 95
        Me.LabelControl15.Text = "Contact No.:"
        '
        'cboSex
        '
        Me.cboSex.Location = New System.Drawing.Point(102, 119)
        Me.cboSex.Margin = New System.Windows.Forms.Padding(2, 3, 2, 3)
        Me.cboSex.Name = "cboSex"
        Me.cboSex.Properties.Appearance.ForeColor = System.Drawing.Color.FromArgb(CType(CType(27, Byte), Integer), CType(CType(41, Byte), Integer), CType(CType(62, Byte), Integer))
        Me.cboSex.Properties.Appearance.Options.UseForeColor = True
        Me.cboSex.Properties.Buttons.AddRange(New DevExpress.XtraEditors.Controls.EditorButton() {New DevExpress.XtraEditors.Controls.EditorButton(DevExpress.XtraEditors.Controls.ButtonPredefines.Combo)})
        Me.cboSex.Properties.Items.AddRange(New Object() {"Male", "Female"})
        Me.cboSex.Properties.LookAndFeel.SkinName = "Visual Studio 2013 Blue"
        Me.cboSex.Properties.LookAndFeel.UseDefaultLookAndFeel = False
        Me.cboSex.Size = New System.Drawing.Size(234, 20)
        Me.cboSex.TabIndex = 5
        '
        'txtContact
        '
        Me.txtContact.Location = New System.Drawing.Point(102, 249)
        Me.txtContact.Margin = New System.Windows.Forms.Padding(2, 3, 2, 3)
        Me.txtContact.Name = "txtContact"
        Me.txtContact.Properties.Appearance.ForeColor = System.Drawing.Color.FromArgb(CType(CType(27, Byte), Integer), CType(CType(41, Byte), Integer), CType(CType(62, Byte), Integer))
        Me.txtContact.Properties.Appearance.Options.UseForeColor = True
        Me.txtContact.Properties.LookAndFeel.SkinName = "Visual Studio 2013 Blue"
        Me.txtContact.Properties.LookAndFeel.UseDefaultLookAndFeel = False
        Me.txtContact.Size = New System.Drawing.Size(234, 20)
        Me.txtContact.TabIndex = 11
        '
        'btnSearch
        '
        Me.btnSearch.Appearance.Font = New System.Drawing.Font("Segoe UI", 9.75!)
        Me.btnSearch.Appearance.Options.UseFont = True
        Me.btnSearch.Cursor = System.Windows.Forms.Cursors.Hand
        Me.btnSearch.ImageOptions.Image = CType(resources.GetObject("btnSearch.ImageOptions.Image"), System.Drawing.Image)
        Me.btnSearch.ImageOptions.Location = DevExpress.XtraEditors.ImageLocation.MiddleCenter
        Me.btnSearch.ImageOptions.SvgImageSize = New System.Drawing.Size(16, 16)
        Me.btnSearch.Location = New System.Drawing.Point(293, 58)
        Me.btnSearch.Margin = New System.Windows.Forms.Padding(2, 3, 2, 3)
        Me.btnSearch.Name = "btnSearch"
        Me.btnSearch.Size = New System.Drawing.Size(44, 20)
        Me.btnSearch.TabIndex = 3
        '
        'LabelControl14
        '
        Me.LabelControl14.Appearance.ForeColor = System.Drawing.Color.FromArgb(CType(CType(27, Byte), Integer), CType(CType(41, Byte), Integer), CType(CType(62, Byte), Integer))
        Me.LabelControl14.Appearance.Options.UseForeColor = True
        Me.LabelControl14.Location = New System.Drawing.Point(11, 213)
        Me.LabelControl14.LookAndFeel.SkinName = "Visual Studio 2013 Blue"
        Me.LabelControl14.LookAndFeel.UseDefaultLookAndFeel = False
        Me.LabelControl14.Margin = New System.Windows.Forms.Padding(2, 3, 2, 3)
        Me.LabelControl14.Name = "LabelControl14"
        Me.LabelControl14.Size = New System.Drawing.Size(44, 13)
        Me.LabelControl14.TabIndex = 94
        Me.LabelControl14.Text = "Address:"
        '
        'txtAddress
        '
        Me.txtAddress.Location = New System.Drawing.Point(102, 211)
        Me.txtAddress.Margin = New System.Windows.Forms.Padding(2, 3, 2, 3)
        Me.txtAddress.Name = "txtAddress"
        Me.txtAddress.Properties.Appearance.ForeColor = System.Drawing.Color.FromArgb(CType(CType(27, Byte), Integer), CType(CType(41, Byte), Integer), CType(CType(62, Byte), Integer))
        Me.txtAddress.Properties.Appearance.Options.UseForeColor = True
        Me.txtAddress.Size = New System.Drawing.Size(234, 35)
        Me.txtAddress.TabIndex = 10
        '
        'txtName
        '
        Me.txtName.Location = New System.Drawing.Point(102, 81)
        Me.txtName.Margin = New System.Windows.Forms.Padding(2, 3, 2, 3)
        Me.txtName.Name = "txtName"
        Me.txtName.Properties.Appearance.ForeColor = System.Drawing.Color.FromArgb(CType(CType(27, Byte), Integer), CType(CType(41, Byte), Integer), CType(CType(62, Byte), Integer))
        Me.txtName.Properties.Appearance.Options.UseForeColor = True
        Me.txtName.Properties.LookAndFeel.SkinName = "Visual Studio 2013 Blue"
        Me.txtName.Properties.LookAndFeel.UseDefaultLookAndFeel = False
        Me.txtName.Size = New System.Drawing.Size(234, 35)
        Me.txtName.TabIndex = 4
        '
        'cboTestName
        '
        Me.cboTestName.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cboTestName.Location = New System.Drawing.Point(283, 1)
        Me.cboTestName.Margin = New System.Windows.Forms.Padding(2, 3, 2, 3)
        Me.cboTestName.Name = "cboTestName"
        Me.cboTestName.Properties.Appearance.Font = New System.Drawing.Font("Segoe UI", 9.75!)
        Me.cboTestName.Properties.Appearance.ForeColor = System.Drawing.Color.FromArgb(CType(CType(27, Byte), Integer), CType(CType(41, Byte), Integer), CType(CType(62, Byte), Integer))
        Me.cboTestName.Properties.Appearance.Options.UseFont = True
        Me.cboTestName.Properties.Appearance.Options.UseForeColor = True
        Me.cboTestName.Properties.Buttons.AddRange(New DevExpress.XtraEditors.Controls.EditorButton() {New DevExpress.XtraEditors.Controls.EditorButton(DevExpress.XtraEditors.Controls.ButtonPredefines.Combo)})
        Me.cboTestName.Properties.LookAndFeel.SkinName = "Visual Studio 2013 Blue"
        Me.cboTestName.Properties.LookAndFeel.UseDefaultLookAndFeel = False
        Me.cboTestName.Properties.TextEditStyle = DevExpress.XtraEditors.Controls.TextEditStyles.DisableTextEditor
        Me.cboTestName.Size = New System.Drawing.Size(138, 24)
        Me.cboTestName.TabIndex = 168
        '
        'LabelControl5
        '
        Me.LabelControl5.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.LabelControl5.Appearance.Font = New System.Drawing.Font("Segoe UI", 10.2!, System.Drawing.FontStyle.Bold)
        Me.LabelControl5.Appearance.ForeColor = System.Drawing.Color.White
        Me.LabelControl5.Appearance.Options.UseFont = True
        Me.LabelControl5.Appearance.Options.UseForeColor = True
        Me.LabelControl5.Location = New System.Drawing.Point(225, 2)
        Me.LabelControl5.LookAndFeel.SkinName = "Visual Studio 2013 Blue"
        Me.LabelControl5.LookAndFeel.UseDefaultLookAndFeel = False
        Me.LabelControl5.Margin = New System.Windows.Forms.Padding(2, 3, 2, 3)
        Me.LabelControl5.Name = "LabelControl5"
        Me.LabelControl5.Size = New System.Drawing.Size(53, 19)
        Me.LabelControl5.TabIndex = 169
        Me.LabelControl5.Text = "Section:"
        Me.LabelControl5.Visible = False
        '
        'lvOrder
        '
        Me.lvOrder.BackColor = System.Drawing.Color.FromArgb(CType(CType(240, Byte), Integer), CType(CType(240, Byte), Integer), CType(CType(240, Byte), Integer))
        Me.lvOrder.BorderStyle = System.Windows.Forms.BorderStyle.None
        Me.lvOrder.Columns.AddRange(New System.Windows.Forms.ColumnHeader() {Me.ColumnHeader6, Me.ColumnHeader12, Me.ColumnHeader13, Me.ColumnHeader4, Me.ColumnHeader7, Me.ColumnHeader19, Me.ColumnHeader21, Me.ColumnHeader20, Me.ColumnHeader26})
        Me.lvOrder.Dock = System.Windows.Forms.DockStyle.Fill
        Me.lvOrder.ForeColor = System.Drawing.Color.FromArgb(CType(CType(27, Byte), Integer), CType(CType(41, Byte), Integer), CType(CType(62, Byte), Integer))
        Me.lvOrder.FullRowSelect = True
        Me.lvOrder.HideSelection = False
        Me.lvOrder.Location = New System.Drawing.Point(2, 27)
        Me.lvOrder.Margin = New System.Windows.Forms.Padding(2, 3, 2, 3)
        Me.lvOrder.Name = "lvOrder"
        Me.lvOrder.Size = New System.Drawing.Size(857, 273)
        Me.lvOrder.SmallImageList = Me.ImageList
        Me.lvOrder.TabIndex = 56
        Me.lvOrder.UseCompatibleStateImageBehavior = False
        Me.lvOrder.View = System.Windows.Forms.View.Details
        '
        'ColumnHeader6
        '
        Me.ColumnHeader6.Text = "Sequence No."
        Me.ColumnHeader6.Width = 120
        '
        'ColumnHeader12
        '
        Me.ColumnHeader12.Text = "Test Name"
        Me.ColumnHeader12.Width = 8000
        '
        'ColumnHeader13
        '
        Me.ColumnHeader13.Text = "Test Code"
        Me.ColumnHeader13.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        Me.ColumnHeader13.Width = 100
        '
        'ColumnHeader4
        '
        Me.ColumnHeader4.Text = "Section"
        Me.ColumnHeader4.Width = 150
        '
        'ColumnHeader7
        '
        Me.ColumnHeader7.Text = "SampleID"
        Me.ColumnHeader7.Width = 100
        '
        'ColumnHeader19
        '
        Me.ColumnHeader19.Text = "Order Code"
        '
        'ColumnHeader21
        '
        Me.ColumnHeader21.Text = "Sub Section"
        '
        'ColumnHeader20
        '
        Me.ColumnHeader20.Text = "HIS Code"
        Me.ColumnHeader20.Width = 100
        '
        'ColumnHeader26
        '
        Me.ColumnHeader26.Text = "Specimen Type"
        '
        'Timer1
        '
        Me.Timer1.Enabled = True
        Me.Timer1.Interval = 1000
        '
        'GroupOrder
        '
        Me.GroupOrder.AppearanceCaption.BackColor = System.Drawing.Color.FromArgb(CType(CType(16, Byte), Integer), CType(CType(110, Byte), Integer), CType(CType(190, Byte), Integer))
        Me.GroupOrder.AppearanceCaption.BackColor2 = System.Drawing.Color.FromArgb(CType(CType(16, Byte), Integer), CType(CType(110, Byte), Integer), CType(CType(190, Byte), Integer))
        Me.GroupOrder.AppearanceCaption.BorderColor = System.Drawing.Color.FromArgb(CType(CType(16, Byte), Integer), CType(CType(110, Byte), Integer), CType(CType(190, Byte), Integer))
        Me.GroupOrder.AppearanceCaption.ForeColor = System.Drawing.Color.White
        Me.GroupOrder.AppearanceCaption.Options.UseBackColor = True
        Me.GroupOrder.AppearanceCaption.Options.UseBorderColor = True
        Me.GroupOrder.AppearanceCaption.Options.UseForeColor = True
        Me.GroupOrder.CaptionImageOptions.AllowGlyphSkinning = True
        Me.GroupOrder.CaptionImageOptions.SvgImage = CType(resources.GetObject("GroupOrder.CaptionImageOptions.SvgImage"), DevExpress.Utils.Svg.SvgImage)
        Me.GroupOrder.CaptionImageOptions.SvgImageSize = New System.Drawing.Size(16, 16)
        Me.GroupOrder.Controls.Add(Me.lvOrder)
        Me.GroupOrder.Controls.Add(Me.SimpleButton1)
        Me.GroupOrder.Controls.Add(Me.btnRemove)
        Me.GroupOrder.Controls.Add(Me.btnDelete)
        Me.GroupOrder.Controls.Add(Me.btnAdd)
        Me.GroupOrder.Controls.Add(Me.lvDup)
        Me.GroupOrder.Controls.Add(Me.lv)
        Me.GroupOrder.Location = New System.Drawing.Point(366, 350)
        Me.GroupOrder.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.GroupOrder.Name = "GroupOrder"
        Me.GroupOrder.Size = New System.Drawing.Size(861, 302)
        Me.GroupOrder.TabIndex = 176
        Me.GroupOrder.Text = "Order List"
        '
        'SimpleButton1
        '
        Me.SimpleButton1.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.SimpleButton1.Appearance.Font = New System.Drawing.Font("Segoe UI", 7.8!, System.Drawing.FontStyle.Bold)
        Me.SimpleButton1.Appearance.Options.UseFont = True
        Me.SimpleButton1.Cursor = System.Windows.Forms.Cursors.Hand
        Me.SimpleButton1.ImageOptions.Location = DevExpress.XtraEditors.ImageLocation.MiddleLeft
        Me.SimpleButton1.ImageOptions.SvgImage = CType(resources.GetObject("SimpleButton1.ImageOptions.SvgImage"), DevExpress.Utils.Svg.SvgImage)
        Me.SimpleButton1.ImageOptions.SvgImageSize = New System.Drawing.Size(16, 16)
        Me.SimpleButton1.Location = New System.Drawing.Point(361, 3)
        Me.SimpleButton1.Margin = New System.Windows.Forms.Padding(2, 3, 2, 3)
        Me.SimpleButton1.Name = "SimpleButton1"
        Me.SimpleButton1.Size = New System.Drawing.Size(111, 20)
        Me.SimpleButton1.TabIndex = 195
        Me.SimpleButton1.Text = "View Packages"
        Me.SimpleButton1.Visible = False
        '
        'btnRemove
        '
        Me.btnRemove.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnRemove.Appearance.Font = New System.Drawing.Font("Segoe UI", 7.8!, System.Drawing.FontStyle.Bold)
        Me.btnRemove.Appearance.Options.UseFont = True
        Me.btnRemove.Cursor = System.Windows.Forms.Cursors.Hand
        Me.btnRemove.ImageOptions.Image = CType(resources.GetObject("btnRemove.ImageOptions.Image"), System.Drawing.Image)
        Me.btnRemove.ImageOptions.Location = DevExpress.XtraEditors.ImageLocation.MiddleLeft
        Me.btnRemove.Location = New System.Drawing.Point(476, 3)
        Me.btnRemove.Margin = New System.Windows.Forms.Padding(2, 3, 2, 3)
        Me.btnRemove.Name = "btnRemove"
        Me.btnRemove.Size = New System.Drawing.Size(173, 20)
        Me.btnRemove.TabIndex = 59
        Me.btnRemove.Text = "Remove Selected Order"
        '
        'btnDelete
        '
        Me.btnDelete.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnDelete.Appearance.Font = New System.Drawing.Font("Segoe UI", 7.8!, System.Drawing.FontStyle.Bold)
        Me.btnDelete.Appearance.Options.UseFont = True
        Me.btnDelete.Cursor = System.Windows.Forms.Cursors.Hand
        Me.btnDelete.ImageOptions.Image = CType(resources.GetObject("btnDelete.ImageOptions.Image"), System.Drawing.Image)
        Me.btnDelete.ImageOptions.Location = DevExpress.XtraEditors.ImageLocation.MiddleLeft
        Me.btnDelete.Location = New System.Drawing.Point(655, 3)
        Me.btnDelete.Margin = New System.Windows.Forms.Padding(2, 3, 2, 3)
        Me.btnDelete.Name = "btnDelete"
        Me.btnDelete.Size = New System.Drawing.Size(103, 20)
        Me.btnDelete.TabIndex = 58
        Me.btnDelete.Text = "Delete Order"
        '
        'btnAdd
        '
        Me.btnAdd.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnAdd.Appearance.Font = New System.Drawing.Font("Segoe UI", 7.8!, System.Drawing.FontStyle.Bold)
        Me.btnAdd.Appearance.Options.UseFont = True
        Me.btnAdd.Cursor = System.Windows.Forms.Cursors.Hand
        Me.btnAdd.ImageOptions.Image = CType(resources.GetObject("btnAdd.ImageOptions.Image"), System.Drawing.Image)
        Me.btnAdd.ImageOptions.Location = DevExpress.XtraEditors.ImageLocation.MiddleLeft
        Me.btnAdd.Location = New System.Drawing.Point(764, 3)
        Me.btnAdd.Margin = New System.Windows.Forms.Padding(2, 3, 2, 3)
        Me.btnAdd.Name = "btnAdd"
        Me.btnAdd.Size = New System.Drawing.Size(93, 20)
        Me.btnAdd.TabIndex = 57
        Me.btnAdd.Text = "Add Order"
        '
        'lvDup
        '
        Me.lvDup.BorderStyle = System.Windows.Forms.BorderStyle.None
        Me.lvDup.CheckBoxes = True
        Me.lvDup.Columns.AddRange(New System.Windows.Forms.ColumnHeader() {Me.ColumnHeader15, Me.ColumnHeader16, Me.ColumnHeader17, Me.ColumnHeader23})
        Me.lvDup.Font = New System.Drawing.Font("Tahoma", 9.75!)
        Me.lvDup.ForeColor = System.Drawing.Color.FromArgb(CType(CType(27, Byte), Integer), CType(CType(41, Byte), Integer), CType(CType(62, Byte), Integer))
        Me.lvDup.FullRowSelect = True
        Me.lvDup.HideSelection = False
        Me.lvDup.Location = New System.Drawing.Point(90, 58)
        Me.lvDup.Margin = New System.Windows.Forms.Padding(2, 3, 2, 3)
        Me.lvDup.Name = "lvDup"
        Me.lvDup.Size = New System.Drawing.Size(331, 121)
        Me.lvDup.TabIndex = 194
        Me.lvDup.UseCompatibleStateImageBehavior = False
        Me.lvDup.View = System.Windows.Forms.View.Details
        '
        'ColumnHeader15
        '
        Me.ColumnHeader15.Text = "Sequence No."
        Me.ColumnHeader15.Width = 129
        '
        'ColumnHeader16
        '
        Me.ColumnHeader16.Text = "Test Name"
        Me.ColumnHeader16.Width = 342
        '
        'ColumnHeader17
        '
        Me.ColumnHeader17.Text = "Section"
        Me.ColumnHeader17.Width = 150
        '
        'ColumnHeader23
        '
        Me.ColumnHeader23.Text = "Sub Section"
        '
        'lv
        '
        Me.lv.BorderStyle = System.Windows.Forms.BorderStyle.None
        Me.lv.CheckBoxes = True
        Me.lv.Columns.AddRange(New System.Windows.Forms.ColumnHeader() {Me.ColumnHeader8, Me.ColumnHeader9, Me.ColumnHeader22, Me.ColumnHeader25})
        Me.lv.Font = New System.Drawing.Font("Tahoma", 9.75!)
        Me.lv.ForeColor = System.Drawing.Color.FromArgb(CType(CType(27, Byte), Integer), CType(CType(41, Byte), Integer), CType(CType(62, Byte), Integer))
        Me.lv.FullRowSelect = True
        Me.lv.HideSelection = False
        Me.lv.Location = New System.Drawing.Point(425, 58)
        Me.lv.Margin = New System.Windows.Forms.Padding(2, 3, 2, 3)
        Me.lv.Name = "lv"
        Me.lv.Size = New System.Drawing.Size(331, 121)
        Me.lv.TabIndex = 189
        Me.lv.UseCompatibleStateImageBehavior = False
        Me.lv.View = System.Windows.Forms.View.Details
        '
        'ColumnHeader8
        '
        Me.ColumnHeader8.Text = "Sequence No."
        Me.ColumnHeader8.Width = 129
        '
        'ColumnHeader9
        '
        Me.ColumnHeader9.Text = "Test Name"
        Me.ColumnHeader9.Width = 342
        '
        'ColumnHeader22
        '
        Me.ColumnHeader22.Text = "Sub Section"
        '
        'ColumnHeader25
        '
        Me.ColumnHeader25.Text = "Specimen Type"
        Me.ColumnHeader25.Width = 100
        '
        'GroupControl1
        '
        Me.GroupControl1.Appearance.BackColor = System.Drawing.Color.White
        Me.GroupControl1.Appearance.Options.UseBackColor = True
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
        Me.GroupControl1.Controls.Add(Me.lvTest)
        Me.GroupControl1.Controls.Add(Me.cboTestName)
        Me.GroupControl1.Controls.Add(Me.LabelControl5)
        Me.GroupControl1.Location = New System.Drawing.Point(366, 42)
        Me.GroupControl1.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.GroupControl1.Name = "GroupControl1"
        Me.GroupControl1.Size = New System.Drawing.Size(423, 304)
        Me.GroupControl1.TabIndex = 177
        Me.GroupControl1.Text = "Test Group"
        '
        'lvTest
        '
        Me.lvTest.BackColor = System.Drawing.Color.FromArgb(CType(CType(240, Byte), Integer), CType(CType(240, Byte), Integer), CType(CType(240, Byte), Integer))
        Me.lvTest.BorderStyle = System.Windows.Forms.BorderStyle.None
        Me.lvTest.Columns.AddRange(New System.Windows.Forms.ColumnHeader() {Me.ColumnHeader2, Me.ColumnHeader5})
        Me.lvTest.Dock = System.Windows.Forms.DockStyle.Fill
        Me.lvTest.ForeColor = System.Drawing.Color.FromArgb(CType(CType(27, Byte), Integer), CType(CType(41, Byte), Integer), CType(CType(62, Byte), Integer))
        Me.lvTest.FullRowSelect = True
        Me.lvTest.HideSelection = False
        Me.lvTest.Location = New System.Drawing.Point(2, 27)
        Me.lvTest.Margin = New System.Windows.Forms.Padding(2, 3, 2, 3)
        Me.lvTest.Name = "lvTest"
        Me.lvTest.Size = New System.Drawing.Size(419, 275)
        Me.lvTest.SmallImageList = Me.ImageList
        Me.lvTest.TabIndex = 170
        Me.lvTest.UseCompatibleStateImageBehavior = False
        Me.lvTest.View = System.Windows.Forms.View.Details
        '
        'ColumnHeader2
        '
        Me.ColumnHeader2.Text = "Package Code"
        Me.ColumnHeader2.Width = 187
        '
        'ColumnHeader5
        '
        Me.ColumnHeader5.Text = "Package Name"
        Me.ColumnHeader5.Width = 342
        '
        'GroupTestDetails
        '
        Me.GroupTestDetails.AppearanceCaption.BackColor = System.Drawing.Color.FromArgb(CType(CType(16, Byte), Integer), CType(CType(110, Byte), Integer), CType(CType(190, Byte), Integer))
        Me.GroupTestDetails.AppearanceCaption.BackColor2 = System.Drawing.Color.FromArgb(CType(CType(16, Byte), Integer), CType(CType(110, Byte), Integer), CType(CType(190, Byte), Integer))
        Me.GroupTestDetails.AppearanceCaption.BorderColor = System.Drawing.Color.FromArgb(CType(CType(16, Byte), Integer), CType(CType(110, Byte), Integer), CType(CType(190, Byte), Integer))
        Me.GroupTestDetails.AppearanceCaption.ForeColor = System.Drawing.Color.White
        Me.GroupTestDetails.AppearanceCaption.Options.UseBackColor = True
        Me.GroupTestDetails.AppearanceCaption.Options.UseBorderColor = True
        Me.GroupTestDetails.AppearanceCaption.Options.UseForeColor = True
        Me.GroupTestDetails.CaptionImageOptions.AllowGlyphSkinning = True
        Me.GroupTestDetails.CaptionImageOptions.SvgImage = CType(resources.GetObject("GroupTestDetails.CaptionImageOptions.SvgImage"), DevExpress.Utils.Svg.SvgImage)
        Me.GroupTestDetails.CaptionImageOptions.SvgImageSize = New System.Drawing.Size(16, 16)
        Me.GroupTestDetails.Controls.Add(Me.chAll)
        Me.GroupTestDetails.Controls.Add(Me.lvList)
        Me.GroupTestDetails.Location = New System.Drawing.Point(794, 42)
        Me.GroupTestDetails.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.GroupTestDetails.Name = "GroupTestDetails"
        Me.GroupTestDetails.Size = New System.Drawing.Size(433, 304)
        Me.GroupTestDetails.TabIndex = 178
        Me.GroupTestDetails.Text = "Test Group Details"
        '
        'chAll
        '
        Me.chAll.AutoSize = True
        Me.chAll.BackColor = System.Drawing.Color.Transparent
        Me.chAll.Font = New System.Drawing.Font("Segoe UI", 10.0!, System.Drawing.FontStyle.Bold)
        Me.chAll.ForeColor = System.Drawing.Color.White
        Me.chAll.Location = New System.Drawing.Point(342, 2)
        Me.chAll.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.chAll.Name = "chAll"
        Me.chAll.Size = New System.Drawing.Size(90, 23)
        Me.chAll.TabIndex = 57
        Me.chAll.Text = "Check All"
        Me.chAll.UseVisualStyleBackColor = False
        '
        'GroupPatient
        '
        Me.GroupPatient.Appearance.BackColor = System.Drawing.Color.White
        Me.GroupPatient.Appearance.Options.UseBackColor = True
        Me.GroupPatient.AppearanceCaption.BackColor = System.Drawing.Color.FromArgb(CType(CType(16, Byte), Integer), CType(CType(110, Byte), Integer), CType(CType(190, Byte), Integer))
        Me.GroupPatient.AppearanceCaption.BackColor2 = System.Drawing.Color.FromArgb(CType(CType(16, Byte), Integer), CType(CType(110, Byte), Integer), CType(CType(190, Byte), Integer))
        Me.GroupPatient.AppearanceCaption.BorderColor = System.Drawing.Color.FromArgb(CType(CType(16, Byte), Integer), CType(CType(110, Byte), Integer), CType(CType(190, Byte), Integer))
        Me.GroupPatient.AppearanceCaption.ForeColor = System.Drawing.Color.White
        Me.GroupPatient.AppearanceCaption.Options.UseBackColor = True
        Me.GroupPatient.AppearanceCaption.Options.UseBorderColor = True
        Me.GroupPatient.AppearanceCaption.Options.UseForeColor = True
        Me.GroupPatient.CaptionImageOptions.AllowGlyphSkinning = True
        Me.GroupPatient.CaptionImageOptions.SvgImage = CType(resources.GetObject("GroupPatient.CaptionImageOptions.SvgImage"), DevExpress.Utils.Svg.SvgImage)
        Me.GroupPatient.CaptionImageOptions.SvgImageSize = New System.Drawing.Size(16, 16)
        Me.GroupPatient.Controls.Add(Me.txtEmail)
        Me.GroupPatient.Controls.Add(Me.LabelControl9)
        Me.GroupPatient.Controls.Add(Me.LabelControl8)
        Me.GroupPatient.Controls.Add(Me.LabelControl7)
        Me.GroupPatient.Controls.Add(Me.txtSampleID)
        Me.GroupPatient.Controls.Add(Me.txtName)
        Me.GroupPatient.Controls.Add(Me.txtClass)
        Me.GroupPatient.Controls.Add(Me.txtAddress)
        Me.GroupPatient.Controls.Add(Me.LabelControl13)
        Me.GroupPatient.Controls.Add(Me.LabelControl14)
        Me.GroupPatient.Controls.Add(Me.txtAge)
        Me.GroupPatient.Controls.Add(Me.btnSearch)
        Me.GroupPatient.Controls.Add(Me.LabelControl6)
        Me.GroupPatient.Controls.Add(Me.txtContact)
        Me.GroupPatient.Controls.Add(Me.LabelControl2)
        Me.GroupPatient.Controls.Add(Me.cboSex)
        Me.GroupPatient.Controls.Add(Me.dtBDate)
        Me.GroupPatient.Controls.Add(Me.LabelControl15)
        Me.GroupPatient.Controls.Add(Me.LabelControl1)
        Me.GroupPatient.Controls.Add(Me.LabelControl39)
        Me.GroupPatient.Controls.Add(Me.txtPatientID)
        Me.GroupPatient.Controls.Add(Me.cboStatus)
        Me.GroupPatient.Location = New System.Drawing.Point(12, 42)
        Me.GroupPatient.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.GroupPatient.Name = "GroupPatient"
        Me.GroupPatient.Size = New System.Drawing.Size(348, 304)
        Me.GroupPatient.TabIndex = 183
        Me.GroupPatient.Text = "Patient Details"
        '
        'txtEmail
        '
        Me.txtEmail.Location = New System.Drawing.Point(102, 272)
        Me.txtEmail.Margin = New System.Windows.Forms.Padding(2, 3, 2, 3)
        Me.txtEmail.Name = "txtEmail"
        Me.txtEmail.Properties.Appearance.ForeColor = System.Drawing.Color.FromArgb(CType(CType(27, Byte), Integer), CType(CType(41, Byte), Integer), CType(CType(62, Byte), Integer))
        Me.txtEmail.Properties.Appearance.Options.UseForeColor = True
        Me.txtEmail.Properties.LookAndFeel.SkinName = "Visual Studio 2013 Blue"
        Me.txtEmail.Properties.LookAndFeel.UseDefaultLookAndFeel = False
        Me.txtEmail.Size = New System.Drawing.Size(234, 20)
        Me.txtEmail.TabIndex = 100
        '
        'LabelControl9
        '
        Me.LabelControl9.Appearance.ForeColor = System.Drawing.Color.FromArgb(CType(CType(27, Byte), Integer), CType(CType(41, Byte), Integer), CType(CType(62, Byte), Integer))
        Me.LabelControl9.Appearance.Options.UseForeColor = True
        Me.LabelControl9.Location = New System.Drawing.Point(11, 275)
        Me.LabelControl9.LookAndFeel.SkinName = "Visual Studio 2013 Blue"
        Me.LabelControl9.LookAndFeel.UseDefaultLookAndFeel = False
        Me.LabelControl9.Margin = New System.Windows.Forms.Padding(2, 3, 2, 3)
        Me.LabelControl9.Name = "LabelControl9"
        Me.LabelControl9.Size = New System.Drawing.Size(74, 13)
        Me.LabelControl9.TabIndex = 101
        Me.LabelControl9.Text = "Email Address:"
        '
        'LabelControl8
        '
        Me.LabelControl8.Appearance.ForeColor = System.Drawing.Color.FromArgb(CType(CType(27, Byte), Integer), CType(CType(41, Byte), Integer), CType(CType(62, Byte), Integer))
        Me.LabelControl8.Appearance.Options.UseForeColor = True
        Me.LabelControl8.Location = New System.Drawing.Point(11, 191)
        Me.LabelControl8.LookAndFeel.SkinName = "Visual Studio 2013 Blue"
        Me.LabelControl8.LookAndFeel.UseDefaultLookAndFeel = False
        Me.LabelControl8.Margin = New System.Windows.Forms.Padding(2, 3, 2, 3)
        Me.LabelControl8.Name = "LabelControl8"
        Me.LabelControl8.Size = New System.Drawing.Size(59, 13)
        Me.LabelControl8.TabIndex = 99
        Me.LabelControl8.Text = "Civil Status:"
        '
        'cboStatus
        '
        Me.cboStatus.Location = New System.Drawing.Point(102, 188)
        Me.cboStatus.Margin = New System.Windows.Forms.Padding(2, 3, 2, 3)
        Me.cboStatus.Name = "cboStatus"
        Me.cboStatus.Properties.Appearance.ForeColor = System.Drawing.Color.FromArgb(CType(CType(27, Byte), Integer), CType(CType(41, Byte), Integer), CType(CType(62, Byte), Integer))
        Me.cboStatus.Properties.Appearance.Options.UseForeColor = True
        Me.cboStatus.Properties.Buttons.AddRange(New DevExpress.XtraEditors.Controls.EditorButton() {New DevExpress.XtraEditors.Controls.EditorButton(DevExpress.XtraEditors.Controls.ButtonPredefines.Combo)})
        Me.cboStatus.Properties.Items.AddRange(New Object() {"Single", "Married"})
        Me.cboStatus.Properties.LookAndFeel.SkinName = "Visual Studio 2013 Blue"
        Me.cboStatus.Properties.LookAndFeel.UseDefaultLookAndFeel = False
        Me.cboStatus.Properties.TextEditStyle = DevExpress.XtraEditors.Controls.TextEditStyles.DisableTextEditor
        Me.cboStatus.Size = New System.Drawing.Size(234, 20)
        Me.cboStatus.TabIndex = 98
        '
        'BarManager
        '
        Me.BarManager.Bars.AddRange(New DevExpress.XtraBars.Bar() {Me.Bar})
        Me.BarManager.Controller = Me.BarAndDockingController
        Me.BarManager.DockControls.Add(Me.barDockControlTop)
        Me.BarManager.DockControls.Add(Me.barDockControlBottom)
        Me.BarManager.DockControls.Add(Me.barDockControlLeft)
        Me.BarManager.DockControls.Add(Me.barDockControlRight)
        Me.BarManager.Form = Me
        Me.BarManager.Items.AddRange(New DevExpress.XtraBars.BarItem() {Me.BarLargeButtonItem1, Me.BarLargeButtonItem2, Me.btnSave, Me.btnSavePrint, Me.btnClear, Me.btnClose})
        Me.BarManager.MainMenu = Me.Bar
        Me.BarManager.MaxItemId = 7
        '
        'Bar
        '
        Me.Bar.BarAppearance.Normal.Font = New System.Drawing.Font("Segoe UI", 9.0!)
        Me.Bar.BarAppearance.Normal.ForeColor = System.Drawing.Color.White
        Me.Bar.BarAppearance.Normal.Options.UseFont = True
        Me.Bar.BarAppearance.Normal.Options.UseForeColor = True
        Me.Bar.BarName = "Main menu"
        Me.Bar.DockCol = 0
        Me.Bar.DockRow = 0
        Me.Bar.DockStyle = DevExpress.XtraBars.BarDockStyle.Top
        Me.Bar.FloatLocation = New System.Drawing.Point(990, 175)
        Me.Bar.LinksPersistInfo.AddRange(New DevExpress.XtraBars.LinkPersistInfo() {New DevExpress.XtraBars.LinkPersistInfo(Me.btnSave, True), New DevExpress.XtraBars.LinkPersistInfo(Me.btnSavePrint, True), New DevExpress.XtraBars.LinkPersistInfo(Me.btnClear, True), New DevExpress.XtraBars.LinkPersistInfo(Me.btnClose)})
        Me.Bar.OptionsBar.MultiLine = True
        Me.Bar.OptionsBar.UseWholeRow = True
        Me.Bar.Text = "Main menu"
        '
        'btnSave
        '
        Me.btnSave.Caption = "Save Order"
        Me.btnSave.CaptionAlignment = DevExpress.XtraBars.BarItemCaptionAlignment.Right
        Me.btnSave.Id = 2
        Me.btnSave.ImageOptions.AllowGlyphSkinning = DevExpress.Utils.DefaultBoolean.[True]
        Me.btnSave.ImageOptions.SvgImage = CType(resources.GetObject("btnSave.ImageOptions.SvgImage"), DevExpress.Utils.Svg.SvgImage)
        Me.btnSave.ImageOptions.SvgImageSize = New System.Drawing.Size(24, 24)
        Me.btnSave.Name = "btnSave"
        '
        'btnSavePrint
        '
        Me.btnSavePrint.Caption = "Save && Print BCode"
        Me.btnSavePrint.CaptionAlignment = DevExpress.XtraBars.BarItemCaptionAlignment.Right
        Me.btnSavePrint.Id = 4
        Me.btnSavePrint.ImageOptions.AllowGlyphSkinning = DevExpress.Utils.DefaultBoolean.[True]
        Me.btnSavePrint.ImageOptions.SvgImage = CType(resources.GetObject("btnSavePrint.ImageOptions.SvgImage"), DevExpress.Utils.Svg.SvgImage)
        Me.btnSavePrint.ImageOptions.SvgImageSize = New System.Drawing.Size(24, 24)
        Me.btnSavePrint.Name = "btnSavePrint"
        '
        'btnClear
        '
        Me.btnClear.Caption = "Clear All"
        Me.btnClear.CaptionAlignment = DevExpress.XtraBars.BarItemCaptionAlignment.Right
        Me.btnClear.Id = 5
        Me.btnClear.ImageOptions.AllowGlyphSkinning = DevExpress.Utils.DefaultBoolean.[True]
        Me.btnClear.ImageOptions.SvgImage = CType(resources.GetObject("btnClear.ImageOptions.SvgImage"), DevExpress.Utils.Svg.SvgImage)
        Me.btnClear.ImageOptions.SvgImageSize = New System.Drawing.Size(24, 24)
        Me.btnClear.Name = "btnClear"
        '
        'btnClose
        '
        Me.btnClose.Alignment = DevExpress.XtraBars.BarItemLinkAlignment.Right
        Me.btnClose.Caption = "Close"
        Me.btnClose.CaptionAlignment = DevExpress.XtraBars.BarItemCaptionAlignment.Right
        Me.btnClose.Id = 6
        Me.btnClose.ImageOptions.AllowGlyphSkinning = DevExpress.Utils.DefaultBoolean.[True]
        Me.btnClose.ImageOptions.SvgImage = CType(resources.GetObject("btnClose.ImageOptions.SvgImage"), DevExpress.Utils.Svg.SvgImage)
        Me.btnClose.ImageOptions.SvgImageSize = New System.Drawing.Size(24, 24)
        Me.btnClose.Name = "btnClose"
        '
        'BarAndDockingController
        '
        Me.BarAndDockingController.AppearancesBar.MainMenuAppearance.Normal.Font = New System.Drawing.Font("Segoe UI", 9.0!)
        Me.BarAndDockingController.AppearancesBar.MainMenuAppearance.Normal.Options.UseFont = True
        Me.BarAndDockingController.PropertiesBar.AllowLinkLighting = False
        '
        'barDockControlTop
        '
        Me.barDockControlTop.CausesValidation = False
        Me.barDockControlTop.Dock = System.Windows.Forms.DockStyle.Top
        Me.barDockControlTop.Location = New System.Drawing.Point(0, 0)
        Me.barDockControlTop.Manager = Me.BarManager
        Me.barDockControlTop.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.barDockControlTop.Size = New System.Drawing.Size(1238, 36)
        '
        'barDockControlBottom
        '
        Me.barDockControlBottom.CausesValidation = False
        Me.barDockControlBottom.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.barDockControlBottom.Location = New System.Drawing.Point(0, 662)
        Me.barDockControlBottom.Manager = Me.BarManager
        Me.barDockControlBottom.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.barDockControlBottom.Size = New System.Drawing.Size(1238, 0)
        '
        'barDockControlLeft
        '
        Me.barDockControlLeft.CausesValidation = False
        Me.barDockControlLeft.Dock = System.Windows.Forms.DockStyle.Left
        Me.barDockControlLeft.Location = New System.Drawing.Point(0, 36)
        Me.barDockControlLeft.Manager = Me.BarManager
        Me.barDockControlLeft.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.barDockControlLeft.Size = New System.Drawing.Size(0, 626)
        '
        'barDockControlRight
        '
        Me.barDockControlRight.CausesValidation = False
        Me.barDockControlRight.Dock = System.Windows.Forms.DockStyle.Right
        Me.barDockControlRight.Location = New System.Drawing.Point(1238, 36)
        Me.barDockControlRight.Manager = Me.BarManager
        Me.barDockControlRight.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.barDockControlRight.Size = New System.Drawing.Size(0, 626)
        '
        'BarLargeButtonItem1
        '
        Me.BarLargeButtonItem1.Caption = "BarLargeButtonItem1"
        Me.BarLargeButtonItem1.Id = 0
        Me.BarLargeButtonItem1.Name = "BarLargeButtonItem1"
        '
        'BarLargeButtonItem2
        '
        Me.BarLargeButtonItem2.Alignment = DevExpress.XtraBars.BarItemLinkAlignment.Right
        Me.BarLargeButtonItem2.Caption = "BarLargeButtonItem2"
        Me.BarLargeButtonItem2.Id = 1
        Me.BarLargeButtonItem2.Name = "BarLargeButtonItem2"
        '
        'PrintPreviewDialog
        '
        Me.PrintPreviewDialog.AutoScrollMargin = New System.Drawing.Size(0, 0)
        Me.PrintPreviewDialog.AutoScrollMinSize = New System.Drawing.Size(0, 0)
        Me.PrintPreviewDialog.ClientSize = New System.Drawing.Size(400, 300)
        Me.PrintPreviewDialog.Enabled = True
        Me.PrintPreviewDialog.Icon = CType(resources.GetObject("PrintPreviewDialog.Icon"), System.Drawing.Icon)
        Me.PrintPreviewDialog.Name = "PrintPreviewDialog"
        Me.PrintPreviewDialog.Visible = False
        '
        'GroupAdditional
        '
        Me.GroupAdditional.AppearanceCaption.BackColor = System.Drawing.Color.FromArgb(CType(CType(16, Byte), Integer), CType(CType(110, Byte), Integer), CType(CType(190, Byte), Integer))
        Me.GroupAdditional.AppearanceCaption.BackColor2 = System.Drawing.Color.FromArgb(CType(CType(16, Byte), Integer), CType(CType(110, Byte), Integer), CType(CType(190, Byte), Integer))
        Me.GroupAdditional.AppearanceCaption.BorderColor = System.Drawing.Color.FromArgb(CType(CType(16, Byte), Integer), CType(CType(110, Byte), Integer), CType(CType(190, Byte), Integer))
        Me.GroupAdditional.AppearanceCaption.ForeColor = System.Drawing.Color.White
        Me.GroupAdditional.AppearanceCaption.Options.UseBackColor = True
        Me.GroupAdditional.AppearanceCaption.Options.UseBorderColor = True
        Me.GroupAdditional.AppearanceCaption.Options.UseForeColor = True
        Me.GroupAdditional.CaptionImageOptions.AllowGlyphSkinning = True
        Me.GroupAdditional.CaptionImageOptions.SvgImage = CType(resources.GetObject("GroupAdditional.CaptionImageOptions.SvgImage"), DevExpress.Utils.Svg.SvgImage)
        Me.GroupAdditional.CaptionImageOptions.SvgImageSize = New System.Drawing.Size(16, 16)
        Me.GroupAdditional.Controls.Add(Me.LabelControl4)
        Me.GroupAdditional.Controls.Add(Me.cboMode)
        Me.GroupAdditional.Controls.Add(Me.txtChargeSlip)
        Me.GroupAdditional.Controls.Add(Me.LabelControl3)
        Me.GroupAdditional.Controls.Add(Me.tmTimeReceived)
        Me.GroupAdditional.Controls.Add(Me.LabelControl32)
        Me.GroupAdditional.Controls.Add(Me.dtReceived)
        Me.GroupAdditional.Controls.Add(Me.LabelControl33)
        Me.GroupAdditional.Controls.Add(Me.txtAccession)
        Me.GroupAdditional.Controls.Add(Me.LabelControl34)
        Me.GroupAdditional.Controls.Add(Me.cboPType)
        Me.GroupAdditional.Controls.Add(Me.LabelControl35)
        Me.GroupAdditional.Controls.Add(Me.txtORNo)
        Me.GroupAdditional.Controls.Add(Me.LabelControl36)
        Me.GroupAdditional.Controls.Add(Me.LabelControl37)
        Me.GroupAdditional.Controls.Add(Me.LabelControl38)
        Me.GroupAdditional.Controls.Add(Me.cboRoom)
        Me.GroupAdditional.Controls.Add(Me.LabelControl41)
        Me.GroupAdditional.Controls.Add(Me.cboPhysician)
        Me.GroupAdditional.Controls.Add(Me.cboRequest)
        Me.GroupAdditional.Location = New System.Drawing.Point(12, 350)
        Me.GroupAdditional.LookAndFeel.SkinName = "The Bezier"
        Me.GroupAdditional.LookAndFeel.UseDefaultLookAndFeel = False
        Me.GroupAdditional.Name = "GroupAdditional"
        Me.GroupAdditional.Size = New System.Drawing.Size(348, 302)
        Me.GroupAdditional.TabIndex = 189
        Me.GroupAdditional.Text = "Additional Details"
        '
        'LabelControl4
        '
        Me.LabelControl4.Appearance.ForeColor = System.Drawing.Color.FromArgb(CType(CType(27, Byte), Integer), CType(CType(41, Byte), Integer), CType(CType(62, Byte), Integer))
        Me.LabelControl4.Appearance.Options.UseForeColor = True
        Me.LabelControl4.Location = New System.Drawing.Point(8, 259)
        Me.LabelControl4.LookAndFeel.SkinName = "Visual Studio 2013 Blue"
        Me.LabelControl4.LookAndFeel.UseDefaultLookAndFeel = False
        Me.LabelControl4.Name = "LabelControl4"
        Me.LabelControl4.Size = New System.Drawing.Size(70, 13)
        Me.LabelControl4.TabIndex = 164
        Me.LabelControl4.Text = "Priority Level::"
        '
        'cboMode
        '
        Me.cboMode.EditValue = "Routine"
        Me.cboMode.Location = New System.Drawing.Point(102, 256)
        Me.cboMode.Name = "cboMode"
        Me.cboMode.Properties.Appearance.ForeColor = System.Drawing.Color.FromArgb(CType(CType(27, Byte), Integer), CType(CType(41, Byte), Integer), CType(CType(62, Byte), Integer))
        Me.cboMode.Properties.Appearance.Options.UseForeColor = True
        Me.cboMode.Properties.Buttons.AddRange(New DevExpress.XtraEditors.Controls.EditorButton() {New DevExpress.XtraEditors.Controls.EditorButton(DevExpress.XtraEditors.Controls.ButtonPredefines.Combo)})
        Me.cboMode.Properties.Items.AddRange(New Object() {"Stat", "Routine"})
        Me.cboMode.Properties.LookAndFeel.SkinName = "Visual Studio 2013 Blue"
        Me.cboMode.Properties.LookAndFeel.UseDefaultLookAndFeel = False
        Me.cboMode.Size = New System.Drawing.Size(234, 20)
        Me.cboMode.TabIndex = 163
        '
        'txtChargeSlip
        '
        Me.txtChargeSlip.Location = New System.Drawing.Point(102, 60)
        Me.txtChargeSlip.Name = "txtChargeSlip"
        Me.txtChargeSlip.Properties.Appearance.ForeColor = System.Drawing.Color.FromArgb(CType(CType(27, Byte), Integer), CType(CType(41, Byte), Integer), CType(CType(62, Byte), Integer))
        Me.txtChargeSlip.Properties.Appearance.Options.UseForeColor = True
        Me.txtChargeSlip.Properties.LookAndFeel.SkinName = "Visual Studio 2013 Blue"
        Me.txtChargeSlip.Properties.LookAndFeel.UseDefaultLookAndFeel = False
        Me.txtChargeSlip.Size = New System.Drawing.Size(234, 20)
        Me.txtChargeSlip.TabIndex = 161
        '
        'LabelControl3
        '
        Me.LabelControl3.Appearance.ForeColor = System.Drawing.Color.FromArgb(CType(CType(27, Byte), Integer), CType(CType(41, Byte), Integer), CType(CType(62, Byte), Integer))
        Me.LabelControl3.Appearance.Options.UseForeColor = True
        Me.LabelControl3.Location = New System.Drawing.Point(8, 63)
        Me.LabelControl3.LookAndFeel.SkinName = "Visual Studio 2013 Blue"
        Me.LabelControl3.LookAndFeel.UseDefaultLookAndFeel = False
        Me.LabelControl3.Name = "LabelControl3"
        Me.LabelControl3.Size = New System.Drawing.Size(72, 13)
        Me.LabelControl3.TabIndex = 162
        Me.LabelControl3.Text = "Charge Slip #:"
        '
        'tmTimeReceived
        '
        Me.tmTimeReceived.CalendarFont = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.tmTimeReceived.Format = System.Windows.Forms.DateTimePickerFormat.Time
        Me.tmTimeReceived.Location = New System.Drawing.Point(102, 134)
        Me.tmTimeReceived.Name = "tmTimeReceived"
        Me.tmTimeReceived.Size = New System.Drawing.Size(234, 22)
        Me.tmTimeReceived.TabIndex = 13
        '
        'LabelControl32
        '
        Me.LabelControl32.Appearance.ForeColor = System.Drawing.Color.FromArgb(CType(CType(27, Byte), Integer), CType(CType(41, Byte), Integer), CType(CType(62, Byte), Integer))
        Me.LabelControl32.Appearance.Options.UseForeColor = True
        Me.LabelControl32.Location = New System.Drawing.Point(8, 235)
        Me.LabelControl32.LookAndFeel.SkinName = "Visual Studio 2013 Blue"
        Me.LabelControl32.LookAndFeel.UseDefaultLookAndFeel = False
        Me.LabelControl32.Name = "LabelControl32"
        Me.LabelControl32.Size = New System.Drawing.Size(66, 13)
        Me.LabelControl32.TabIndex = 143
        Me.LabelControl32.Text = "Patient Type:"
        '
        'dtReceived
        '
        Me.dtReceived.CalendarFont = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.dtReceived.Format = System.Windows.Forms.DateTimePickerFormat.[Short]
        Me.dtReceived.Location = New System.Drawing.Point(102, 108)
        Me.dtReceived.Name = "dtReceived"
        Me.dtReceived.Size = New System.Drawing.Size(234, 22)
        Me.dtReceived.TabIndex = 12
        '
        'LabelControl33
        '
        Me.LabelControl33.Appearance.ForeColor = System.Drawing.Color.FromArgb(CType(CType(27, Byte), Integer), CType(CType(41, Byte), Integer), CType(CType(62, Byte), Integer))
        Me.LabelControl33.Appearance.Options.UseForeColor = True
        Me.LabelControl33.Location = New System.Drawing.Point(8, 211)
        Me.LabelControl33.LookAndFeel.SkinName = "Visual Studio 2013 Blue"
        Me.LabelControl33.LookAndFeel.UseDefaultLookAndFeel = False
        Me.LabelControl33.Name = "LabelControl33"
        Me.LabelControl33.Size = New System.Drawing.Size(65, 13)
        Me.LabelControl33.TabIndex = 28
        Me.LabelControl33.Text = "Room/Ward:"
        '
        'txtAccession
        '
        Me.txtAccession.Location = New System.Drawing.Point(102, 36)
        Me.txtAccession.Name = "txtAccession"
        Me.txtAccession.Properties.Appearance.ForeColor = System.Drawing.Color.FromArgb(CType(CType(27, Byte), Integer), CType(CType(41, Byte), Integer), CType(CType(62, Byte), Integer))
        Me.txtAccession.Properties.Appearance.Options.UseForeColor = True
        Me.txtAccession.Properties.LookAndFeel.SkinName = "Visual Studio 2013 Blue"
        Me.txtAccession.Properties.LookAndFeel.UseDefaultLookAndFeel = False
        Me.txtAccession.Size = New System.Drawing.Size(234, 20)
        Me.txtAccession.TabIndex = 10
        '
        'LabelControl34
        '
        Me.LabelControl34.Appearance.ForeColor = System.Drawing.Color.FromArgb(CType(CType(27, Byte), Integer), CType(CType(41, Byte), Integer), CType(CType(62, Byte), Integer))
        Me.LabelControl34.Appearance.Options.UseForeColor = True
        Me.LabelControl34.Location = New System.Drawing.Point(8, 87)
        Me.LabelControl34.LookAndFeel.SkinName = "Visual Studio 2013 Blue"
        Me.LabelControl34.LookAndFeel.UseDefaultLookAndFeel = False
        Me.LabelControl34.Name = "LabelControl34"
        Me.LabelControl34.Size = New System.Drawing.Size(40, 13)
        Me.LabelControl34.TabIndex = 160
        Me.LabelControl34.Text = "OR No.:"
        '
        'cboPType
        '
        Me.cboPType.EditValue = "Out Patient"
        Me.cboPType.Location = New System.Drawing.Point(102, 232)
        Me.cboPType.Name = "cboPType"
        Me.cboPType.Properties.Appearance.ForeColor = System.Drawing.Color.FromArgb(CType(CType(27, Byte), Integer), CType(CType(41, Byte), Integer), CType(CType(62, Byte), Integer))
        Me.cboPType.Properties.Appearance.Options.UseForeColor = True
        Me.cboPType.Properties.Buttons.AddRange(New DevExpress.XtraEditors.Controls.EditorButton() {New DevExpress.XtraEditors.Controls.EditorButton(DevExpress.XtraEditors.Controls.ButtonPredefines.Combo)})
        Me.cboPType.Properties.Items.AddRange(New Object() {"In Patient", "Out Patient"})
        Me.cboPType.Properties.LookAndFeel.SkinName = "Visual Studio 2013 Blue"
        Me.cboPType.Properties.LookAndFeel.UseDefaultLookAndFeel = False
        Me.cboPType.Size = New System.Drawing.Size(234, 20)
        Me.cboPType.TabIndex = 17
        '
        'LabelControl35
        '
        Me.LabelControl35.Appearance.ForeColor = System.Drawing.Color.FromArgb(CType(CType(27, Byte), Integer), CType(CType(41, Byte), Integer), CType(CType(62, Byte), Integer))
        Me.LabelControl35.Appearance.Options.UseForeColor = True
        Me.LabelControl35.Location = New System.Drawing.Point(8, 187)
        Me.LabelControl35.LookAndFeel.SkinName = "Visual Studio 2013 Blue"
        Me.LabelControl35.LookAndFeel.UseDefaultLookAndFeel = False
        Me.LabelControl35.Name = "LabelControl35"
        Me.LabelControl35.Size = New System.Drawing.Size(72, 13)
        Me.LabelControl35.TabIndex = 30
        Me.LabelControl35.Text = "Requested By:"
        '
        'txtORNo
        '
        Me.txtORNo.Location = New System.Drawing.Point(102, 84)
        Me.txtORNo.Name = "txtORNo"
        Me.txtORNo.Properties.Appearance.ForeColor = System.Drawing.Color.FromArgb(CType(CType(27, Byte), Integer), CType(CType(41, Byte), Integer), CType(CType(62, Byte), Integer))
        Me.txtORNo.Properties.Appearance.Options.UseForeColor = True
        Me.txtORNo.Properties.LookAndFeel.SkinName = "Visual Studio 2013 Blue"
        Me.txtORNo.Properties.LookAndFeel.UseDefaultLookAndFeel = False
        Me.txtORNo.Size = New System.Drawing.Size(234, 20)
        Me.txtORNo.TabIndex = 11
        '
        'LabelControl36
        '
        Me.LabelControl36.Appearance.ForeColor = System.Drawing.Color.FromArgb(CType(CType(27, Byte), Integer), CType(CType(41, Byte), Integer), CType(CType(62, Byte), Integer))
        Me.LabelControl36.Appearance.Options.UseForeColor = True
        Me.LabelControl36.Location = New System.Drawing.Point(8, 163)
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
        Me.LabelControl37.Location = New System.Drawing.Point(8, 141)
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
        Me.LabelControl38.Location = New System.Drawing.Point(8, 39)
        Me.LabelControl38.LookAndFeel.SkinName = "Visual Studio 2013 Blue"
        Me.LabelControl38.LookAndFeel.UseDefaultLookAndFeel = False
        Me.LabelControl38.Name = "LabelControl38"
        Me.LabelControl38.Size = New System.Drawing.Size(77, 13)
        Me.LabelControl38.TabIndex = 158
        Me.LabelControl38.Text = "HIS Tracking #:"
        '
        'cboRoom
        '
        Me.cboRoom.Location = New System.Drawing.Point(102, 208)
        Me.cboRoom.Name = "cboRoom"
        Me.cboRoom.Properties.Appearance.ForeColor = System.Drawing.Color.FromArgb(CType(CType(27, Byte), Integer), CType(CType(41, Byte), Integer), CType(CType(62, Byte), Integer))
        Me.cboRoom.Properties.Appearance.Options.UseForeColor = True
        Me.cboRoom.Properties.Buttons.AddRange(New DevExpress.XtraEditors.Controls.EditorButton() {New DevExpress.XtraEditors.Controls.EditorButton(DevExpress.XtraEditors.Controls.ButtonPredefines.Combo)})
        Me.cboRoom.Properties.LookAndFeel.SkinName = "Visual Studio 2013 Blue"
        Me.cboRoom.Properties.LookAndFeel.UseDefaultLookAndFeel = False
        Me.cboRoom.Size = New System.Drawing.Size(234, 20)
        Me.cboRoom.TabIndex = 16
        '
        'LabelControl41
        '
        Me.LabelControl41.Appearance.ForeColor = System.Drawing.Color.FromArgb(CType(CType(27, Byte), Integer), CType(CType(41, Byte), Integer), CType(CType(62, Byte), Integer))
        Me.LabelControl41.Appearance.Options.UseForeColor = True
        Me.LabelControl41.Location = New System.Drawing.Point(8, 115)
        Me.LabelControl41.LookAndFeel.SkinName = "Visual Studio 2013 Blue"
        Me.LabelControl41.LookAndFeel.UseDefaultLookAndFeel = False
        Me.LabelControl41.Name = "LabelControl41"
        Me.LabelControl41.Size = New System.Drawing.Size(75, 13)
        Me.LabelControl41.TabIndex = 136
        Me.LabelControl41.Text = "Date Received:"
        '
        'cboPhysician
        '
        Me.cboPhysician.Location = New System.Drawing.Point(102, 184)
        Me.cboPhysician.Name = "cboPhysician"
        Me.cboPhysician.Properties.Appearance.ForeColor = System.Drawing.Color.FromArgb(CType(CType(27, Byte), Integer), CType(CType(41, Byte), Integer), CType(CType(62, Byte), Integer))
        Me.cboPhysician.Properties.Appearance.Options.UseForeColor = True
        Me.cboPhysician.Properties.Buttons.AddRange(New DevExpress.XtraEditors.Controls.EditorButton() {New DevExpress.XtraEditors.Controls.EditorButton(DevExpress.XtraEditors.Controls.ButtonPredefines.Combo)})
        Me.cboPhysician.Properties.LookAndFeel.SkinName = "Visual Studio 2013 Blue"
        Me.cboPhysician.Properties.LookAndFeel.UseDefaultLookAndFeel = False
        Me.cboPhysician.Size = New System.Drawing.Size(234, 20)
        Me.cboPhysician.TabIndex = 15
        '
        'cboRequest
        '
        Me.cboRequest.Location = New System.Drawing.Point(102, 160)
        Me.cboRequest.Name = "cboRequest"
        Me.cboRequest.Properties.Appearance.ForeColor = System.Drawing.Color.FromArgb(CType(CType(27, Byte), Integer), CType(CType(41, Byte), Integer), CType(CType(62, Byte), Integer))
        Me.cboRequest.Properties.Appearance.Options.UseForeColor = True
        Me.cboRequest.Properties.Buttons.AddRange(New DevExpress.XtraEditors.Controls.EditorButton() {New DevExpress.XtraEditors.Controls.EditorButton(DevExpress.XtraEditors.Controls.ButtonPredefines.Combo)})
        Me.cboRequest.Properties.LookAndFeel.SkinName = "Visual Studio 2013 Blue"
        Me.cboRequest.Properties.LookAndFeel.UseDefaultLookAndFeel = False
        Me.cboRequest.Size = New System.Drawing.Size(234, 20)
        Me.cboRequest.TabIndex = 14
        '
        'frmPatientOrderAE
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(1238, 662)
        Me.Controls.Add(Me.GroupAdditional)
        Me.Controls.Add(Me.GroupPatient)
        Me.Controls.Add(Me.bcCode)
        Me.Controls.Add(Me.GroupTestDetails)
        Me.Controls.Add(Me.GroupControl1)
        Me.Controls.Add(Me.picCode)
        Me.Controls.Add(Me.GroupOrder)
        Me.Controls.Add(Me.barDockControlLeft)
        Me.Controls.Add(Me.barDockControlRight)
        Me.Controls.Add(Me.barDockControlBottom)
        Me.Controls.Add(Me.barDockControlTop)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "frmPatientOrderAE"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Patient Order Details"
        CType(Me.picCode, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.txtSampleID.Properties, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.txtClass.Properties, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.txtAge.Properties, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.dtBDate.Properties.CalendarTimeProperties, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.dtBDate.Properties, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.txtPatientID.Properties, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.cboSex.Properties, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.txtContact.Properties, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.txtAddress.Properties, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.txtName.Properties, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.cboTestName.Properties, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.GroupOrder, System.ComponentModel.ISupportInitialize).EndInit()
        Me.GroupOrder.ResumeLayout(False)
        CType(Me.GroupControl1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.GroupControl1.ResumeLayout(False)
        Me.GroupControl1.PerformLayout()
        CType(Me.GroupTestDetails, System.ComponentModel.ISupportInitialize).EndInit()
        Me.GroupTestDetails.ResumeLayout(False)
        Me.GroupTestDetails.PerformLayout()
        CType(Me.GroupPatient, System.ComponentModel.ISupportInitialize).EndInit()
        Me.GroupPatient.ResumeLayout(False)
        Me.GroupPatient.PerformLayout()
        CType(Me.txtEmail.Properties, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.cboStatus.Properties, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.BarManager, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.BarAndDockingController, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.GroupAdditional, System.ComponentModel.ISupportInitialize).EndInit()
        Me.GroupAdditional.ResumeLayout(False)
        Me.GroupAdditional.PerformLayout()
        CType(Me.cboMode.Properties, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.txtChargeSlip.Properties, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.txtAccession.Properties, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.cboPType.Properties, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.txtORNo.Properties, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.cboRoom.Properties, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.cboPhysician.Properties, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.cboRequest.Properties, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents lvList As ListView
    Friend WithEvents ColumnHeader10 As ColumnHeader
    Friend WithEvents ColumnHeader1 As ColumnHeader
    Friend WithEvents ColumnHeader3 As ColumnHeader
    Friend WithEvents bcCode As DevExpress.XtraEditors.BarCodeControl
    Friend WithEvents picCode As PictureBox
    Friend WithEvents LabelControl7 As DevExpress.XtraEditors.LabelControl
    Friend WithEvents txtSampleID As DevExpress.XtraEditors.TextEdit
    Friend WithEvents txtClass As DevExpress.XtraEditors.ComboBoxEdit
    Friend WithEvents LabelControl13 As DevExpress.XtraEditors.LabelControl
    Friend WithEvents txtAge As DevExpress.XtraEditors.TextEdit
    Friend WithEvents LabelControl6 As DevExpress.XtraEditors.LabelControl
    Friend WithEvents LabelControl2 As DevExpress.XtraEditors.LabelControl
    Friend WithEvents dtBDate As DevExpress.XtraEditors.DateEdit
    Friend WithEvents LabelControl1 As DevExpress.XtraEditors.LabelControl
    Friend WithEvents txtPatientID As DevExpress.XtraEditors.TextEdit
    Friend WithEvents LabelControl39 As DevExpress.XtraEditors.LabelControl
    Friend WithEvents LabelControl15 As DevExpress.XtraEditors.LabelControl
    Friend WithEvents cboSex As DevExpress.XtraEditors.ComboBoxEdit
    Friend WithEvents txtContact As DevExpress.XtraEditors.TextEdit
    Friend WithEvents btnSearch As DevExpress.XtraEditors.SimpleButton
    Friend WithEvents LabelControl14 As DevExpress.XtraEditors.LabelControl
    Friend WithEvents txtAddress As DevExpress.XtraEditors.MemoEdit
    Friend WithEvents txtName As DevExpress.XtraEditors.MemoEdit
    Friend WithEvents cboTestName As DevExpress.XtraEditors.ComboBoxEdit
    Friend WithEvents LabelControl5 As DevExpress.XtraEditors.LabelControl
    Friend WithEvents lvOrder As ListView
    Friend WithEvents ColumnHeader6 As ColumnHeader
    Friend WithEvents ColumnHeader12 As ColumnHeader
    Friend WithEvents ColumnHeader13 As ColumnHeader
    Friend WithEvents ColumnHeader4 As ColumnHeader
    Friend WithEvents Timer1 As Timer
    Friend WithEvents GroupTestDetails As DevExpress.XtraEditors.GroupControl
    Friend WithEvents GroupControl1 As DevExpress.XtraEditors.GroupControl
    Friend WithEvents GroupOrder As DevExpress.XtraEditors.GroupControl
    Friend WithEvents lvTest As ListView
    Friend WithEvents ColumnHeader2 As ColumnHeader
    Friend WithEvents ColumnHeader5 As ColumnHeader
    Friend WithEvents chAll As CheckBox
    Friend WithEvents btnRemove As DevExpress.XtraEditors.SimpleButton
    Friend WithEvents btnDelete As DevExpress.XtraEditors.SimpleButton
    Friend WithEvents btnAdd As DevExpress.XtraEditors.SimpleButton
    Friend WithEvents GroupPatient As DevExpress.XtraEditors.GroupControl
    Friend WithEvents BarManager As DevExpress.XtraBars.BarManager
    Friend WithEvents Bar As DevExpress.XtraBars.Bar
    Friend WithEvents barDockControlTop As DevExpress.XtraBars.BarDockControl
    Friend WithEvents barDockControlBottom As DevExpress.XtraBars.BarDockControl
    Friend WithEvents barDockControlLeft As DevExpress.XtraBars.BarDockControl
    Friend WithEvents barDockControlRight As DevExpress.XtraBars.BarDockControl
    Friend WithEvents BarLargeButtonItem1 As DevExpress.XtraBars.BarLargeButtonItem
    Friend WithEvents BarLargeButtonItem2 As DevExpress.XtraBars.BarLargeButtonItem
    Friend WithEvents btnSave As DevExpress.XtraBars.BarLargeButtonItem
    Friend WithEvents btnSavePrint As DevExpress.XtraBars.BarLargeButtonItem
    Friend WithEvents BarAndDockingController As DevExpress.XtraBars.BarAndDockingController
    Friend WithEvents btnClear As DevExpress.XtraBars.BarLargeButtonItem
    Friend WithEvents btnClose As DevExpress.XtraBars.BarLargeButtonItem
    Friend WithEvents ColumnHeader7 As ColumnHeader
    Friend WithEvents lv As ListView
    Friend WithEvents ColumnHeader8 As ColumnHeader
    Friend WithEvents ColumnHeader9 As ColumnHeader
    Friend WithEvents ImageList As ImageList
    Friend WithEvents PrintPreviewDialog As PrintPreviewDialog
    Friend WithEvents PrintDocument As Printing.PrintDocument
    Friend WithEvents lvDup As ListView
    Friend WithEvents ColumnHeader15 As ColumnHeader
    Friend WithEvents ColumnHeader16 As ColumnHeader
    Friend WithEvents ColumnHeader17 As ColumnHeader
    Friend WithEvents ColumnHeader18 As ColumnHeader
    Friend WithEvents ColumnHeader19 As ColumnHeader
    Friend WithEvents ColumnHeader21 As ColumnHeader
    Friend WithEvents ColumnHeader22 As ColumnHeader
    Friend WithEvents ColumnHeader23 As ColumnHeader
    Friend WithEvents ColumnHeader24 As ColumnHeader
    Friend WithEvents GroupAdditional As DevExpress.XtraEditors.GroupControl
    Friend WithEvents txtChargeSlip As DevExpress.XtraEditors.TextEdit
    Friend WithEvents LabelControl3 As DevExpress.XtraEditors.LabelControl
    Friend WithEvents tmTimeReceived As DateTimePicker
    Friend WithEvents LabelControl32 As DevExpress.XtraEditors.LabelControl
    Friend WithEvents dtReceived As DateTimePicker
    Friend WithEvents LabelControl33 As DevExpress.XtraEditors.LabelControl
    Friend WithEvents txtAccession As DevExpress.XtraEditors.TextEdit
    Friend WithEvents LabelControl34 As DevExpress.XtraEditors.LabelControl
    Friend WithEvents cboPType As DevExpress.XtraEditors.ComboBoxEdit
    Friend WithEvents LabelControl35 As DevExpress.XtraEditors.LabelControl
    Friend WithEvents txtORNo As DevExpress.XtraEditors.TextEdit
    Friend WithEvents LabelControl36 As DevExpress.XtraEditors.LabelControl
    Friend WithEvents LabelControl37 As DevExpress.XtraEditors.LabelControl
    Friend WithEvents LabelControl38 As DevExpress.XtraEditors.LabelControl
    Friend WithEvents cboRoom As DevExpress.XtraEditors.ComboBoxEdit
    Friend WithEvents LabelControl41 As DevExpress.XtraEditors.LabelControl
    Friend WithEvents cboPhysician As DevExpress.XtraEditors.ComboBoxEdit
    Friend WithEvents cboRequest As DevExpress.XtraEditors.ComboBoxEdit
    Friend WithEvents LabelControl4 As DevExpress.XtraEditors.LabelControl
    Friend WithEvents cboMode As DevExpress.XtraEditors.ComboBoxEdit
    Friend WithEvents SimpleButton1 As DevExpress.XtraEditors.SimpleButton
    Friend WithEvents ColumnHeader11 As ColumnHeader
    Friend WithEvents ColumnHeader20 As ColumnHeader
    Friend WithEvents ColumnHeader14 As ColumnHeader
    Friend WithEvents ColumnHeader25 As ColumnHeader
    Friend WithEvents txtEmail As DevExpress.XtraEditors.TextEdit
    Friend WithEvents LabelControl9 As DevExpress.XtraEditors.LabelControl
    Friend WithEvents LabelControl8 As DevExpress.XtraEditors.LabelControl
    Friend WithEvents cboStatus As DevExpress.XtraEditors.ComboBoxEdit
    Friend WithEvents ColumnHeader26 As ColumnHeader
End Class
