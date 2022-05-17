<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class frmMicroWorklist
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmMicroWorklist))
        Dim Code128Generator1 As DevExpress.XtraPrinting.BarCode.Code128Generator = New DevExpress.XtraPrinting.BarCode.Code128Generator()
        Me.pnlBackground = New DevExpress.XtraEditors.PanelControl()
        Me.PanelControl4 = New DevExpress.XtraEditors.PanelControl()
        Me.lblCountQueue = New DevExpress.XtraEditors.LabelControl()
        Me.SeparatorControl3 = New DevExpress.XtraEditors.SeparatorControl()
        Me.txtSearch1 = New DevExpress.XtraTab.XtraTabControl()
        Me.XTPage1 = New DevExpress.XtraTab.XtraTabPage()
        Me.GroupControl2 = New DevExpress.XtraEditors.GroupControl()
        Me.dtList = New DevExpress.XtraGrid.GridControl()
        Me.GridView = New DevExpress.XtraGrid.Views.Grid.GridView()
        Me.BarManager = New DevExpress.XtraBars.BarManager(Me.components)
        Me.Bar = New DevExpress.XtraBars.Bar()
        Me.btnBarcode = New DevExpress.XtraBars.BarLargeButtonItem()
        Me.btnPrint = New DevExpress.XtraBars.BarLargeButtonItem()
        Me.btnView = New DevExpress.XtraBars.BarLargeButtonItem()
        Me.btnDelete = New DevExpress.XtraBars.BarLargeButtonItem()
        Me.btnReject = New DevExpress.XtraBars.BarLargeButtonItem()
        Me.btnRefresh = New DevExpress.XtraBars.BarLargeButtonItem()
        Me.btnClose = New DevExpress.XtraBars.BarLargeButtonItem()
        Me.BarAndDockingController = New DevExpress.XtraBars.BarAndDockingController(Me.components)
        Me.barDockControlTop = New DevExpress.XtraBars.BarDockControl()
        Me.barDockControlBottom = New DevExpress.XtraBars.BarDockControl()
        Me.barDockControlLeft = New DevExpress.XtraBars.BarDockControl()
        Me.barDockControlRight = New DevExpress.XtraBars.BarDockControl()
        Me.BarButtonItem1 = New DevExpress.XtraBars.BarButtonItem()
        Me.BarButtonItem2 = New DevExpress.XtraBars.BarButtonItem()
        Me.BarButtonItem3 = New DevExpress.XtraBars.BarButtonItem()
        Me.BarButtonItem4 = New DevExpress.XtraBars.BarButtonItem()
        Me.BarButtonItem5 = New DevExpress.XtraBars.BarButtonItem()
        Me.BarButtonItem6 = New DevExpress.XtraBars.BarButtonItem()
        Me.btnNewOrder = New DevExpress.XtraBars.BarLargeButtonItem()
        Me.PanelControl2 = New DevExpress.XtraEditors.PanelControl()
        Me.LabelControl4 = New DevExpress.XtraEditors.LabelControl()
        Me.txtSearch = New DevExpress.XtraEditors.SearchControl()
        Me.LabelControl6 = New DevExpress.XtraEditors.LabelControl()
        Me.cboLocation = New DevExpress.XtraEditors.ComboBoxEdit()
        Me.bcCode = New DevExpress.XtraEditors.BarCodeControl()
        Me.picCode = New System.Windows.Forms.PictureBox()
        Me.dtTo = New System.Windows.Forms.DateTimePicker()
        Me.LabelControl2 = New DevExpress.XtraEditors.LabelControl()
        Me.dtFrom = New System.Windows.Forms.DateTimePicker()
        Me.XTPage2 = New DevExpress.XtraTab.XtraTabPage()
        Me.GroupControl1 = New DevExpress.XtraEditors.GroupControl()
        Me.dtCompleted = New DevExpress.XtraGrid.GridControl()
        Me.GridCompleted = New DevExpress.XtraGrid.Views.Grid.GridView()
        Me.PanelControl1 = New DevExpress.XtraEditors.PanelControl()
        Me.LabelControl5 = New DevExpress.XtraEditors.LabelControl()
        Me.SearchControl1 = New DevExpress.XtraEditors.SearchControl()
        Me.LabelControl1 = New DevExpress.XtraEditors.LabelControl()
        Me.cboLocation1 = New DevExpress.XtraEditors.ComboBoxEdit()
        Me.pnlHeader = New DevExpress.XtraEditors.PanelControl()
        Me.btnSearch = New DevExpress.XtraEditors.SimpleButton()
        Me.dtFrom1 = New System.Windows.Forms.DateTimePicker()
        Me.dtTo1 = New System.Windows.Forms.DateTimePicker()
        Me.LabelControl3 = New DevExpress.XtraEditors.LabelControl()
        Me.ImageList = New System.Windows.Forms.ImageList(Me.components)
        Me.tmLoadNew = New System.Windows.Forms.Timer(Me.components)
        Me.tmAutoValidate = New System.Windows.Forms.Timer(Me.components)
        Me.PrintDocument = New System.Drawing.Printing.PrintDocument()
        Me.PrintPreviewDialog = New System.Windows.Forms.PrintPreviewDialog()
        Me.LabelControl7 = New DevExpress.XtraEditors.LabelControl()
        Me.rgSelect = New DevExpress.XtraEditors.RadioGroup()
        Me.LabelControl8 = New DevExpress.XtraEditors.LabelControl()
        Me.rgSelect1 = New DevExpress.XtraEditors.RadioGroup()
        CType(Me.pnlBackground, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.pnlBackground.SuspendLayout()
        CType(Me.PanelControl4, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.PanelControl4.SuspendLayout()
        CType(Me.SeparatorControl3, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.txtSearch1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.txtSearch1.SuspendLayout()
        Me.XTPage1.SuspendLayout()
        CType(Me.GroupControl2, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.GroupControl2.SuspendLayout()
        CType(Me.dtList, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.GridView, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.BarManager, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.BarAndDockingController, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.PanelControl2, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.PanelControl2.SuspendLayout()
        CType(Me.txtSearch.Properties, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.cboLocation.Properties, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.picCode, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.XTPage2.SuspendLayout()
        CType(Me.GroupControl1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.GroupControl1.SuspendLayout()
        CType(Me.dtCompleted, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.GridCompleted, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.PanelControl1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.PanelControl1.SuspendLayout()
        CType(Me.SearchControl1.Properties, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.cboLocation1.Properties, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.pnlHeader, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.pnlHeader.SuspendLayout()
        CType(Me.rgSelect.Properties, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.rgSelect1.Properties, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'pnlBackground
        '
        Me.pnlBackground.Appearance.BackColor = System.Drawing.Color.White
        Me.pnlBackground.Appearance.Options.UseBackColor = True
        Me.pnlBackground.BorderStyle = DevExpress.XtraEditors.Controls.BorderStyles.NoBorder
        Me.pnlBackground.Controls.Add(Me.PanelControl4)
        Me.pnlBackground.Controls.Add(Me.txtSearch1)
        Me.pnlBackground.Controls.Add(Me.barDockControlLeft)
        Me.pnlBackground.Controls.Add(Me.barDockControlRight)
        Me.pnlBackground.Controls.Add(Me.barDockControlBottom)
        Me.pnlBackground.Controls.Add(Me.barDockControlTop)
        Me.pnlBackground.Dock = System.Windows.Forms.DockStyle.Fill
        Me.pnlBackground.Location = New System.Drawing.Point(0, 0)
        Me.pnlBackground.Name = "pnlBackground"
        Me.pnlBackground.Size = New System.Drawing.Size(1271, 610)
        Me.pnlBackground.TabIndex = 6
        '
        'PanelControl4
        '
        Me.PanelControl4.Appearance.BackColor = System.Drawing.Color.FromArgb(CType(CType(16, Byte), Integer), CType(CType(110, Byte), Integer), CType(CType(190, Byte), Integer))
        Me.PanelControl4.Appearance.Options.UseBackColor = True
        Me.PanelControl4.BorderStyle = DevExpress.XtraEditors.Controls.BorderStyles.NoBorder
        Me.PanelControl4.Controls.Add(Me.lblCountQueue)
        Me.PanelControl4.Controls.Add(Me.SeparatorControl3)
        Me.PanelControl4.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.PanelControl4.Location = New System.Drawing.Point(0, 580)
        Me.PanelControl4.Name = "PanelControl4"
        Me.PanelControl4.Size = New System.Drawing.Size(1271, 30)
        Me.PanelControl4.TabIndex = 81
        '
        'lblCountQueue
        '
        Me.lblCountQueue.Appearance.Font = New System.Drawing.Font("Segoe UI", 8.25!, System.Drawing.FontStyle.Bold)
        Me.lblCountQueue.Appearance.ForeColor = System.Drawing.Color.White
        Me.lblCountQueue.Appearance.Options.UseFont = True
        Me.lblCountQueue.Appearance.Options.UseForeColor = True
        Me.lblCountQueue.Location = New System.Drawing.Point(9, 9)
        Me.lblCountQueue.Name = "lblCountQueue"
        Me.lblCountQueue.Size = New System.Drawing.Size(74, 13)
        Me.lblCountQueue.TabIndex = 0
        Me.lblCountQueue.Text = "Record Count:"
        '
        'SeparatorControl3
        '
        Me.SeparatorControl3.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.SeparatorControl3.AutoSizeMode = True
        Me.SeparatorControl3.Location = New System.Drawing.Point(-10, -9)
        Me.SeparatorControl3.Name = "SeparatorControl3"
        Me.SeparatorControl3.Size = New System.Drawing.Size(1575, 19)
        Me.SeparatorControl3.TabIndex = 70
        '
        'txtSearch1
        '
        Me.txtSearch1.AllowGlyphSkinning = DevExpress.Utils.DefaultBoolean.[True]
        Me.txtSearch1.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtSearch1.AppearancePage.Header.Font = New System.Drawing.Font("Segoe UI", 8.25!, System.Drawing.FontStyle.Bold)
        Me.txtSearch1.AppearancePage.Header.Options.UseFont = True
        Me.txtSearch1.AppearancePage.HeaderActive.BackColor = System.Drawing.Color.FromArgb(CType(CType(16, Byte), Integer), CType(CType(110, Byte), Integer), CType(CType(190, Byte), Integer))
        Me.txtSearch1.AppearancePage.HeaderActive.BackColor2 = System.Drawing.Color.FromArgb(CType(CType(16, Byte), Integer), CType(CType(110, Byte), Integer), CType(CType(190, Byte), Integer))
        Me.txtSearch1.AppearancePage.HeaderActive.BorderColor = System.Drawing.Color.FromArgb(CType(CType(16, Byte), Integer), CType(CType(110, Byte), Integer), CType(CType(190, Byte), Integer))
        Me.txtSearch1.AppearancePage.HeaderActive.Options.UseBackColor = True
        Me.txtSearch1.AppearancePage.HeaderActive.Options.UseBorderColor = True
        Me.txtSearch1.HeaderAutoFill = DevExpress.Utils.DefaultBoolean.[True]
        Me.txtSearch1.HeaderButtonsShowMode = DevExpress.XtraTab.TabButtonShowMode.Always
        Me.txtSearch1.Location = New System.Drawing.Point(10, 42)
        Me.txtSearch1.Name = "txtSearch1"
        Me.txtSearch1.SelectedTabPage = Me.XTPage1
        Me.txtSearch1.Size = New System.Drawing.Size(1251, 532)
        Me.txtSearch1.TabIndex = 25
        Me.txtSearch1.TabPages.AddRange(New DevExpress.XtraTab.XtraTabPage() {Me.XTPage1, Me.XTPage2})
        '
        'XTPage1
        '
        Me.XTPage1.Controls.Add(Me.GroupControl2)
        Me.XTPage1.Controls.Add(Me.PanelControl2)
        Me.XTPage1.Controls.Add(Me.bcCode)
        Me.XTPage1.Controls.Add(Me.picCode)
        Me.XTPage1.Controls.Add(Me.dtTo)
        Me.XTPage1.Controls.Add(Me.LabelControl2)
        Me.XTPage1.Controls.Add(Me.dtFrom)
        Me.XTPage1.ImageOptions.SvgImage = CType(resources.GetObject("XTPage1.ImageOptions.SvgImage"), DevExpress.Utils.Svg.SvgImage)
        Me.XTPage1.ImageOptions.SvgImageSize = New System.Drawing.Size(16, 16)
        Me.XTPage1.Name = "XTPage1"
        Me.XTPage1.Size = New System.Drawing.Size(1249, 500)
        Me.XTPage1.Text = "Worklist"
        '
        'GroupControl2
        '
        Me.GroupControl2.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.GroupControl2.AppearanceCaption.BackColor = System.Drawing.Color.FromArgb(CType(CType(16, Byte), Integer), CType(CType(110, Byte), Integer), CType(CType(190, Byte), Integer))
        Me.GroupControl2.AppearanceCaption.BackColor2 = System.Drawing.Color.FromArgb(CType(CType(16, Byte), Integer), CType(CType(110, Byte), Integer), CType(CType(190, Byte), Integer))
        Me.GroupControl2.AppearanceCaption.BorderColor = System.Drawing.Color.FromArgb(CType(CType(16, Byte), Integer), CType(CType(110, Byte), Integer), CType(CType(190, Byte), Integer))
        Me.GroupControl2.AppearanceCaption.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.GroupControl2.AppearanceCaption.Options.UseBackColor = True
        Me.GroupControl2.AppearanceCaption.Options.UseBorderColor = True
        Me.GroupControl2.AppearanceCaption.Options.UseFont = True
        Me.GroupControl2.Controls.Add(Me.dtList)
        Me.GroupControl2.Location = New System.Drawing.Point(6, 44)
        Me.GroupControl2.Name = "GroupControl2"
        Me.GroupControl2.Size = New System.Drawing.Size(1240, 453)
        Me.GroupControl2.TabIndex = 176
        Me.GroupControl2.Text = "List"
        '
        'dtList
        '
        Me.dtList.Dock = System.Windows.Forms.DockStyle.Fill
        Me.dtList.Location = New System.Drawing.Point(2, 27)
        Me.dtList.MainView = Me.GridView
        Me.dtList.MenuManager = Me.BarManager
        Me.dtList.Name = "dtList"
        Me.dtList.Size = New System.Drawing.Size(1236, 424)
        Me.dtList.TabIndex = 146
        Me.dtList.ViewCollection.AddRange(New DevExpress.XtraGrid.Views.Base.BaseView() {Me.GridView})
        '
        'GridView
        '
        Me.GridView.Appearance.FocusedRow.Font = New System.Drawing.Font("Segoe UI", 8.25!, System.Drawing.FontStyle.Bold)
        Me.GridView.Appearance.FocusedRow.Options.UseFont = True
        Me.GridView.GridControl = Me.dtList
        Me.GridView.Name = "GridView"
        Me.GridView.OptionsBehavior.AutoExpandAllGroups = True
        Me.GridView.OptionsCustomization.AllowGroup = False
        Me.GridView.OptionsCustomization.AllowQuickHideColumns = False
        Me.GridView.OptionsSelection.CheckBoxSelectorColumnWidth = 30
        Me.GridView.OptionsView.ShowGroupPanel = False
        Me.GridView.OptionsView.ShowVerticalLines = DevExpress.Utils.DefaultBoolean.[False]
        '
        'BarManager
        '
        Me.BarManager.AllowCustomization = False
        Me.BarManager.AllowMoveBarOnToolbar = False
        Me.BarManager.AllowQuickCustomization = False
        Me.BarManager.AllowShowToolbarsPopup = False
        Me.BarManager.Bars.AddRange(New DevExpress.XtraBars.Bar() {Me.Bar})
        Me.BarManager.Controller = Me.BarAndDockingController
        Me.BarManager.DockControls.Add(Me.barDockControlTop)
        Me.BarManager.DockControls.Add(Me.barDockControlBottom)
        Me.BarManager.DockControls.Add(Me.barDockControlLeft)
        Me.BarManager.DockControls.Add(Me.barDockControlRight)
        Me.BarManager.DockWindowTabFont = New System.Drawing.Font("Tahoma", 10.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.BarManager.Form = Me.pnlBackground
        Me.BarManager.Items.AddRange(New DevExpress.XtraBars.BarItem() {Me.BarButtonItem1, Me.BarButtonItem2, Me.BarButtonItem3, Me.BarButtonItem4, Me.BarButtonItem5, Me.BarButtonItem6, Me.btnView, Me.btnDelete, Me.btnRefresh, Me.btnClose, Me.btnNewOrder, Me.btnPrint, Me.btnReject, Me.btnBarcode})
        Me.BarManager.MainMenu = Me.Bar
        Me.BarManager.MaxItemId = 22
        '
        'Bar
        '
        Me.Bar.BarName = "Main menu"
        Me.Bar.DockCol = 0
        Me.Bar.DockRow = 0
        Me.Bar.DockStyle = DevExpress.XtraBars.BarDockStyle.Top
        Me.Bar.LinksPersistInfo.AddRange(New DevExpress.XtraBars.LinkPersistInfo() {New DevExpress.XtraBars.LinkPersistInfo(Me.btnBarcode, True), New DevExpress.XtraBars.LinkPersistInfo(Me.btnPrint, True), New DevExpress.XtraBars.LinkPersistInfo(Me.btnView, True), New DevExpress.XtraBars.LinkPersistInfo(Me.btnDelete, True), New DevExpress.XtraBars.LinkPersistInfo(Me.btnReject, True), New DevExpress.XtraBars.LinkPersistInfo(Me.btnRefresh, True), New DevExpress.XtraBars.LinkPersistInfo(Me.btnClose)})
        Me.Bar.OptionsBar.AllowQuickCustomization = False
        Me.Bar.OptionsBar.MultiLine = True
        Me.Bar.OptionsBar.UseWholeRow = True
        Me.Bar.Text = "Main menu"
        '
        'btnBarcode
        '
        Me.btnBarcode.Caption = "Re-Print Barcode"
        Me.btnBarcode.CaptionAlignment = DevExpress.XtraBars.BarItemCaptionAlignment.Right
        Me.btnBarcode.Id = 20
        Me.btnBarcode.ImageOptions.AllowGlyphSkinning = DevExpress.Utils.DefaultBoolean.[True]
        Me.btnBarcode.ImageOptions.SvgImage = CType(resources.GetObject("btnBarcode.ImageOptions.SvgImage"), DevExpress.Utils.Svg.SvgImage)
        Me.btnBarcode.ImageOptions.SvgImageSize = New System.Drawing.Size(24, 24)
        Me.btnBarcode.Name = "btnBarcode"
        '
        'btnPrint
        '
        Me.btnPrint.Caption = "&Print (F4)"
        Me.btnPrint.CaptionAlignment = DevExpress.XtraBars.BarItemCaptionAlignment.Right
        Me.btnPrint.Id = 17
        Me.btnPrint.ImageOptions.AllowGlyphSkinning = DevExpress.Utils.DefaultBoolean.[True]
        Me.btnPrint.ImageOptions.SvgImage = CType(resources.GetObject("btnPrint.ImageOptions.SvgImage"), DevExpress.Utils.Svg.SvgImage)
        Me.btnPrint.ImageOptions.SvgImageSize = New System.Drawing.Size(24, 24)
        Me.btnPrint.ItemShortcut = New DevExpress.XtraBars.BarShortcut(System.Windows.Forms.Keys.F4)
        Me.btnPrint.Name = "btnPrint"
        Me.btnPrint.Visibility = DevExpress.XtraBars.BarItemVisibility.Never
        '
        'btnView
        '
        Me.btnView.Caption = "View (F1)"
        Me.btnView.CaptionAlignment = DevExpress.XtraBars.BarItemCaptionAlignment.Right
        Me.btnView.Id = 6
        Me.btnView.ImageOptions.AllowGlyphSkinning = DevExpress.Utils.DefaultBoolean.[True]
        Me.btnView.ImageOptions.SvgImage = CType(resources.GetObject("btnView.ImageOptions.SvgImage"), DevExpress.Utils.Svg.SvgImage)
        Me.btnView.ImageOptions.SvgImageSize = New System.Drawing.Size(24, 24)
        Me.btnView.ItemShortcut = New DevExpress.XtraBars.BarShortcut(System.Windows.Forms.Keys.F1)
        Me.btnView.Name = "btnView"
        '
        'btnDelete
        '
        Me.btnDelete.Caption = "Reject Order"
        Me.btnDelete.CaptionAlignment = DevExpress.XtraBars.BarItemCaptionAlignment.Right
        Me.btnDelete.Id = 9
        Me.btnDelete.ImageOptions.AllowGlyphSkinning = DevExpress.Utils.DefaultBoolean.[True]
        Me.btnDelete.ImageOptions.SvgImage = CType(resources.GetObject("btnDelete.ImageOptions.SvgImage"), DevExpress.Utils.Svg.SvgImage)
        Me.btnDelete.ImageOptions.SvgImageSize = New System.Drawing.Size(24, 24)
        Me.btnDelete.Name = "btnDelete"
        '
        'btnReject
        '
        Me.btnReject.Caption = "Cancel Order"
        Me.btnReject.CaptionAlignment = DevExpress.XtraBars.BarItemCaptionAlignment.Right
        Me.btnReject.Id = 19
        Me.btnReject.ImageOptions.AllowGlyphSkinning = DevExpress.Utils.DefaultBoolean.[True]
        Me.btnReject.ImageOptions.SvgImage = CType(resources.GetObject("btnReject.ImageOptions.SvgImage"), DevExpress.Utils.Svg.SvgImage)
        Me.btnReject.ImageOptions.SvgImageSize = New System.Drawing.Size(24, 24)
        Me.btnReject.Name = "btnReject"
        '
        'btnRefresh
        '
        Me.btnRefresh.Caption = "Refresh (F5)"
        Me.btnRefresh.CaptionAlignment = DevExpress.XtraBars.BarItemCaptionAlignment.Right
        Me.btnRefresh.Id = 10
        Me.btnRefresh.ImageOptions.AllowGlyphSkinning = DevExpress.Utils.DefaultBoolean.[True]
        Me.btnRefresh.ImageOptions.SvgImage = CType(resources.GetObject("btnRefresh.ImageOptions.SvgImage"), DevExpress.Utils.Svg.SvgImage)
        Me.btnRefresh.ImageOptions.SvgImageSize = New System.Drawing.Size(24, 24)
        Me.btnRefresh.ItemShortcut = New DevExpress.XtraBars.BarShortcut(System.Windows.Forms.Keys.F5)
        Me.btnRefresh.Name = "btnRefresh"
        '
        'btnClose
        '
        Me.btnClose.Alignment = DevExpress.XtraBars.BarItemLinkAlignment.Right
        Me.btnClose.Caption = "Close"
        Me.btnClose.CaptionAlignment = DevExpress.XtraBars.BarItemCaptionAlignment.Right
        Me.btnClose.Id = 11
        Me.btnClose.ImageOptions.AllowGlyphSkinning = DevExpress.Utils.DefaultBoolean.[True]
        Me.btnClose.ImageOptions.SvgImage = CType(resources.GetObject("btnClose.ImageOptions.SvgImage"), DevExpress.Utils.Svg.SvgImage)
        Me.btnClose.ImageOptions.SvgImageSize = New System.Drawing.Size(24, 24)
        Me.btnClose.Name = "btnClose"
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
        Me.barDockControlTop.Size = New System.Drawing.Size(1271, 36)
        '
        'barDockControlBottom
        '
        Me.barDockControlBottom.CausesValidation = False
        Me.barDockControlBottom.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.barDockControlBottom.Location = New System.Drawing.Point(0, 610)
        Me.barDockControlBottom.Manager = Me.BarManager
        Me.barDockControlBottom.Size = New System.Drawing.Size(1271, 0)
        '
        'barDockControlLeft
        '
        Me.barDockControlLeft.CausesValidation = False
        Me.barDockControlLeft.Dock = System.Windows.Forms.DockStyle.Left
        Me.barDockControlLeft.Location = New System.Drawing.Point(0, 36)
        Me.barDockControlLeft.Manager = Me.BarManager
        Me.barDockControlLeft.Size = New System.Drawing.Size(0, 574)
        '
        'barDockControlRight
        '
        Me.barDockControlRight.CausesValidation = False
        Me.barDockControlRight.Dock = System.Windows.Forms.DockStyle.Right
        Me.barDockControlRight.Location = New System.Drawing.Point(1271, 36)
        Me.barDockControlRight.Manager = Me.BarManager
        Me.barDockControlRight.Size = New System.Drawing.Size(0, 574)
        '
        'BarButtonItem1
        '
        Me.BarButtonItem1.Caption = "View"
        Me.BarButtonItem1.Id = 0
        Me.BarButtonItem1.ImageOptions.Image = CType(resources.GetObject("BarButtonItem1.ImageOptions.Image"), System.Drawing.Image)
        Me.BarButtonItem1.ImageOptions.LargeImage = CType(resources.GetObject("BarButtonItem1.ImageOptions.LargeImage"), System.Drawing.Image)
        Me.BarButtonItem1.Name = "BarButtonItem1"
        '
        'BarButtonItem2
        '
        Me.BarButtonItem2.Caption = "Test Order"
        Me.BarButtonItem2.Id = 1
        Me.BarButtonItem2.ImageOptions.Image = CType(resources.GetObject("BarButtonItem2.ImageOptions.Image"), System.Drawing.Image)
        Me.BarButtonItem2.ImageOptions.LargeImage = CType(resources.GetObject("BarButtonItem2.ImageOptions.LargeImage"), System.Drawing.Image)
        Me.BarButtonItem2.Name = "BarButtonItem2"
        '
        'BarButtonItem3
        '
        Me.BarButtonItem3.Caption = "Patient Information"
        Me.BarButtonItem3.Id = 2
        Me.BarButtonItem3.Name = "BarButtonItem3"
        '
        'BarButtonItem4
        '
        Me.BarButtonItem4.Caption = "Delete"
        Me.BarButtonItem4.Id = 3
        Me.BarButtonItem4.Name = "BarButtonItem4"
        '
        'BarButtonItem5
        '
        Me.BarButtonItem5.Caption = "Refresh"
        Me.BarButtonItem5.Id = 4
        Me.BarButtonItem5.Name = "BarButtonItem5"
        '
        'BarButtonItem6
        '
        Me.BarButtonItem6.Caption = "Close"
        Me.BarButtonItem6.Id = 5
        Me.BarButtonItem6.ImageOptions.Image = CType(resources.GetObject("BarButtonItem6.ImageOptions.Image"), System.Drawing.Image)
        Me.BarButtonItem6.ImageOptions.LargeImage = CType(resources.GetObject("BarButtonItem6.ImageOptions.LargeImage"), System.Drawing.Image)
        Me.BarButtonItem6.Name = "BarButtonItem6"
        '
        'btnNewOrder
        '
        Me.btnNewOrder.Caption = "&New Order"
        Me.btnNewOrder.CaptionAlignment = DevExpress.XtraBars.BarItemCaptionAlignment.Right
        Me.btnNewOrder.Id = 15
        Me.btnNewOrder.ImageOptions.AllowGlyphSkinning = DevExpress.Utils.DefaultBoolean.[True]
        Me.btnNewOrder.ImageOptions.Image = CType(resources.GetObject("btnNewOrder.ImageOptions.Image"), System.Drawing.Image)
        Me.btnNewOrder.ImageOptions.LargeImage = CType(resources.GetObject("btnNewOrder.ImageOptions.LargeImage"), System.Drawing.Image)
        Me.btnNewOrder.Name = "btnNewOrder"
        '
        'PanelControl2
        '
        Me.PanelControl2.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.PanelControl2.Controls.Add(Me.LabelControl7)
        Me.PanelControl2.Controls.Add(Me.rgSelect)
        Me.PanelControl2.Controls.Add(Me.LabelControl4)
        Me.PanelControl2.Controls.Add(Me.txtSearch)
        Me.PanelControl2.Controls.Add(Me.LabelControl6)
        Me.PanelControl2.Controls.Add(Me.cboLocation)
        Me.PanelControl2.Location = New System.Drawing.Point(6, 6)
        Me.PanelControl2.Name = "PanelControl2"
        Me.PanelControl2.Size = New System.Drawing.Size(1240, 32)
        Me.PanelControl2.TabIndex = 144
        '
        'LabelControl4
        '
        Me.LabelControl4.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.LabelControl4.Appearance.Font = New System.Drawing.Font("Segoe UI", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.LabelControl4.Appearance.Options.UseFont = True
        Me.LabelControl4.Location = New System.Drawing.Point(971, 9)
        Me.LabelControl4.Name = "LabelControl4"
        Me.LabelControl4.Size = New System.Drawing.Size(63, 13)
        Me.LabelControl4.TabIndex = 154
        Me.LabelControl4.Text = "Search here:"
        '
        'txtSearch
        '
        Me.txtSearch.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtSearch.Location = New System.Drawing.Point(1036, 5)
        Me.txtSearch.Name = "txtSearch"
        Me.txtSearch.Properties.Appearance.ForeColor = System.Drawing.Color.FromArgb(CType(CType(27, Byte), Integer), CType(CType(41, Byte), Integer), CType(CType(62, Byte), Integer))
        Me.txtSearch.Properties.Appearance.Options.UseForeColor = True
        Me.txtSearch.Properties.Buttons.AddRange(New DevExpress.XtraEditors.Controls.EditorButton() {New DevExpress.XtraEditors.Repository.ClearButton(), New DevExpress.XtraEditors.Repository.SearchButton()})
        Me.txtSearch.Properties.LookAndFeel.SkinName = "Visual Studio 2013 Blue"
        Me.txtSearch.Properties.LookAndFeel.UseDefaultLookAndFeel = False
        Me.txtSearch.Size = New System.Drawing.Size(199, 20)
        Me.txtSearch.TabIndex = 153
        '
        'LabelControl6
        '
        Me.LabelControl6.Appearance.Font = New System.Drawing.Font("Segoe UI", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.LabelControl6.Appearance.ForeColor = System.Drawing.Color.FromArgb(CType(CType(27, Byte), Integer), CType(CType(41, Byte), Integer), CType(CType(62, Byte), Integer))
        Me.LabelControl6.Appearance.Options.UseFont = True
        Me.LabelControl6.Appearance.Options.UseForeColor = True
        Me.LabelControl6.Location = New System.Drawing.Point(8, 8)
        Me.LabelControl6.Name = "LabelControl6"
        Me.LabelControl6.Size = New System.Drawing.Size(48, 13)
        Me.LabelControl6.TabIndex = 73
        Me.LabelControl6.Text = "Location:"
        '
        'cboLocation
        '
        Me.cboLocation.Location = New System.Drawing.Point(66, 6)
        Me.cboLocation.Name = "cboLocation"
        Me.cboLocation.Properties.Appearance.ForeColor = System.Drawing.Color.FromArgb(CType(CType(27, Byte), Integer), CType(CType(41, Byte), Integer), CType(CType(62, Byte), Integer))
        Me.cboLocation.Properties.Appearance.Options.UseForeColor = True
        Me.cboLocation.Properties.Buttons.AddRange(New DevExpress.XtraEditors.Controls.EditorButton() {New DevExpress.XtraEditors.Controls.EditorButton(DevExpress.XtraEditors.Controls.ButtonPredefines.Combo)})
        Me.cboLocation.Properties.LookAndFeel.SkinName = "Visual Studio 2013 Blue"
        Me.cboLocation.Properties.LookAndFeel.UseDefaultLookAndFeel = False
        Me.cboLocation.Properties.TextEditStyle = DevExpress.XtraEditors.Controls.TextEditStyles.DisableTextEditor
        Me.cboLocation.Properties.UseReadOnlyAppearance = False
        Me.cboLocation.Size = New System.Drawing.Size(206, 20)
        Me.cboLocation.TabIndex = 70
        '
        'bcCode
        '
        Me.bcCode.AutoModule = True
        Me.bcCode.BackColor = System.Drawing.Color.White
        Me.bcCode.BorderStyle = DevExpress.XtraEditors.Controls.BorderStyles.NoBorder
        Me.bcCode.Font = New System.Drawing.Font("Tahoma", 13.8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.bcCode.HorizontalAlignment = DevExpress.Utils.HorzAlignment.Center
        Me.bcCode.HorizontalTextAlignment = DevExpress.Utils.HorzAlignment.Center
        Me.bcCode.Location = New System.Drawing.Point(860, 99)
        Me.bcCode.Margin = New System.Windows.Forms.Padding(2)
        Me.bcCode.Name = "bcCode"
        Me.bcCode.Padding = New System.Windows.Forms.Padding(8, 2, 8, 0)
        Me.bcCode.ShowText = False
        Me.bcCode.Size = New System.Drawing.Size(343, 92)
        Me.bcCode.Symbology = Code128Generator1
        Me.bcCode.TabIndex = 143
        '
        'picCode
        '
        Me.picCode.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.picCode.BackColor = System.Drawing.Color.White
        Me.picCode.Location = New System.Drawing.Point(887, 196)
        Me.picCode.Margin = New System.Windows.Forms.Padding(2, 3, 2, 3)
        Me.picCode.Name = "picCode"
        Me.picCode.Size = New System.Drawing.Size(340, 72)
        Me.picCode.TabIndex = 142
        Me.picCode.TabStop = False
        Me.picCode.Visible = False
        '
        'dtTo
        '
        Me.dtTo.CalendarForeColor = System.Drawing.Color.FromArgb(CType(CType(27, Byte), Integer), CType(CType(41, Byte), Integer), CType(CType(62, Byte), Integer))
        Me.dtTo.Font = New System.Drawing.Font("Tahoma", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.dtTo.Format = System.Windows.Forms.DateTimePickerFormat.[Short]
        Me.dtTo.Location = New System.Drawing.Point(978, 117)
        Me.dtTo.Name = "dtTo"
        Me.dtTo.Size = New System.Drawing.Size(152, 23)
        Me.dtTo.TabIndex = 8
        '
        'LabelControl2
        '
        Me.LabelControl2.Appearance.Font = New System.Drawing.Font("Segoe UI", 9.75!)
        Me.LabelControl2.Appearance.ForeColor = System.Drawing.Color.FromArgb(CType(CType(27, Byte), Integer), CType(CType(41, Byte), Integer), CType(CType(62, Byte), Integer))
        Me.LabelControl2.Appearance.Options.UseFont = True
        Me.LabelControl2.Appearance.Options.UseForeColor = True
        Me.LabelControl2.Location = New System.Drawing.Point(733, 120)
        Me.LabelControl2.Name = "LabelControl2"
        Me.LabelControl2.Size = New System.Drawing.Size(71, 17)
        Me.LabelControl2.TabIndex = 7
        Me.LabelControl2.Text = "Date Range:"
        '
        'dtFrom
        '
        Me.dtFrom.CalendarForeColor = System.Drawing.Color.FromArgb(CType(CType(27, Byte), Integer), CType(CType(41, Byte), Integer), CType(CType(62, Byte), Integer))
        Me.dtFrom.Font = New System.Drawing.Font("Tahoma", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.dtFrom.Format = System.Windows.Forms.DateTimePickerFormat.[Short]
        Me.dtFrom.Location = New System.Drawing.Point(820, 117)
        Me.dtFrom.Name = "dtFrom"
        Me.dtFrom.Size = New System.Drawing.Size(152, 23)
        Me.dtFrom.TabIndex = 9
        '
        'XTPage2
        '
        Me.XTPage2.Controls.Add(Me.GroupControl1)
        Me.XTPage2.Controls.Add(Me.PanelControl1)
        Me.XTPage2.Controls.Add(Me.pnlHeader)
        Me.XTPage2.ImageOptions.SvgImage = CType(resources.GetObject("XTPage2.ImageOptions.SvgImage"), DevExpress.Utils.Svg.SvgImage)
        Me.XTPage2.ImageOptions.SvgImageSize = New System.Drawing.Size(16, 16)
        Me.XTPage2.Name = "XTPage2"
        Me.XTPage2.Size = New System.Drawing.Size(1249, 500)
        Me.XTPage2.Text = "Archive"
        '
        'GroupControl1
        '
        Me.GroupControl1.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.GroupControl1.AppearanceCaption.BackColor = System.Drawing.Color.FromArgb(CType(CType(16, Byte), Integer), CType(CType(110, Byte), Integer), CType(CType(190, Byte), Integer))
        Me.GroupControl1.AppearanceCaption.BackColor2 = System.Drawing.Color.FromArgb(CType(CType(16, Byte), Integer), CType(CType(110, Byte), Integer), CType(CType(190, Byte), Integer))
        Me.GroupControl1.AppearanceCaption.BorderColor = System.Drawing.Color.FromArgb(CType(CType(16, Byte), Integer), CType(CType(110, Byte), Integer), CType(CType(190, Byte), Integer))
        Me.GroupControl1.AppearanceCaption.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.GroupControl1.AppearanceCaption.Options.UseBackColor = True
        Me.GroupControl1.AppearanceCaption.Options.UseBorderColor = True
        Me.GroupControl1.AppearanceCaption.Options.UseFont = True
        Me.GroupControl1.Controls.Add(Me.dtCompleted)
        Me.GroupControl1.Location = New System.Drawing.Point(6, 79)
        Me.GroupControl1.Name = "GroupControl1"
        Me.GroupControl1.Size = New System.Drawing.Size(1236, 413)
        Me.GroupControl1.TabIndex = 175
        Me.GroupControl1.Text = "List"
        '
        'dtCompleted
        '
        Me.dtCompleted.Dock = System.Windows.Forms.DockStyle.Fill
        Me.dtCompleted.Location = New System.Drawing.Point(2, 27)
        Me.dtCompleted.MainView = Me.GridCompleted
        Me.dtCompleted.MenuManager = Me.BarManager
        Me.dtCompleted.Name = "dtCompleted"
        Me.dtCompleted.Size = New System.Drawing.Size(1232, 384)
        Me.dtCompleted.TabIndex = 148
        Me.dtCompleted.ViewCollection.AddRange(New DevExpress.XtraGrid.Views.Base.BaseView() {Me.GridCompleted})
        '
        'GridCompleted
        '
        Me.GridCompleted.Appearance.FocusedRow.Font = New System.Drawing.Font("Segoe UI", 8.25!, System.Drawing.FontStyle.Bold)
        Me.GridCompleted.Appearance.FocusedRow.Options.UseFont = True
        Me.GridCompleted.GridControl = Me.dtCompleted
        Me.GridCompleted.Name = "GridCompleted"
        Me.GridCompleted.OptionsBehavior.AutoExpandAllGroups = True
        Me.GridCompleted.OptionsCustomization.AllowGroup = False
        Me.GridCompleted.OptionsCustomization.AllowQuickHideColumns = False
        Me.GridCompleted.OptionsSelection.CheckBoxSelectorColumnWidth = 30
        Me.GridCompleted.OptionsView.ShowGroupPanel = False
        Me.GridCompleted.OptionsView.ShowVerticalLines = DevExpress.Utils.DefaultBoolean.[False]
        '
        'PanelControl1
        '
        Me.PanelControl1.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.PanelControl1.Controls.Add(Me.LabelControl8)
        Me.PanelControl1.Controls.Add(Me.rgSelect1)
        Me.PanelControl1.Controls.Add(Me.LabelControl5)
        Me.PanelControl1.Controls.Add(Me.SearchControl1)
        Me.PanelControl1.Controls.Add(Me.LabelControl1)
        Me.PanelControl1.Controls.Add(Me.cboLocation1)
        Me.PanelControl1.Location = New System.Drawing.Point(6, 43)
        Me.PanelControl1.Name = "PanelControl1"
        Me.PanelControl1.Size = New System.Drawing.Size(1236, 30)
        Me.PanelControl1.TabIndex = 15
        '
        'LabelControl5
        '
        Me.LabelControl5.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.LabelControl5.Appearance.Font = New System.Drawing.Font("Segoe UI", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.LabelControl5.Appearance.Options.UseFont = True
        Me.LabelControl5.Location = New System.Drawing.Point(966, 9)
        Me.LabelControl5.Name = "LabelControl5"
        Me.LabelControl5.Size = New System.Drawing.Size(63, 13)
        Me.LabelControl5.TabIndex = 154
        Me.LabelControl5.Text = "Search here:"
        '
        'SearchControl1
        '
        Me.SearchControl1.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.SearchControl1.Location = New System.Drawing.Point(1032, 5)
        Me.SearchControl1.Name = "SearchControl1"
        Me.SearchControl1.Properties.Appearance.ForeColor = System.Drawing.Color.FromArgb(CType(CType(27, Byte), Integer), CType(CType(41, Byte), Integer), CType(CType(62, Byte), Integer))
        Me.SearchControl1.Properties.Appearance.Options.UseForeColor = True
        Me.SearchControl1.Properties.Buttons.AddRange(New DevExpress.XtraEditors.Controls.EditorButton() {New DevExpress.XtraEditors.Repository.ClearButton(), New DevExpress.XtraEditors.Repository.SearchButton()})
        Me.SearchControl1.Properties.LookAndFeel.SkinName = "Visual Studio 2013 Blue"
        Me.SearchControl1.Properties.LookAndFeel.UseDefaultLookAndFeel = False
        Me.SearchControl1.Size = New System.Drawing.Size(199, 20)
        Me.SearchControl1.TabIndex = 153
        '
        'LabelControl1
        '
        Me.LabelControl1.Appearance.Font = New System.Drawing.Font("Segoe UI", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.LabelControl1.Appearance.ForeColor = System.Drawing.Color.FromArgb(CType(CType(27, Byte), Integer), CType(CType(41, Byte), Integer), CType(CType(62, Byte), Integer))
        Me.LabelControl1.Appearance.Options.UseFont = True
        Me.LabelControl1.Appearance.Options.UseForeColor = True
        Me.LabelControl1.Location = New System.Drawing.Point(8, 8)
        Me.LabelControl1.Name = "LabelControl1"
        Me.LabelControl1.Size = New System.Drawing.Size(48, 13)
        Me.LabelControl1.TabIndex = 73
        Me.LabelControl1.Text = "Location:"
        '
        'cboLocation1
        '
        Me.cboLocation1.Location = New System.Drawing.Point(74, 5)
        Me.cboLocation1.Name = "cboLocation1"
        Me.cboLocation1.Properties.Appearance.ForeColor = System.Drawing.Color.FromArgb(CType(CType(27, Byte), Integer), CType(CType(41, Byte), Integer), CType(CType(62, Byte), Integer))
        Me.cboLocation1.Properties.Appearance.Options.UseForeColor = True
        Me.cboLocation1.Properties.Buttons.AddRange(New DevExpress.XtraEditors.Controls.EditorButton() {New DevExpress.XtraEditors.Controls.EditorButton(DevExpress.XtraEditors.Controls.ButtonPredefines.Combo)})
        Me.cboLocation1.Properties.LookAndFeel.SkinName = "Visual Studio 2013 Blue"
        Me.cboLocation1.Properties.LookAndFeel.UseDefaultLookAndFeel = False
        Me.cboLocation1.Properties.TextEditStyle = DevExpress.XtraEditors.Controls.TextEditStyles.DisableTextEditor
        Me.cboLocation1.Properties.UseReadOnlyAppearance = False
        Me.cboLocation1.Size = New System.Drawing.Size(310, 20)
        Me.cboLocation1.TabIndex = 72
        '
        'pnlHeader
        '
        Me.pnlHeader.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.pnlHeader.Controls.Add(Me.btnSearch)
        Me.pnlHeader.Controls.Add(Me.dtFrom1)
        Me.pnlHeader.Controls.Add(Me.dtTo1)
        Me.pnlHeader.Controls.Add(Me.LabelControl3)
        Me.pnlHeader.Location = New System.Drawing.Point(6, 7)
        Me.pnlHeader.Name = "pnlHeader"
        Me.pnlHeader.Size = New System.Drawing.Size(1236, 30)
        Me.pnlHeader.TabIndex = 10
        '
        'btnSearch
        '
        Me.btnSearch.ImageOptions.SvgImage = CType(resources.GetObject("btnSearch.ImageOptions.SvgImage"), DevExpress.Utils.Svg.SvgImage)
        Me.btnSearch.ImageOptions.SvgImageSize = New System.Drawing.Size(24, 24)
        Me.btnSearch.Location = New System.Drawing.Point(316, 3)
        Me.btnSearch.Name = "btnSearch"
        Me.btnSearch.Size = New System.Drawing.Size(68, 23)
        Me.btnSearch.TabIndex = 12
        Me.btnSearch.Text = "&Apply"
        '
        'dtFrom1
        '
        Me.dtFrom1.CalendarForeColor = System.Drawing.Color.FromArgb(CType(CType(27, Byte), Integer), CType(CType(41, Byte), Integer), CType(CType(62, Byte), Integer))
        Me.dtFrom1.Format = System.Windows.Forms.DateTimePickerFormat.[Short]
        Me.dtFrom1.Location = New System.Drawing.Point(74, 4)
        Me.dtFrom1.Name = "dtFrom1"
        Me.dtFrom1.Size = New System.Drawing.Size(114, 22)
        Me.dtFrom1.TabIndex = 9
        '
        'dtTo1
        '
        Me.dtTo1.CalendarFont = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.dtTo1.CalendarForeColor = System.Drawing.Color.FromArgb(CType(CType(27, Byte), Integer), CType(CType(41, Byte), Integer), CType(CType(62, Byte), Integer))
        Me.dtTo1.Format = System.Windows.Forms.DateTimePickerFormat.[Short]
        Me.dtTo1.Location = New System.Drawing.Point(194, 4)
        Me.dtTo1.Name = "dtTo1"
        Me.dtTo1.Size = New System.Drawing.Size(115, 22)
        Me.dtTo1.TabIndex = 8
        '
        'LabelControl3
        '
        Me.LabelControl3.Appearance.Font = New System.Drawing.Font("Segoe UI", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.LabelControl3.Appearance.ForeColor = System.Drawing.Color.FromArgb(CType(CType(27, Byte), Integer), CType(CType(41, Byte), Integer), CType(CType(62, Byte), Integer))
        Me.LabelControl3.Appearance.Options.UseFont = True
        Me.LabelControl3.Appearance.Options.UseForeColor = True
        Me.LabelControl3.Location = New System.Drawing.Point(8, 7)
        Me.LabelControl3.Name = "LabelControl3"
        Me.LabelControl3.Size = New System.Drawing.Size(63, 13)
        Me.LabelControl3.TabIndex = 7
        Me.LabelControl3.Text = "Date Range:"
        '
        'ImageList
        '
        Me.ImageList.ImageStream = CType(resources.GetObject("ImageList.ImageStream"), System.Windows.Forms.ImageListStreamer)
        Me.ImageList.TransparentColor = System.Drawing.Color.Transparent
        Me.ImageList.Images.SetKeyName(0, "Folder.ico")
        '
        'tmLoadNew
        '
        Me.tmLoadNew.Enabled = True
        '
        'tmAutoValidate
        '
        Me.tmAutoValidate.Enabled = True
        Me.tmAutoValidate.Interval = 1000
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
        'LabelControl7
        '
        Me.LabelControl7.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.LabelControl7.Appearance.Font = New System.Drawing.Font("Segoe UI", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.LabelControl7.Appearance.Options.UseFont = True
        Me.LabelControl7.Location = New System.Drawing.Point(606, 10)
        Me.LabelControl7.Name = "LabelControl7"
        Me.LabelControl7.Size = New System.Drawing.Size(75, 13)
        Me.LabelControl7.TabIndex = 156
        Me.LabelControl7.Text = "Search option:"
        '
        'rgSelect
        '
        Me.rgSelect.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.rgSelect.Location = New System.Drawing.Point(685, 5)
        Me.rgSelect.MenuManager = Me.BarManager
        Me.rgSelect.Name = "rgSelect"
        Me.rgSelect.Properties.Appearance.BackColor = System.Drawing.SystemColors.ButtonFace
        Me.rgSelect.Properties.Appearance.Options.UseBackColor = True
        Me.rgSelect.Properties.BorderStyle = DevExpress.XtraEditors.Controls.BorderStyles.NoBorder
        Me.rgSelect.Properties.Items.AddRange(New DevExpress.XtraEditors.Controls.RadioGroupItem() {New DevExpress.XtraEditors.Controls.RadioGroupItem(Nothing, "Sample ID"), New DevExpress.XtraEditors.Controls.RadioGroupItem(Nothing, "Hospital No."), New DevExpress.XtraEditors.Controls.RadioGroupItem(Nothing, "Patient Name")})
        Me.rgSelect.Size = New System.Drawing.Size(280, 22)
        Me.rgSelect.TabIndex = 155
        '
        'LabelControl8
        '
        Me.LabelControl8.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.LabelControl8.Appearance.Font = New System.Drawing.Font("Segoe UI", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.LabelControl8.Appearance.Options.UseFont = True
        Me.LabelControl8.Location = New System.Drawing.Point(602, 8)
        Me.LabelControl8.Name = "LabelControl8"
        Me.LabelControl8.Size = New System.Drawing.Size(75, 13)
        Me.LabelControl8.TabIndex = 158
        Me.LabelControl8.Text = "Search option:"
        '
        'rgSelect1
        '
        Me.rgSelect1.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.rgSelect1.Location = New System.Drawing.Point(680, 2)
        Me.rgSelect1.MenuManager = Me.BarManager
        Me.rgSelect1.Name = "rgSelect1"
        Me.rgSelect1.Properties.Appearance.BackColor = System.Drawing.SystemColors.ButtonFace
        Me.rgSelect1.Properties.Appearance.Options.UseBackColor = True
        Me.rgSelect1.Properties.BorderStyle = DevExpress.XtraEditors.Controls.BorderStyles.NoBorder
        Me.rgSelect1.Properties.Items.AddRange(New DevExpress.XtraEditors.Controls.RadioGroupItem() {New DevExpress.XtraEditors.Controls.RadioGroupItem(Nothing, "Sample ID", True, 0), New DevExpress.XtraEditors.Controls.RadioGroupItem(Nothing, "Hospital No.", True, 1), New DevExpress.XtraEditors.Controls.RadioGroupItem(Nothing, "Patient Name", True, 2)})
        Me.rgSelect1.Size = New System.Drawing.Size(280, 25)
        Me.rgSelect1.TabIndex = 157
        '
        'frmMicroWorklist
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(1271, 610)
        Me.Controls.Add(Me.pnlBackground)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Name = "frmMicroWorklist"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Microbiology Worklist Window"
        Me.WindowState = System.Windows.Forms.FormWindowState.Maximized
        CType(Me.pnlBackground, System.ComponentModel.ISupportInitialize).EndInit()
        Me.pnlBackground.ResumeLayout(False)
        Me.pnlBackground.PerformLayout()
        CType(Me.PanelControl4, System.ComponentModel.ISupportInitialize).EndInit()
        Me.PanelControl4.ResumeLayout(False)
        Me.PanelControl4.PerformLayout()
        CType(Me.SeparatorControl3, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.txtSearch1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.txtSearch1.ResumeLayout(False)
        Me.XTPage1.ResumeLayout(False)
        Me.XTPage1.PerformLayout()
        CType(Me.GroupControl2, System.ComponentModel.ISupportInitialize).EndInit()
        Me.GroupControl2.ResumeLayout(False)
        CType(Me.dtList, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.GridView, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.BarManager, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.BarAndDockingController, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.PanelControl2, System.ComponentModel.ISupportInitialize).EndInit()
        Me.PanelControl2.ResumeLayout(False)
        Me.PanelControl2.PerformLayout()
        CType(Me.txtSearch.Properties, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.cboLocation.Properties, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.picCode, System.ComponentModel.ISupportInitialize).EndInit()
        Me.XTPage2.ResumeLayout(False)
        CType(Me.GroupControl1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.GroupControl1.ResumeLayout(False)
        CType(Me.dtCompleted, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.GridCompleted, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.PanelControl1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.PanelControl1.ResumeLayout(False)
        Me.PanelControl1.PerformLayout()
        CType(Me.SearchControl1.Properties, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.cboLocation1.Properties, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.pnlHeader, System.ComponentModel.ISupportInitialize).EndInit()
        Me.pnlHeader.ResumeLayout(False)
        Me.pnlHeader.PerformLayout()
        CType(Me.rgSelect.Properties, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.rgSelect1.Properties, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents pnlBackground As DevExpress.XtraEditors.PanelControl
    Friend WithEvents BarManager As DevExpress.XtraBars.BarManager
    Friend WithEvents Bar As DevExpress.XtraBars.Bar
    Friend WithEvents BarButtonItem1 As DevExpress.XtraBars.BarButtonItem
    Friend WithEvents BarButtonItem2 As DevExpress.XtraBars.BarButtonItem
    Friend WithEvents BarButtonItem3 As DevExpress.XtraBars.BarButtonItem
    Friend WithEvents BarButtonItem4 As DevExpress.XtraBars.BarButtonItem
    Friend WithEvents BarButtonItem5 As DevExpress.XtraBars.BarButtonItem
    Friend WithEvents BarButtonItem6 As DevExpress.XtraBars.BarButtonItem
    Friend WithEvents barDockControlTop As DevExpress.XtraBars.BarDockControl
    Friend WithEvents barDockControlBottom As DevExpress.XtraBars.BarDockControl
    Friend WithEvents barDockControlLeft As DevExpress.XtraBars.BarDockControl
    Friend WithEvents barDockControlRight As DevExpress.XtraBars.BarDockControl
    Friend WithEvents btnView As DevExpress.XtraBars.BarLargeButtonItem
    Friend WithEvents btnDelete As DevExpress.XtraBars.BarLargeButtonItem
    Friend WithEvents btnRefresh As DevExpress.XtraBars.BarLargeButtonItem
    Friend WithEvents btnClose As DevExpress.XtraBars.BarLargeButtonItem
    Friend WithEvents BarAndDockingController As DevExpress.XtraBars.BarAndDockingController
    Friend WithEvents ImageList As System.Windows.Forms.ImageList
    Friend WithEvents dtFrom As System.Windows.Forms.DateTimePicker
    Friend WithEvents dtTo As System.Windows.Forms.DateTimePicker
    Friend WithEvents LabelControl2 As DevExpress.XtraEditors.LabelControl
    Friend WithEvents txtSearch1 As DevExpress.XtraTab.XtraTabControl
    Friend WithEvents XTPage1 As DevExpress.XtraTab.XtraTabPage
    Friend WithEvents XTPage2 As DevExpress.XtraTab.XtraTabPage
    Friend WithEvents btnNewOrder As DevExpress.XtraBars.BarLargeButtonItem
    Friend WithEvents tmLoadNew As System.Windows.Forms.Timer
    Friend WithEvents btnPrint As DevExpress.XtraBars.BarLargeButtonItem
    Friend WithEvents btnReject As DevExpress.XtraBars.BarLargeButtonItem
    Friend WithEvents cboLocation As DevExpress.XtraEditors.ComboBoxEdit
    Friend WithEvents tmAutoValidate As System.Windows.Forms.Timer
    Friend WithEvents btnBarcode As DevExpress.XtraBars.BarLargeButtonItem
    Friend WithEvents bcCode As DevExpress.XtraEditors.BarCodeControl
    Friend WithEvents picCode As System.Windows.Forms.PictureBox
    Friend WithEvents PrintDocument As System.Drawing.Printing.PrintDocument
    Friend WithEvents PrintPreviewDialog As System.Windows.Forms.PrintPreviewDialog
    Friend WithEvents PanelControl4 As DevExpress.XtraEditors.PanelControl
    Friend WithEvents lblCountQueue As DevExpress.XtraEditors.LabelControl
    Friend WithEvents SeparatorControl3 As DevExpress.XtraEditors.SeparatorControl
    Friend WithEvents GroupControl1 As DevExpress.XtraEditors.GroupControl
    Friend WithEvents PanelControl1 As DevExpress.XtraEditors.PanelControl
    Friend WithEvents LabelControl1 As DevExpress.XtraEditors.LabelControl
    Friend WithEvents cboLocation1 As DevExpress.XtraEditors.ComboBoxEdit
    Friend WithEvents pnlHeader As DevExpress.XtraEditors.PanelControl
    Friend WithEvents btnSearch As DevExpress.XtraEditors.SimpleButton
    Friend WithEvents dtFrom1 As DateTimePicker
    Friend WithEvents dtTo1 As DateTimePicker
    Friend WithEvents LabelControl3 As DevExpress.XtraEditors.LabelControl
    Friend WithEvents GroupControl2 As DevExpress.XtraEditors.GroupControl
    Friend WithEvents PanelControl2 As DevExpress.XtraEditors.PanelControl
    Friend WithEvents LabelControl6 As DevExpress.XtraEditors.LabelControl
    Friend WithEvents dtList As DevExpress.XtraGrid.GridControl
    Friend WithEvents GridView As DevExpress.XtraGrid.Views.Grid.GridView
    Friend WithEvents dtCompleted As DevExpress.XtraGrid.GridControl
    Friend WithEvents GridCompleted As DevExpress.XtraGrid.Views.Grid.GridView
    Friend WithEvents LabelControl4 As DevExpress.XtraEditors.LabelControl
    Friend WithEvents txtSearch As DevExpress.XtraEditors.SearchControl
    Friend WithEvents LabelControl5 As DevExpress.XtraEditors.LabelControl
    Friend WithEvents SearchControl1 As DevExpress.XtraEditors.SearchControl
    Friend WithEvents LabelControl7 As DevExpress.XtraEditors.LabelControl
    Friend WithEvents rgSelect As DevExpress.XtraEditors.RadioGroup
    Friend WithEvents LabelControl8 As DevExpress.XtraEditors.LabelControl
    Friend WithEvents rgSelect1 As DevExpress.XtraEditors.RadioGroup
End Class
