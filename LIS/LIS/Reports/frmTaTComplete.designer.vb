<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class frmTatComplete
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmTatComplete))
        Me.ImageList = New System.Windows.Forms.ImageList(Me.components)
        Me.BarManager = New DevExpress.XtraBars.BarManager(Me.components)
        Me.Bar2 = New DevExpress.XtraBars.Bar()
        Me.btnPrint = New DevExpress.XtraBars.BarLargeButtonItem()
        Me.btnRefresh = New DevExpress.XtraBars.BarLargeButtonItem()
        Me.btnClose = New DevExpress.XtraBars.BarLargeButtonItem()
        Me.BarAndDockingController = New DevExpress.XtraBars.BarAndDockingController(Me.components)
        Me.barDockControlTop = New DevExpress.XtraBars.BarDockControl()
        Me.barDockControlBottom = New DevExpress.XtraBars.BarDockControl()
        Me.barDockControlLeft = New DevExpress.XtraBars.BarDockControl()
        Me.barDockControlRight = New DevExpress.XtraBars.BarDockControl()
        Me.XtraTabControl1 = New DevExpress.XtraTab.XtraTabControl()
        Me.XtraTabPage1 = New DevExpress.XtraTab.XtraTabPage()
        Me.dtResult = New DevExpress.XtraGrid.GridControl()
        Me.GridView = New DevExpress.XtraGrid.Views.Grid.GridView()
        Me.lblDiffTime = New DevExpress.XtraEditors.LabelControl()
        Me.XtraTabPage3 = New DevExpress.XtraTab.XtraTabPage()
        Me.dtTaTQuant = New DevExpress.XtraGrid.GridControl()
        Me.GridTaTQuant = New DevExpress.XtraGrid.Views.Grid.GridView()
        Me.lblTotalCount = New DevExpress.XtraEditors.LabelControl()
        Me.LabelControl1 = New DevExpress.XtraEditors.LabelControl()
        Me.btnSearch = New DevExpress.XtraEditors.SimpleButton()
        Me.cboLimit = New DevExpress.XtraEditors.ComboBoxEdit()
        Me.dtFrom = New System.Windows.Forms.DateTimePicker()
        Me.dtTo = New System.Windows.Forms.DateTimePicker()
        Me.LabelControl4 = New DevExpress.XtraEditors.LabelControl()
        Me.LabelControl7 = New DevExpress.XtraEditors.LabelControl()
        Me.cboMedtech = New DevExpress.XtraEditors.ComboBoxEdit()
        Me.pnlHeader = New DevExpress.XtraEditors.PanelControl()
        CType(Me.BarManager, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.BarAndDockingController, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.XtraTabControl1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.XtraTabControl1.SuspendLayout()
        Me.XtraTabPage1.SuspendLayout()
        CType(Me.dtResult, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.GridView, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.XtraTabPage3.SuspendLayout()
        CType(Me.dtTaTQuant, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.GridTaTQuant, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.cboLimit.Properties, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.cboMedtech.Properties, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.pnlHeader, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.pnlHeader.SuspendLayout()
        Me.SuspendLayout()
        '
        'ImageList
        '
        Me.ImageList.ImageStream = CType(resources.GetObject("ImageList.ImageStream"), System.Windows.Forms.ImageListStreamer)
        Me.ImageList.TransparentColor = System.Drawing.Color.Transparent
        Me.ImageList.Images.SetKeyName(0, "Folder.ico")
        '
        'BarManager
        '
        Me.BarManager.AllowCustomization = False
        Me.BarManager.AllowQuickCustomization = False
        Me.BarManager.Bars.AddRange(New DevExpress.XtraBars.Bar() {Me.Bar2})
        Me.BarManager.Controller = Me.BarAndDockingController
        Me.BarManager.DockControls.Add(Me.barDockControlTop)
        Me.BarManager.DockControls.Add(Me.barDockControlBottom)
        Me.BarManager.DockControls.Add(Me.barDockControlLeft)
        Me.BarManager.DockControls.Add(Me.barDockControlRight)
        Me.BarManager.Form = Me
        Me.BarManager.Items.AddRange(New DevExpress.XtraBars.BarItem() {Me.btnPrint, Me.btnRefresh, Me.btnClose})
        Me.BarManager.MainMenu = Me.Bar2
        Me.BarManager.MaxItemId = 3
        '
        'Bar2
        '
        Me.Bar2.BarName = "Main menu"
        Me.Bar2.DockCol = 0
        Me.Bar2.DockRow = 0
        Me.Bar2.DockStyle = DevExpress.XtraBars.BarDockStyle.Top
        Me.Bar2.LinksPersistInfo.AddRange(New DevExpress.XtraBars.LinkPersistInfo() {New DevExpress.XtraBars.LinkPersistInfo(Me.btnPrint), New DevExpress.XtraBars.LinkPersistInfo(Me.btnRefresh, True), New DevExpress.XtraBars.LinkPersistInfo(Me.btnClose)})
        Me.Bar2.OptionsBar.MultiLine = True
        Me.Bar2.OptionsBar.UseWholeRow = True
        Me.Bar2.Text = "Main menu"
        '
        'btnPrint
        '
        Me.btnPrint.Caption = "&Print"
        Me.btnPrint.CaptionAlignment = DevExpress.XtraBars.BarItemCaptionAlignment.Right
        Me.btnPrint.Id = 0
        Me.btnPrint.ImageOptions.AllowGlyphSkinning = DevExpress.Utils.DefaultBoolean.[True]
        Me.btnPrint.ImageOptions.SvgImage = CType(resources.GetObject("btnPrint.ImageOptions.SvgImage"), DevExpress.Utils.Svg.SvgImage)
        Me.btnPrint.ImageOptions.SvgImageSize = New System.Drawing.Size(24, 24)
        Me.btnPrint.Name = "btnPrint"
        '
        'btnRefresh
        '
        Me.btnRefresh.Caption = "&Refresh"
        Me.btnRefresh.CaptionAlignment = DevExpress.XtraBars.BarItemCaptionAlignment.Right
        Me.btnRefresh.Id = 1
        Me.btnRefresh.ImageOptions.AllowGlyphSkinning = DevExpress.Utils.DefaultBoolean.[True]
        Me.btnRefresh.ImageOptions.SvgImage = CType(resources.GetObject("btnRefresh.ImageOptions.SvgImage"), DevExpress.Utils.Svg.SvgImage)
        Me.btnRefresh.ImageOptions.SvgImageSize = New System.Drawing.Size(24, 24)
        Me.btnRefresh.Name = "btnRefresh"
        '
        'btnClose
        '
        Me.btnClose.Alignment = DevExpress.XtraBars.BarItemLinkAlignment.Right
        Me.btnClose.Caption = "&Close"
        Me.btnClose.CaptionAlignment = DevExpress.XtraBars.BarItemCaptionAlignment.Right
        Me.btnClose.Id = 2
        Me.btnClose.ImageOptions.AllowGlyphSkinning = DevExpress.Utils.DefaultBoolean.[True]
        Me.btnClose.ImageOptions.SvgImage = CType(resources.GetObject("btnClose.ImageOptions.SvgImage"), DevExpress.Utils.Svg.SvgImage)
        Me.btnClose.ImageOptions.SvgImageSize = New System.Drawing.Size(24, 24)
        Me.btnClose.Name = "btnClose"
        '
        'BarAndDockingController
        '
        Me.BarAndDockingController.AppearancesBar.MainMenuAppearance.Normal.Font = New System.Drawing.Font("Segoe UI", 9.75!)
        Me.BarAndDockingController.AppearancesBar.MainMenuAppearance.Normal.ForeColor = System.Drawing.Color.White
        Me.BarAndDockingController.AppearancesBar.MainMenuAppearance.Normal.Options.UseFont = True
        Me.BarAndDockingController.AppearancesBar.MainMenuAppearance.Normal.Options.UseForeColor = True
        Me.BarAndDockingController.LookAndFeel.SkinName = "VS2010"
        Me.BarAndDockingController.LookAndFeel.Style = DevExpress.LookAndFeel.LookAndFeelStyle.Office2003
        Me.BarAndDockingController.PropertiesBar.AllowLinkLighting = False
        '
        'barDockControlTop
        '
        Me.barDockControlTop.CausesValidation = False
        Me.barDockControlTop.Dock = System.Windows.Forms.DockStyle.Top
        Me.barDockControlTop.Location = New System.Drawing.Point(0, 0)
        Me.barDockControlTop.Manager = Me.BarManager
        Me.barDockControlTop.Size = New System.Drawing.Size(1244, 36)
        '
        'barDockControlBottom
        '
        Me.barDockControlBottom.CausesValidation = False
        Me.barDockControlBottom.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.barDockControlBottom.Location = New System.Drawing.Point(0, 647)
        Me.barDockControlBottom.Manager = Me.BarManager
        Me.barDockControlBottom.Size = New System.Drawing.Size(1244, 0)
        '
        'barDockControlLeft
        '
        Me.barDockControlLeft.CausesValidation = False
        Me.barDockControlLeft.Dock = System.Windows.Forms.DockStyle.Left
        Me.barDockControlLeft.Location = New System.Drawing.Point(0, 36)
        Me.barDockControlLeft.Manager = Me.BarManager
        Me.barDockControlLeft.Size = New System.Drawing.Size(0, 611)
        '
        'barDockControlRight
        '
        Me.barDockControlRight.CausesValidation = False
        Me.barDockControlRight.Dock = System.Windows.Forms.DockStyle.Right
        Me.barDockControlRight.Location = New System.Drawing.Point(1244, 36)
        Me.barDockControlRight.Manager = Me.BarManager
        Me.barDockControlRight.Size = New System.Drawing.Size(0, 611)
        '
        'XtraTabControl1
        '
        Me.XtraTabControl1.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.XtraTabControl1.Location = New System.Drawing.Point(8, 77)
        Me.XtraTabControl1.Name = "XtraTabControl1"
        Me.XtraTabControl1.SelectedTabPage = Me.XtraTabPage1
        Me.XtraTabControl1.Size = New System.Drawing.Size(1224, 559)
        Me.XtraTabControl1.TabIndex = 15
        Me.XtraTabControl1.TabPages.AddRange(New DevExpress.XtraTab.XtraTabPage() {Me.XtraTabPage1, Me.XtraTabPage3})
        '
        'XtraTabPage1
        '
        Me.XtraTabPage1.Controls.Add(Me.dtResult)
        Me.XtraTabPage1.Controls.Add(Me.lblDiffTime)
        Me.XtraTabPage1.ImageOptions.Image = CType(resources.GetObject("XtraTabPage1.ImageOptions.Image"), System.Drawing.Image)
        Me.XtraTabPage1.Name = "XtraTabPage1"
        Me.XtraTabPage1.Size = New System.Drawing.Size(1222, 527)
        Me.XtraTabPage1.Text = "TaT Time Difference"
        '
        'dtResult
        '
        Me.dtResult.Dock = System.Windows.Forms.DockStyle.Fill
        Me.dtResult.Location = New System.Drawing.Point(0, 0)
        Me.dtResult.MainView = Me.GridView
        Me.dtResult.MenuManager = Me.BarManager
        Me.dtResult.Name = "dtResult"
        Me.dtResult.Size = New System.Drawing.Size(1222, 510)
        Me.dtResult.TabIndex = 0
        Me.dtResult.ViewCollection.AddRange(New DevExpress.XtraGrid.Views.Base.BaseView() {Me.GridView})
        '
        'GridView
        '
        Me.GridView.GridControl = Me.dtResult
        Me.GridView.Name = "GridView"
        Me.GridView.OptionsView.ShowGroupPanel = False
        '
        'lblDiffTime
        '
        Me.lblDiffTime.Appearance.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblDiffTime.Appearance.Options.UseFont = True
        Me.lblDiffTime.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.lblDiffTime.Location = New System.Drawing.Point(0, 510)
        Me.lblDiffTime.Name = "lblDiffTime"
        Me.lblDiffTime.Size = New System.Drawing.Size(86, 17)
        Me.lblDiffTime.TabIndex = 3
        Me.lblDiffTime.Text = "LabelControl2"
        '
        'XtraTabPage3
        '
        Me.XtraTabPage3.Controls.Add(Me.dtTaTQuant)
        Me.XtraTabPage3.Controls.Add(Me.lblTotalCount)
        Me.XtraTabPage3.Name = "XtraTabPage3"
        Me.XtraTabPage3.Size = New System.Drawing.Size(1222, 527)
        Me.XtraTabPage3.Text = "TaT Quantitative"
        '
        'dtTaTQuant
        '
        Me.dtTaTQuant.Dock = System.Windows.Forms.DockStyle.Fill
        Me.dtTaTQuant.Location = New System.Drawing.Point(0, 0)
        Me.dtTaTQuant.MainView = Me.GridTaTQuant
        Me.dtTaTQuant.MenuManager = Me.BarManager
        Me.dtTaTQuant.Name = "dtTaTQuant"
        Me.dtTaTQuant.Size = New System.Drawing.Size(1222, 510)
        Me.dtTaTQuant.TabIndex = 1
        Me.dtTaTQuant.ViewCollection.AddRange(New DevExpress.XtraGrid.Views.Base.BaseView() {Me.GridTaTQuant})
        '
        'GridTaTQuant
        '
        Me.GridTaTQuant.GridControl = Me.dtTaTQuant
        Me.GridTaTQuant.Name = "GridTaTQuant"
        Me.GridTaTQuant.OptionsView.ShowGroupPanel = False
        '
        'lblTotalCount
        '
        Me.lblTotalCount.Appearance.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblTotalCount.Appearance.Options.UseFont = True
        Me.lblTotalCount.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.lblTotalCount.Location = New System.Drawing.Point(0, 510)
        Me.lblTotalCount.Name = "lblTotalCount"
        Me.lblTotalCount.Size = New System.Drawing.Size(86, 17)
        Me.lblTotalCount.TabIndex = 2
        Me.lblTotalCount.Text = "LabelControl2"
        '
        'LabelControl1
        '
        Me.LabelControl1.Appearance.ForeColor = System.Drawing.Color.FromArgb(CType(CType(27, Byte), Integer), CType(CType(41, Byte), Integer), CType(CType(62, Byte), Integer))
        Me.LabelControl1.Appearance.Options.UseForeColor = True
        Me.LabelControl1.Location = New System.Drawing.Point(302, 9)
        Me.LabelControl1.Name = "LabelControl1"
        Me.LabelControl1.Size = New System.Drawing.Size(41, 13)
        Me.LabelControl1.TabIndex = 10
        Me.LabelControl1.Text = "Section:"
        '
        'btnSearch
        '
        Me.btnSearch.ImageOptions.SvgImage = CType(resources.GetObject("btnSearch.ImageOptions.SvgImage"), DevExpress.Utils.Svg.SvgImage)
        Me.btnSearch.ImageOptions.SvgImageSize = New System.Drawing.Size(16, 16)
        Me.btnSearch.Location = New System.Drawing.Point(768, 4)
        Me.btnSearch.Name = "btnSearch"
        Me.btnSearch.Size = New System.Drawing.Size(76, 24)
        Me.btnSearch.TabIndex = 11
        Me.btnSearch.Text = "&Apply"
        '
        'cboLimit
        '
        Me.cboLimit.Location = New System.Drawing.Point(350, 6)
        Me.cboLimit.Name = "cboLimit"
        Me.cboLimit.Properties.Appearance.ForeColor = System.Drawing.Color.FromArgb(CType(CType(27, Byte), Integer), CType(CType(41, Byte), Integer), CType(CType(62, Byte), Integer))
        Me.cboLimit.Properties.Appearance.Options.UseForeColor = True
        Me.cboLimit.Properties.Buttons.AddRange(New DevExpress.XtraEditors.Controls.EditorButton() {New DevExpress.XtraEditors.Controls.EditorButton(DevExpress.XtraEditors.Controls.ButtonPredefines.Combo)})
        Me.cboLimit.Properties.Items.AddRange(New Object() {"Chemistry", "Hematology"})
        Me.cboLimit.Properties.LookAndFeel.SkinName = "Visual Studio 2013 Blue"
        Me.cboLimit.Properties.LookAndFeel.UseDefaultLookAndFeel = False
        Me.cboLimit.Properties.TextEditStyle = DevExpress.XtraEditors.Controls.TextEditStyles.DisableTextEditor
        Me.cboLimit.Properties.UseReadOnlyAppearance = False
        Me.cboLimit.Size = New System.Drawing.Size(160, 20)
        Me.cboLimit.TabIndex = 9
        '
        'dtFrom
        '
        Me.dtFrom.CalendarForeColor = System.Drawing.Color.FromArgb(CType(CType(27, Byte), Integer), CType(CType(41, Byte), Integer), CType(CType(62, Byte), Integer))
        Me.dtFrom.Format = System.Windows.Forms.DateTimePickerFormat.[Short]
        Me.dtFrom.Location = New System.Drawing.Point(73, 5)
        Me.dtFrom.Name = "dtFrom"
        Me.dtFrom.Size = New System.Drawing.Size(107, 22)
        Me.dtFrom.TabIndex = 6
        '
        'dtTo
        '
        Me.dtTo.CalendarForeColor = System.Drawing.Color.FromArgb(CType(CType(27, Byte), Integer), CType(CType(41, Byte), Integer), CType(CType(62, Byte), Integer))
        Me.dtTo.Format = System.Windows.Forms.DateTimePickerFormat.[Short]
        Me.dtTo.Location = New System.Drawing.Point(186, 5)
        Me.dtTo.Name = "dtTo"
        Me.dtTo.Size = New System.Drawing.Size(107, 22)
        Me.dtTo.TabIndex = 5
        '
        'LabelControl4
        '
        Me.LabelControl4.Appearance.ForeColor = System.Drawing.Color.FromArgb(CType(CType(27, Byte), Integer), CType(CType(41, Byte), Integer), CType(CType(62, Byte), Integer))
        Me.LabelControl4.Appearance.Options.UseForeColor = True
        Me.LabelControl4.Location = New System.Drawing.Point(7, 9)
        Me.LabelControl4.Name = "LabelControl4"
        Me.LabelControl4.Size = New System.Drawing.Size(63, 13)
        Me.LabelControl4.TabIndex = 4
        Me.LabelControl4.Text = "Date Range:"
        '
        'LabelControl7
        '
        Me.LabelControl7.Appearance.ForeColor = System.Drawing.Color.FromArgb(CType(CType(27, Byte), Integer), CType(CType(41, Byte), Integer), CType(CType(62, Byte), Integer))
        Me.LabelControl7.Appearance.Options.UseForeColor = True
        Me.LabelControl7.Location = New System.Drawing.Point(525, 10)
        Me.LabelControl7.Name = "LabelControl7"
        Me.LabelControl7.Size = New System.Drawing.Size(70, 13)
        Me.LabelControl7.TabIndex = 41
        Me.LabelControl7.Text = "Performed By:"
        '
        'cboMedtech
        '
        Me.cboMedtech.Location = New System.Drawing.Point(603, 6)
        Me.cboMedtech.Name = "cboMedtech"
        Me.cboMedtech.Properties.Appearance.ForeColor = System.Drawing.Color.FromArgb(CType(CType(27, Byte), Integer), CType(CType(41, Byte), Integer), CType(CType(62, Byte), Integer))
        Me.cboMedtech.Properties.Appearance.Options.UseForeColor = True
        Me.cboMedtech.Properties.Buttons.AddRange(New DevExpress.XtraEditors.Controls.EditorButton() {New DevExpress.XtraEditors.Controls.EditorButton(DevExpress.XtraEditors.Controls.ButtonPredefines.Combo)})
        Me.cboMedtech.Properties.LookAndFeel.SkinName = "Visual Studio 2013 Blue"
        Me.cboMedtech.Properties.LookAndFeel.UseDefaultLookAndFeel = False
        Me.cboMedtech.Properties.TextEditStyle = DevExpress.XtraEditors.Controls.TextEditStyles.DisableTextEditor
        Me.cboMedtech.Properties.UseReadOnlyAppearance = False
        Me.cboMedtech.Size = New System.Drawing.Size(160, 20)
        Me.cboMedtech.TabIndex = 40
        '
        'pnlHeader
        '
        Me.pnlHeader.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.pnlHeader.Appearance.BackColor = System.Drawing.Color.FromArgb(CType(CType(42, Byte), Integer), CType(CType(88, Byte), Integer), CType(CType(173, Byte), Integer))
        Me.pnlHeader.Appearance.Options.UseBackColor = True
        Me.pnlHeader.Controls.Add(Me.cboLimit)
        Me.pnlHeader.Controls.Add(Me.LabelControl1)
        Me.pnlHeader.Controls.Add(Me.LabelControl7)
        Me.pnlHeader.Controls.Add(Me.btnSearch)
        Me.pnlHeader.Controls.Add(Me.cboMedtech)
        Me.pnlHeader.Controls.Add(Me.LabelControl4)
        Me.pnlHeader.Controls.Add(Me.dtFrom)
        Me.pnlHeader.Controls.Add(Me.dtTo)
        Me.pnlHeader.Location = New System.Drawing.Point(8, 42)
        Me.pnlHeader.Name = "pnlHeader"
        Me.pnlHeader.Size = New System.Drawing.Size(1223, 32)
        Me.pnlHeader.TabIndex = 35
        '
        'frmTatComplete
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(1244, 647)
        Me.Controls.Add(Me.pnlHeader)
        Me.Controls.Add(Me.XtraTabControl1)
        Me.Controls.Add(Me.barDockControlLeft)
        Me.Controls.Add(Me.barDockControlRight)
        Me.Controls.Add(Me.barDockControlBottom)
        Me.Controls.Add(Me.barDockControlTop)
        Me.FormBorderEffect = DevExpress.XtraEditors.FormBorderEffect.Glow
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Name = "frmTatComplete"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Worksheet"
        Me.WindowState = System.Windows.Forms.FormWindowState.Maximized
        CType(Me.BarManager, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.BarAndDockingController, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.XtraTabControl1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.XtraTabControl1.ResumeLayout(False)
        Me.XtraTabPage1.ResumeLayout(False)
        Me.XtraTabPage1.PerformLayout()
        CType(Me.dtResult, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.GridView, System.ComponentModel.ISupportInitialize).EndInit()
        Me.XtraTabPage3.ResumeLayout(False)
        Me.XtraTabPage3.PerformLayout()
        CType(Me.dtTaTQuant, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.GridTaTQuant, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.cboLimit.Properties, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.cboMedtech.Properties, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.pnlHeader, System.ComponentModel.ISupportInitialize).EndInit()
        Me.pnlHeader.ResumeLayout(False)
        Me.pnlHeader.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents BarManager As DevExpress.XtraBars.BarManager
    Friend WithEvents Bar2 As DevExpress.XtraBars.Bar
    Friend WithEvents btnPrint As DevExpress.XtraBars.BarLargeButtonItem
    Friend WithEvents btnRefresh As DevExpress.XtraBars.BarLargeButtonItem
    Friend WithEvents btnClose As DevExpress.XtraBars.BarLargeButtonItem
    Friend WithEvents BarAndDockingController As DevExpress.XtraBars.BarAndDockingController
    Friend WithEvents barDockControlTop As DevExpress.XtraBars.BarDockControl
    Friend WithEvents barDockControlBottom As DevExpress.XtraBars.BarDockControl
    Friend WithEvents barDockControlLeft As DevExpress.XtraBars.BarDockControl
    Friend WithEvents barDockControlRight As DevExpress.XtraBars.BarDockControl
    Friend WithEvents ImageList As System.Windows.Forms.ImageList
    Friend WithEvents XtraTabControl1 As DevExpress.XtraTab.XtraTabControl
    Friend WithEvents XtraTabPage1 As DevExpress.XtraTab.XtraTabPage
    Friend WithEvents btnSearch As DevExpress.XtraEditors.SimpleButton
    Friend WithEvents LabelControl1 As DevExpress.XtraEditors.LabelControl
    Friend WithEvents cboLimit As DevExpress.XtraEditors.ComboBoxEdit
    Friend WithEvents dtFrom As System.Windows.Forms.DateTimePicker
    Friend WithEvents dtTo As System.Windows.Forms.DateTimePicker
    Friend WithEvents LabelControl4 As DevExpress.XtraEditors.LabelControl
    Friend WithEvents LabelControl7 As DevExpress.XtraEditors.LabelControl
    Friend WithEvents cboMedtech As DevExpress.XtraEditors.ComboBoxEdit
    Friend WithEvents dtResult As DevExpress.XtraGrid.GridControl
    Friend WithEvents GridView As DevExpress.XtraGrid.Views.Grid.GridView
    Friend WithEvents XtraTabPage3 As DevExpress.XtraTab.XtraTabPage
    Friend WithEvents dtTaTQuant As DevExpress.XtraGrid.GridControl
    Friend WithEvents GridTaTQuant As DevExpress.XtraGrid.Views.Grid.GridView
    Friend WithEvents pnlHeader As DevExpress.XtraEditors.PanelControl
    Friend WithEvents lblDiffTime As DevExpress.XtraEditors.LabelControl
    Friend WithEvents lblTotalCount As DevExpress.XtraEditors.LabelControl
End Class
