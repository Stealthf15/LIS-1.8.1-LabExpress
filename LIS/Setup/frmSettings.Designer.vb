<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmSettings
    Inherits System.Windows.Forms.Form

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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmSettings))
        Me.SimpleButton3 = New DevExpress.XtraEditors.SimpleButton()
        Me.ImageCollection1 = New DevExpress.Utils.ImageCollection(Me.components)
        Me.pnlBackground = New DevExpress.XtraEditors.PanelControl()
        Me.SeparatorControl2 = New DevExpress.XtraEditors.SeparatorControl()
        Me.SeparatorControl1 = New DevExpress.XtraEditors.SeparatorControl()
        Me.SimpleButton1 = New DevExpress.XtraEditors.SimpleButton()
        Me.LabelControl3 = New DevExpress.XtraEditors.LabelControl()
        Me.PanelControl26 = New DevExpress.XtraEditors.PanelControl()
        Me.LabelControl11 = New DevExpress.XtraEditors.LabelControl()
        Me.PictureBox7 = New System.Windows.Forms.PictureBox()
        Me.PanelControl19 = New DevExpress.XtraEditors.PanelControl()
        Me.btnAddAccount = New DevExpress.XtraEditors.SimpleButton()
        Me.btnViewAccount = New DevExpress.XtraEditors.SimpleButton()
        Me.ShapeContainer2 = New Microsoft.VisualBasic.PowerPacks.ShapeContainer()
        Me.LineShape2 = New Microsoft.VisualBasic.PowerPacks.LineShape()
        Me.PanelControl18 = New DevExpress.XtraEditors.PanelControl()
        Me.PictureBox15 = New System.Windows.Forms.PictureBox()
        Me.PanelControl21 = New DevExpress.XtraEditors.PanelControl()
        Me.btnViewReference = New DevExpress.XtraEditors.SimpleButton()
        Me.btnAddReference = New DevExpress.XtraEditors.SimpleButton()
        Me.ShapeContainer9 = New Microsoft.VisualBasic.PowerPacks.ShapeContainer()
        Me.LineShape9 = New Microsoft.VisualBasic.PowerPacks.LineShape()
        Me.LabelControl13 = New DevExpress.XtraEditors.LabelControl()
        Me.pnlFooter = New DevExpress.XtraEditors.PanelControl()
        Me.btnClose = New DevExpress.XtraEditors.SimpleButton()
        Me.btnSave = New DevExpress.XtraEditors.SimpleButton()
        Me.ShapeContainer1 = New Microsoft.VisualBasic.PowerPacks.ShapeContainer()
        Me.LineShape1 = New Microsoft.VisualBasic.PowerPacks.LineShape()
        CType(Me.ImageCollection1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.pnlBackground, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.pnlBackground.SuspendLayout()
        CType(Me.SeparatorControl2, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.SeparatorControl1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.PanelControl26, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.PanelControl26.SuspendLayout()
        CType(Me.PictureBox7, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.PanelControl19, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.PanelControl19.SuspendLayout()
        CType(Me.PanelControl18, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.PanelControl18.SuspendLayout()
        CType(Me.PictureBox15, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.PanelControl21, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.PanelControl21.SuspendLayout()
        CType(Me.pnlFooter, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.pnlFooter.SuspendLayout()
        Me.SuspendLayout()
        '
        'SimpleButton3
        '
        Me.SimpleButton3.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.SimpleButton3.Appearance.Font = New System.Drawing.Font("Segoe UI", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.SimpleButton3.Appearance.Options.UseFont = True
        Me.SimpleButton3.Location = New System.Drawing.Point(-153, -308)
        Me.SimpleButton3.LookAndFeel.SkinName = "DevExpress Dark Style"
        Me.SimpleButton3.LookAndFeel.UseDefaultLookAndFeel = False
        Me.SimpleButton3.Name = "SimpleButton3"
        Me.SimpleButton3.Size = New System.Drawing.Size(49, 28)
        Me.SimpleButton3.TabIndex = 65
        Me.SimpleButton3.Text = "..."
        '
        'ImageCollection1
        '
        Me.ImageCollection1.ImageSize = New System.Drawing.Size(32, 32)
        Me.ImageCollection1.ImageStream = CType(resources.GetObject("ImageCollection1.ImageStream"), DevExpress.Utils.ImageCollectionStreamer)
        Me.ImageCollection1.InsertGalleryImage("backward_32x32.png", "grayscaleimages/navigation/backward_32x32.png", DevExpress.Images.ImageResourceCache.Default.GetImage("grayscaleimages/navigation/backward_32x32.png"), 0)
        Me.ImageCollection1.Images.SetKeyName(0, "backward_32x32.png")
        '
        'pnlBackground
        '
        Me.pnlBackground.AutoSize = True
        Me.pnlBackground.Controls.Add(Me.SeparatorControl2)
        Me.pnlBackground.Controls.Add(Me.SeparatorControl1)
        Me.pnlBackground.Controls.Add(Me.SimpleButton1)
        Me.pnlBackground.Controls.Add(Me.LabelControl3)
        Me.pnlBackground.Controls.Add(Me.PanelControl26)
        Me.pnlBackground.Controls.Add(Me.PanelControl18)
        Me.pnlBackground.Controls.Add(Me.pnlFooter)
        Me.pnlBackground.Controls.Add(Me.SimpleButton3)
        Me.pnlBackground.Dock = System.Windows.Forms.DockStyle.Fill
        Me.pnlBackground.Location = New System.Drawing.Point(0, 0)
        Me.pnlBackground.LookAndFeel.Style = DevExpress.LookAndFeel.LookAndFeelStyle.Office2003
        Me.pnlBackground.LookAndFeel.UseDefaultLookAndFeel = False
        Me.pnlBackground.Name = "pnlBackground"
        Me.pnlBackground.Size = New System.Drawing.Size(843, 356)
        Me.pnlBackground.TabIndex = 1
        '
        'SeparatorControl2
        '
        Me.SeparatorControl2.Anchor = CType((System.Windows.Forms.AnchorStyles.Left Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.SeparatorControl2.AutoSizeMode = True
        Me.SeparatorControl2.Location = New System.Drawing.Point(12, 255)
        Me.SeparatorControl2.Name = "SeparatorControl2"
        Me.SeparatorControl2.Size = New System.Drawing.Size(819, 20)
        Me.SeparatorControl2.TabIndex = 98
        '
        'SeparatorControl1
        '
        Me.SeparatorControl1.Anchor = CType((System.Windows.Forms.AnchorStyles.Left Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.SeparatorControl1.AutoSizeMode = True
        Me.SeparatorControl1.Location = New System.Drawing.Point(12, 42)
        Me.SeparatorControl1.Name = "SeparatorControl1"
        Me.SeparatorControl1.Size = New System.Drawing.Size(819, 20)
        Me.SeparatorControl1.TabIndex = 97
        '
        'SimpleButton1
        '
        Me.SimpleButton1.Appearance.BackColor = System.Drawing.Color.Transparent
        Me.SimpleButton1.Appearance.BackColor2 = System.Drawing.Color.Transparent
        Me.SimpleButton1.Appearance.BorderColor = System.Drawing.Color.Transparent
        Me.SimpleButton1.Appearance.Font = New System.Drawing.Font("Segoe UI", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.SimpleButton1.Appearance.ForeColor = System.Drawing.Color.Gray
        Me.SimpleButton1.Appearance.Options.UseBackColor = True
        Me.SimpleButton1.Appearance.Options.UseBorderColor = True
        Me.SimpleButton1.Appearance.Options.UseFont = True
        Me.SimpleButton1.Appearance.Options.UseForeColor = True
        Me.SimpleButton1.AppearanceHovered.BackColor = System.Drawing.SystemColors.Control
        Me.SimpleButton1.AppearanceHovered.Options.UseBackColor = True
        Me.SimpleButton1.ImageOptions.AllowGlyphSkinning = DevExpress.Utils.DefaultBoolean.[True]
        Me.SimpleButton1.ImageOptions.Image = CType(resources.GetObject("SimpleButton1.ImageOptions.Image"), System.Drawing.Image)
        Me.SimpleButton1.ImageOptions.Location = DevExpress.XtraEditors.ImageLocation.MiddleCenter
        Me.SimpleButton1.Location = New System.Drawing.Point(811, 5)
        Me.SimpleButton1.LookAndFeel.SkinName = "Visual Studio 2013 Blue"
        Me.SimpleButton1.LookAndFeel.Style = DevExpress.LookAndFeel.LookAndFeelStyle.UltraFlat
        Me.SimpleButton1.LookAndFeel.UseDefaultLookAndFeel = False
        Me.SimpleButton1.Name = "SimpleButton1"
        Me.SimpleButton1.Size = New System.Drawing.Size(27, 26)
        Me.SimpleButton1.TabIndex = 73
        Me.SimpleButton1.Text = "X"
        '
        'LabelControl3
        '
        Me.LabelControl3.Appearance.Font = New System.Drawing.Font("Segoe UI Semibold", 18.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.LabelControl3.Appearance.ForeColor = System.Drawing.Color.FromArgb(CType(CType(42, Byte), Integer), CType(CType(88, Byte), Integer), CType(CType(173, Byte), Integer))
        Me.LabelControl3.Appearance.Options.UseFont = True
        Me.LabelControl3.Appearance.Options.UseForeColor = True
        Me.LabelControl3.AutoSizeMode = DevExpress.XtraEditors.LabelAutoSizeMode.None
        Me.LabelControl3.Location = New System.Drawing.Point(12, 7)
        Me.LabelControl3.Name = "LabelControl3"
        Me.LabelControl3.Size = New System.Drawing.Size(257, 34)
        Me.LabelControl3.TabIndex = 72
        Me.LabelControl3.Text = "Settings Window"
        '
        'PanelControl26
        '
        Me.PanelControl26.Appearance.BackColor = System.Drawing.SystemColors.Control
        Me.PanelControl26.Appearance.Options.UseBackColor = True
        Me.PanelControl26.BorderStyle = DevExpress.XtraEditors.Controls.BorderStyles.NoBorder
        Me.PanelControl26.Controls.Add(Me.LabelControl11)
        Me.PanelControl26.Controls.Add(Me.PictureBox7)
        Me.PanelControl26.Controls.Add(Me.PanelControl19)
        Me.PanelControl26.Location = New System.Drawing.Point(422, 71)
        Me.PanelControl26.LookAndFeel.SkinName = "Office 2013 Light Gray"
        Me.PanelControl26.LookAndFeel.UseDefaultLookAndFeel = False
        Me.PanelControl26.Name = "PanelControl26"
        Me.PanelControl26.Size = New System.Drawing.Size(243, 170)
        Me.PanelControl26.TabIndex = 71
        '
        'LabelControl11
        '
        Me.LabelControl11.Appearance.Font = New System.Drawing.Font("Segoe UI Semibold", 18.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.LabelControl11.Appearance.ForeColor = System.Drawing.Color.FromArgb(CType(CType(42, Byte), Integer), CType(CType(88, Byte), Integer), CType(CType(173, Byte), Integer))
        Me.LabelControl11.Appearance.Options.UseFont = True
        Me.LabelControl11.Appearance.Options.UseForeColor = True
        Me.LabelControl11.Appearance.Options.UseTextOptions = True
        Me.LabelControl11.Appearance.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center
        Me.LabelControl11.AutoSizeMode = DevExpress.XtraEditors.LabelAutoSizeMode.None
        Me.LabelControl11.Location = New System.Drawing.Point(3, 129)
        Me.LabelControl11.Name = "LabelControl11"
        Me.LabelControl11.Size = New System.Drawing.Size(165, 32)
        Me.LabelControl11.TabIndex = 10
        Me.LabelControl11.Text = "Result Order"
        '
        'PictureBox7
        '
        Me.PictureBox7.Image = CType(resources.GetObject("PictureBox7.Image"), System.Drawing.Image)
        Me.PictureBox7.Location = New System.Drawing.Point(34, 18)
        Me.PictureBox7.Name = "PictureBox7"
        Me.PictureBox7.Size = New System.Drawing.Size(100, 98)
        Me.PictureBox7.TabIndex = 9
        Me.PictureBox7.TabStop = False
        '
        'PanelControl19
        '
        Me.PanelControl19.Appearance.BackColor = System.Drawing.Color.FromArgb(CType(CType(42, Byte), Integer), CType(CType(88, Byte), Integer), CType(CType(173, Byte), Integer))
        Me.PanelControl19.Appearance.Options.UseBackColor = True
        Me.PanelControl19.BorderStyle = DevExpress.XtraEditors.Controls.BorderStyles.NoBorder
        Me.PanelControl19.Controls.Add(Me.btnAddAccount)
        Me.PanelControl19.Controls.Add(Me.btnViewAccount)
        Me.PanelControl19.Controls.Add(Me.ShapeContainer2)
        Me.PanelControl19.Dock = System.Windows.Forms.DockStyle.Right
        Me.PanelControl19.Location = New System.Drawing.Point(171, 0)
        Me.PanelControl19.Name = "PanelControl19"
        Me.PanelControl19.Size = New System.Drawing.Size(72, 170)
        Me.PanelControl19.TabIndex = 6
        '
        'btnAddAccount
        '
        Me.btnAddAccount.Appearance.BackColor = System.Drawing.Color.Transparent
        Me.btnAddAccount.Appearance.BackColor2 = System.Drawing.Color.Transparent
        Me.btnAddAccount.Appearance.BorderColor = System.Drawing.Color.Transparent
        Me.btnAddAccount.Appearance.Font = New System.Drawing.Font("Segoe UI", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnAddAccount.Appearance.ForeColor = System.Drawing.Color.White
        Me.btnAddAccount.Appearance.Options.UseBackColor = True
        Me.btnAddAccount.Appearance.Options.UseBorderColor = True
        Me.btnAddAccount.Appearance.Options.UseFont = True
        Me.btnAddAccount.Appearance.Options.UseForeColor = True
        Me.btnAddAccount.AppearanceHovered.BackColor = System.Drawing.Color.CornflowerBlue
        Me.btnAddAccount.AppearanceHovered.Options.UseBackColor = True
        Me.btnAddAccount.Cursor = System.Windows.Forms.Cursors.Hand
        Me.btnAddAccount.Dock = System.Windows.Forms.DockStyle.Top
        Me.btnAddAccount.ImageOptions.AllowGlyphSkinning = DevExpress.Utils.DefaultBoolean.[True]
        Me.btnAddAccount.ImageOptions.Image = CType(resources.GetObject("btnAddAccount.ImageOptions.Image"), System.Drawing.Image)
        Me.btnAddAccount.ImageOptions.Location = DevExpress.XtraEditors.ImageLocation.MiddleCenter
        Me.btnAddAccount.Location = New System.Drawing.Point(0, 0)
        Me.btnAddAccount.LookAndFeel.SkinName = "Visual Studio 2013 Blue"
        Me.btnAddAccount.LookAndFeel.Style = DevExpress.LookAndFeel.LookAndFeelStyle.UltraFlat
        Me.btnAddAccount.LookAndFeel.UseDefaultLookAndFeel = False
        Me.btnAddAccount.Name = "btnAddAccount"
        Me.btnAddAccount.Size = New System.Drawing.Size(72, 82)
        Me.btnAddAccount.TabIndex = 0
        '
        'btnViewAccount
        '
        Me.btnViewAccount.Appearance.Font = New System.Drawing.Font("Segoe UI", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnViewAccount.Appearance.ForeColor = System.Drawing.Color.White
        Me.btnViewAccount.Appearance.Options.UseFont = True
        Me.btnViewAccount.Appearance.Options.UseForeColor = True
        Me.btnViewAccount.AppearanceHovered.BackColor = System.Drawing.Color.CornflowerBlue
        Me.btnViewAccount.AppearanceHovered.Options.UseBackColor = True
        Me.btnViewAccount.ButtonStyle = DevExpress.XtraEditors.Controls.BorderStyles.NoBorder
        Me.btnViewAccount.Cursor = System.Windows.Forms.Cursors.Hand
        Me.btnViewAccount.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.btnViewAccount.ImageOptions.AllowGlyphSkinning = DevExpress.Utils.DefaultBoolean.[True]
        Me.btnViewAccount.ImageOptions.Image = CType(resources.GetObject("btnViewAccount.ImageOptions.Image"), System.Drawing.Image)
        Me.btnViewAccount.ImageOptions.Location = DevExpress.XtraEditors.ImageLocation.MiddleCenter
        Me.btnViewAccount.Location = New System.Drawing.Point(0, 88)
        Me.btnViewAccount.Name = "btnViewAccount"
        Me.btnViewAccount.Size = New System.Drawing.Size(72, 82)
        Me.btnViewAccount.TabIndex = 1
        '
        'ShapeContainer2
        '
        Me.ShapeContainer2.Location = New System.Drawing.Point(0, 0)
        Me.ShapeContainer2.Margin = New System.Windows.Forms.Padding(0)
        Me.ShapeContainer2.Name = "ShapeContainer2"
        Me.ShapeContainer2.Shapes.AddRange(New Microsoft.VisualBasic.PowerPacks.Shape() {Me.LineShape2})
        Me.ShapeContainer2.Size = New System.Drawing.Size(72, 170)
        Me.ShapeContainer2.TabIndex = 2
        Me.ShapeContainer2.TabStop = False
        '
        'LineShape2
        '
        Me.LineShape2.BorderColor = System.Drawing.SystemColors.ButtonFace
        Me.LineShape2.Name = "LineShape2"
        Me.LineShape2.X1 = -2
        Me.LineShape2.X2 = 74
        Me.LineShape2.Y1 = 85
        Me.LineShape2.Y2 = 85
        '
        'PanelControl18
        '
        Me.PanelControl18.Appearance.BackColor = System.Drawing.SystemColors.Control
        Me.PanelControl18.Appearance.Options.UseBackColor = True
        Me.PanelControl18.BorderStyle = DevExpress.XtraEditors.Controls.BorderStyles.NoBorder
        Me.PanelControl18.Controls.Add(Me.PictureBox15)
        Me.PanelControl18.Controls.Add(Me.PanelControl21)
        Me.PanelControl18.Controls.Add(Me.LabelControl13)
        Me.PanelControl18.Cursor = System.Windows.Forms.Cursors.Hand
        Me.PanelControl18.Location = New System.Drawing.Point(168, 71)
        Me.PanelControl18.LookAndFeel.SkinName = "Office 2013 Light Gray"
        Me.PanelControl18.LookAndFeel.UseDefaultLookAndFeel = False
        Me.PanelControl18.Name = "PanelControl18"
        Me.PanelControl18.Size = New System.Drawing.Size(243, 170)
        Me.PanelControl18.TabIndex = 70
        '
        'PictureBox15
        '
        Me.PictureBox15.Image = CType(resources.GetObject("PictureBox15.Image"), System.Drawing.Image)
        Me.PictureBox15.Location = New System.Drawing.Point(32, 18)
        Me.PictureBox15.Name = "PictureBox15"
        Me.PictureBox15.Size = New System.Drawing.Size(100, 98)
        Me.PictureBox15.TabIndex = 8
        Me.PictureBox15.TabStop = False
        '
        'PanelControl21
        '
        Me.PanelControl21.Appearance.BackColor = System.Drawing.Color.FromArgb(CType(CType(42, Byte), Integer), CType(CType(88, Byte), Integer), CType(CType(173, Byte), Integer))
        Me.PanelControl21.Appearance.Options.UseBackColor = True
        Me.PanelControl21.BorderStyle = DevExpress.XtraEditors.Controls.BorderStyles.NoBorder
        Me.PanelControl21.Controls.Add(Me.btnViewReference)
        Me.PanelControl21.Controls.Add(Me.btnAddReference)
        Me.PanelControl21.Controls.Add(Me.ShapeContainer9)
        Me.PanelControl21.Dock = System.Windows.Forms.DockStyle.Right
        Me.PanelControl21.Location = New System.Drawing.Point(171, 0)
        Me.PanelControl21.Name = "PanelControl21"
        Me.PanelControl21.Size = New System.Drawing.Size(72, 170)
        Me.PanelControl21.TabIndex = 6
        '
        'btnViewReference
        '
        Me.btnViewReference.Appearance.Font = New System.Drawing.Font("Segoe UI", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnViewReference.Appearance.ForeColor = System.Drawing.Color.White
        Me.btnViewReference.Appearance.Options.UseFont = True
        Me.btnViewReference.Appearance.Options.UseForeColor = True
        Me.btnViewReference.AppearanceHovered.BackColor = System.Drawing.Color.CornflowerBlue
        Me.btnViewReference.AppearanceHovered.Options.UseBackColor = True
        Me.btnViewReference.ButtonStyle = DevExpress.XtraEditors.Controls.BorderStyles.NoBorder
        Me.btnViewReference.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.btnViewReference.ImageOptions.AllowGlyphSkinning = DevExpress.Utils.DefaultBoolean.[True]
        Me.btnViewReference.ImageOptions.Image = CType(resources.GetObject("btnViewReference.ImageOptions.Image"), System.Drawing.Image)
        Me.btnViewReference.ImageOptions.Location = DevExpress.XtraEditors.ImageLocation.MiddleCenter
        Me.btnViewReference.Location = New System.Drawing.Point(0, 88)
        Me.btnViewReference.Name = "btnViewReference"
        Me.btnViewReference.Size = New System.Drawing.Size(72, 82)
        Me.btnViewReference.TabIndex = 1
        '
        'btnAddReference
        '
        Me.btnAddReference.Appearance.BackColor = System.Drawing.Color.Transparent
        Me.btnAddReference.Appearance.BackColor2 = System.Drawing.Color.Transparent
        Me.btnAddReference.Appearance.BorderColor = System.Drawing.Color.Transparent
        Me.btnAddReference.Appearance.Font = New System.Drawing.Font("Segoe UI", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnAddReference.Appearance.ForeColor = System.Drawing.Color.White
        Me.btnAddReference.Appearance.Options.UseBackColor = True
        Me.btnAddReference.Appearance.Options.UseBorderColor = True
        Me.btnAddReference.Appearance.Options.UseFont = True
        Me.btnAddReference.Appearance.Options.UseForeColor = True
        Me.btnAddReference.AppearanceHovered.BackColor = System.Drawing.Color.CornflowerBlue
        Me.btnAddReference.AppearanceHovered.Options.UseBackColor = True
        Me.btnAddReference.Dock = System.Windows.Forms.DockStyle.Top
        Me.btnAddReference.ImageOptions.AllowGlyphSkinning = DevExpress.Utils.DefaultBoolean.[True]
        Me.btnAddReference.ImageOptions.Image = CType(resources.GetObject("btnAddReference.ImageOptions.Image"), System.Drawing.Image)
        Me.btnAddReference.ImageOptions.Location = DevExpress.XtraEditors.ImageLocation.MiddleCenter
        Me.btnAddReference.Location = New System.Drawing.Point(0, 0)
        Me.btnAddReference.LookAndFeel.SkinName = "Visual Studio 2013 Blue"
        Me.btnAddReference.LookAndFeel.Style = DevExpress.LookAndFeel.LookAndFeelStyle.UltraFlat
        Me.btnAddReference.LookAndFeel.UseDefaultLookAndFeel = False
        Me.btnAddReference.Name = "btnAddReference"
        Me.btnAddReference.Size = New System.Drawing.Size(72, 82)
        Me.btnAddReference.TabIndex = 0
        '
        'ShapeContainer9
        '
        Me.ShapeContainer9.Location = New System.Drawing.Point(0, 0)
        Me.ShapeContainer9.Margin = New System.Windows.Forms.Padding(0)
        Me.ShapeContainer9.Name = "ShapeContainer9"
        Me.ShapeContainer9.Shapes.AddRange(New Microsoft.VisualBasic.PowerPacks.Shape() {Me.LineShape9})
        Me.ShapeContainer9.Size = New System.Drawing.Size(72, 170)
        Me.ShapeContainer9.TabIndex = 2
        Me.ShapeContainer9.TabStop = False
        '
        'LineShape9
        '
        Me.LineShape9.BorderColor = System.Drawing.SystemColors.ButtonFace
        Me.LineShape9.Name = "LineShape1"
        Me.LineShape9.X1 = -2
        Me.LineShape9.X2 = 74
        Me.LineShape9.Y1 = 85
        Me.LineShape9.Y2 = 85
        '
        'LabelControl13
        '
        Me.LabelControl13.Appearance.Font = New System.Drawing.Font("Segoe UI Semibold", 14.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.LabelControl13.Appearance.ForeColor = System.Drawing.Color.FromArgb(CType(CType(42, Byte), Integer), CType(CType(88, Byte), Integer), CType(CType(173, Byte), Integer))
        Me.LabelControl13.Appearance.Options.UseFont = True
        Me.LabelControl13.Appearance.Options.UseForeColor = True
        Me.LabelControl13.Appearance.Options.UseTextOptions = True
        Me.LabelControl13.Appearance.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center
        Me.LabelControl13.AutoSizeMode = DevExpress.XtraEditors.LabelAutoSizeMode.None
        Me.LabelControl13.Location = New System.Drawing.Point(3, 129)
        Me.LabelControl13.Name = "LabelControl13"
        Me.LabelControl13.Size = New System.Drawing.Size(165, 32)
        Me.LabelControl13.TabIndex = 5
        Me.LabelControl13.Text = "Reference Range"
        '
        'pnlFooter
        '
        Me.pnlFooter.Controls.Add(Me.btnClose)
        Me.pnlFooter.Controls.Add(Me.btnSave)
        Me.pnlFooter.Controls.Add(Me.ShapeContainer1)
        Me.pnlFooter.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.pnlFooter.Location = New System.Drawing.Point(2, 287)
        Me.pnlFooter.LookAndFeel.SkinName = "DevExpress Dark Style"
        Me.pnlFooter.LookAndFeel.UseDefaultLookAndFeel = False
        Me.pnlFooter.Name = "pnlFooter"
        Me.pnlFooter.Size = New System.Drawing.Size(839, 67)
        Me.pnlFooter.TabIndex = 69
        '
        'btnClose
        '
        Me.btnClose.Appearance.Font = New System.Drawing.Font("Segoe UI Light", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnClose.Appearance.ForeColor = System.Drawing.Color.White
        Me.btnClose.Appearance.Options.UseFont = True
        Me.btnClose.Appearance.Options.UseForeColor = True
        Me.btnClose.AppearanceHovered.BackColor = System.Drawing.Color.Gray
        Me.btnClose.AppearanceHovered.BackColor2 = System.Drawing.Color.Gray
        Me.btnClose.AppearanceHovered.BorderColor = System.Drawing.Color.Gray
        Me.btnClose.AppearanceHovered.Options.UseBackColor = True
        Me.btnClose.AppearanceHovered.Options.UseBorderColor = True
        Me.btnClose.ButtonStyle = DevExpress.XtraEditors.Controls.BorderStyles.NoBorder
        Me.btnClose.Cursor = System.Windows.Forms.Cursors.Hand
        Me.btnClose.ImageOptions.AllowGlyphSkinning = DevExpress.Utils.DefaultBoolean.[True]
        Me.btnClose.ImageOptions.Image = CType(resources.GetObject("btnClose.ImageOptions.Image"), System.Drawing.Image)
        Me.btnClose.ImageOptions.ImageToTextAlignment = DevExpress.XtraEditors.ImageAlignToText.TopCenter
        Me.btnClose.Location = New System.Drawing.Point(424, 5)
        Me.btnClose.LookAndFeel.Style = DevExpress.LookAndFeel.LookAndFeelStyle.Flat
        Me.btnClose.Name = "btnClose"
        Me.btnClose.Size = New System.Drawing.Size(114, 57)
        Me.btnClose.TabIndex = 1
        Me.btnClose.Text = "&Close"
        '
        'btnSave
        '
        Me.btnSave.Appearance.Font = New System.Drawing.Font("Segoe UI Light", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnSave.Appearance.ForeColor = System.Drawing.Color.White
        Me.btnSave.Appearance.Options.UseFont = True
        Me.btnSave.Appearance.Options.UseForeColor = True
        Me.btnSave.AppearanceHovered.BackColor = System.Drawing.Color.Gray
        Me.btnSave.AppearanceHovered.BackColor2 = System.Drawing.Color.Gray
        Me.btnSave.AppearanceHovered.BorderColor = System.Drawing.Color.Gray
        Me.btnSave.AppearanceHovered.Options.UseBackColor = True
        Me.btnSave.AppearanceHovered.Options.UseBorderColor = True
        Me.btnSave.ButtonStyle = DevExpress.XtraEditors.Controls.BorderStyles.NoBorder
        Me.btnSave.Cursor = System.Windows.Forms.Cursors.Hand
        Me.btnSave.ImageOptions.AllowGlyphSkinning = DevExpress.Utils.DefaultBoolean.[True]
        Me.btnSave.ImageOptions.Image = CType(resources.GetObject("btnSave.ImageOptions.Image"), System.Drawing.Image)
        Me.btnSave.ImageOptions.ImageToTextAlignment = DevExpress.XtraEditors.ImageAlignToText.TopCenter
        Me.btnSave.Location = New System.Drawing.Point(301, 5)
        Me.btnSave.LookAndFeel.Style = DevExpress.LookAndFeel.LookAndFeelStyle.Flat
        Me.btnSave.Name = "btnSave"
        Me.btnSave.Size = New System.Drawing.Size(114, 57)
        Me.btnSave.TabIndex = 0
        Me.btnSave.Text = "&Save"
        '
        'ShapeContainer1
        '
        Me.ShapeContainer1.Location = New System.Drawing.Point(2, 2)
        Me.ShapeContainer1.Margin = New System.Windows.Forms.Padding(0)
        Me.ShapeContainer1.Name = "ShapeContainer1"
        Me.ShapeContainer1.Shapes.AddRange(New Microsoft.VisualBasic.PowerPacks.Shape() {Me.LineShape1})
        Me.ShapeContainer1.Size = New System.Drawing.Size(835, 63)
        Me.ShapeContainer1.TabIndex = 5
        Me.ShapeContainer1.TabStop = False
        '
        'LineShape1
        '
        Me.LineShape1.BorderColor = System.Drawing.SystemColors.Control
        Me.LineShape1.Name = "LineShape1"
        Me.LineShape1.X1 = 417
        Me.LineShape1.X2 = 417
        Me.LineShape1.Y1 = 2
        Me.LineShape1.Y2 = 59
        '
        'frmSettings
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(843, 356)
        Me.Controls.Add(Me.pnlBackground)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None
        Me.name = "frmSettings"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        CType(Me.ImageCollection1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.pnlBackground, System.ComponentModel.ISupportInitialize).EndInit()
        Me.pnlBackground.ResumeLayout(False)
        CType(Me.SeparatorControl2, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.SeparatorControl1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.PanelControl26, System.ComponentModel.ISupportInitialize).EndInit()
        Me.PanelControl26.ResumeLayout(False)
        CType(Me.PictureBox7, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.PanelControl19, System.ComponentModel.ISupportInitialize).EndInit()
        Me.PanelControl19.ResumeLayout(False)
        CType(Me.PanelControl18, System.ComponentModel.ISupportInitialize).EndInit()
        Me.PanelControl18.ResumeLayout(False)
        CType(Me.PictureBox15, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.PanelControl21, System.ComponentModel.ISupportInitialize).EndInit()
        Me.PanelControl21.ResumeLayout(False)
        CType(Me.pnlFooter, System.ComponentModel.ISupportInitialize).EndInit()
        Me.pnlFooter.ResumeLayout(False)
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents SimpleButton3 As DevExpress.XtraEditors.SimpleButton
    Friend WithEvents ImageCollection1 As DevExpress.Utils.ImageCollection
    Friend WithEvents pnlBackground As DevExpress.XtraEditors.PanelControl
    Friend WithEvents pnlFooter As DevExpress.XtraEditors.PanelControl
    Friend WithEvents btnClose As DevExpress.XtraEditors.SimpleButton
    Friend WithEvents btnSave As DevExpress.XtraEditors.SimpleButton
    Friend WithEvents ShapeContainer1 As Microsoft.VisualBasic.PowerPacks.ShapeContainer
    Friend WithEvents LineShape1 As Microsoft.VisualBasic.PowerPacks.LineShape
    Friend WithEvents PanelControl26 As DevExpress.XtraEditors.PanelControl
    Friend WithEvents LabelControl11 As DevExpress.XtraEditors.LabelControl
    Friend WithEvents PictureBox7 As System.Windows.Forms.PictureBox
    Friend WithEvents PanelControl19 As DevExpress.XtraEditors.PanelControl
    Friend WithEvents btnAddAccount As DevExpress.XtraEditors.SimpleButton
    Friend WithEvents btnViewAccount As DevExpress.XtraEditors.SimpleButton
    Friend WithEvents ShapeContainer2 As Microsoft.VisualBasic.PowerPacks.ShapeContainer
    Friend WithEvents LineShape2 As Microsoft.VisualBasic.PowerPacks.LineShape
    Friend WithEvents PanelControl18 As DevExpress.XtraEditors.PanelControl
    Friend WithEvents PictureBox15 As System.Windows.Forms.PictureBox
    Friend WithEvents PanelControl21 As DevExpress.XtraEditors.PanelControl
    Friend WithEvents btnViewReference As DevExpress.XtraEditors.SimpleButton
    Friend WithEvents btnAddReference As DevExpress.XtraEditors.SimpleButton
    Friend WithEvents ShapeContainer9 As Microsoft.VisualBasic.PowerPacks.ShapeContainer
    Friend WithEvents LineShape9 As Microsoft.VisualBasic.PowerPacks.LineShape
    Friend WithEvents LabelControl13 As DevExpress.XtraEditors.LabelControl
    Friend WithEvents SimpleButton1 As DevExpress.XtraEditors.SimpleButton
    Friend WithEvents LabelControl3 As DevExpress.XtraEditors.LabelControl
    Friend WithEvents SeparatorControl1 As DevExpress.XtraEditors.SeparatorControl
    Friend WithEvents SeparatorControl2 As DevExpress.XtraEditors.SeparatorControl
End Class
