<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmCompanyProfile
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmCompanyProfile))
        Me.LabelControl1 = New DevExpress.XtraEditors.LabelControl()
        Me.SeparatorControl1 = New DevExpress.XtraEditors.SeparatorControl()
        Me.btnSelect = New DevExpress.XtraEditors.SimpleButton()
        Me.btnClear = New DevExpress.XtraEditors.SimpleButton()
        Me.txtAddress = New DevExpress.XtraEditors.MemoEdit()
        Me.txtName = New DevExpress.XtraEditors.MemoEdit()
        Me.picLogo_primary = New System.Windows.Forms.PictureBox()
        Me.txtFax = New DevExpress.XtraEditors.TextEdit()
        Me.LabelControl5 = New DevExpress.XtraEditors.LabelControl()
        Me.txtTel = New DevExpress.XtraEditors.TextEdit()
        Me.LabelControl2 = New DevExpress.XtraEditors.LabelControl()
        Me.LabelControl4 = New DevExpress.XtraEditors.LabelControl()
        Me.btnSave = New DevExpress.XtraEditors.SimpleButton()
        Me.btnClose = New DevExpress.XtraEditors.SimpleButton()
        Me.OpenFileDialog = New System.Windows.Forms.OpenFileDialog()
        Me.ImageCollection1 = New DevExpress.Utils.ImageCollection(Me.components)
        Me.picLogo_secondary = New System.Windows.Forms.PictureBox()
        Me.btnSelect1 = New DevExpress.XtraEditors.SimpleButton()
        Me.btnClear1 = New DevExpress.XtraEditors.SimpleButton()
        Me.LabelControl10 = New DevExpress.XtraEditors.LabelControl()
        CType(Me.SeparatorControl1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.txtAddress.Properties, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.txtName.Properties, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.picLogo_primary, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.txtFax.Properties, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.txtTel.Properties, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.ImageCollection1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.picLogo_secondary, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'LabelControl1
        '
        Me.LabelControl1.Appearance.Font = New System.Drawing.Font("Segoe UI", 9.75!)
        Me.LabelControl1.Appearance.ForeColor = System.Drawing.Color.FromArgb(CType(CType(27, Byte), Integer), CType(CType(41, Byte), Integer), CType(CType(62, Byte), Integer))
        Me.LabelControl1.Appearance.Options.UseFont = True
        Me.LabelControl1.Appearance.Options.UseForeColor = True
        Me.LabelControl1.Location = New System.Drawing.Point(304, 47)
        Me.LabelControl1.LookAndFeel.SkinName = "Visual Studio 2013 Blue"
        Me.LabelControl1.LookAndFeel.UseDefaultLookAndFeel = False
        Me.LabelControl1.Name = "LabelControl1"
        Me.LabelControl1.Size = New System.Drawing.Size(106, 17)
        Me.LabelControl1.TabIndex = 12
        Me.LabelControl1.Text = "Company Name: *"
        '
        'SeparatorControl1
        '
        Me.SeparatorControl1.AutoSizeMode = True
        Me.SeparatorControl1.Location = New System.Drawing.Point(0, 504)
        Me.SeparatorControl1.Name = "SeparatorControl1"
        Me.SeparatorControl1.Size = New System.Drawing.Size(724, 19)
        Me.SeparatorControl1.TabIndex = 103
        '
        'btnSelect
        '
        Me.btnSelect.Appearance.Font = New System.Drawing.Font("Segoe UI", 9.75!)
        Me.btnSelect.Appearance.ForeColor = System.Drawing.Color.FromArgb(CType(CType(27, Byte), Integer), CType(CType(41, Byte), Integer), CType(CType(62, Byte), Integer))
        Me.btnSelect.Appearance.Options.UseFont = True
        Me.btnSelect.Appearance.Options.UseForeColor = True
        Me.btnSelect.ImageOptions.AllowGlyphSkinning = DevExpress.Utils.DefaultBoolean.[True]
        Me.btnSelect.ImageOptions.Image = CType(resources.GetObject("btnSelect.ImageOptions.Image"), System.Drawing.Image)
        Me.btnSelect.Location = New System.Drawing.Point(24, 227)
        Me.btnSelect.Name = "btnSelect"
        Me.btnSelect.Size = New System.Drawing.Size(128, 27)
        Me.btnSelect.TabIndex = 101
        Me.btnSelect.Text = "Choose Logo"
        '
        'btnClear
        '
        Me.btnClear.Appearance.Font = New System.Drawing.Font("Segoe UI", 9.75!)
        Me.btnClear.Appearance.ForeColor = System.Drawing.Color.FromArgb(CType(CType(27, Byte), Integer), CType(CType(41, Byte), Integer), CType(CType(62, Byte), Integer))
        Me.btnClear.Appearance.Options.UseFont = True
        Me.btnClear.Appearance.Options.UseForeColor = True
        Me.btnClear.ImageOptions.AllowGlyphSkinning = DevExpress.Utils.DefaultBoolean.[True]
        Me.btnClear.ImageOptions.Image = CType(resources.GetObject("btnClear.ImageOptions.Image"), System.Drawing.Image)
        Me.btnClear.Location = New System.Drawing.Point(158, 227)
        Me.btnClear.Name = "btnClear"
        Me.btnClear.Size = New System.Drawing.Size(96, 27)
        Me.btnClear.TabIndex = 100
        Me.btnClear.Text = "Clear"
        '
        'txtAddress
        '
        Me.txtAddress.Location = New System.Drawing.Point(417, 131)
        Me.txtAddress.Name = "txtAddress"
        Me.txtAddress.Properties.Appearance.Font = New System.Drawing.Font("Segoe UI", 9.75!)
        Me.txtAddress.Properties.Appearance.ForeColor = System.Drawing.Color.FromArgb(CType(CType(27, Byte), Integer), CType(CType(41, Byte), Integer), CType(CType(62, Byte), Integer))
        Me.txtAddress.Properties.Appearance.Options.UseFont = True
        Me.txtAddress.Properties.Appearance.Options.UseForeColor = True
        Me.txtAddress.Size = New System.Drawing.Size(284, 82)
        Me.txtAddress.TabIndex = 99
        '
        'txtName
        '
        Me.txtName.Location = New System.Drawing.Point(417, 45)
        Me.txtName.Name = "txtName"
        Me.txtName.Properties.Appearance.Font = New System.Drawing.Font("Segoe UI", 9.75!)
        Me.txtName.Properties.Appearance.ForeColor = System.Drawing.Color.FromArgb(CType(CType(27, Byte), Integer), CType(CType(41, Byte), Integer), CType(CType(62, Byte), Integer))
        Me.txtName.Properties.Appearance.Options.UseFont = True
        Me.txtName.Properties.Appearance.Options.UseForeColor = True
        Me.txtName.Size = New System.Drawing.Size(284, 80)
        Me.txtName.TabIndex = 98
        '
        'picLogo_primary
        '
        Me.picLogo_primary.BackColor = System.Drawing.Color.White
        Me.picLogo_primary.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.picLogo_primary.Location = New System.Drawing.Point(24, 12)
        Me.picLogo_primary.Name = "picLogo_primary"
        Me.picLogo_primary.Size = New System.Drawing.Size(230, 210)
        Me.picLogo_primary.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        Me.picLogo_primary.TabIndex = 97
        Me.picLogo_primary.TabStop = False
        '
        'txtFax
        '
        Me.txtFax.Location = New System.Drawing.Point(417, 247)
        Me.txtFax.Name = "txtFax"
        Me.txtFax.Properties.Appearance.Font = New System.Drawing.Font("Segoe UI", 9.75!)
        Me.txtFax.Properties.Appearance.ForeColor = System.Drawing.Color.FromArgb(CType(CType(27, Byte), Integer), CType(CType(41, Byte), Integer), CType(CType(62, Byte), Integer))
        Me.txtFax.Properties.Appearance.Options.UseFont = True
        Me.txtFax.Properties.Appearance.Options.UseForeColor = True
        Me.txtFax.Properties.LookAndFeel.SkinName = "Visual Studio 2013 Blue"
        Me.txtFax.Properties.LookAndFeel.UseDefaultLookAndFeel = False
        Me.txtFax.Size = New System.Drawing.Size(284, 24)
        Me.txtFax.TabIndex = 4
        '
        'LabelControl5
        '
        Me.LabelControl5.Appearance.Font = New System.Drawing.Font("Segoe UI", 9.75!)
        Me.LabelControl5.Appearance.ForeColor = System.Drawing.Color.FromArgb(CType(CType(27, Byte), Integer), CType(CType(41, Byte), Integer), CType(CType(62, Byte), Integer))
        Me.LabelControl5.Appearance.Options.UseFont = True
        Me.LabelControl5.Appearance.Options.UseForeColor = True
        Me.LabelControl5.Location = New System.Drawing.Point(304, 251)
        Me.LabelControl5.LookAndFeel.SkinName = "Visual Studio 2013 Blue"
        Me.LabelControl5.LookAndFeel.UseDefaultLookAndFeel = False
        Me.LabelControl5.Name = "LabelControl5"
        Me.LabelControl5.Size = New System.Drawing.Size(83, 17)
        Me.LabelControl5.TabIndex = 84
        Me.LabelControl5.Text = "Fax No./Email:"
        '
        'txtTel
        '
        Me.txtTel.Location = New System.Drawing.Point(417, 219)
        Me.txtTel.Name = "txtTel"
        Me.txtTel.Properties.Appearance.Font = New System.Drawing.Font("Segoe UI", 9.75!)
        Me.txtTel.Properties.Appearance.ForeColor = System.Drawing.Color.FromArgb(CType(CType(27, Byte), Integer), CType(CType(41, Byte), Integer), CType(CType(62, Byte), Integer))
        Me.txtTel.Properties.Appearance.Options.UseFont = True
        Me.txtTel.Properties.Appearance.Options.UseForeColor = True
        Me.txtTel.Properties.LookAndFeel.SkinName = "Visual Studio 2013 Blue"
        Me.txtTel.Properties.LookAndFeel.UseDefaultLookAndFeel = False
        Me.txtTel.Size = New System.Drawing.Size(284, 24)
        Me.txtTel.TabIndex = 3
        '
        'LabelControl2
        '
        Me.LabelControl2.Appearance.Font = New System.Drawing.Font("Segoe UI", 9.75!)
        Me.LabelControl2.Appearance.ForeColor = System.Drawing.Color.FromArgb(CType(CType(27, Byte), Integer), CType(CType(41, Byte), Integer), CType(CType(62, Byte), Integer))
        Me.LabelControl2.Appearance.Options.UseFont = True
        Me.LabelControl2.Appearance.Options.UseForeColor = True
        Me.LabelControl2.Location = New System.Drawing.Point(304, 223)
        Me.LabelControl2.LookAndFeel.SkinName = "Visual Studio 2013 Blue"
        Me.LabelControl2.LookAndFeel.UseDefaultLookAndFeel = False
        Me.LabelControl2.Name = "LabelControl2"
        Me.LabelControl2.Size = New System.Drawing.Size(94, 17)
        Me.LabelControl2.TabIndex = 78
        Me.LabelControl2.Text = "Tel./Mobile No.:"
        '
        'LabelControl4
        '
        Me.LabelControl4.Appearance.Font = New System.Drawing.Font("Segoe UI", 9.75!)
        Me.LabelControl4.Appearance.ForeColor = System.Drawing.Color.FromArgb(CType(CType(27, Byte), Integer), CType(CType(41, Byte), Integer), CType(CType(62, Byte), Integer))
        Me.LabelControl4.Appearance.Options.UseFont = True
        Me.LabelControl4.Appearance.Options.UseForeColor = True
        Me.LabelControl4.Location = New System.Drawing.Point(304, 133)
        Me.LabelControl4.LookAndFeel.SkinName = "Visual Studio 2013 Blue"
        Me.LabelControl4.LookAndFeel.UseDefaultLookAndFeel = False
        Me.LabelControl4.Name = "LabelControl4"
        Me.LabelControl4.Size = New System.Drawing.Size(60, 17)
        Me.LabelControl4.TabIndex = 80
        Me.LabelControl4.Text = "Address: *"
        '
        'btnSave
        '
        Me.btnSave.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnSave.AppearanceHovered.BackColor = System.Drawing.Color.Gray
        Me.btnSave.AppearanceHovered.BackColor2 = System.Drawing.Color.Gray
        Me.btnSave.AppearanceHovered.BorderColor = System.Drawing.Color.Gray
        Me.btnSave.AppearanceHovered.Options.UseBackColor = True
        Me.btnSave.AppearanceHovered.Options.UseBorderColor = True
        Me.btnSave.Cursor = System.Windows.Forms.Cursors.Hand
        Me.btnSave.ImageOptions.AllowGlyphSkinning = DevExpress.Utils.DefaultBoolean.[True]
        Me.btnSave.ImageOptions.Image = CType(resources.GetObject("btnSave.ImageOptions.Image"), System.Drawing.Image)
        Me.btnSave.ImageOptions.ImageToTextAlignment = DevExpress.XtraEditors.ImageAlignToText.LeftCenter
        Me.btnSave.Location = New System.Drawing.Point(509, 534)
        Me.btnSave.Name = "btnSave"
        Me.btnSave.Size = New System.Drawing.Size(108, 27)
        Me.btnSave.TabIndex = 104
        Me.btnSave.Text = "&Save"
        '
        'btnClose
        '
        Me.btnClose.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnClose.AppearanceHovered.BackColor = System.Drawing.Color.Gray
        Me.btnClose.AppearanceHovered.BackColor2 = System.Drawing.Color.Gray
        Me.btnClose.AppearanceHovered.BorderColor = System.Drawing.Color.Gray
        Me.btnClose.AppearanceHovered.Options.UseBackColor = True
        Me.btnClose.AppearanceHovered.Options.UseBorderColor = True
        Me.btnClose.Cursor = System.Windows.Forms.Cursors.Hand
        Me.btnClose.ImageOptions.AllowGlyphSkinning = DevExpress.Utils.DefaultBoolean.[True]
        Me.btnClose.ImageOptions.Image = CType(resources.GetObject("btnClose.ImageOptions.Image"), System.Drawing.Image)
        Me.btnClose.ImageOptions.ImageToTextAlignment = DevExpress.XtraEditors.ImageAlignToText.LeftCenter
        Me.btnClose.Location = New System.Drawing.Point(623, 534)
        Me.btnClose.Name = "btnClose"
        Me.btnClose.Size = New System.Drawing.Size(108, 27)
        Me.btnClose.TabIndex = 105
        Me.btnClose.Text = "&Close"
        '
        'OpenFileDialog
        '
        Me.OpenFileDialog.FileName = "OpenFileDialog1"
        '
        'ImageCollection1
        '
        Me.ImageCollection1.ImageSize = New System.Drawing.Size(32, 32)
        Me.ImageCollection1.ImageStream = CType(resources.GetObject("ImageCollection1.ImageStream"), DevExpress.Utils.ImageCollectionStreamer)
        Me.ImageCollection1.InsertGalleryImage("backward_32x32.png", "grayscaleimages/navigation/backward_32x32.png", DevExpress.Images.ImageResourceCache.Default.GetImage("grayscaleimages/navigation/backward_32x32.png"), 0)
        Me.ImageCollection1.Images.SetKeyName(0, "backward_32x32.png")
        '
        'picLogo_secondary
        '
        Me.picLogo_secondary.BackColor = System.Drawing.Color.White
        Me.picLogo_secondary.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.picLogo_secondary.Location = New System.Drawing.Point(24, 258)
        Me.picLogo_secondary.Name = "picLogo_secondary"
        Me.picLogo_secondary.Size = New System.Drawing.Size(230, 210)
        Me.picLogo_secondary.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        Me.picLogo_secondary.TabIndex = 106
        Me.picLogo_secondary.TabStop = False
        '
        'btnSelect1
        '
        Me.btnSelect1.Appearance.Font = New System.Drawing.Font("Segoe UI", 9.75!)
        Me.btnSelect1.Appearance.ForeColor = System.Drawing.Color.FromArgb(CType(CType(27, Byte), Integer), CType(CType(41, Byte), Integer), CType(CType(62, Byte), Integer))
        Me.btnSelect1.Appearance.Options.UseFont = True
        Me.btnSelect1.Appearance.Options.UseForeColor = True
        Me.btnSelect1.ImageOptions.AllowGlyphSkinning = DevExpress.Utils.DefaultBoolean.[True]
        Me.btnSelect1.ImageOptions.Image = CType(resources.GetObject("btnSelect1.ImageOptions.Image"), System.Drawing.Image)
        Me.btnSelect1.Location = New System.Drawing.Point(24, 473)
        Me.btnSelect1.Name = "btnSelect1"
        Me.btnSelect1.Size = New System.Drawing.Size(128, 27)
        Me.btnSelect1.TabIndex = 108
        Me.btnSelect1.Text = "Choose Logo"
        '
        'btnClear1
        '
        Me.btnClear1.Appearance.Font = New System.Drawing.Font("Segoe UI", 9.75!)
        Me.btnClear1.Appearance.ForeColor = System.Drawing.Color.FromArgb(CType(CType(27, Byte), Integer), CType(CType(41, Byte), Integer), CType(CType(62, Byte), Integer))
        Me.btnClear1.Appearance.Options.UseFont = True
        Me.btnClear1.Appearance.Options.UseForeColor = True
        Me.btnClear1.ImageOptions.AllowGlyphSkinning = DevExpress.Utils.DefaultBoolean.[True]
        Me.btnClear1.ImageOptions.Image = CType(resources.GetObject("btnClear1.ImageOptions.Image"), System.Drawing.Image)
        Me.btnClear1.Location = New System.Drawing.Point(158, 473)
        Me.btnClear1.Name = "btnClear1"
        Me.btnClear1.Size = New System.Drawing.Size(96, 27)
        Me.btnClear1.TabIndex = 107
        Me.btnClear1.Text = "Clear"
        '
        'LabelControl10
        '
        Me.LabelControl10.Appearance.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Bold)
        Me.LabelControl10.Appearance.ForeColor = System.Drawing.Color.FromArgb(CType(CType(27, Byte), Integer), CType(CType(41, Byte), Integer), CType(CType(62, Byte), Integer))
        Me.LabelControl10.Appearance.Options.UseFont = True
        Me.LabelControl10.Appearance.Options.UseForeColor = True
        Me.LabelControl10.Location = New System.Drawing.Point(272, 12)
        Me.LabelControl10.LookAndFeel.SkinName = "Visual Studio 2013 Blue"
        Me.LabelControl10.LookAndFeel.UseDefaultLookAndFeel = False
        Me.LabelControl10.Name = "LabelControl10"
        Me.LabelControl10.Size = New System.Drawing.Size(103, 17)
        Me.LabelControl10.TabIndex = 119
        Me.LabelControl10.Text = "Company Profile"
        '
        'frmCompanyProfile
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(743, 569)
        Me.Controls.Add(Me.LabelControl10)
        Me.Controls.Add(Me.picLogo_secondary)
        Me.Controls.Add(Me.btnSelect1)
        Me.Controls.Add(Me.btnClear1)
        Me.Controls.Add(Me.btnSave)
        Me.Controls.Add(Me.btnClose)
        Me.Controls.Add(Me.SeparatorControl1)
        Me.Controls.Add(Me.picLogo_primary)
        Me.Controls.Add(Me.btnSelect)
        Me.Controls.Add(Me.LabelControl1)
        Me.Controls.Add(Me.btnClear)
        Me.Controls.Add(Me.LabelControl4)
        Me.Controls.Add(Me.txtAddress)
        Me.Controls.Add(Me.LabelControl2)
        Me.Controls.Add(Me.txtName)
        Me.Controls.Add(Me.txtTel)
        Me.Controls.Add(Me.LabelControl5)
        Me.Controls.Add(Me.txtFax)
        Me.FormBorderEffect = DevExpress.XtraEditors.FormBorderEffect.Glow
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "frmCompanyProfile"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Company Profile"
        CType(Me.SeparatorControl1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.txtAddress.Properties, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.txtName.Properties, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.picLogo_primary, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.txtFax.Properties, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.txtTel.Properties, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.ImageCollection1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.picLogo_secondary, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents LabelControl1 As DevExpress.XtraEditors.LabelControl
    Friend WithEvents LabelControl2 As DevExpress.XtraEditors.LabelControl
    Friend WithEvents LabelControl4 As DevExpress.XtraEditors.LabelControl
    Friend WithEvents txtFax As DevExpress.XtraEditors.TextEdit
    Friend WithEvents LabelControl5 As DevExpress.XtraEditors.LabelControl
    Friend WithEvents txtTel As DevExpress.XtraEditors.TextEdit
    Friend WithEvents picLogo_primary As System.Windows.Forms.PictureBox
    Friend WithEvents txtAddress As DevExpress.XtraEditors.MemoEdit
    Friend WithEvents txtName As DevExpress.XtraEditors.MemoEdit
    Friend WithEvents btnSelect As DevExpress.XtraEditors.SimpleButton
    Friend WithEvents btnClear As DevExpress.XtraEditors.SimpleButton
    Friend WithEvents SeparatorControl1 As DevExpress.XtraEditors.SeparatorControl
    Friend WithEvents btnSave As DevExpress.XtraEditors.SimpleButton
    Friend WithEvents btnClose As DevExpress.XtraEditors.SimpleButton
    Friend WithEvents OpenFileDialog As System.Windows.Forms.OpenFileDialog
    Friend WithEvents ImageCollection1 As DevExpress.Utils.ImageCollection
    Friend WithEvents picLogo_secondary As System.Windows.Forms.PictureBox
    Friend WithEvents btnSelect1 As DevExpress.XtraEditors.SimpleButton
    Friend WithEvents btnClear1 As DevExpress.XtraEditors.SimpleButton
    Friend WithEvents LabelControl10 As DevExpress.XtraEditors.LabelControl
End Class
