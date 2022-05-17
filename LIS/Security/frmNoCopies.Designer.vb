<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmNoCopies
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmNoCopies))
        Me.cboPrinter = New DevExpress.XtraEditors.ComboBoxEdit()
        Me.LabelControl19 = New DevExpress.XtraEditors.LabelControl()
        Me.LabelControl13 = New DevExpress.XtraEditors.LabelControl()
        Me.txtCopies = New DevExpress.XtraEditors.TextEdit()
        Me.btnPrint = New DevExpress.XtraEditors.SimpleButton()
        Me.btnClose = New DevExpress.XtraEditors.SimpleButton()
        Me.SeparatorControl1 = New DevExpress.XtraEditors.SeparatorControl()
        Me.PanelHeader = New DevExpress.XtraEditors.SidePanel()
        CType(Me.cboPrinter.Properties, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.txtCopies.Properties, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.SeparatorControl1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'cboPrinter
        '
        Me.cboPrinter.Location = New System.Drawing.Point(133, 12)
        Me.cboPrinter.Name = "cboPrinter"
        Me.cboPrinter.Properties.Appearance.Font = New System.Drawing.Font("Segoe UI", 9.75!)
        Me.cboPrinter.Properties.Appearance.ForeColor = System.Drawing.Color.FromArgb(CType(CType(27, Byte), Integer), CType(CType(41, Byte), Integer), CType(CType(62, Byte), Integer))
        Me.cboPrinter.Properties.Appearance.Options.UseFont = True
        Me.cboPrinter.Properties.Appearance.Options.UseForeColor = True
        Me.cboPrinter.Properties.Buttons.AddRange(New DevExpress.XtraEditors.Controls.EditorButton() {New DevExpress.XtraEditors.Controls.EditorButton(DevExpress.XtraEditors.Controls.ButtonPredefines.Combo)})
        Me.cboPrinter.Properties.Items.AddRange(New Object() {"Day", "Week", "Month", "Year"})
        Me.cboPrinter.Properties.LookAndFeel.SkinName = "Visual Studio 2013 Blue"
        Me.cboPrinter.Properties.LookAndFeel.UseDefaultLookAndFeel = False
        Me.cboPrinter.Size = New System.Drawing.Size(266, 24)
        Me.cboPrinter.TabIndex = 68
        '
        'LabelControl19
        '
        Me.LabelControl19.Appearance.Font = New System.Drawing.Font("Segoe UI", 9.75!)
        Me.LabelControl19.Appearance.ForeColor = System.Drawing.Color.FromArgb(CType(CType(27, Byte), Integer), CType(CType(41, Byte), Integer), CType(CType(62, Byte), Integer))
        Me.LabelControl19.Appearance.Options.UseFont = True
        Me.LabelControl19.Appearance.Options.UseForeColor = True
        Me.LabelControl19.Location = New System.Drawing.Point(28, 18)
        Me.LabelControl19.LookAndFeel.SkinName = "Visual Studio 2013 Blue"
        Me.LabelControl19.LookAndFeel.UseDefaultLookAndFeel = False
        Me.LabelControl19.Name = "LabelControl19"
        Me.LabelControl19.Size = New System.Drawing.Size(86, 17)
        Me.LabelControl19.TabIndex = 67
        Me.LabelControl19.Text = "Default Printer:"
        '
        'LabelControl13
        '
        Me.LabelControl13.Appearance.Font = New System.Drawing.Font("Segoe UI", 9.75!)
        Me.LabelControl13.Appearance.ForeColor = System.Drawing.Color.FromArgb(CType(CType(27, Byte), Integer), CType(CType(41, Byte), Integer), CType(CType(62, Byte), Integer))
        Me.LabelControl13.Appearance.Options.UseFont = True
        Me.LabelControl13.Appearance.Options.UseForeColor = True
        Me.LabelControl13.Location = New System.Drawing.Point(27, 46)
        Me.LabelControl13.LookAndFeel.SkinName = "Visual Studio 2013 Blue"
        Me.LabelControl13.LookAndFeel.UseDefaultLookAndFeel = False
        Me.LabelControl13.Name = "LabelControl13"
        Me.LabelControl13.Size = New System.Drawing.Size(43, 17)
        Me.LabelControl13.TabIndex = 66
        Me.LabelControl13.Text = "Copies:"
        '
        'txtCopies
        '
        Me.txtCopies.EditValue = "1"
        Me.txtCopies.Location = New System.Drawing.Point(133, 40)
        Me.txtCopies.Name = "txtCopies"
        Me.txtCopies.Properties.Appearance.Font = New System.Drawing.Font("Segoe UI", 9.75!)
        Me.txtCopies.Properties.Appearance.ForeColor = System.Drawing.Color.FromArgb(CType(CType(27, Byte), Integer), CType(CType(41, Byte), Integer), CType(CType(62, Byte), Integer))
        Me.txtCopies.Properties.Appearance.Options.UseFont = True
        Me.txtCopies.Properties.Appearance.Options.UseForeColor = True
        Me.txtCopies.Properties.LookAndFeel.SkinName = "Visual Studio 2013 Blue"
        Me.txtCopies.Properties.LookAndFeel.UseDefaultLookAndFeel = False
        Me.txtCopies.Size = New System.Drawing.Size(180, 24)
        Me.txtCopies.TabIndex = 65
        '
        'btnPrint
        '
        Me.btnPrint.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnPrint.Appearance.BackColor = System.Drawing.Color.FromArgb(CType(CType(88, Byte), Integer), CType(CType(88, Byte), Integer), CType(CType(88, Byte), Integer))
        Me.btnPrint.Appearance.BackColor2 = System.Drawing.Color.FromArgb(CType(CType(88, Byte), Integer), CType(CType(88, Byte), Integer), CType(CType(88, Byte), Integer))
        Me.btnPrint.Appearance.BorderColor = System.Drawing.Color.FromArgb(CType(CType(88, Byte), Integer), CType(CType(88, Byte), Integer), CType(CType(88, Byte), Integer))
        Me.btnPrint.Appearance.Font = New System.Drawing.Font("Segoe UI", 9.75!)
        Me.btnPrint.Appearance.ForeColor = System.Drawing.Color.FromArgb(CType(CType(27, Byte), Integer), CType(CType(41, Byte), Integer), CType(CType(62, Byte), Integer))
        Me.btnPrint.Appearance.Options.UseBackColor = True
        Me.btnPrint.Appearance.Options.UseBorderColor = True
        Me.btnPrint.Appearance.Options.UseFont = True
        Me.btnPrint.Appearance.Options.UseForeColor = True
        Me.btnPrint.AppearanceHovered.BackColor = System.Drawing.Color.Gray
        Me.btnPrint.AppearanceHovered.BackColor2 = System.Drawing.Color.Gray
        Me.btnPrint.AppearanceHovered.BorderColor = System.Drawing.Color.Gray
        Me.btnPrint.AppearanceHovered.Options.UseBackColor = True
        Me.btnPrint.AppearanceHovered.Options.UseBorderColor = True
        Me.btnPrint.Cursor = System.Windows.Forms.Cursors.Hand
        Me.btnPrint.ImageOptions.AllowGlyphSkinning = DevExpress.Utils.DefaultBoolean.[True]
        Me.btnPrint.ImageOptions.Image = CType(resources.GetObject("btnPrint.ImageOptions.Image"), System.Drawing.Image)
        Me.btnPrint.ImageOptions.ImageToTextAlignment = DevExpress.XtraEditors.ImageAlignToText.LeftCenter
        Me.btnPrint.Location = New System.Drawing.Point(203, 295)
        Me.btnPrint.LookAndFeel.SkinName = "The Asphalt World"
        Me.btnPrint.LookAndFeel.UseDefaultLookAndFeel = False
        Me.btnPrint.Name = "btnPrint"
        Me.btnPrint.Size = New System.Drawing.Size(108, 27)
        Me.btnPrint.TabIndex = 69
        Me.btnPrint.Text = "&Print"
        '
        'btnClose
        '
        Me.btnClose.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnClose.Appearance.BackColor = System.Drawing.Color.FromArgb(CType(CType(88, Byte), Integer), CType(CType(88, Byte), Integer), CType(CType(88, Byte), Integer))
        Me.btnClose.Appearance.BackColor2 = System.Drawing.Color.FromArgb(CType(CType(88, Byte), Integer), CType(CType(88, Byte), Integer), CType(CType(88, Byte), Integer))
        Me.btnClose.Appearance.BorderColor = System.Drawing.Color.FromArgb(CType(CType(88, Byte), Integer), CType(CType(88, Byte), Integer), CType(CType(88, Byte), Integer))
        Me.btnClose.Appearance.Font = New System.Drawing.Font("Segoe UI", 9.75!)
        Me.btnClose.Appearance.ForeColor = System.Drawing.Color.FromArgb(CType(CType(27, Byte), Integer), CType(CType(41, Byte), Integer), CType(CType(62, Byte), Integer))
        Me.btnClose.Appearance.Options.UseBackColor = True
        Me.btnClose.Appearance.Options.UseBorderColor = True
        Me.btnClose.Appearance.Options.UseFont = True
        Me.btnClose.Appearance.Options.UseForeColor = True
        Me.btnClose.AppearanceHovered.BackColor = System.Drawing.Color.Gray
        Me.btnClose.AppearanceHovered.BackColor2 = System.Drawing.Color.Gray
        Me.btnClose.AppearanceHovered.BorderColor = System.Drawing.Color.Gray
        Me.btnClose.AppearanceHovered.Options.UseBackColor = True
        Me.btnClose.AppearanceHovered.Options.UseBorderColor = True
        Me.btnClose.Cursor = System.Windows.Forms.Cursors.Hand
        Me.btnClose.ImageOptions.AllowGlyphSkinning = DevExpress.Utils.DefaultBoolean.[True]
        Me.btnClose.ImageOptions.Image = CType(resources.GetObject("btnClose.ImageOptions.Image"), System.Drawing.Image)
        Me.btnClose.ImageOptions.ImageToTextAlignment = DevExpress.XtraEditors.ImageAlignToText.LeftCenter
        Me.btnClose.Location = New System.Drawing.Point(317, 295)
        Me.btnClose.LookAndFeel.SkinName = "The Asphalt World"
        Me.btnClose.LookAndFeel.UseDefaultLookAndFeel = False
        Me.btnClose.Name = "btnClose"
        Me.btnClose.Size = New System.Drawing.Size(108, 27)
        Me.btnClose.TabIndex = 70
        Me.btnClose.Text = "&Close"
        '
        'SeparatorControl1
        '
        Me.SeparatorControl1.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.SeparatorControl1.AutoSizeMode = True
        Me.SeparatorControl1.Location = New System.Drawing.Point(2, 66)
        Me.SeparatorControl1.Name = "SeparatorControl1"
        Me.SeparatorControl1.Size = New System.Drawing.Size(433, 19)
        Me.SeparatorControl1.TabIndex = 94
        '
        'PanelHeader
        '
        Me.PanelHeader.Appearance.BackColor = System.Drawing.Color.FromArgb(CType(CType(6, Byte), Integer), CType(CType(31, Byte), Integer), CType(CType(71, Byte), Integer))
        Me.PanelHeader.Appearance.Options.UseBackColor = True
        Me.PanelHeader.BackgroundImage = CType(resources.GetObject("PanelHeader.BackgroundImage"), System.Drawing.Image)
        Me.PanelHeader.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch
        Me.PanelHeader.BorderThickness = 0
        Me.PanelHeader.Location = New System.Drawing.Point(41, 109)
        Me.PanelHeader.Name = "PanelHeader"
        Me.PanelHeader.Padding = New System.Windows.Forms.Padding(10)
        Me.PanelHeader.Size = New System.Drawing.Size(300, 124)
        Me.PanelHeader.TabIndex = 95
        Me.PanelHeader.Text = "SidePanel2"
        '
        'frmNoCopies
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(437, 332)
        Me.Controls.Add(Me.PanelHeader)
        Me.Controls.Add(Me.SeparatorControl1)
        Me.Controls.Add(Me.btnPrint)
        Me.Controls.Add(Me.btnClose)
        Me.Controls.Add(Me.cboPrinter)
        Me.Controls.Add(Me.LabelControl19)
        Me.Controls.Add(Me.LabelControl13)
        Me.Controls.Add(Me.txtCopies)
        Me.FormBorderEffect = DevExpress.XtraEditors.FormBorderEffect.Glow
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.LookAndFeel.SkinName = "Visual Studio 2013 Blue"
        Me.LookAndFeel.UseDefaultLookAndFeel = False
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "frmNoCopies"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Set Default Printer"
        CType(Me.cboPrinter.Properties, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.txtCopies.Properties, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.SeparatorControl1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents cboPrinter As DevExpress.XtraEditors.ComboBoxEdit
    Friend WithEvents LabelControl19 As DevExpress.XtraEditors.LabelControl
    Friend WithEvents LabelControl13 As DevExpress.XtraEditors.LabelControl
    Friend WithEvents txtCopies As DevExpress.XtraEditors.TextEdit
    Friend WithEvents btnPrint As DevExpress.XtraEditors.SimpleButton
    Friend WithEvents btnClose As DevExpress.XtraEditors.SimpleButton
    Friend WithEvents SeparatorControl1 As DevExpress.XtraEditors.SeparatorControl
    Friend WithEvents PanelHeader As DevExpress.XtraEditors.SidePanel
End Class
