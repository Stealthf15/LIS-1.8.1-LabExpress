<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmLogin
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmLogin))
        Me.txtPassword = New DevExpress.XtraEditors.TextEdit()
        Me.btnSave = New DevExpress.XtraEditors.SimpleButton()
        Me.btnClose = New DevExpress.XtraEditors.SimpleButton()
        Me.txtUsername = New DevExpress.XtraEditors.TextEdit()
        Me.chkNewWorklist = New DevExpress.XtraEditors.CheckEdit()
        Me.cboLocation = New DevExpress.XtraEditors.ComboBoxEdit()
        Me.SeparatorControl2 = New DevExpress.XtraEditors.SeparatorControl()
        Me.SeparatorControl3 = New DevExpress.XtraEditors.SeparatorControl()
        Me.SeparatorControl4 = New DevExpress.XtraEditors.SeparatorControl()
        Me.LabelControl1 = New DevExpress.XtraEditors.LabelControl()
        Me.lblDB = New DevExpress.XtraEditors.LabelControl()
        CType(Me.txtPassword.Properties, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.txtUsername.Properties, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.chkNewWorklist.Properties, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.cboLocation.Properties, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.SeparatorControl2, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.SeparatorControl3, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.SeparatorControl4, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'txtPassword
        '
        Me.txtPassword.Location = New System.Drawing.Point(20, 129)
        Me.txtPassword.Margin = New System.Windows.Forms.Padding(2, 3, 2, 3)
        Me.txtPassword.Name = "txtPassword"
        Me.txtPassword.Properties.Appearance.BackColor = System.Drawing.Color.FromArgb(CType(CType(6, Byte), Integer), CType(CType(31, Byte), Integer), CType(CType(71, Byte), Integer))
        Me.txtPassword.Properties.Appearance.Font = New System.Drawing.Font("Segoe UI Light", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtPassword.Properties.Appearance.ForeColor = System.Drawing.Color.White
        Me.txtPassword.Properties.Appearance.Options.UseBackColor = True
        Me.txtPassword.Properties.Appearance.Options.UseFont = True
        Me.txtPassword.Properties.Appearance.Options.UseForeColor = True
        Me.txtPassword.Properties.Appearance.Options.UseTextOptions = True
        Me.txtPassword.Properties.Appearance.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center
        Me.txtPassword.Properties.BorderStyle = DevExpress.XtraEditors.Controls.BorderStyles.NoBorder
        Me.txtPassword.Properties.LookAndFeel.SkinName = "Visual Studio 2013 Blue"
        Me.txtPassword.Properties.LookAndFeel.UseDefaultLookAndFeel = False
        Me.txtPassword.Properties.NullText = "Password"
        Me.txtPassword.Properties.PasswordChar = Global.Microsoft.VisualBasic.ChrW(42)
        Me.txtPassword.Size = New System.Drawing.Size(408, 26)
        Me.txtPassword.TabIndex = 2
        '
        'btnSave
        '
        Me.btnSave.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnSave.Appearance.BackColor = System.Drawing.Color.FromArgb(CType(CType(7, Byte), Integer), CType(CType(46, Byte), Integer), CType(CType(110, Byte), Integer))
        Me.btnSave.Appearance.BackColor2 = System.Drawing.Color.FromArgb(CType(CType(7, Byte), Integer), CType(CType(46, Byte), Integer), CType(CType(110, Byte), Integer))
        Me.btnSave.Appearance.BorderColor = System.Drawing.Color.White
        Me.btnSave.Appearance.Font = New System.Drawing.Font("Segoe UI", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnSave.Appearance.ForeColor = System.Drawing.Color.White
        Me.btnSave.Appearance.Options.UseBackColor = True
        Me.btnSave.Appearance.Options.UseBorderColor = True
        Me.btnSave.Appearance.Options.UseFont = True
        Me.btnSave.Appearance.Options.UseForeColor = True
        Me.btnSave.ImageOptions.AllowGlyphSkinning = DevExpress.Utils.DefaultBoolean.[True]
        Me.btnSave.ImageOptions.Image = CType(resources.GetObject("btnSave.ImageOptions.Image"), System.Drawing.Image)
        Me.btnSave.ImageOptions.ImageToTextAlignment = DevExpress.XtraEditors.ImageAlignToText.LeftCenter
        Me.btnSave.Location = New System.Drawing.Point(20, 229)
        Me.btnSave.LookAndFeel.SkinName = "VS2010"
        Me.btnSave.LookAndFeel.UseDefaultLookAndFeel = False
        Me.btnSave.Margin = New System.Windows.Forms.Padding(2, 3, 2, 3)
        Me.btnSave.Name = "btnSave"
        Me.btnSave.Size = New System.Drawing.Size(408, 38)
        Me.btnSave.TabIndex = 5
        Me.btnSave.Text = "Login"
        '
        'btnClose
        '
        Me.btnClose.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnClose.Appearance.BackColor = System.Drawing.Color.FromArgb(CType(CType(6, Byte), Integer), CType(CType(31, Byte), Integer), CType(CType(71, Byte), Integer))
        Me.btnClose.Appearance.BackColor2 = System.Drawing.Color.FromArgb(CType(CType(6, Byte), Integer), CType(CType(31, Byte), Integer), CType(CType(71, Byte), Integer))
        Me.btnClose.Appearance.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnClose.Appearance.ForeColor = System.Drawing.Color.White
        Me.btnClose.Appearance.Options.UseBackColor = True
        Me.btnClose.Appearance.Options.UseFont = True
        Me.btnClose.Appearance.Options.UseForeColor = True
        Me.btnClose.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.btnClose.ImageOptions.AllowGlyphSkinning = DevExpress.Utils.DefaultBoolean.[True]
        Me.btnClose.ImageOptions.Image = CType(resources.GetObject("btnClose.ImageOptions.Image"), System.Drawing.Image)
        Me.btnClose.ImageOptions.ImageToTextAlignment = DevExpress.XtraEditors.ImageAlignToText.LeftCenter
        Me.btnClose.Location = New System.Drawing.Point(406, 4)
        Me.btnClose.LookAndFeel.SkinName = "Office 2010 Blue"
        Me.btnClose.LookAndFeel.Style = DevExpress.LookAndFeel.LookAndFeelStyle.UltraFlat
        Me.btnClose.LookAndFeel.UseDefaultLookAndFeel = False
        Me.btnClose.Margin = New System.Windows.Forms.Padding(2, 3, 2, 3)
        Me.btnClose.Name = "btnClose"
        Me.btnClose.Size = New System.Drawing.Size(38, 33)
        Me.btnClose.TabIndex = 8
        '
        'txtUsername
        '
        Me.txtUsername.Location = New System.Drawing.Point(20, 94)
        Me.txtUsername.Margin = New System.Windows.Forms.Padding(2, 3, 2, 3)
        Me.txtUsername.Name = "txtUsername"
        Me.txtUsername.Properties.Appearance.BackColor = System.Drawing.Color.FromArgb(CType(CType(6, Byte), Integer), CType(CType(31, Byte), Integer), CType(CType(71, Byte), Integer))
        Me.txtUsername.Properties.Appearance.Font = New System.Drawing.Font("Segoe UI Light", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtUsername.Properties.Appearance.ForeColor = System.Drawing.Color.White
        Me.txtUsername.Properties.Appearance.Options.UseBackColor = True
        Me.txtUsername.Properties.Appearance.Options.UseFont = True
        Me.txtUsername.Properties.Appearance.Options.UseForeColor = True
        Me.txtUsername.Properties.Appearance.Options.UseTextOptions = True
        Me.txtUsername.Properties.Appearance.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center
        Me.txtUsername.Properties.BorderStyle = DevExpress.XtraEditors.Controls.BorderStyles.NoBorder
        Me.txtUsername.Properties.LookAndFeel.SkinName = "Visual Studio 2013 Blue"
        Me.txtUsername.Properties.LookAndFeel.UseDefaultLookAndFeel = False
        Me.txtUsername.Properties.NullText = "Username"
        Me.txtUsername.Size = New System.Drawing.Size(408, 26)
        Me.txtUsername.TabIndex = 1
        '
        'chkNewWorklist
        '
        Me.chkNewWorklist.Location = New System.Drawing.Point(29, 200)
        Me.chkNewWorklist.Name = "chkNewWorklist"
        Me.chkNewWorklist.Properties.Appearance.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Bold)
        Me.chkNewWorklist.Properties.Appearance.ForeColor = System.Drawing.Color.White
        Me.chkNewWorklist.Properties.Appearance.Options.UseFont = True
        Me.chkNewWorklist.Properties.Appearance.Options.UseForeColor = True
        Me.chkNewWorklist.Properties.Caption = "New Worklist"
        Me.chkNewWorklist.Size = New System.Drawing.Size(106, 21)
        Me.chkNewWorklist.TabIndex = 4
        Me.chkNewWorklist.Visible = False
        '
        'cboLocation
        '
        Me.cboLocation.EditValue = "Laboratory"
        Me.cboLocation.Location = New System.Drawing.Point(20, 164)
        Me.cboLocation.Margin = New System.Windows.Forms.Padding(2, 3, 2, 3)
        Me.cboLocation.Name = "cboLocation"
        Me.cboLocation.Properties.Appearance.BackColor = System.Drawing.Color.FromArgb(CType(CType(6, Byte), Integer), CType(CType(31, Byte), Integer), CType(CType(71, Byte), Integer))
        Me.cboLocation.Properties.Appearance.Font = New System.Drawing.Font("Segoe UI Light", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cboLocation.Properties.Appearance.ForeColor = System.Drawing.Color.White
        Me.cboLocation.Properties.Appearance.Options.UseBackColor = True
        Me.cboLocation.Properties.Appearance.Options.UseFont = True
        Me.cboLocation.Properties.Appearance.Options.UseForeColor = True
        Me.cboLocation.Properties.Appearance.Options.UseTextOptions = True
        Me.cboLocation.Properties.Appearance.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center
        Me.cboLocation.Properties.BorderStyle = DevExpress.XtraEditors.Controls.BorderStyles.NoBorder
        Me.cboLocation.Properties.Buttons.AddRange(New DevExpress.XtraEditors.Controls.EditorButton() {New DevExpress.XtraEditors.Controls.EditorButton(DevExpress.XtraEditors.Controls.ButtonPredefines.Combo)})
        Me.cboLocation.Properties.LookAndFeel.SkinName = "Visual Studio 2013 Blue"
        Me.cboLocation.Properties.LookAndFeel.UseDefaultLookAndFeel = False
        Me.cboLocation.Properties.NullText = "Location"
        Me.cboLocation.Properties.TextEditStyle = DevExpress.XtraEditors.Controls.TextEditStyles.DisableTextEditor
        Me.cboLocation.Size = New System.Drawing.Size(424, 26)
        Me.cboLocation.TabIndex = 3
        '
        'SeparatorControl2
        '
        Me.SeparatorControl2.AutoSizeMode = True
        Me.SeparatorControl2.Location = New System.Drawing.Point(20, 112)
        Me.SeparatorControl2.Margin = New System.Windows.Forms.Padding(2, 3, 2, 3)
        Me.SeparatorControl2.Name = "SeparatorControl2"
        Me.SeparatorControl2.Padding = New System.Windows.Forms.Padding(8, 9, 8, 9)
        Me.SeparatorControl2.Size = New System.Drawing.Size(408, 19)
        Me.SeparatorControl2.TabIndex = 104
        '
        'SeparatorControl3
        '
        Me.SeparatorControl3.AutoSizeMode = True
        Me.SeparatorControl3.Location = New System.Drawing.Point(20, 146)
        Me.SeparatorControl3.Margin = New System.Windows.Forms.Padding(2, 3, 2, 3)
        Me.SeparatorControl3.Name = "SeparatorControl3"
        Me.SeparatorControl3.Padding = New System.Windows.Forms.Padding(8, 9, 8, 9)
        Me.SeparatorControl3.Size = New System.Drawing.Size(408, 19)
        Me.SeparatorControl3.TabIndex = 105
        '
        'SeparatorControl4
        '
        Me.SeparatorControl4.AutoSizeMode = True
        Me.SeparatorControl4.Location = New System.Drawing.Point(20, 182)
        Me.SeparatorControl4.Margin = New System.Windows.Forms.Padding(2, 3, 2, 3)
        Me.SeparatorControl4.Name = "SeparatorControl4"
        Me.SeparatorControl4.Padding = New System.Windows.Forms.Padding(8, 9, 8, 9)
        Me.SeparatorControl4.Size = New System.Drawing.Size(408, 19)
        Me.SeparatorControl4.TabIndex = 106
        '
        'LabelControl1
        '
        Me.LabelControl1.Appearance.Font = New System.Drawing.Font("Segoe UI Semibold", 18.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.LabelControl1.Appearance.ForeColor = System.Drawing.Color.White
        Me.LabelControl1.Appearance.Options.UseFont = True
        Me.LabelControl1.Appearance.Options.UseForeColor = True
        Me.LabelControl1.Location = New System.Drawing.Point(96, 38)
        Me.LabelControl1.Name = "LabelControl1"
        Me.LabelControl1.Size = New System.Drawing.Size(265, 32)
        Me.LabelControl1.TabIndex = 107
        Me.LabelControl1.Text = "USER AUTHENTICATION"
        '
        'lblDB
        '
        Me.lblDB.Appearance.Font = New System.Drawing.Font("Segoe UI", 12.0!, System.Drawing.FontStyle.Underline, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblDB.Appearance.ForeColor = System.Drawing.Color.White
        Me.lblDB.Appearance.Options.UseFont = True
        Me.lblDB.Appearance.Options.UseForeColor = True
        Me.lblDB.Appearance.Options.UseTextOptions = True
        Me.lblDB.Appearance.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center
        Me.lblDB.AutoSizeMode = DevExpress.XtraEditors.LabelAutoSizeMode.None
        Me.lblDB.Cursor = System.Windows.Forms.Cursors.Hand
        Me.lblDB.Location = New System.Drawing.Point(20, 271)
        Me.lblDB.Name = "lblDB"
        Me.lblDB.Size = New System.Drawing.Size(408, 32)
        Me.lblDB.TabIndex = 6
        Me.lblDB.Text = "Configure Database"
        '
        'frmLogin
        '
        Me.AcceptButton = Me.btnSave
        Me.Appearance.BackColor = System.Drawing.Color.FromArgb(CType(CType(6, Byte), Integer), CType(CType(31, Byte), Integer), CType(CType(71, Byte), Integer))
        Me.Appearance.BackColor2 = System.Drawing.Color.FromArgb(CType(CType(7, Byte), Integer), CType(CType(46, Byte), Integer), CType(CType(110, Byte), Integer))
        Me.Appearance.Options.UseBackColor = True
        Me.AutoScaleDimensions = New System.Drawing.SizeF(96.0!, 96.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Dpi
        Me.CancelButton = Me.btnClose
        Me.ClientSize = New System.Drawing.Size(446, 315)
        Me.Controls.Add(Me.lblDB)
        Me.Controls.Add(Me.LabelControl1)
        Me.Controls.Add(Me.SeparatorControl3)
        Me.Controls.Add(Me.chkNewWorklist)
        Me.Controls.Add(Me.txtUsername)
        Me.Controls.Add(Me.btnClose)
        Me.Controls.Add(Me.txtPassword)
        Me.Controls.Add(Me.btnSave)
        Me.Controls.Add(Me.cboLocation)
        Me.Controls.Add(Me.SeparatorControl2)
        Me.Controls.Add(Me.SeparatorControl4)
        Me.FormBorderEffect = DevExpress.XtraEditors.FormBorderEffect.Glow
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Margin = New System.Windows.Forms.Padding(2, 3, 2, 3)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "frmLogin"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Authentication Window"
        CType(Me.txtPassword.Properties, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.txtUsername.Properties, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.chkNewWorklist.Properties, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.cboLocation.Properties, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.SeparatorControl2, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.SeparatorControl3, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.SeparatorControl4, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents txtUsername As DevExpress.XtraEditors.TextEdit
    Friend WithEvents btnSave As DevExpress.XtraEditors.SimpleButton
    Friend WithEvents btnClose As DevExpress.XtraEditors.SimpleButton
    Friend WithEvents txtPassword As DevExpress.XtraEditors.TextEdit
    Friend WithEvents chkNewWorklist As DevExpress.XtraEditors.CheckEdit
    Friend WithEvents cboLocation As DevExpress.XtraEditors.ComboBoxEdit
    Friend WithEvents SeparatorControl2 As DevExpress.XtraEditors.SeparatorControl
    Friend WithEvents SeparatorControl3 As DevExpress.XtraEditors.SeparatorControl
    Friend WithEvents SeparatorControl4 As DevExpress.XtraEditors.SeparatorControl
    Friend WithEvents LabelControl1 As DevExpress.XtraEditors.LabelControl
    Friend WithEvents lblDB As DevExpress.XtraEditors.LabelControl
End Class
