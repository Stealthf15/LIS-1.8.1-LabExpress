<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmSpecimenAE
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmSpecimenAE))
        Me.LabelControl1 = New DevExpress.XtraEditors.LabelControl()
        Me.LabelControl6 = New DevExpress.XtraEditors.LabelControl()
        Me.txtTestCode = New DevExpress.XtraEditors.TextEdit()
        Me.cboTestName = New DevExpress.XtraEditors.ComboBoxEdit()
        Me.LabelControl4 = New DevExpress.XtraEditors.LabelControl()
        Me.btnClose = New DevExpress.XtraEditors.SimpleButton()
        Me.btnSave = New DevExpress.XtraEditors.SimpleButton()
        Me.LabelControl3 = New DevExpress.XtraEditors.LabelControl()
        Me.cboUnit = New DevExpress.XtraEditors.ComboBoxEdit()
        Me.cboTest = New DevExpress.XtraEditors.ComboBoxEdit()
        Me.LabelControl9 = New DevExpress.XtraEditors.LabelControl()
        Me.cboConv_Unit = New DevExpress.XtraEditors.ComboBoxEdit()
        Me.LabelControl11 = New DevExpress.XtraEditors.LabelControl()
        Me.cboSpecimen = New DevExpress.XtraEditors.ComboBoxEdit()
        Me.LabelControl7 = New DevExpress.XtraEditors.LabelControl()
        Me.txtOrderNo = New DevExpress.XtraEditors.TextEdit()
        Me.txtHISField = New DevExpress.XtraEditors.TextEdit()
        Me.LabelControl5 = New DevExpress.XtraEditors.LabelControl()
        Me.txtHISCode = New DevExpress.XtraEditors.TextEdit()
        Me.LabelControl13 = New DevExpress.XtraEditors.LabelControl()
        Me.LabelControl8 = New DevExpress.XtraEditors.LabelControl()
        Me.cboTestGroup = New DevExpress.XtraEditors.ComboBoxEdit()
        CType(Me.txtTestCode.Properties, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.cboTestName.Properties, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.cboUnit.Properties, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.cboTest.Properties, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.cboConv_Unit.Properties, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.cboSpecimen.Properties, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.txtOrderNo.Properties, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.txtHISField.Properties, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.txtHISCode.Properties, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.cboTestGroup.Properties, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'LabelControl1
        '
        Me.LabelControl1.Appearance.ForeColor = System.Drawing.Color.FromArgb(CType(CType(27, Byte), Integer), CType(CType(41, Byte), Integer), CType(CType(62, Byte), Integer))
        Me.LabelControl1.Appearance.Options.UseForeColor = True
        Me.LabelControl1.Location = New System.Drawing.Point(19, 44)
        Me.LabelControl1.LookAndFeel.SkinName = "Visual Studio 2013 Blue"
        Me.LabelControl1.LookAndFeel.UseDefaultLookAndFeel = False
        Me.LabelControl1.Name = "LabelControl1"
        Me.LabelControl1.Size = New System.Drawing.Size(56, 13)
        Me.LabelControl1.TabIndex = 12
        Me.LabelControl1.Text = "Test Name:"
        '
        'LabelControl6
        '
        Me.LabelControl6.Appearance.ForeColor = System.Drawing.Color.FromArgb(CType(CType(27, Byte), Integer), CType(CType(41, Byte), Integer), CType(CType(62, Byte), Integer))
        Me.LabelControl6.Appearance.Options.UseForeColor = True
        Me.LabelControl6.Location = New System.Drawing.Point(19, 148)
        Me.LabelControl6.LookAndFeel.SkinName = "Visual Studio 2013 Blue"
        Me.LabelControl6.LookAndFeel.UseDefaultLookAndFeel = False
        Me.LabelControl6.Name = "LabelControl6"
        Me.LabelControl6.Size = New System.Drawing.Size(54, 13)
        Me.LabelControl6.TabIndex = 86
        Me.LabelControl6.Text = "Test Code:"
        '
        'txtTestCode
        '
        Me.txtTestCode.Location = New System.Drawing.Point(136, 142)
        Me.txtTestCode.Name = "txtTestCode"
        Me.txtTestCode.Properties.Appearance.ForeColor = System.Drawing.Color.FromArgb(CType(CType(27, Byte), Integer), CType(CType(41, Byte), Integer), CType(CType(62, Byte), Integer))
        Me.txtTestCode.Properties.Appearance.Options.UseForeColor = True
        Me.txtTestCode.Properties.LookAndFeel.SkinName = "Visual Studio 2013 Blue"
        Me.txtTestCode.Properties.LookAndFeel.UseDefaultLookAndFeel = False
        Me.txtTestCode.Size = New System.Drawing.Size(316, 20)
        Me.txtTestCode.TabIndex = 10
        '
        'cboTestName
        '
        Me.cboTestName.Enabled = False
        Me.cboTestName.Location = New System.Drawing.Point(136, 12)
        Me.cboTestName.Name = "cboTestName"
        Me.cboTestName.Properties.Appearance.ForeColor = System.Drawing.Color.FromArgb(CType(CType(27, Byte), Integer), CType(CType(41, Byte), Integer), CType(CType(62, Byte), Integer))
        Me.cboTestName.Properties.Appearance.Options.UseForeColor = True
        Me.cboTestName.Properties.Buttons.AddRange(New DevExpress.XtraEditors.Controls.EditorButton() {New DevExpress.XtraEditors.Controls.EditorButton(DevExpress.XtraEditors.Controls.ButtonPredefines.Combo)})
        Me.cboTestName.Properties.LookAndFeel.SkinName = "Visual Studio 2013 Blue"
        Me.cboTestName.Properties.LookAndFeel.UseDefaultLookAndFeel = False
        Me.cboTestName.Size = New System.Drawing.Size(316, 20)
        Me.cboTestName.TabIndex = 0
        '
        'LabelControl4
        '
        Me.LabelControl4.Appearance.ForeColor = System.Drawing.Color.FromArgb(CType(CType(27, Byte), Integer), CType(CType(41, Byte), Integer), CType(CType(62, Byte), Integer))
        Me.LabelControl4.Appearance.Options.UseForeColor = True
        Me.LabelControl4.Location = New System.Drawing.Point(19, 18)
        Me.LabelControl4.LookAndFeel.SkinName = "Visual Studio 2013 Blue"
        Me.LabelControl4.LookAndFeel.UseDefaultLookAndFeel = False
        Me.LabelControl4.Name = "LabelControl4"
        Me.LabelControl4.Size = New System.Drawing.Size(41, 13)
        Me.LabelControl4.TabIndex = 80
        Me.LabelControl4.Text = "Section:"
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
        Me.btnClose.Location = New System.Drawing.Point(345, 276)
        Me.btnClose.Name = "btnClose"
        Me.btnClose.Size = New System.Drawing.Size(108, 27)
        Me.btnClose.TabIndex = 102
        Me.btnClose.Text = "&Close"
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
        Me.btnSave.Location = New System.Drawing.Point(231, 276)
        Me.btnSave.Name = "btnSave"
        Me.btnSave.Size = New System.Drawing.Size(108, 27)
        Me.btnSave.TabIndex = 101
        Me.btnSave.Text = "&Save"
        '
        'LabelControl3
        '
        Me.LabelControl3.Appearance.ForeColor = System.Drawing.Color.FromArgb(CType(CType(27, Byte), Integer), CType(CType(41, Byte), Integer), CType(CType(62, Byte), Integer))
        Me.LabelControl3.Appearance.Options.UseForeColor = True
        Me.LabelControl3.Location = New System.Drawing.Point(19, 96)
        Me.LabelControl3.LookAndFeel.SkinName = "Visual Studio 2013 Blue"
        Me.LabelControl3.LookAndFeel.UseDefaultLookAndFeel = False
        Me.LabelControl3.Name = "LabelControl3"
        Me.LabelControl3.Size = New System.Drawing.Size(37, 13)
        Me.LabelControl3.TabIndex = 104
        Me.LabelControl3.Text = "SI Unit:"
        '
        'cboUnit
        '
        Me.cboUnit.Location = New System.Drawing.Point(136, 90)
        Me.cboUnit.Name = "cboUnit"
        Me.cboUnit.Properties.Appearance.ForeColor = System.Drawing.Color.FromArgb(CType(CType(27, Byte), Integer), CType(CType(41, Byte), Integer), CType(CType(62, Byte), Integer))
        Me.cboUnit.Properties.Appearance.Options.UseForeColor = True
        Me.cboUnit.Properties.Buttons.AddRange(New DevExpress.XtraEditors.Controls.EditorButton() {New DevExpress.XtraEditors.Controls.EditorButton(DevExpress.XtraEditors.Controls.ButtonPredefines.Combo)})
        Me.cboUnit.Properties.LookAndFeel.SkinName = "Visual Studio 2013 Blue"
        Me.cboUnit.Properties.LookAndFeel.UseDefaultLookAndFeel = False
        Me.cboUnit.Size = New System.Drawing.Size(316, 20)
        Me.cboUnit.TabIndex = 5
        '
        'cboTest
        '
        Me.cboTest.Location = New System.Drawing.Point(136, 38)
        Me.cboTest.Name = "cboTest"
        Me.cboTest.Properties.Appearance.ForeColor = System.Drawing.Color.FromArgb(CType(CType(27, Byte), Integer), CType(CType(41, Byte), Integer), CType(CType(62, Byte), Integer))
        Me.cboTest.Properties.Appearance.Options.UseForeColor = True
        Me.cboTest.Properties.Buttons.AddRange(New DevExpress.XtraEditors.Controls.EditorButton() {New DevExpress.XtraEditors.Controls.EditorButton(DevExpress.XtraEditors.Controls.ButtonPredefines.Combo)})
        Me.cboTest.Properties.LookAndFeel.SkinName = "Visual Studio 2013 Blue"
        Me.cboTest.Properties.LookAndFeel.UseDefaultLookAndFeel = False
        Me.cboTest.Size = New System.Drawing.Size(316, 20)
        Me.cboTest.TabIndex = 2
        '
        'LabelControl9
        '
        Me.LabelControl9.Appearance.ForeColor = System.Drawing.Color.FromArgb(CType(CType(27, Byte), Integer), CType(CType(41, Byte), Integer), CType(CType(62, Byte), Integer))
        Me.LabelControl9.Appearance.Options.UseForeColor = True
        Me.LabelControl9.Location = New System.Drawing.Point(19, 122)
        Me.LabelControl9.LookAndFeel.SkinName = "Visual Studio 2013 Blue"
        Me.LabelControl9.LookAndFeel.UseDefaultLookAndFeel = False
        Me.LabelControl9.Name = "LabelControl9"
        Me.LabelControl9.Size = New System.Drawing.Size(97, 13)
        Me.LabelControl9.TabIndex = 117
        Me.LabelControl9.Text = "Conventional Unit:"
        '
        'cboConv_Unit
        '
        Me.cboConv_Unit.Location = New System.Drawing.Point(136, 116)
        Me.cboConv_Unit.Name = "cboConv_Unit"
        Me.cboConv_Unit.Properties.Appearance.ForeColor = System.Drawing.Color.FromArgb(CType(CType(27, Byte), Integer), CType(CType(41, Byte), Integer), CType(CType(62, Byte), Integer))
        Me.cboConv_Unit.Properties.Appearance.Options.UseForeColor = True
        Me.cboConv_Unit.Properties.Buttons.AddRange(New DevExpress.XtraEditors.Controls.EditorButton() {New DevExpress.XtraEditors.Controls.EditorButton(DevExpress.XtraEditors.Controls.ButtonPredefines.Combo)})
        Me.cboConv_Unit.Properties.LookAndFeel.SkinName = "Visual Studio 2013 Blue"
        Me.cboConv_Unit.Properties.LookAndFeel.UseDefaultLookAndFeel = False
        Me.cboConv_Unit.Size = New System.Drawing.Size(316, 20)
        Me.cboConv_Unit.TabIndex = 6
        '
        'LabelControl11
        '
        Me.LabelControl11.Appearance.ForeColor = System.Drawing.Color.FromArgb(CType(CType(27, Byte), Integer), CType(CType(41, Byte), Integer), CType(CType(62, Byte), Integer))
        Me.LabelControl11.Appearance.Options.UseForeColor = True
        Me.LabelControl11.Location = New System.Drawing.Point(19, 71)
        Me.LabelControl11.LookAndFeel.SkinName = "Visual Studio 2013 Blue"
        Me.LabelControl11.LookAndFeel.UseDefaultLookAndFeel = False
        Me.LabelControl11.Name = "LabelControl11"
        Me.LabelControl11.Size = New System.Drawing.Size(79, 13)
        Me.LabelControl11.TabIndex = 114
        Me.LabelControl11.Text = "Specimen Type:"
        '
        'cboSpecimen
        '
        Me.cboSpecimen.Location = New System.Drawing.Point(136, 64)
        Me.cboSpecimen.Name = "cboSpecimen"
        Me.cboSpecimen.Properties.Appearance.ForeColor = System.Drawing.Color.FromArgb(CType(CType(27, Byte), Integer), CType(CType(41, Byte), Integer), CType(CType(62, Byte), Integer))
        Me.cboSpecimen.Properties.Appearance.Options.UseForeColor = True
        Me.cboSpecimen.Properties.Buttons.AddRange(New DevExpress.XtraEditors.Controls.EditorButton() {New DevExpress.XtraEditors.Controls.EditorButton(DevExpress.XtraEditors.Controls.ButtonPredefines.Combo)})
        Me.cboSpecimen.Properties.LookAndFeel.SkinName = "Visual Studio 2013 Blue"
        Me.cboSpecimen.Properties.LookAndFeel.UseDefaultLookAndFeel = False
        Me.cboSpecimen.Size = New System.Drawing.Size(316, 20)
        Me.cboSpecimen.TabIndex = 4
        '
        'LabelControl7
        '
        Me.LabelControl7.Appearance.ForeColor = System.Drawing.Color.FromArgb(CType(CType(27, Byte), Integer), CType(CType(41, Byte), Integer), CType(CType(62, Byte), Integer))
        Me.LabelControl7.Appearance.Options.UseForeColor = True
        Me.LabelControl7.Location = New System.Drawing.Point(19, 226)
        Me.LabelControl7.LookAndFeel.SkinName = "Visual Studio 2013 Blue"
        Me.LabelControl7.LookAndFeel.UseDefaultLookAndFeel = False
        Me.LabelControl7.Name = "LabelControl7"
        Me.LabelControl7.Size = New System.Drawing.Size(58, 13)
        Me.LabelControl7.TabIndex = 119
        Me.LabelControl7.Text = "Display No:"
        '
        'txtOrderNo
        '
        Me.txtOrderNo.Location = New System.Drawing.Point(136, 220)
        Me.txtOrderNo.Name = "txtOrderNo"
        Me.txtOrderNo.Properties.Appearance.ForeColor = System.Drawing.Color.FromArgb(CType(CType(27, Byte), Integer), CType(CType(41, Byte), Integer), CType(CType(62, Byte), Integer))
        Me.txtOrderNo.Properties.Appearance.Options.UseForeColor = True
        Me.txtOrderNo.Properties.LookAndFeel.SkinName = "Visual Studio 2013 Blue"
        Me.txtOrderNo.Properties.LookAndFeel.UseDefaultLookAndFeel = False
        Me.txtOrderNo.Size = New System.Drawing.Size(316, 20)
        Me.txtOrderNo.TabIndex = 11
        '
        'txtHISField
        '
        Me.txtHISField.Location = New System.Drawing.Point(136, 194)
        Me.txtHISField.Name = "txtHISField"
        Me.txtHISField.Properties.Appearance.ForeColor = System.Drawing.Color.FromArgb(CType(CType(27, Byte), Integer), CType(CType(41, Byte), Integer), CType(CType(62, Byte), Integer))
        Me.txtHISField.Properties.Appearance.Options.UseForeColor = True
        Me.txtHISField.Properties.LookAndFeel.SkinName = "Visual Studio 2013 Blue"
        Me.txtHISField.Properties.LookAndFeel.UseDefaultLookAndFeel = False
        Me.txtHISField.Size = New System.Drawing.Size(316, 20)
        Me.txtHISField.TabIndex = 127
        '
        'LabelControl5
        '
        Me.LabelControl5.Appearance.ForeColor = System.Drawing.Color.FromArgb(CType(CType(27, Byte), Integer), CType(CType(41, Byte), Integer), CType(CType(62, Byte), Integer))
        Me.LabelControl5.Appearance.Options.UseForeColor = True
        Me.LabelControl5.Location = New System.Drawing.Point(19, 200)
        Me.LabelControl5.LookAndFeel.SkinName = "Visual Studio 2013 Blue"
        Me.LabelControl5.LookAndFeel.UseDefaultLookAndFeel = False
        Me.LabelControl5.Name = "LabelControl5"
        Me.LabelControl5.Size = New System.Drawing.Size(80, 13)
        Me.LabelControl5.TabIndex = 128
        Me.LabelControl5.Text = "HIS Field Name:"
        '
        'txtHISCode
        '
        Me.txtHISCode.Enabled = False
        Me.txtHISCode.Location = New System.Drawing.Point(136, 168)
        Me.txtHISCode.Name = "txtHISCode"
        Me.txtHISCode.Properties.Appearance.ForeColor = System.Drawing.Color.FromArgb(CType(CType(27, Byte), Integer), CType(CType(41, Byte), Integer), CType(CType(62, Byte), Integer))
        Me.txtHISCode.Properties.Appearance.Options.UseForeColor = True
        Me.txtHISCode.Properties.LookAndFeel.SkinName = "Visual Studio 2013 Blue"
        Me.txtHISCode.Properties.LookAndFeel.UseDefaultLookAndFeel = False
        Me.txtHISCode.Size = New System.Drawing.Size(316, 20)
        Me.txtHISCode.TabIndex = 125
        '
        'LabelControl13
        '
        Me.LabelControl13.Appearance.ForeColor = System.Drawing.Color.FromArgb(CType(CType(27, Byte), Integer), CType(CType(41, Byte), Integer), CType(CType(62, Byte), Integer))
        Me.LabelControl13.Appearance.Options.UseForeColor = True
        Me.LabelControl13.Location = New System.Drawing.Point(19, 174)
        Me.LabelControl13.LookAndFeel.SkinName = "Visual Studio 2013 Blue"
        Me.LabelControl13.LookAndFeel.UseDefaultLookAndFeel = False
        Me.LabelControl13.Name = "LabelControl13"
        Me.LabelControl13.Size = New System.Drawing.Size(50, 13)
        Me.LabelControl13.TabIndex = 126
        Me.LabelControl13.Text = "HIS Code:"
        '
        'LabelControl8
        '
        Me.LabelControl8.Appearance.ForeColor = System.Drawing.Color.FromArgb(CType(CType(27, Byte), Integer), CType(CType(41, Byte), Integer), CType(CType(62, Byte), Integer))
        Me.LabelControl8.Appearance.Options.UseForeColor = True
        Me.LabelControl8.Location = New System.Drawing.Point(19, 254)
        Me.LabelControl8.LookAndFeel.SkinName = "Visual Studio 2013 Blue"
        Me.LabelControl8.LookAndFeel.UseDefaultLookAndFeel = False
        Me.LabelControl8.Name = "LabelControl8"
        Me.LabelControl8.Size = New System.Drawing.Size(60, 13)
        Me.LabelControl8.TabIndex = 121
        Me.LabelControl8.Text = "Test Group:"
        '
        'cboTestGroup
        '
        Me.cboTestGroup.Location = New System.Drawing.Point(136, 246)
        Me.cboTestGroup.Name = "cboTestGroup"
        Me.cboTestGroup.Properties.Appearance.ForeColor = System.Drawing.Color.FromArgb(CType(CType(27, Byte), Integer), CType(CType(41, Byte), Integer), CType(CType(62, Byte), Integer))
        Me.cboTestGroup.Properties.Appearance.Options.UseForeColor = True
        Me.cboTestGroup.Properties.Buttons.AddRange(New DevExpress.XtraEditors.Controls.EditorButton() {New DevExpress.XtraEditors.Controls.EditorButton(DevExpress.XtraEditors.Controls.ButtonPredefines.Combo)})
        Me.cboTestGroup.Properties.Items.AddRange(New Object() {"Complete Blood Count", "Differential Count", "Absolute Differential Count", "Blood Typing", "Lipid Profile", "Bilirubin", "Enzymes", "Other Test"})
        Me.cboTestGroup.Properties.LookAndFeel.SkinName = "Visual Studio 2013 Blue"
        Me.cboTestGroup.Properties.LookAndFeel.UseDefaultLookAndFeel = False
        Me.cboTestGroup.Size = New System.Drawing.Size(316, 20)
        Me.cboTestGroup.TabIndex = 12
        '
        'frmSpecimenAE
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(465, 315)
        Me.Controls.Add(Me.txtHISField)
        Me.Controls.Add(Me.LabelControl5)
        Me.Controls.Add(Me.btnClose)
        Me.Controls.Add(Me.txtHISCode)
        Me.Controls.Add(Me.btnSave)
        Me.Controls.Add(Me.LabelControl13)
        Me.Controls.Add(Me.cboTestName)
        Me.Controls.Add(Me.LabelControl8)
        Me.Controls.Add(Me.cboTestGroup)
        Me.Controls.Add(Me.cboUnit)
        Me.Controls.Add(Me.LabelControl7)
        Me.Controls.Add(Me.LabelControl6)
        Me.Controls.Add(Me.cboSpecimen)
        Me.Controls.Add(Me.LabelControl3)
        Me.Controls.Add(Me.txtOrderNo)
        Me.Controls.Add(Me.LabelControl1)
        Me.Controls.Add(Me.cboTest)
        Me.Controls.Add(Me.LabelControl9)
        Me.Controls.Add(Me.txtTestCode)
        Me.Controls.Add(Me.LabelControl11)
        Me.Controls.Add(Me.cboConv_Unit)
        Me.Controls.Add(Me.LabelControl4)
        Me.FormBorderEffect = DevExpress.XtraEditors.FormBorderEffect.Glow
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "frmSpecimenAE"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Specimen Details"
        CType(Me.txtTestCode.Properties, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.cboTestName.Properties, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.cboUnit.Properties, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.cboTest.Properties, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.cboConv_Unit.Properties, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.cboSpecimen.Properties, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.txtOrderNo.Properties, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.txtHISField.Properties, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.txtHISCode.Properties, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.cboTestGroup.Properties, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents LabelControl1 As DevExpress.XtraEditors.LabelControl
    Friend WithEvents LabelControl4 As DevExpress.XtraEditors.LabelControl
    Friend WithEvents cboTestName As DevExpress.XtraEditors.ComboBoxEdit
    Friend WithEvents LabelControl6 As DevExpress.XtraEditors.LabelControl
    Friend WithEvents txtTestCode As DevExpress.XtraEditors.TextEdit
    Friend WithEvents btnClose As DevExpress.XtraEditors.SimpleButton
    Friend WithEvents btnSave As DevExpress.XtraEditors.SimpleButton
    Friend WithEvents LabelControl3 As DevExpress.XtraEditors.LabelControl
    Friend WithEvents cboUnit As DevExpress.XtraEditors.ComboBoxEdit
    Friend WithEvents cboTest As DevExpress.XtraEditors.ComboBoxEdit
    Friend WithEvents LabelControl9 As DevExpress.XtraEditors.LabelControl
    Friend WithEvents cboConv_Unit As DevExpress.XtraEditors.ComboBoxEdit
    Friend WithEvents LabelControl11 As DevExpress.XtraEditors.LabelControl
    Friend WithEvents cboSpecimen As DevExpress.XtraEditors.ComboBoxEdit
    Friend WithEvents LabelControl7 As DevExpress.XtraEditors.LabelControl
    Friend WithEvents txtOrderNo As DevExpress.XtraEditors.TextEdit
    Friend WithEvents LabelControl8 As DevExpress.XtraEditors.LabelControl
    Friend WithEvents cboTestGroup As DevExpress.XtraEditors.ComboBoxEdit
    Friend WithEvents txtHISCode As DevExpress.XtraEditors.TextEdit
    Friend WithEvents LabelControl13 As DevExpress.XtraEditors.LabelControl
    Friend WithEvents txtHISField As DevExpress.XtraEditors.TextEdit
    Friend WithEvents LabelControl5 As DevExpress.XtraEditors.LabelControl
End Class
