<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmAccountA
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmAccountA))
        Me.LabelControl1 = New DevExpress.XtraEditors.LabelControl()
        Me.txtUsername = New DevExpress.XtraEditors.TextEdit()
        Me.LabelControl4 = New DevExpress.XtraEditors.LabelControl()
        Me.LabelControl2 = New DevExpress.XtraEditors.LabelControl()
        Me.LabelControl5 = New DevExpress.XtraEditors.LabelControl()
        Me.txtPassword = New DevExpress.XtraEditors.TextEdit()
        Me.cboUserType = New DevExpress.XtraEditors.ComboBoxEdit()
        Me.btnSave = New DevExpress.XtraEditors.SimpleButton()
        Me.btnClose = New DevExpress.XtraEditors.SimpleButton()
        Me.cboMedTech = New DevExpress.XtraEditors.ComboBoxEdit()
        Me.cboLocation = New DevExpress.XtraEditors.ComboBoxEdit()
        Me.LabelControl3 = New DevExpress.XtraEditors.LabelControl()
        Me.TabPane1 = New DevExpress.XtraBars.Navigation.TabPane()
        Me.TabNavigationPage1 = New DevExpress.XtraBars.Navigation.TabNavigationPage()
        Me.TabNavigationPage2 = New DevExpress.XtraBars.Navigation.TabNavigationPage()
        Me.dtResult = New DevExpress.XtraGrid.GridControl()
        Me.GridView = New DevExpress.XtraGrid.Views.Grid.GridView()
        Me.TabNavigationPage3 = New DevExpress.XtraBars.Navigation.TabNavigationPage()
        Me.dtPermissions = New DevExpress.XtraGrid.GridControl()
        Me.GridPermissions = New DevExpress.XtraGrid.Views.Grid.GridView()
        CType(Me.txtUsername.Properties, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.txtPassword.Properties, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.cboUserType.Properties, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.cboMedTech.Properties, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.cboLocation.Properties, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.TabPane1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.TabPane1.SuspendLayout()
        Me.TabNavigationPage1.SuspendLayout()
        Me.TabNavigationPage2.SuspendLayout()
        CType(Me.dtResult, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.GridView, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.TabNavigationPage3.SuspendLayout()
        CType(Me.dtPermissions, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.GridPermissions, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'LabelControl1
        '
        Me.LabelControl1.Appearance.ForeColor = System.Drawing.Color.FromArgb(CType(CType(27, Byte), Integer), CType(CType(41, Byte), Integer), CType(CType(62, Byte), Integer))
        Me.LabelControl1.Appearance.Options.UseForeColor = True
        Me.LabelControl1.Location = New System.Drawing.Point(33, 56)
        Me.LabelControl1.LookAndFeel.SkinName = "Visual Studio 2013 Blue"
        Me.LabelControl1.LookAndFeel.UseDefaultLookAndFeel = False
        Me.LabelControl1.Name = "LabelControl1"
        Me.LabelControl1.Size = New System.Drawing.Size(32, 13)
        Me.LabelControl1.TabIndex = 12
        Me.LabelControl1.Text = "Name:"
        '
        'txtUsername
        '
        Me.txtUsername.Location = New System.Drawing.Point(120, 79)
        Me.txtUsername.Name = "txtUsername"
        Me.txtUsername.Properties.Appearance.ForeColor = System.Drawing.Color.FromArgb(CType(CType(27, Byte), Integer), CType(CType(41, Byte), Integer), CType(CType(62, Byte), Integer))
        Me.txtUsername.Properties.Appearance.Options.UseForeColor = True
        Me.txtUsername.Properties.LookAndFeel.SkinName = "Visual Studio 2013 Blue"
        Me.txtUsername.Properties.LookAndFeel.UseDefaultLookAndFeel = False
        Me.txtUsername.Size = New System.Drawing.Size(281, 20)
        Me.txtUsername.TabIndex = 2
        '
        'LabelControl4
        '
        Me.LabelControl4.Appearance.ForeColor = System.Drawing.Color.FromArgb(CType(CType(27, Byte), Integer), CType(CType(41, Byte), Integer), CType(CType(62, Byte), Integer))
        Me.LabelControl4.Appearance.Options.UseForeColor = True
        Me.LabelControl4.Location = New System.Drawing.Point(33, 82)
        Me.LabelControl4.LookAndFeel.SkinName = "Visual Studio 2013 Blue"
        Me.LabelControl4.LookAndFeel.UseDefaultLookAndFeel = False
        Me.LabelControl4.Name = "LabelControl4"
        Me.LabelControl4.Size = New System.Drawing.Size(54, 13)
        Me.LabelControl4.TabIndex = 80
        Me.LabelControl4.Text = "Username:"
        '
        'LabelControl2
        '
        Me.LabelControl2.Appearance.ForeColor = System.Drawing.Color.FromArgb(CType(CType(27, Byte), Integer), CType(CType(41, Byte), Integer), CType(CType(62, Byte), Integer))
        Me.LabelControl2.Appearance.Options.UseForeColor = True
        Me.LabelControl2.Location = New System.Drawing.Point(33, 108)
        Me.LabelControl2.LookAndFeel.SkinName = "Visual Studio 2013 Blue"
        Me.LabelControl2.LookAndFeel.UseDefaultLookAndFeel = False
        Me.LabelControl2.Name = "LabelControl2"
        Me.LabelControl2.Size = New System.Drawing.Size(52, 13)
        Me.LabelControl2.TabIndex = 78
        Me.LabelControl2.Text = "Password:"
        '
        'LabelControl5
        '
        Me.LabelControl5.Appearance.ForeColor = System.Drawing.Color.FromArgb(CType(CType(27, Byte), Integer), CType(CType(41, Byte), Integer), CType(CType(62, Byte), Integer))
        Me.LabelControl5.Appearance.Options.UseForeColor = True
        Me.LabelControl5.Location = New System.Drawing.Point(33, 30)
        Me.LabelControl5.LookAndFeel.SkinName = "Visual Studio 2013 Blue"
        Me.LabelControl5.LookAndFeel.UseDefaultLookAndFeel = False
        Me.LabelControl5.Name = "LabelControl5"
        Me.LabelControl5.Size = New System.Drawing.Size(53, 13)
        Me.LabelControl5.TabIndex = 84
        Me.LabelControl5.Text = "User Type:"
        '
        'txtPassword
        '
        Me.txtPassword.Location = New System.Drawing.Point(120, 105)
        Me.txtPassword.Name = "txtPassword"
        Me.txtPassword.Properties.Appearance.ForeColor = System.Drawing.Color.FromArgb(CType(CType(27, Byte), Integer), CType(CType(41, Byte), Integer), CType(CType(62, Byte), Integer))
        Me.txtPassword.Properties.Appearance.Options.UseForeColor = True
        Me.txtPassword.Properties.LookAndFeel.SkinName = "Visual Studio 2013 Blue"
        Me.txtPassword.Properties.LookAndFeel.UseDefaultLookAndFeel = False
        Me.txtPassword.Properties.PasswordChar = Global.Microsoft.VisualBasic.ChrW(42)
        Me.txtPassword.Size = New System.Drawing.Size(281, 20)
        Me.txtPassword.TabIndex = 3
        '
        'cboUserType
        '
        Me.cboUserType.Location = New System.Drawing.Point(120, 27)
        Me.cboUserType.Name = "cboUserType"
        Me.cboUserType.Properties.Appearance.ForeColor = System.Drawing.Color.FromArgb(CType(CType(27, Byte), Integer), CType(CType(41, Byte), Integer), CType(CType(62, Byte), Integer))
        Me.cboUserType.Properties.Appearance.Options.UseForeColor = True
        Me.cboUserType.Properties.Buttons.AddRange(New DevExpress.XtraEditors.Controls.EditorButton() {New DevExpress.XtraEditors.Controls.EditorButton(DevExpress.XtraEditors.Controls.ButtonPredefines.Combo)})
        Me.cboUserType.Properties.Items.AddRange(New Object() {"Receptionist", "Medical Technologist", "Pathologist", "Doctor", "Nurse", "Administrator"})
        Me.cboUserType.Properties.LookAndFeel.SkinName = "Visual Studio 2013 Blue"
        Me.cboUserType.Properties.LookAndFeel.UseDefaultLookAndFeel = False
        Me.cboUserType.Size = New System.Drawing.Size(281, 20)
        Me.cboUserType.TabIndex = 4
        '
        'btnSave
        '
        Me.btnSave.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnSave.Appearance.ForeColor = System.Drawing.Color.FromArgb(CType(CType(27, Byte), Integer), CType(CType(41, Byte), Integer), CType(CType(62, Byte), Integer))
        Me.btnSave.Appearance.Options.UseForeColor = True
        Me.btnSave.AppearanceHovered.BackColor = System.Drawing.Color.Gray
        Me.btnSave.AppearanceHovered.BackColor2 = System.Drawing.Color.Gray
        Me.btnSave.AppearanceHovered.BorderColor = System.Drawing.Color.Gray
        Me.btnSave.AppearanceHovered.Options.UseBackColor = True
        Me.btnSave.AppearanceHovered.Options.UseBorderColor = True
        Me.btnSave.Cursor = System.Windows.Forms.Cursors.Hand
        Me.btnSave.ImageOptions.AllowGlyphSkinning = DevExpress.Utils.DefaultBoolean.[True]
        Me.btnSave.ImageOptions.Image = CType(resources.GetObject("btnSave.ImageOptions.Image"), System.Drawing.Image)
        Me.btnSave.ImageOptions.ImageToTextAlignment = DevExpress.XtraEditors.ImageAlignToText.LeftCenter
        Me.btnSave.Location = New System.Drawing.Point(224, 451)
        Me.btnSave.Name = "btnSave"
        Me.btnSave.Size = New System.Drawing.Size(108, 27)
        Me.btnSave.TabIndex = 6
        Me.btnSave.Text = "&Save"
        '
        'btnClose
        '
        Me.btnClose.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnClose.Appearance.ForeColor = System.Drawing.Color.FromArgb(CType(CType(27, Byte), Integer), CType(CType(41, Byte), Integer), CType(CType(62, Byte), Integer))
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
        Me.btnClose.Location = New System.Drawing.Point(338, 451)
        Me.btnClose.Name = "btnClose"
        Me.btnClose.Size = New System.Drawing.Size(108, 27)
        Me.btnClose.TabIndex = 7
        Me.btnClose.Text = "&Close"
        '
        'cboMedTech
        '
        Me.cboMedTech.Location = New System.Drawing.Point(120, 53)
        Me.cboMedTech.Name = "cboMedTech"
        Me.cboMedTech.Properties.Appearance.ForeColor = System.Drawing.Color.FromArgb(CType(CType(27, Byte), Integer), CType(CType(41, Byte), Integer), CType(CType(62, Byte), Integer))
        Me.cboMedTech.Properties.Appearance.Options.UseForeColor = True
        Me.cboMedTech.Properties.Buttons.AddRange(New DevExpress.XtraEditors.Controls.EditorButton() {New DevExpress.XtraEditors.Controls.EditorButton(DevExpress.XtraEditors.Controls.ButtonPredefines.Combo)})
        Me.cboMedTech.Properties.LookAndFeel.SkinName = "Visual Studio 2013 Blue"
        Me.cboMedTech.Properties.LookAndFeel.UseDefaultLookAndFeel = False
        Me.cboMedTech.Size = New System.Drawing.Size(281, 20)
        Me.cboMedTech.TabIndex = 1
        '
        'cboLocation
        '
        Me.cboLocation.Location = New System.Drawing.Point(120, 131)
        Me.cboLocation.Name = "cboLocation"
        Me.cboLocation.Properties.Appearance.ForeColor = System.Drawing.Color.FromArgb(CType(CType(27, Byte), Integer), CType(CType(41, Byte), Integer), CType(CType(62, Byte), Integer))
        Me.cboLocation.Properties.Appearance.Options.UseForeColor = True
        Me.cboLocation.Properties.Buttons.AddRange(New DevExpress.XtraEditors.Controls.EditorButton() {New DevExpress.XtraEditors.Controls.EditorButton(DevExpress.XtraEditors.Controls.ButtonPredefines.Combo)})
        Me.cboLocation.Properties.Items.AddRange(New Object() {"Receptionist", "MedTech", "Cashier", "Administrator", "Super Admin"})
        Me.cboLocation.Properties.LookAndFeel.SkinName = "Visual Studio 2013 Blue"
        Me.cboLocation.Properties.LookAndFeel.UseDefaultLookAndFeel = False
        Me.cboLocation.Size = New System.Drawing.Size(281, 20)
        Me.cboLocation.TabIndex = 5
        '
        'LabelControl3
        '
        Me.LabelControl3.Appearance.ForeColor = System.Drawing.Color.FromArgb(CType(CType(27, Byte), Integer), CType(CType(41, Byte), Integer), CType(CType(62, Byte), Integer))
        Me.LabelControl3.Appearance.Options.UseForeColor = True
        Me.LabelControl3.Location = New System.Drawing.Point(33, 134)
        Me.LabelControl3.LookAndFeel.SkinName = "Visual Studio 2013 Blue"
        Me.LabelControl3.LookAndFeel.UseDefaultLookAndFeel = False
        Me.LabelControl3.Name = "LabelControl3"
        Me.LabelControl3.Size = New System.Drawing.Size(64, 13)
        Me.LabelControl3.TabIndex = 98
        Me.LabelControl3.Text = "Department:"
        '
        'TabPane1
        '
        Me.TabPane1.Controls.Add(Me.TabNavigationPage1)
        Me.TabPane1.Controls.Add(Me.TabNavigationPage2)
        Me.TabPane1.Controls.Add(Me.TabNavigationPage3)
        Me.TabPane1.Location = New System.Drawing.Point(12, 12)
        Me.TabPane1.Name = "TabPane1"
        Me.TabPane1.Pages.AddRange(New DevExpress.XtraBars.Navigation.NavigationPageBase() {Me.TabNavigationPage1, Me.TabNavigationPage2, Me.TabNavigationPage3})
        Me.TabPane1.RegularSize = New System.Drawing.Size(433, 431)
        Me.TabPane1.SelectedPage = Me.TabNavigationPage1
        Me.TabPane1.Size = New System.Drawing.Size(433, 431)
        Me.TabPane1.TabIndex = 102
        Me.TabPane1.Text = "Genral"
        '
        'TabNavigationPage1
        '
        Me.TabNavigationPage1.Caption = "General"
        Me.TabNavigationPage1.Controls.Add(Me.cboUserType)
        Me.TabNavigationPage1.Controls.Add(Me.cboMedTech)
        Me.TabNavigationPage1.Controls.Add(Me.LabelControl2)
        Me.TabNavigationPage1.Controls.Add(Me.cboLocation)
        Me.TabNavigationPage1.Controls.Add(Me.LabelControl5)
        Me.TabNavigationPage1.Controls.Add(Me.LabelControl3)
        Me.TabNavigationPage1.Controls.Add(Me.LabelControl4)
        Me.TabNavigationPage1.Controls.Add(Me.txtPassword)
        Me.TabNavigationPage1.Controls.Add(Me.txtUsername)
        Me.TabNavigationPage1.Controls.Add(Me.LabelControl1)
        Me.TabNavigationPage1.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold)
        Me.TabNavigationPage1.Name = "TabNavigationPage1"
        Me.TabNavigationPage1.Size = New System.Drawing.Size(433, 402)
        '
        'TabNavigationPage2
        '
        Me.TabNavigationPage2.Caption = "Modules"
        Me.TabNavigationPage2.Controls.Add(Me.dtResult)
        Me.TabNavigationPage2.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold)
        Me.TabNavigationPage2.Name = "TabNavigationPage2"
        Me.TabNavigationPage2.Size = New System.Drawing.Size(433, 402)
        '
        'dtResult
        '
        Me.dtResult.Dock = System.Windows.Forms.DockStyle.Fill
        Me.dtResult.Location = New System.Drawing.Point(0, 0)
        Me.dtResult.MainView = Me.GridView
        Me.dtResult.Name = "dtResult"
        Me.dtResult.Size = New System.Drawing.Size(433, 402)
        Me.dtResult.TabIndex = 59
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
        'TabNavigationPage3
        '
        Me.TabNavigationPage3.Caption = "Permissions"
        Me.TabNavigationPage3.Controls.Add(Me.dtPermissions)
        Me.TabNavigationPage3.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold)
        Me.TabNavigationPage3.Name = "TabNavigationPage3"
        Me.TabNavigationPage3.Size = New System.Drawing.Size(433, 402)
        '
        'dtPermissions
        '
        Me.dtPermissions.Dock = System.Windows.Forms.DockStyle.Fill
        Me.dtPermissions.Location = New System.Drawing.Point(0, 0)
        Me.dtPermissions.MainView = Me.GridPermissions
        Me.dtPermissions.Name = "dtPermissions"
        Me.dtPermissions.Size = New System.Drawing.Size(433, 402)
        Me.dtPermissions.TabIndex = 60
        Me.dtPermissions.ViewCollection.AddRange(New DevExpress.XtraGrid.Views.Base.BaseView() {Me.GridPermissions})
        '
        'GridPermissions
        '
        Me.GridPermissions.Appearance.SelectedRow.BackColor = System.Drawing.Color.WhiteSmoke
        Me.GridPermissions.Appearance.SelectedRow.BackColor2 = System.Drawing.Color.WhiteSmoke
        Me.GridPermissions.Appearance.SelectedRow.BorderColor = System.Drawing.Color.WhiteSmoke
        Me.GridPermissions.Appearance.SelectedRow.ForeColor = System.Drawing.Color.FromArgb(CType(CType(38, Byte), Integer), CType(CType(38, Byte), Integer), CType(CType(38, Byte), Integer))
        Me.GridPermissions.Appearance.SelectedRow.Options.UseBackColor = True
        Me.GridPermissions.Appearance.SelectedRow.Options.UseBorderColor = True
        Me.GridPermissions.Appearance.SelectedRow.Options.UseForeColor = True
        Me.GridPermissions.GridControl = Me.dtPermissions
        Me.GridPermissions.Name = "GridPermissions"
        Me.GridPermissions.OptionsBehavior.AutoExpandAllGroups = True
        Me.GridPermissions.OptionsCustomization.AllowColumnMoving = False
        Me.GridPermissions.OptionsCustomization.AllowColumnResizing = False
        Me.GridPermissions.OptionsCustomization.AllowFilter = False
        Me.GridPermissions.OptionsCustomization.AllowGroup = False
        Me.GridPermissions.OptionsCustomization.AllowQuickHideColumns = False
        Me.GridPermissions.OptionsCustomization.AllowSort = False
        Me.GridPermissions.OptionsSelection.CheckBoxSelectorColumnWidth = 30
        Me.GridPermissions.OptionsSelection.MultiSelect = True
        Me.GridPermissions.OptionsSelection.MultiSelectMode = DevExpress.XtraGrid.Views.Grid.GridMultiSelectMode.CheckBoxRowSelect
        Me.GridPermissions.OptionsSelection.ShowCheckBoxSelectorInColumnHeader = DevExpress.Utils.DefaultBoolean.[True]
        Me.GridPermissions.OptionsView.ShowGroupPanel = False
        Me.GridPermissions.OptionsView.ShowVerticalLines = DevExpress.Utils.DefaultBoolean.[False]
        '
        'frmAccountA
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(457, 488)
        Me.Controls.Add(Me.TabPane1)
        Me.Controls.Add(Me.btnSave)
        Me.Controls.Add(Me.btnClose)
        Me.FormBorderEffect = DevExpress.XtraEditors.FormBorderEffect.Glow
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "frmAccountA"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Add User Account"
        CType(Me.txtUsername.Properties, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.txtPassword.Properties, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.cboUserType.Properties, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.cboMedTech.Properties, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.cboLocation.Properties, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.TabPane1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.TabPane1.ResumeLayout(False)
        Me.TabNavigationPage1.ResumeLayout(False)
        Me.TabNavigationPage1.PerformLayout()
        Me.TabNavigationPage2.ResumeLayout(False)
        CType(Me.dtResult, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.GridView, System.ComponentModel.ISupportInitialize).EndInit()
        Me.TabNavigationPage3.ResumeLayout(False)
        CType(Me.dtPermissions, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.GridPermissions, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents LabelControl1 As DevExpress.XtraEditors.LabelControl
    Friend WithEvents txtUsername As DevExpress.XtraEditors.TextEdit
    Friend WithEvents LabelControl4 As DevExpress.XtraEditors.LabelControl
    Friend WithEvents LabelControl2 As DevExpress.XtraEditors.LabelControl
    Friend WithEvents LabelControl5 As DevExpress.XtraEditors.LabelControl
    Friend WithEvents txtPassword As DevExpress.XtraEditors.TextEdit
    Friend WithEvents cboUserType As DevExpress.XtraEditors.ComboBoxEdit
    Friend WithEvents btnSave As DevExpress.XtraEditors.SimpleButton
    Friend WithEvents btnClose As DevExpress.XtraEditors.SimpleButton
    Friend WithEvents cboMedTech As DevExpress.XtraEditors.ComboBoxEdit
    Friend WithEvents cboLocation As DevExpress.XtraEditors.ComboBoxEdit
    Friend WithEvents LabelControl3 As DevExpress.XtraEditors.LabelControl
    Friend WithEvents TabPane1 As DevExpress.XtraBars.Navigation.TabPane
    Friend WithEvents TabNavigationPage1 As DevExpress.XtraBars.Navigation.TabNavigationPage
    Friend WithEvents TabNavigationPage2 As DevExpress.XtraBars.Navigation.TabNavigationPage
    Friend WithEvents dtResult As DevExpress.XtraGrid.GridControl
    Friend WithEvents GridView As DevExpress.XtraGrid.Views.Grid.GridView
    Friend WithEvents TabNavigationPage3 As DevExpress.XtraBars.Navigation.TabNavigationPage
    Friend WithEvents dtPermissions As DevExpress.XtraGrid.GridControl
    Friend WithEvents GridPermissions As DevExpress.XtraGrid.Views.Grid.GridView
End Class
