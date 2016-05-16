<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class RoleEdit
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(RoleEdit))
        Me.Label1 = New System.Windows.Forms.Label
        Me.ID = New System.Windows.Forms.TextBox
        Me.RoleName = New System.Windows.Forms.TextBox
        Me.Label2 = New System.Windows.Forms.Label
        Me.TabControl1 = New System.Windows.Forms.TabControl
        Me.TabPage1 = New System.Windows.Forms.TabPage
        Me.ViewAPX = New System.Windows.Forms.CheckBox
        Me.Login = New System.Windows.Forms.CheckBox
        Me.Launch = New System.Windows.Forms.CheckBox
        Me.Active = New System.Windows.Forms.CheckBox
        Me.TabPage2 = New System.Windows.Forms.TabPage
        Me.ViewDist = New System.Windows.Forms.CheckBox
        Me.AllAccountAccess = New System.Windows.Forms.CheckBox
        Me.ViewTeam = New System.Windows.Forms.CheckBox
        Me.ViewPositions = New System.Windows.Forms.CheckBox
        Me.ViewAccount = New System.Windows.Forms.CheckBox
        Me.TabPage3 = New System.Windows.Forms.TabPage
        Me.ViewReps = New System.Windows.Forms.CheckBox
        Me.ViewFirms = New System.Windows.Forms.CheckBox
        Me.TabPage4 = New System.Windows.Forms.TabPage
        Me.AddUser = New System.Windows.Forms.CheckBox
        Me.ViewAdmin = New System.Windows.Forms.CheckBox
        Me.SetDefaultFee = New System.Windows.Forms.CheckBox
        Me.EditFirmFee = New System.Windows.Forms.CheckBox
        Me.SpoofUser = New System.Windows.Forms.CheckBox
        Me.EditUser = New System.Windows.Forms.CheckBox
        Me.ViewUsers = New System.Windows.Forms.CheckBox
        Me.TabPage5 = New System.Windows.Forms.TabPage
        Me.EditAccountMovesSettings = New System.Windows.Forms.CheckBox
        Me.ViewAccountMoves = New System.Windows.Forms.CheckBox
        Me.AMPSecCodes = New System.Windows.Forms.CheckBox
        Me.AMPAccess = New System.Windows.Forms.CheckBox
        Me.GroupBox4 = New System.Windows.Forms.GroupBox
        Me.LaunchRevenueCenter = New System.Windows.Forms.CheckBox
        Me.ETFPriceImport = New System.Windows.Forms.CheckBox
        Me.ViewReports = New System.Windows.Forms.CheckBox
        Me.GroupBox3 = New System.Windows.Forms.GroupBox
        Me.MAPGenerateReports = New System.Windows.Forms.CheckBox
        Me.MasterMappingCenterView = New System.Windows.Forms.CheckBox
        Me.MAPRefreshApprovals = New System.Windows.Forms.CheckBox
        Me.MAPAddDatabase = New System.Windows.Forms.CheckBox
        Me.MAPAssociatePlatform = New System.Windows.Forms.CheckBox
        Me.MAPEditWLPlatform = New System.Windows.Forms.CheckBox
        Me.MAPAddWLPlatform = New System.Windows.Forms.CheckBox
        Me.MAPEditPlatform = New System.Windows.Forms.CheckBox
        Me.MAPAddPlatform = New System.Windows.Forms.CheckBox
        Me.MAPEditObjective = New System.Windows.Forms.CheckBox
        Me.MAPEditAgreements = New System.Windows.Forms.CheckBox
        Me.MAPEditFirms = New System.Windows.Forms.CheckBox
        Me.MAPEditProductTypes = New System.Windows.Forms.CheckBox
        Me.MAPEditProducts = New System.Windows.Forms.CheckBox
        Me.MAPAddObjective = New System.Windows.Forms.CheckBox
        Me.MAPAddProductType = New System.Windows.Forms.CheckBox
        Me.MAPAddProducts = New System.Windows.Forms.CheckBox
        Me.MAPAddFirm = New System.Windows.Forms.CheckBox
        Me.MAPAddAgreement = New System.Windows.Forms.CheckBox
        Me.GroupBox2 = New System.Windows.Forms.GroupBox
        Me.UMAAutoRecon = New System.Windows.Forms.CheckBox
        Me.UMASystemFunctions = New System.Windows.Forms.CheckBox
        Me.UMATradeFunctions = New System.Windows.Forms.CheckBox
        Me.UMAPortfolioFunctions = New System.Windows.Forms.CheckBox
        Me.UMASkipCheck = New System.Windows.Forms.CheckBox
        Me.UMATranslate = New System.Windows.Forms.CheckBox
        Me.UMAImport = New System.Windows.Forms.CheckBox
        Me.UMALaunch = New System.Windows.Forms.CheckBox
        Me.Label3 = New System.Windows.Forms.Label
        Me.GroupBox1 = New System.Windows.Forms.GroupBox
        Me.RFPView = New System.Windows.Forms.CheckBox
        Me.RFPSubmit = New System.Windows.Forms.CheckBox
        Me.RFPWork = New System.Windows.Forms.CheckBox
        Me.RFPQuestionView = New System.Windows.Forms.CheckBox
        Me.RFPQuestionEdit = New System.Windows.Forms.CheckBox
        Me.RFPAddFirm = New System.Windows.Forms.CheckBox
        Me.RFPQuestionDelete = New System.Windows.Forms.CheckBox
        Me.ViewCompositeVer = New System.Windows.Forms.CheckBox
        Me.EditCompositeVer = New System.Windows.Forms.CheckBox
        Me.EditCompositeReasons = New System.Windows.Forms.CheckBox
        Me.FinishRecon = New System.Windows.Forms.CheckBox
        Me.ViewOps = New System.Windows.Forms.CheckBox
        Me.TabPage6 = New System.Windows.Forms.TabPage
        Me.LeadManagerAccess = New System.Windows.Forms.CheckBox
        Me.SalesManagementAccess = New System.Windows.Forms.CheckBox
        Me.TeamManagerAccess = New System.Windows.Forms.CheckBox
        Me.ViewRepBusinessReport = New System.Windows.Forms.CheckBox
        Me.cmdSave = New System.Windows.Forms.Button
        Me.cmdCancel = New System.Windows.Forms.Button
        Me.EditWhatsNew = New System.Windows.Forms.CheckBox
        Me.TabControl1.SuspendLayout()
        Me.TabPage1.SuspendLayout()
        Me.TabPage2.SuspendLayout()
        Me.TabPage3.SuspendLayout()
        Me.TabPage4.SuspendLayout()
        Me.TabPage5.SuspendLayout()
        Me.GroupBox4.SuspendLayout()
        Me.GroupBox3.SuspendLayout()
        Me.GroupBox2.SuspendLayout()
        Me.GroupBox1.SuspendLayout()
        Me.TabPage6.SuspendLayout()
        Me.SuspendLayout()
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(29, 9)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(43, 13)
        Me.Label1.TabIndex = 0
        Me.Label1.Text = "Role ID"
        '
        'ID
        '
        Me.ID.Location = New System.Drawing.Point(78, 6)
        Me.ID.Name = "ID"
        Me.ID.ReadOnly = True
        Me.ID.Size = New System.Drawing.Size(90, 20)
        Me.ID.TabIndex = 1
        '
        'RoleName
        '
        Me.RoleName.Location = New System.Drawing.Point(78, 32)
        Me.RoleName.Name = "RoleName"
        Me.RoleName.Size = New System.Drawing.Size(315, 20)
        Me.RoleName.TabIndex = 3
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(12, 35)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(60, 13)
        Me.Label2.TabIndex = 2
        Me.Label2.Text = "Role Name"
        '
        'TabControl1
        '
        Me.TabControl1.Controls.Add(Me.TabPage1)
        Me.TabControl1.Controls.Add(Me.TabPage2)
        Me.TabControl1.Controls.Add(Me.TabPage3)
        Me.TabControl1.Controls.Add(Me.TabPage4)
        Me.TabControl1.Controls.Add(Me.TabPage5)
        Me.TabControl1.Controls.Add(Me.TabPage6)
        Me.TabControl1.Location = New System.Drawing.Point(15, 70)
        Me.TabControl1.Name = "TabControl1"
        Me.TabControl1.SelectedIndex = 0
        Me.TabControl1.Size = New System.Drawing.Size(944, 546)
        Me.TabControl1.TabIndex = 4
        '
        'TabPage1
        '
        Me.TabPage1.Controls.Add(Me.EditWhatsNew)
        Me.TabPage1.Controls.Add(Me.ViewAPX)
        Me.TabPage1.Controls.Add(Me.Login)
        Me.TabPage1.Controls.Add(Me.Launch)
        Me.TabPage1.Controls.Add(Me.Active)
        Me.TabPage1.Location = New System.Drawing.Point(4, 22)
        Me.TabPage1.Name = "TabPage1"
        Me.TabPage1.Padding = New System.Windows.Forms.Padding(3)
        Me.TabPage1.Size = New System.Drawing.Size(936, 520)
        Me.TabPage1.TabIndex = 0
        Me.TabPage1.Text = "Application Rights"
        Me.TabPage1.UseVisualStyleBackColor = True
        '
        'ViewAPX
        '
        Me.ViewAPX.AutoSize = True
        Me.ViewAPX.Location = New System.Drawing.Point(6, 75)
        Me.ViewAPX.Name = "ViewAPX"
        Me.ViewAPX.Size = New System.Drawing.Size(145, 17)
        Me.ViewAPX.TabIndex = 12
        Me.ViewAPX.Text = "Allow user to launch APX"
        Me.ViewAPX.UseVisualStyleBackColor = True
        '
        'Login
        '
        Me.Login.AutoSize = True
        Me.Login.Location = New System.Drawing.Point(6, 52)
        Me.Login.Name = "Login"
        Me.Login.Size = New System.Drawing.Size(164, 17)
        Me.Login.TabIndex = 2
        Me.Login.Text = "Allow user to login to program"
        Me.Login.UseVisualStyleBackColor = True
        '
        'Launch
        '
        Me.Launch.AutoSize = True
        Me.Launch.Location = New System.Drawing.Point(6, 29)
        Me.Launch.Name = "Launch"
        Me.Launch.Size = New System.Drawing.Size(162, 17)
        Me.Launch.TabIndex = 1
        Me.Launch.Text = "Allow user to launch program"
        Me.Launch.UseVisualStyleBackColor = True
        '
        'Active
        '
        Me.Active.AutoSize = True
        Me.Active.Location = New System.Drawing.Point(6, 6)
        Me.Active.Name = "Active"
        Me.Active.Size = New System.Drawing.Size(56, 17)
        Me.Active.TabIndex = 0
        Me.Active.Text = "Active"
        Me.Active.UseVisualStyleBackColor = True
        '
        'TabPage2
        '
        Me.TabPage2.Controls.Add(Me.ViewDist)
        Me.TabPage2.Controls.Add(Me.AllAccountAccess)
        Me.TabPage2.Controls.Add(Me.ViewTeam)
        Me.TabPage2.Controls.Add(Me.ViewPositions)
        Me.TabPage2.Controls.Add(Me.ViewAccount)
        Me.TabPage2.Location = New System.Drawing.Point(4, 22)
        Me.TabPage2.Name = "TabPage2"
        Me.TabPage2.Padding = New System.Windows.Forms.Padding(3)
        Me.TabPage2.Size = New System.Drawing.Size(936, 520)
        Me.TabPage2.TabIndex = 1
        Me.TabPage2.Text = "Client Rights"
        Me.TabPage2.UseVisualStyleBackColor = True
        '
        'ViewDist
        '
        Me.ViewDist.AutoSize = True
        Me.ViewDist.Location = New System.Drawing.Point(6, 52)
        Me.ViewDist.Name = "ViewDist"
        Me.ViewDist.Size = New System.Drawing.Size(198, 17)
        Me.ViewDist.TabIndex = 10
        Me.ViewDist.Text = "Allow user to view Cash Distributions"
        Me.ViewDist.UseVisualStyleBackColor = True
        '
        'AllAccountAccess
        '
        Me.AllAccountAccess.AutoSize = True
        Me.AllAccountAccess.Location = New System.Drawing.Point(762, 6)
        Me.AllAccountAccess.Name = "AllAccountAccess"
        Me.AllAccountAccess.Size = New System.Drawing.Size(171, 17)
        Me.AllAccountAccess.TabIndex = 6
        Me.AllAccountAccess.Text = "Allow user to view all accounts"
        Me.AllAccountAccess.UseVisualStyleBackColor = True
        '
        'ViewTeam
        '
        Me.ViewTeam.AutoSize = True
        Me.ViewTeam.Location = New System.Drawing.Point(312, 6)
        Me.ViewTeam.Name = "ViewTeam"
        Me.ViewTeam.Size = New System.Drawing.Size(235, 17)
        Me.ViewTeam.TabIndex = 5
        Me.ViewTeam.Text = "Allow user to view all accounts on their team"
        Me.ViewTeam.UseVisualStyleBackColor = True
        '
        'ViewPositions
        '
        Me.ViewPositions.AutoSize = True
        Me.ViewPositions.Location = New System.Drawing.Point(6, 29)
        Me.ViewPositions.Name = "ViewPositions"
        Me.ViewPositions.Size = New System.Drawing.Size(174, 17)
        Me.ViewPositions.TabIndex = 4
        Me.ViewPositions.Text = "Allow user to view position data"
        Me.ViewPositions.UseVisualStyleBackColor = True
        '
        'ViewAccount
        '
        Me.ViewAccount.AutoSize = True
        Me.ViewAccount.Location = New System.Drawing.Point(6, 6)
        Me.ViewAccount.Name = "ViewAccount"
        Me.ViewAccount.Size = New System.Drawing.Size(158, 17)
        Me.ViewAccount.TabIndex = 3
        Me.ViewAccount.Text = "Allow user to view accounts"
        Me.ViewAccount.UseVisualStyleBackColor = True
        '
        'TabPage3
        '
        Me.TabPage3.Controls.Add(Me.ViewReps)
        Me.TabPage3.Controls.Add(Me.ViewFirms)
        Me.TabPage3.Location = New System.Drawing.Point(4, 22)
        Me.TabPage3.Name = "TabPage3"
        Me.TabPage3.Size = New System.Drawing.Size(936, 520)
        Me.TabPage3.TabIndex = 2
        Me.TabPage3.Text = "Firm Rights"
        Me.TabPage3.UseVisualStyleBackColor = True
        '
        'ViewReps
        '
        Me.ViewReps.AutoSize = True
        Me.ViewReps.Location = New System.Drawing.Point(3, 26)
        Me.ViewReps.Name = "ViewReps"
        Me.ViewReps.Size = New System.Drawing.Size(139, 17)
        Me.ViewReps.TabIndex = 5
        Me.ViewReps.Text = "Allow user to view Reps"
        Me.ViewReps.UseVisualStyleBackColor = True
        '
        'ViewFirms
        '
        Me.ViewFirms.AutoSize = True
        Me.ViewFirms.Location = New System.Drawing.Point(3, 3)
        Me.ViewFirms.Name = "ViewFirms"
        Me.ViewFirms.Size = New System.Drawing.Size(138, 17)
        Me.ViewFirms.TabIndex = 4
        Me.ViewFirms.Text = "Allow user to view Firms"
        Me.ViewFirms.UseVisualStyleBackColor = True
        '
        'TabPage4
        '
        Me.TabPage4.Controls.Add(Me.AddUser)
        Me.TabPage4.Controls.Add(Me.ViewAdmin)
        Me.TabPage4.Controls.Add(Me.SetDefaultFee)
        Me.TabPage4.Controls.Add(Me.EditFirmFee)
        Me.TabPage4.Controls.Add(Me.SpoofUser)
        Me.TabPage4.Controls.Add(Me.EditUser)
        Me.TabPage4.Controls.Add(Me.ViewUsers)
        Me.TabPage4.Location = New System.Drawing.Point(4, 22)
        Me.TabPage4.Name = "TabPage4"
        Me.TabPage4.Size = New System.Drawing.Size(936, 520)
        Me.TabPage4.TabIndex = 3
        Me.TabPage4.Text = "Functions"
        Me.TabPage4.UseVisualStyleBackColor = True
        '
        'AddUser
        '
        Me.AddUser.AutoSize = True
        Me.AddUser.Location = New System.Drawing.Point(288, 72)
        Me.AddUser.Name = "AddUser"
        Me.AddUser.Size = New System.Drawing.Size(125, 17)
        Me.AddUser.TabIndex = 13
        Me.AddUser.Text = "Allow user add Users"
        Me.AddUser.UseVisualStyleBackColor = True
        '
        'ViewAdmin
        '
        Me.ViewAdmin.AutoSize = True
        Me.ViewAdmin.Location = New System.Drawing.Point(288, 3)
        Me.ViewAdmin.Name = "ViewAdmin"
        Me.ViewAdmin.Size = New System.Drawing.Size(192, 17)
        Me.ViewAdmin.TabIndex = 12
        Me.ViewAdmin.Text = "Allow user to view Admin Functions"
        Me.ViewAdmin.UseVisualStyleBackColor = True
        '
        'SetDefaultFee
        '
        Me.SetDefaultFee.AutoSize = True
        Me.SetDefaultFee.Location = New System.Drawing.Point(3, 49)
        Me.SetDefaultFee.Name = "SetDefaultFee"
        Me.SetDefaultFee.Size = New System.Drawing.Size(214, 17)
        Me.SetDefaultFee.TabIndex = 11
        Me.SetDefaultFee.Text = "Allow user to set Default Fee Schedules"
        Me.SetDefaultFee.UseVisualStyleBackColor = True
        '
        'EditFirmFee
        '
        Me.EditFirmFee.AutoSize = True
        Me.EditFirmFee.Location = New System.Drawing.Point(3, 26)
        Me.EditFirmFee.Name = "EditFirmFee"
        Me.EditFirmFee.Size = New System.Drawing.Size(232, 17)
        Me.EditFirmFee.TabIndex = 10
        Me.EditFirmFee.Text = "Allow user to Edit Firm Level Fee Schedules"
        Me.EditFirmFee.UseVisualStyleBackColor = True
        '
        'SpoofUser
        '
        Me.SpoofUser.AutoSize = True
        Me.SpoofUser.Location = New System.Drawing.Point(3, 3)
        Me.SpoofUser.Name = "SpoofUser"
        Me.SpoofUser.Size = New System.Drawing.Size(185, 17)
        Me.SpoofUser.TabIndex = 9
        Me.SpoofUser.Text = "Allow user to spoof user accounts"
        Me.SpoofUser.UseVisualStyleBackColor = True
        '
        'EditUser
        '
        Me.EditUser.AutoSize = True
        Me.EditUser.Location = New System.Drawing.Point(288, 49)
        Me.EditUser.Name = "EditUser"
        Me.EditUser.Size = New System.Drawing.Size(187, 17)
        Me.EditUser.TabIndex = 8
        Me.EditUser.Text = "Allow user to edit Users and Roles"
        Me.EditUser.UseVisualStyleBackColor = True
        '
        'ViewUsers
        '
        Me.ViewUsers.AutoSize = True
        Me.ViewUsers.Location = New System.Drawing.Point(288, 26)
        Me.ViewUsers.Name = "ViewUsers"
        Me.ViewUsers.Size = New System.Drawing.Size(192, 17)
        Me.ViewUsers.TabIndex = 7
        Me.ViewUsers.Text = "Allow user to view Users and Roles"
        Me.ViewUsers.UseVisualStyleBackColor = True
        '
        'TabPage5
        '
        Me.TabPage5.Controls.Add(Me.EditAccountMovesSettings)
        Me.TabPage5.Controls.Add(Me.ViewAccountMoves)
        Me.TabPage5.Controls.Add(Me.AMPSecCodes)
        Me.TabPage5.Controls.Add(Me.AMPAccess)
        Me.TabPage5.Controls.Add(Me.GroupBox4)
        Me.TabPage5.Controls.Add(Me.ETFPriceImport)
        Me.TabPage5.Controls.Add(Me.ViewReports)
        Me.TabPage5.Controls.Add(Me.GroupBox3)
        Me.TabPage5.Controls.Add(Me.GroupBox2)
        Me.TabPage5.Controls.Add(Me.Label3)
        Me.TabPage5.Controls.Add(Me.GroupBox1)
        Me.TabPage5.Controls.Add(Me.ViewCompositeVer)
        Me.TabPage5.Controls.Add(Me.EditCompositeVer)
        Me.TabPage5.Controls.Add(Me.EditCompositeReasons)
        Me.TabPage5.Controls.Add(Me.FinishRecon)
        Me.TabPage5.Controls.Add(Me.ViewOps)
        Me.TabPage5.Location = New System.Drawing.Point(4, 22)
        Me.TabPage5.Name = "TabPage5"
        Me.TabPage5.Size = New System.Drawing.Size(936, 520)
        Me.TabPage5.TabIndex = 4
        Me.TabPage5.Text = "Operations"
        Me.TabPage5.UseVisualStyleBackColor = True
        '
        'EditAccountMovesSettings
        '
        Me.EditAccountMovesSettings.AutoSize = True
        Me.EditAccountMovesSettings.Location = New System.Drawing.Point(3, 260)
        Me.EditAccountMovesSettings.Name = "EditAccountMovesSettings"
        Me.EditAccountMovesSettings.Size = New System.Drawing.Size(221, 17)
        Me.EditAccountMovesSettings.TabIndex = 26
        Me.EditAccountMovesSettings.Text = "Allow user to edit account moves settings"
        Me.EditAccountMovesSettings.UseVisualStyleBackColor = True
        '
        'ViewAccountMoves
        '
        Me.ViewAccountMoves.AutoSize = True
        Me.ViewAccountMoves.Location = New System.Drawing.Point(3, 237)
        Me.ViewAccountMoves.Name = "ViewAccountMoves"
        Me.ViewAccountMoves.Size = New System.Drawing.Size(187, 17)
        Me.ViewAccountMoves.TabIndex = 25
        Me.ViewAccountMoves.Text = "Allow user to view account moves"
        Me.ViewAccountMoves.UseVisualStyleBackColor = True
        '
        'AMPSecCodes
        '
        Me.AMPSecCodes.AutoSize = True
        Me.AMPSecCodes.Location = New System.Drawing.Point(3, 216)
        Me.AMPSecCodes.Name = "AMPSecCodes"
        Me.AMPSecCodes.Size = New System.Drawing.Size(256, 17)
        Me.AMPSecCodes.TabIndex = 24
        Me.AMPSecCodes.Text = "Allow user to edit Ameriprise Step Out Sec Types"
        Me.AMPSecCodes.UseVisualStyleBackColor = True
        '
        'AMPAccess
        '
        Me.AMPAccess.AutoSize = True
        Me.AMPAccess.Location = New System.Drawing.Point(3, 193)
        Me.AMPAccess.Name = "AMPAccess"
        Me.AMPAccess.Size = New System.Drawing.Size(219, 17)
        Me.AMPAccess.TabIndex = 23
        Me.AMPAccess.Text = "Allow user to access Ameriprise Step Out"
        Me.AMPAccess.UseVisualStyleBackColor = True
        '
        'GroupBox4
        '
        Me.GroupBox4.Controls.Add(Me.LaunchRevenueCenter)
        Me.GroupBox4.Location = New System.Drawing.Point(480, 237)
        Me.GroupBox4.Name = "GroupBox4"
        Me.GroupBox4.Size = New System.Drawing.Size(216, 268)
        Me.GroupBox4.TabIndex = 22
        Me.GroupBox4.TabStop = False
        Me.GroupBox4.Text = "Revenue Center"
        '
        'LaunchRevenueCenter
        '
        Me.LaunchRevenueCenter.AutoSize = True
        Me.LaunchRevenueCenter.Location = New System.Drawing.Point(6, 19)
        Me.LaunchRevenueCenter.Name = "LaunchRevenueCenter"
        Me.LaunchRevenueCenter.Size = New System.Drawing.Size(192, 17)
        Me.LaunchRevenueCenter.TabIndex = 6
        Me.LaunchRevenueCenter.Text = "Allow user to view Revenue Center"
        Me.LaunchRevenueCenter.UseVisualStyleBackColor = True
        '
        'ETFPriceImport
        '
        Me.ETFPriceImport.AutoSize = True
        Me.ETFPriceImport.Location = New System.Drawing.Point(3, 170)
        Me.ETFPriceImport.Name = "ETFPriceImport"
        Me.ETFPriceImport.Size = New System.Drawing.Size(210, 17)
        Me.ETFPriceImport.TabIndex = 21
        Me.ETFPriceImport.Text = "Allow user to Access ETF Pricing Utility"
        Me.ETFPriceImport.UseVisualStyleBackColor = True
        '
        'ViewReports
        '
        Me.ViewReports.AutoSize = True
        Me.ViewReports.Location = New System.Drawing.Point(3, 147)
        Me.ViewReports.Name = "ViewReports"
        Me.ViewReports.Size = New System.Drawing.Size(152, 17)
        Me.ViewReports.TabIndex = 20
        Me.ViewReports.Text = "Allow user to View Reports"
        Me.ViewReports.UseVisualStyleBackColor = True
        '
        'GroupBox3
        '
        Me.GroupBox3.Controls.Add(Me.MAPGenerateReports)
        Me.GroupBox3.Controls.Add(Me.MasterMappingCenterView)
        Me.GroupBox3.Controls.Add(Me.MAPRefreshApprovals)
        Me.GroupBox3.Controls.Add(Me.MAPAddDatabase)
        Me.GroupBox3.Controls.Add(Me.MAPAssociatePlatform)
        Me.GroupBox3.Controls.Add(Me.MAPEditWLPlatform)
        Me.GroupBox3.Controls.Add(Me.MAPAddWLPlatform)
        Me.GroupBox3.Controls.Add(Me.MAPEditPlatform)
        Me.GroupBox3.Controls.Add(Me.MAPAddPlatform)
        Me.GroupBox3.Controls.Add(Me.MAPEditObjective)
        Me.GroupBox3.Controls.Add(Me.MAPEditAgreements)
        Me.GroupBox3.Controls.Add(Me.MAPEditFirms)
        Me.GroupBox3.Controls.Add(Me.MAPEditProductTypes)
        Me.GroupBox3.Controls.Add(Me.MAPEditProducts)
        Me.GroupBox3.Controls.Add(Me.MAPAddObjective)
        Me.GroupBox3.Controls.Add(Me.MAPAddProductType)
        Me.GroupBox3.Controls.Add(Me.MAPAddProducts)
        Me.GroupBox3.Controls.Add(Me.MAPAddFirm)
        Me.GroupBox3.Controls.Add(Me.MAPAddAgreement)
        Me.GroupBox3.Location = New System.Drawing.Point(702, 25)
        Me.GroupBox3.Name = "GroupBox3"
        Me.GroupBox3.Size = New System.Drawing.Size(231, 480)
        Me.GroupBox3.TabIndex = 19
        Me.GroupBox3.TabStop = False
        Me.GroupBox3.Text = "Master Mapping"
        '
        'MAPGenerateReports
        '
        Me.MAPGenerateReports.AutoSize = True
        Me.MAPGenerateReports.Location = New System.Drawing.Point(9, 433)
        Me.MAPGenerateReports.Name = "MAPGenerateReports"
        Me.MAPGenerateReports.Size = New System.Drawing.Size(110, 17)
        Me.MAPGenerateReports.TabIndex = 117
        Me.MAPGenerateReports.Text = "Generate Reports"
        Me.MAPGenerateReports.UseVisualStyleBackColor = True
        '
        'MasterMappingCenterView
        '
        Me.MasterMappingCenterView.AutoSize = True
        Me.MasterMappingCenterView.Location = New System.Drawing.Point(9, 19)
        Me.MasterMappingCenterView.Name = "MasterMappingCenterView"
        Me.MasterMappingCenterView.Size = New System.Drawing.Size(191, 17)
        Me.MasterMappingCenterView.TabIndex = 18
        Me.MasterMappingCenterView.Text = "Allow user to View Master Mapping"
        Me.MasterMappingCenterView.UseVisualStyleBackColor = True
        '
        'MAPRefreshApprovals
        '
        Me.MAPRefreshApprovals.AutoSize = True
        Me.MAPRefreshApprovals.Location = New System.Drawing.Point(9, 410)
        Me.MAPRefreshApprovals.Name = "MAPRefreshApprovals"
        Me.MAPRefreshApprovals.Size = New System.Drawing.Size(113, 17)
        Me.MAPRefreshApprovals.TabIndex = 116
        Me.MAPRefreshApprovals.Text = "Refresh Approvals"
        Me.MAPRefreshApprovals.UseVisualStyleBackColor = True
        '
        'MAPAddDatabase
        '
        Me.MAPAddDatabase.AutoSize = True
        Me.MAPAddDatabase.Location = New System.Drawing.Point(9, 387)
        Me.MAPAddDatabase.Name = "MAPAddDatabase"
        Me.MAPAddDatabase.Size = New System.Drawing.Size(94, 17)
        Me.MAPAddDatabase.TabIndex = 115
        Me.MAPAddDatabase.Text = "Add Database"
        Me.MAPAddDatabase.UseVisualStyleBackColor = True
        '
        'MAPAssociatePlatform
        '
        Me.MAPAssociatePlatform.AutoSize = True
        Me.MAPAssociatePlatform.Location = New System.Drawing.Point(9, 364)
        Me.MAPAssociatePlatform.Name = "MAPAssociatePlatform"
        Me.MAPAssociatePlatform.Size = New System.Drawing.Size(113, 17)
        Me.MAPAssociatePlatform.TabIndex = 114
        Me.MAPAssociatePlatform.Text = "Associate Platform"
        Me.MAPAssociatePlatform.UseVisualStyleBackColor = True
        '
        'MAPEditWLPlatform
        '
        Me.MAPEditWLPlatform.AutoSize = True
        Me.MAPEditWLPlatform.Location = New System.Drawing.Point(9, 341)
        Me.MAPEditWLPlatform.Name = "MAPEditWLPlatform"
        Me.MAPEditWLPlatform.Size = New System.Drawing.Size(108, 17)
        Me.MAPEditWLPlatform.TabIndex = 113
        Me.MAPEditWLPlatform.Text = " Edit WL Platform"
        Me.MAPEditWLPlatform.UseVisualStyleBackColor = True
        '
        'MAPAddWLPlatform
        '
        Me.MAPAddWLPlatform.AutoSize = True
        Me.MAPAddWLPlatform.Location = New System.Drawing.Point(9, 318)
        Me.MAPAddWLPlatform.Name = "MAPAddWLPlatform"
        Me.MAPAddWLPlatform.Size = New System.Drawing.Size(109, 17)
        Me.MAPAddWLPlatform.TabIndex = 112
        Me.MAPAddWLPlatform.Text = " Add WL Platform"
        Me.MAPAddWLPlatform.UseVisualStyleBackColor = True
        '
        'MAPEditPlatform
        '
        Me.MAPEditPlatform.AutoSize = True
        Me.MAPEditPlatform.Location = New System.Drawing.Point(9, 295)
        Me.MAPEditPlatform.Name = "MAPEditPlatform"
        Me.MAPEditPlatform.Size = New System.Drawing.Size(85, 17)
        Me.MAPEditPlatform.TabIndex = 111
        Me.MAPEditPlatform.Text = "Edit Platform"
        Me.MAPEditPlatform.UseVisualStyleBackColor = True
        '
        'MAPAddPlatform
        '
        Me.MAPAddPlatform.AutoSize = True
        Me.MAPAddPlatform.Location = New System.Drawing.Point(9, 272)
        Me.MAPAddPlatform.Name = "MAPAddPlatform"
        Me.MAPAddPlatform.Size = New System.Drawing.Size(86, 17)
        Me.MAPAddPlatform.TabIndex = 110
        Me.MAPAddPlatform.Text = "Add Platform"
        Me.MAPAddPlatform.UseVisualStyleBackColor = True
        '
        'MAPEditObjective
        '
        Me.MAPEditObjective.AutoSize = True
        Me.MAPEditObjective.Location = New System.Drawing.Point(9, 249)
        Me.MAPEditObjective.Name = "MAPEditObjective"
        Me.MAPEditObjective.Size = New System.Drawing.Size(92, 17)
        Me.MAPEditObjective.TabIndex = 109
        Me.MAPEditObjective.Text = "Edit Objective"
        Me.MAPEditObjective.UseVisualStyleBackColor = True
        '
        'MAPEditAgreements
        '
        Me.MAPEditAgreements.AutoSize = True
        Me.MAPEditAgreements.Location = New System.Drawing.Point(9, 226)
        Me.MAPEditAgreements.Name = "MAPEditAgreements"
        Me.MAPEditAgreements.Size = New System.Drawing.Size(103, 17)
        Me.MAPEditAgreements.TabIndex = 108
        Me.MAPEditAgreements.Text = "Edit Agreements"
        Me.MAPEditAgreements.UseVisualStyleBackColor = True
        '
        'MAPEditFirms
        '
        Me.MAPEditFirms.AutoSize = True
        Me.MAPEditFirms.Location = New System.Drawing.Point(9, 203)
        Me.MAPEditFirms.Name = "MAPEditFirms"
        Me.MAPEditFirms.Size = New System.Drawing.Size(71, 17)
        Me.MAPEditFirms.TabIndex = 107
        Me.MAPEditFirms.Text = "Edit Firms"
        Me.MAPEditFirms.UseVisualStyleBackColor = True
        '
        'MAPEditProductTypes
        '
        Me.MAPEditProductTypes.AutoSize = True
        Me.MAPEditProductTypes.Location = New System.Drawing.Point(9, 180)
        Me.MAPEditProductTypes.Name = "MAPEditProductTypes"
        Me.MAPEditProductTypes.Size = New System.Drawing.Size(116, 17)
        Me.MAPEditProductTypes.TabIndex = 106
        Me.MAPEditProductTypes.Text = "Edit Product Types"
        Me.MAPEditProductTypes.UseVisualStyleBackColor = True
        '
        'MAPEditProducts
        '
        Me.MAPEditProducts.AutoSize = True
        Me.MAPEditProducts.Location = New System.Drawing.Point(9, 157)
        Me.MAPEditProducts.Name = "MAPEditProducts"
        Me.MAPEditProducts.Size = New System.Drawing.Size(89, 17)
        Me.MAPEditProducts.TabIndex = 105
        Me.MAPEditProducts.Text = "Edit Products"
        Me.MAPEditProducts.UseVisualStyleBackColor = True
        '
        'MAPAddObjective
        '
        Me.MAPAddObjective.AutoSize = True
        Me.MAPAddObjective.Location = New System.Drawing.Point(9, 134)
        Me.MAPAddObjective.Name = "MAPAddObjective"
        Me.MAPAddObjective.Size = New System.Drawing.Size(93, 17)
        Me.MAPAddObjective.TabIndex = 104
        Me.MAPAddObjective.Text = "Add Objective"
        Me.MAPAddObjective.UseVisualStyleBackColor = True
        '
        'MAPAddProductType
        '
        Me.MAPAddProductType.AutoSize = True
        Me.MAPAddProductType.Location = New System.Drawing.Point(9, 111)
        Me.MAPAddProductType.Name = "MAPAddProductType"
        Me.MAPAddProductType.Size = New System.Drawing.Size(112, 17)
        Me.MAPAddProductType.TabIndex = 103
        Me.MAPAddProductType.Text = "Add Product Type"
        Me.MAPAddProductType.UseVisualStyleBackColor = True
        '
        'MAPAddProducts
        '
        Me.MAPAddProducts.AutoSize = True
        Me.MAPAddProducts.Location = New System.Drawing.Point(9, 88)
        Me.MAPAddProducts.Name = "MAPAddProducts"
        Me.MAPAddProducts.Size = New System.Drawing.Size(90, 17)
        Me.MAPAddProducts.TabIndex = 102
        Me.MAPAddProducts.Text = "Add Products"
        Me.MAPAddProducts.UseVisualStyleBackColor = True
        '
        'MAPAddFirm
        '
        Me.MAPAddFirm.AutoSize = True
        Me.MAPAddFirm.Location = New System.Drawing.Point(9, 65)
        Me.MAPAddFirm.Name = "MAPAddFirm"
        Me.MAPAddFirm.Size = New System.Drawing.Size(67, 17)
        Me.MAPAddFirm.TabIndex = 101
        Me.MAPAddFirm.Text = "Add Firm"
        Me.MAPAddFirm.UseVisualStyleBackColor = True
        '
        'MAPAddAgreement
        '
        Me.MAPAddAgreement.AutoSize = True
        Me.MAPAddAgreement.Location = New System.Drawing.Point(9, 42)
        Me.MAPAddAgreement.Name = "MAPAddAgreement"
        Me.MAPAddAgreement.Size = New System.Drawing.Size(99, 17)
        Me.MAPAddAgreement.TabIndex = 100
        Me.MAPAddAgreement.Text = "Add Agreement"
        Me.MAPAddAgreement.UseVisualStyleBackColor = True
        '
        'GroupBox2
        '
        Me.GroupBox2.Controls.Add(Me.UMAAutoRecon)
        Me.GroupBox2.Controls.Add(Me.UMASystemFunctions)
        Me.GroupBox2.Controls.Add(Me.UMATradeFunctions)
        Me.GroupBox2.Controls.Add(Me.UMAPortfolioFunctions)
        Me.GroupBox2.Controls.Add(Me.UMASkipCheck)
        Me.GroupBox2.Controls.Add(Me.UMATranslate)
        Me.GroupBox2.Controls.Add(Me.UMAImport)
        Me.GroupBox2.Controls.Add(Me.UMALaunch)
        Me.GroupBox2.Location = New System.Drawing.Point(480, 25)
        Me.GroupBox2.Name = "GroupBox2"
        Me.GroupBox2.Size = New System.Drawing.Size(216, 209)
        Me.GroupBox2.TabIndex = 17
        Me.GroupBox2.TabStop = False
        Me.GroupBox2.Text = "UMA Functions"
        '
        'UMAAutoRecon
        '
        Me.UMAAutoRecon.AutoSize = True
        Me.UMAAutoRecon.Location = New System.Drawing.Point(6, 180)
        Me.UMAAutoRecon.Name = "UMAAutoRecon"
        Me.UMAAutoRecon.Size = New System.Drawing.Size(207, 17)
        Me.UMAAutoRecon.TabIndex = 12
        Me.UMAAutoRecon.Text = "Allow user to use Auto Recon function"
        Me.UMAAutoRecon.UseVisualStyleBackColor = True
        '
        'UMASystemFunctions
        '
        Me.UMASystemFunctions.AutoSize = True
        Me.UMASystemFunctions.Location = New System.Drawing.Point(6, 157)
        Me.UMASystemFunctions.Name = "UMASystemFunctions"
        Me.UMASystemFunctions.Size = New System.Drawing.Size(186, 17)
        Me.UMASystemFunctions.TabIndex = 11
        Me.UMASystemFunctions.Text = "Allow user to edit System Fuctions"
        Me.UMASystemFunctions.UseVisualStyleBackColor = True
        '
        'UMATradeFunctions
        '
        Me.UMATradeFunctions.AutoSize = True
        Me.UMATradeFunctions.Location = New System.Drawing.Point(6, 134)
        Me.UMATradeFunctions.Name = "UMATradeFunctions"
        Me.UMATradeFunctions.Size = New System.Drawing.Size(186, 17)
        Me.UMATradeFunctions.TabIndex = 10
        Me.UMATradeFunctions.Text = "Allow user to edit Trade Functions"
        Me.UMATradeFunctions.UseVisualStyleBackColor = True
        '
        'UMAPortfolioFunctions
        '
        Me.UMAPortfolioFunctions.AutoSize = True
        Me.UMAPortfolioFunctions.Location = New System.Drawing.Point(6, 111)
        Me.UMAPortfolioFunctions.Name = "UMAPortfolioFunctions"
        Me.UMAPortfolioFunctions.Size = New System.Drawing.Size(196, 17)
        Me.UMAPortfolioFunctions.TabIndex = 9
        Me.UMAPortfolioFunctions.Text = "Allow user to edit Portfolio Functions"
        Me.UMAPortfolioFunctions.UseVisualStyleBackColor = True
        '
        'UMASkipCheck
        '
        Me.UMASkipCheck.AutoSize = True
        Me.UMASkipCheck.Location = New System.Drawing.Point(6, 88)
        Me.UMASkipCheck.Name = "UMASkipCheck"
        Me.UMASkipCheck.Size = New System.Drawing.Size(186, 17)
        Me.UMASkipCheck.TabIndex = 8
        Me.UMASkipCheck.Text = "Allow user to skip data verification"
        Me.UMASkipCheck.UseVisualStyleBackColor = True
        '
        'UMATranslate
        '
        Me.UMATranslate.AutoSize = True
        Me.UMATranslate.Location = New System.Drawing.Point(6, 65)
        Me.UMATranslate.Name = "UMATranslate"
        Me.UMATranslate.Size = New System.Drawing.Size(181, 17)
        Me.UMATranslate.TabIndex = 7
        Me.UMATranslate.Text = "Allow user to Translate UMA files"
        Me.UMATranslate.UseVisualStyleBackColor = True
        '
        'UMAImport
        '
        Me.UMAImport.AutoSize = True
        Me.UMAImport.Location = New System.Drawing.Point(6, 42)
        Me.UMAImport.Name = "UMAImport"
        Me.UMAImport.Size = New System.Drawing.Size(166, 17)
        Me.UMAImport.TabIndex = 6
        Me.UMAImport.Text = "Allow user to Import UMA files"
        Me.UMAImport.UseVisualStyleBackColor = True
        '
        'UMALaunch
        '
        Me.UMALaunch.AutoSize = True
        Me.UMALaunch.Location = New System.Drawing.Point(6, 19)
        Me.UMALaunch.Name = "UMALaunch"
        Me.UMALaunch.Size = New System.Drawing.Size(175, 17)
        Me.UMALaunch.TabIndex = 5
        Me.UMALaunch.Text = "Allow user to view UMA System"
        Me.UMALaunch.UseVisualStyleBackColor = True
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label3.Location = New System.Drawing.Point(5, 9)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(634, 13)
        Me.Label3.TabIndex = 16
        Me.Label3.Text = "You must grant permission to View Operations Tab before the user can launch any o" & _
            "f the functions on this tab."
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.RFPView)
        Me.GroupBox1.Controls.Add(Me.RFPSubmit)
        Me.GroupBox1.Controls.Add(Me.RFPWork)
        Me.GroupBox1.Controls.Add(Me.RFPQuestionView)
        Me.GroupBox1.Controls.Add(Me.RFPQuestionEdit)
        Me.GroupBox1.Controls.Add(Me.RFPAddFirm)
        Me.GroupBox1.Controls.Add(Me.RFPQuestionDelete)
        Me.GroupBox1.Location = New System.Drawing.Point(247, 25)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(227, 181)
        Me.GroupBox1.TabIndex = 15
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "RFP Central"
        '
        'RFPView
        '
        Me.RFPView.AutoSize = True
        Me.RFPView.Location = New System.Drawing.Point(6, 19)
        Me.RFPView.Name = "RFPView"
        Me.RFPView.Size = New System.Drawing.Size(172, 17)
        Me.RFPView.TabIndex = 4
        Me.RFPView.Text = "Allow user to view RFP System"
        Me.RFPView.UseVisualStyleBackColor = True
        '
        'RFPSubmit
        '
        Me.RFPSubmit.AutoSize = True
        Me.RFPSubmit.Location = New System.Drawing.Point(6, 42)
        Me.RFPSubmit.Name = "RFPSubmit"
        Me.RFPSubmit.Size = New System.Drawing.Size(160, 17)
        Me.RFPSubmit.TabIndex = 10
        Me.RFPSubmit.Text = "Allow user to Submit an RFP"
        Me.RFPSubmit.UseVisualStyleBackColor = True
        '
        'RFPWork
        '
        Me.RFPWork.AutoSize = True
        Me.RFPWork.Location = New System.Drawing.Point(6, 65)
        Me.RFPWork.Name = "RFPWork"
        Me.RFPWork.Size = New System.Drawing.Size(191, 17)
        Me.RFPWork.TabIndex = 9
        Me.RFPWork.Text = "Allow user to Work submited RFP's"
        Me.RFPWork.UseVisualStyleBackColor = True
        '
        'RFPQuestionView
        '
        Me.RFPQuestionView.AutoSize = True
        Me.RFPQuestionView.Location = New System.Drawing.Point(6, 88)
        Me.RFPQuestionView.Name = "RFPQuestionView"
        Me.RFPQuestionView.Size = New System.Drawing.Size(185, 17)
        Me.RFPQuestionView.TabIndex = 6
        Me.RFPQuestionView.Text = "Allow user to View Question Bank"
        Me.RFPQuestionView.UseVisualStyleBackColor = True
        '
        'RFPQuestionEdit
        '
        Me.RFPQuestionEdit.AutoSize = True
        Me.RFPQuestionEdit.Location = New System.Drawing.Point(6, 111)
        Me.RFPQuestionEdit.Name = "RFPQuestionEdit"
        Me.RFPQuestionEdit.Size = New System.Drawing.Size(157, 17)
        Me.RFPQuestionEdit.TabIndex = 5
        Me.RFPQuestionEdit.Text = "Allow user to Edit Questions"
        Me.RFPQuestionEdit.UseVisualStyleBackColor = True
        '
        'RFPAddFirm
        '
        Me.RFPAddFirm.AutoSize = True
        Me.RFPAddFirm.Location = New System.Drawing.Point(6, 157)
        Me.RFPAddFirm.Name = "RFPAddFirm"
        Me.RFPAddFirm.Size = New System.Drawing.Size(182, 17)
        Me.RFPAddFirm.TabIndex = 7
        Me.RFPAddFirm.Text = "Allow user to Maintain Firm RFP's"
        Me.RFPAddFirm.UseVisualStyleBackColor = True
        '
        'RFPQuestionDelete
        '
        Me.RFPQuestionDelete.AutoSize = True
        Me.RFPQuestionDelete.Location = New System.Drawing.Point(6, 134)
        Me.RFPQuestionDelete.Name = "RFPQuestionDelete"
        Me.RFPQuestionDelete.Size = New System.Drawing.Size(170, 17)
        Me.RFPQuestionDelete.TabIndex = 8
        Me.RFPQuestionDelete.Text = "Allow user to Delete Questions"
        Me.RFPQuestionDelete.UseVisualStyleBackColor = True
        '
        'ViewCompositeVer
        '
        Me.ViewCompositeVer.AutoSize = True
        Me.ViewCompositeVer.Location = New System.Drawing.Point(3, 55)
        Me.ViewCompositeVer.Name = "ViewCompositeVer"
        Me.ViewCompositeVer.Size = New System.Drawing.Size(219, 17)
        Me.ViewCompositeVer.TabIndex = 14
        Me.ViewCompositeVer.Text = "Allow user to View Composite Verification"
        Me.ViewCompositeVer.UseVisualStyleBackColor = True
        '
        'EditCompositeVer
        '
        Me.EditCompositeVer.AutoSize = True
        Me.EditCompositeVer.Location = New System.Drawing.Point(3, 78)
        Me.EditCompositeVer.Name = "EditCompositeVer"
        Me.EditCompositeVer.Size = New System.Drawing.Size(214, 17)
        Me.EditCompositeVer.TabIndex = 13
        Me.EditCompositeVer.Text = "Allow user to Edit Composite Verification"
        Me.EditCompositeVer.UseVisualStyleBackColor = True
        '
        'EditCompositeReasons
        '
        Me.EditCompositeReasons.AutoSize = True
        Me.EditCompositeReasons.Location = New System.Drawing.Point(3, 101)
        Me.EditCompositeReasons.Name = "EditCompositeReasons"
        Me.EditCompositeReasons.Size = New System.Drawing.Size(204, 17)
        Me.EditCompositeReasons.TabIndex = 12
        Me.EditCompositeReasons.Text = "Allow user to Edit Composite Reasons"
        Me.EditCompositeReasons.UseVisualStyleBackColor = True
        '
        'FinishRecon
        '
        Me.FinishRecon.AutoSize = True
        Me.FinishRecon.Location = New System.Drawing.Point(3, 124)
        Me.FinishRecon.Name = "FinishRecon"
        Me.FinishRecon.Size = New System.Drawing.Size(151, 17)
        Me.FinishRecon.TabIndex = 11
        Me.FinishRecon.Text = "Allow user to Finish Recon"
        Me.FinishRecon.UseVisualStyleBackColor = True
        '
        'ViewOps
        '
        Me.ViewOps.AutoSize = True
        Me.ViewOps.Location = New System.Drawing.Point(3, 32)
        Me.ViewOps.Name = "ViewOps"
        Me.ViewOps.Size = New System.Drawing.Size(187, 17)
        Me.ViewOps.TabIndex = 3
        Me.ViewOps.Text = "Allow user to view Operations Tab"
        Me.ViewOps.UseVisualStyleBackColor = True
        '
        'TabPage6
        '
        Me.TabPage6.Controls.Add(Me.LeadManagerAccess)
        Me.TabPage6.Controls.Add(Me.SalesManagementAccess)
        Me.TabPage6.Controls.Add(Me.TeamManagerAccess)
        Me.TabPage6.Controls.Add(Me.ViewRepBusinessReport)
        Me.TabPage6.Location = New System.Drawing.Point(4, 22)
        Me.TabPage6.Name = "TabPage6"
        Me.TabPage6.Size = New System.Drawing.Size(936, 520)
        Me.TabPage6.TabIndex = 5
        Me.TabPage6.Text = "Sales"
        Me.TabPage6.UseVisualStyleBackColor = True
        '
        'LeadManagerAccess
        '
        Me.LeadManagerAccess.AutoSize = True
        Me.LeadManagerAccess.Location = New System.Drawing.Point(13, 83)
        Me.LeadManagerAccess.Name = "LeadManagerAccess"
        Me.LeadManagerAccess.Size = New System.Drawing.Size(154, 17)
        Me.LeadManagerAccess.TabIndex = 27
        Me.LeadManagerAccess.Text = "Can access Lead Manager"
        Me.LeadManagerAccess.UseVisualStyleBackColor = True
        '
        'SalesManagementAccess
        '
        Me.SalesManagementAccess.AutoSize = True
        Me.SalesManagementAccess.Location = New System.Drawing.Point(13, 60)
        Me.SalesManagementAccess.Name = "SalesManagementAccess"
        Me.SalesManagementAccess.Size = New System.Drawing.Size(166, 17)
        Me.SalesManagementAccess.TabIndex = 26
        Me.SalesManagementAccess.Text = "Can view Sales Manager Tab"
        Me.SalesManagementAccess.UseVisualStyleBackColor = True
        '
        'TeamManagerAccess
        '
        Me.TeamManagerAccess.AutoSize = True
        Me.TeamManagerAccess.Location = New System.Drawing.Point(13, 37)
        Me.TeamManagerAccess.Name = "TeamManagerAccess"
        Me.TeamManagerAccess.Size = New System.Drawing.Size(157, 17)
        Me.TeamManagerAccess.TabIndex = 25
        Me.TeamManagerAccess.Text = "Can access Team Manager"
        Me.TeamManagerAccess.UseVisualStyleBackColor = True
        '
        'ViewRepBusinessReport
        '
        Me.ViewRepBusinessReport.AutoSize = True
        Me.ViewRepBusinessReport.Location = New System.Drawing.Point(13, 14)
        Me.ViewRepBusinessReport.Name = "ViewRepBusinessReport"
        Me.ViewRepBusinessReport.Size = New System.Drawing.Size(228, 17)
        Me.ViewRepBusinessReport.TabIndex = 24
        Me.ViewRepBusinessReport.Text = "Can run Rep Business Breakdown Reports"
        Me.ViewRepBusinessReport.UseVisualStyleBackColor = True
        '
        'cmdSave
        '
        Me.cmdSave.Location = New System.Drawing.Point(861, 12)
        Me.cmdSave.Name = "cmdSave"
        Me.cmdSave.Size = New System.Drawing.Size(94, 23)
        Me.cmdSave.TabIndex = 21
        Me.cmdSave.Text = "&Save"
        '
        'cmdCancel
        '
        Me.cmdCancel.Location = New System.Drawing.Point(861, 41)
        Me.cmdCancel.Name = "cmdCancel"
        Me.cmdCancel.Size = New System.Drawing.Size(94, 23)
        Me.cmdCancel.TabIndex = 20
        Me.cmdCancel.Text = "&Cancel"
        '
        'EditWhatsNew
        '
        Me.EditWhatsNew.AutoSize = True
        Me.EditWhatsNew.Location = New System.Drawing.Point(6, 98)
        Me.EditWhatsNew.Name = "EditWhatsNew"
        Me.EditWhatsNew.Size = New System.Drawing.Size(191, 17)
        Me.EditWhatsNew.TabIndex = 13
        Me.EditWhatsNew.Text = "Allow user Edit Whats New Screen"
        Me.EditWhatsNew.UseVisualStyleBackColor = True
        '
        'RoleEdit
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.Color.White
        Me.ClientSize = New System.Drawing.Size(971, 628)
        Me.Controls.Add(Me.cmdSave)
        Me.Controls.Add(Me.cmdCancel)
        Me.Controls.Add(Me.TabControl1)
        Me.Controls.Add(Me.RoleName)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.ID)
        Me.Controls.Add(Me.Label1)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.Fixed3D
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "RoleEdit"
        Me.Text = "RoleEdit"
        Me.TabControl1.ResumeLayout(False)
        Me.TabPage1.ResumeLayout(False)
        Me.TabPage1.PerformLayout()
        Me.TabPage2.ResumeLayout(False)
        Me.TabPage2.PerformLayout()
        Me.TabPage3.ResumeLayout(False)
        Me.TabPage3.PerformLayout()
        Me.TabPage4.ResumeLayout(False)
        Me.TabPage4.PerformLayout()
        Me.TabPage5.ResumeLayout(False)
        Me.TabPage5.PerformLayout()
        Me.GroupBox4.ResumeLayout(False)
        Me.GroupBox4.PerformLayout()
        Me.GroupBox3.ResumeLayout(False)
        Me.GroupBox3.PerformLayout()
        Me.GroupBox2.ResumeLayout(False)
        Me.GroupBox2.PerformLayout()
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox1.PerformLayout()
        Me.TabPage6.ResumeLayout(False)
        Me.TabPage6.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents ID As System.Windows.Forms.TextBox
    Friend WithEvents RoleName As System.Windows.Forms.TextBox
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents TabControl1 As System.Windows.Forms.TabControl
    Friend WithEvents TabPage1 As System.Windows.Forms.TabPage
    Friend WithEvents TabPage2 As System.Windows.Forms.TabPage
    Friend WithEvents TabPage3 As System.Windows.Forms.TabPage
    Friend WithEvents TabPage4 As System.Windows.Forms.TabPage
    Friend WithEvents cmdSave As System.Windows.Forms.Button
    Friend WithEvents cmdCancel As System.Windows.Forms.Button
    Friend WithEvents Login As System.Windows.Forms.CheckBox
    Friend WithEvents Launch As System.Windows.Forms.CheckBox
    Friend WithEvents Active As System.Windows.Forms.CheckBox
    Friend WithEvents AllAccountAccess As System.Windows.Forms.CheckBox
    Friend WithEvents ViewTeam As System.Windows.Forms.CheckBox
    Friend WithEvents ViewPositions As System.Windows.Forms.CheckBox
    Friend WithEvents ViewAccount As System.Windows.Forms.CheckBox
    Friend WithEvents SpoofUser As System.Windows.Forms.CheckBox
    Friend WithEvents EditUser As System.Windows.Forms.CheckBox
    Friend WithEvents ViewUsers As System.Windows.Forms.CheckBox
    Friend WithEvents ViewFirms As System.Windows.Forms.CheckBox
    Friend WithEvents ViewDist As System.Windows.Forms.CheckBox
    Friend WithEvents ViewReps As System.Windows.Forms.CheckBox
    Friend WithEvents EditFirmFee As System.Windows.Forms.CheckBox
    Friend WithEvents ViewAPX As System.Windows.Forms.CheckBox
    Friend WithEvents SetDefaultFee As System.Windows.Forms.CheckBox
    Friend WithEvents AddUser As System.Windows.Forms.CheckBox
    Friend WithEvents ViewAdmin As System.Windows.Forms.CheckBox
    Friend WithEvents TabPage5 As System.Windows.Forms.TabPage
    Friend WithEvents ViewOps As System.Windows.Forms.CheckBox
    Friend WithEvents ViewCompositeVer As System.Windows.Forms.CheckBox
    Friend WithEvents EditCompositeVer As System.Windows.Forms.CheckBox
    Friend WithEvents EditCompositeReasons As System.Windows.Forms.CheckBox
    Friend WithEvents FinishRecon As System.Windows.Forms.CheckBox
    Friend WithEvents RFPSubmit As System.Windows.Forms.CheckBox
    Friend WithEvents RFPWork As System.Windows.Forms.CheckBox
    Friend WithEvents RFPQuestionDelete As System.Windows.Forms.CheckBox
    Friend WithEvents RFPAddFirm As System.Windows.Forms.CheckBox
    Friend WithEvents RFPQuestionView As System.Windows.Forms.CheckBox
    Friend WithEvents RFPQuestionEdit As System.Windows.Forms.CheckBox
    Friend WithEvents RFPView As System.Windows.Forms.CheckBox
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents GroupBox2 As System.Windows.Forms.GroupBox
    Friend WithEvents UMASkipCheck As System.Windows.Forms.CheckBox
    Friend WithEvents UMATranslate As System.Windows.Forms.CheckBox
    Friend WithEvents UMAImport As System.Windows.Forms.CheckBox
    Friend WithEvents UMALaunch As System.Windows.Forms.CheckBox
    Friend WithEvents UMASystemFunctions As System.Windows.Forms.CheckBox
    Friend WithEvents UMATradeFunctions As System.Windows.Forms.CheckBox
    Friend WithEvents UMAPortfolioFunctions As System.Windows.Forms.CheckBox
    Friend WithEvents UMAAutoRecon As System.Windows.Forms.CheckBox
    Friend WithEvents MasterMappingCenterView As System.Windows.Forms.CheckBox
    Friend WithEvents GroupBox3 As System.Windows.Forms.GroupBox
    Friend WithEvents MAPGenerateReports As System.Windows.Forms.CheckBox
    Friend WithEvents MAPRefreshApprovals As System.Windows.Forms.CheckBox
    Friend WithEvents MAPAddDatabase As System.Windows.Forms.CheckBox
    Friend WithEvents MAPAssociatePlatform As System.Windows.Forms.CheckBox
    Friend WithEvents MAPEditWLPlatform As System.Windows.Forms.CheckBox
    Friend WithEvents MAPAddWLPlatform As System.Windows.Forms.CheckBox
    Friend WithEvents MAPEditPlatform As System.Windows.Forms.CheckBox
    Friend WithEvents MAPAddPlatform As System.Windows.Forms.CheckBox
    Friend WithEvents MAPEditObjective As System.Windows.Forms.CheckBox
    Friend WithEvents MAPEditAgreements As System.Windows.Forms.CheckBox
    Friend WithEvents MAPEditFirms As System.Windows.Forms.CheckBox
    Friend WithEvents MAPEditProductTypes As System.Windows.Forms.CheckBox
    Friend WithEvents MAPEditProducts As System.Windows.Forms.CheckBox
    Friend WithEvents MAPAddObjective As System.Windows.Forms.CheckBox
    Friend WithEvents MAPAddProductType As System.Windows.Forms.CheckBox
    Friend WithEvents MAPAddProducts As System.Windows.Forms.CheckBox
    Friend WithEvents MAPAddFirm As System.Windows.Forms.CheckBox
    Friend WithEvents MAPAddAgreement As System.Windows.Forms.CheckBox
    Friend WithEvents ViewReports As System.Windows.Forms.CheckBox
    Friend WithEvents ETFPriceImport As System.Windows.Forms.CheckBox
    Friend WithEvents GroupBox4 As System.Windows.Forms.GroupBox
    Friend WithEvents LaunchRevenueCenter As System.Windows.Forms.CheckBox
    Friend WithEvents TabPage6 As System.Windows.Forms.TabPage
    Friend WithEvents LeadManagerAccess As System.Windows.Forms.CheckBox
    Friend WithEvents SalesManagementAccess As System.Windows.Forms.CheckBox
    Friend WithEvents TeamManagerAccess As System.Windows.Forms.CheckBox
    Friend WithEvents ViewRepBusinessReport As System.Windows.Forms.CheckBox
    Friend WithEvents AMPSecCodes As System.Windows.Forms.CheckBox
    Friend WithEvents AMPAccess As System.Windows.Forms.CheckBox
    Friend WithEvents EditAccountMovesSettings As System.Windows.Forms.CheckBox
    Friend WithEvents ViewAccountMoves As System.Windows.Forms.CheckBox
    Friend WithEvents EditWhatsNew As System.Windows.Forms.CheckBox
End Class
