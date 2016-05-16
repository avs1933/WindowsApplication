<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FirmOverview
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
        Me.components = New System.ComponentModel.Container
        Dim DataGridViewCellStyle1 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(FirmOverview))
        Me.ContactID = New System.Windows.Forms.TextBox
        Me.Label2 = New System.Windows.Forms.Label
        Me.Label3 = New System.Windows.Forms.Label
        Me.ContactCode = New System.Windows.Forms.TextBox
        Me.GroupBox1 = New System.Windows.Forms.GroupBox
        Me.GroupBox2 = New System.Windows.Forms.GroupBox
        Me.GroupBox3 = New System.Windows.Forms.GroupBox
        Me.GroupBox4 = New System.Windows.Forms.GroupBox
        Me.FirmName = New System.Windows.Forms.TextBox
        Me.PlatformOnly = New System.Windows.Forms.CheckBox
        Me.AdditionalPlatformApproval = New System.Windows.Forms.CheckBox
        Me.UMA = New System.Windows.Forms.CheckBox
        Me.AgreementSigned = New System.Windows.Forms.CheckBox
        Me.EPaperwork = New System.Windows.Forms.CheckBox
        Me.SubAdvised = New System.Windows.Forms.CheckBox
        Me.Solicited = New System.Windows.Forms.CheckBox
        Me.Advised = New System.Windows.Forms.CheckBox
        Me.CustomFeeSchedule = New System.Windows.Forms.TextBox
        Me.Timer1 = New System.Windows.Forms.Timer(Me.components)
        Me.ConservativeTaxable = New System.Windows.Forms.CheckBox
        Me.CorePlus = New System.Windows.Forms.CheckBox
        Me.MortgageInvestmentPortfolio = New System.Windows.Forms.CheckBox
        Me.CreditOpportunities = New System.Windows.Forms.CheckBox
        Me.ShortTermTaxExempt = New System.Windows.Forms.CheckBox
        Me.CoreTaxExempt = New System.Windows.Forms.CheckBox
        Me.High50Dividend = New System.Windows.Forms.CheckBox
        Me.PeroniMethod = New System.Windows.Forms.CheckBox
        Me.ERISA = New System.Windows.Forms.CheckBox
        Me.TAMTaxable = New System.Windows.Forms.CheckBox
        Me.TAMTaxExempt = New System.Windows.Forms.CheckBox
        Me.Label1 = New System.Windows.Forms.Label
        Me.AgreementVersion = New System.Windows.Forms.TextBox
        Me.Label4 = New System.Windows.Forms.Label
        Me.DateExecuted = New System.Windows.Forms.TextBox
        Me.Label5 = New System.Windows.Forms.Label
        Me.Platforms = New System.Windows.Forms.TextBox
        Me.Label6 = New System.Windows.Forms.Label
        Me.PlatformsWhiteLabel = New System.Windows.Forms.TextBox
        Me.Label7 = New System.Windows.Forms.Label
        Me.Address = New System.Windows.Forms.RichTextBox
        Me.Label8 = New System.Windows.Forms.Label
        Me.URL = New System.Windows.Forms.TextBox
        Me.Label9 = New System.Windows.Forms.Label
        Me.Phone = New System.Windows.Forms.TextBox
        Me.Label10 = New System.Windows.Forms.Label
        Me.Fax = New System.Windows.Forms.TextBox
        Me.GroupBox6 = New System.Windows.Forms.GroupBox
        Me.Notes = New System.Windows.Forms.RichTextBox
        Me.Label11 = New System.Windows.Forms.Label
        Me.Label12 = New System.Windows.Forms.Label
        Me.Label13 = New System.Windows.Forms.Label
        Me.Label14 = New System.Windows.Forms.Label
        Me.Label15 = New System.Windows.Forms.Label
        Me.CheckBox1 = New System.Windows.Forms.CheckBox
        Me.DataGridView1 = New System.Windows.Forms.DataGridView
        Me.OK = New System.Windows.Forms.Button
        Me.ContextMenuStrip1 = New System.Windows.Forms.ContextMenuStrip(Me.components)
        Me.EditToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem
        Me.DeleteToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem
        Me.CheckBox2 = New System.Windows.Forms.CheckBox
        Me.ViewAPX = New System.Windows.Forms.Button
        Me.Button1 = New System.Windows.Forms.Button
        Me.GroupBox1.SuspendLayout()
        Me.GroupBox2.SuspendLayout()
        Me.GroupBox3.SuspendLayout()
        Me.GroupBox4.SuspendLayout()
        Me.GroupBox6.SuspendLayout()
        CType(Me.DataGridView1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.ContextMenuStrip1.SuspendLayout()
        Me.SuspendLayout()
        '
        'ContactID
        '
        Me.ContactID.BackColor = System.Drawing.Color.White
        Me.ContactID.Location = New System.Drawing.Point(90, 43)
        Me.ContactID.Name = "ContactID"
        Me.ContactID.ReadOnly = True
        Me.ContactID.Size = New System.Drawing.Size(126, 20)
        Me.ContactID.TabIndex = 1
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(26, 46)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(58, 13)
        Me.Label2.TabIndex = 2
        Me.Label2.Text = "Advent ID:"
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(9, 72)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(75, 13)
        Me.Label3.TabIndex = 4
        Me.Label3.Text = "Contact Code:"
        '
        'ContactCode
        '
        Me.ContactCode.BackColor = System.Drawing.Color.White
        Me.ContactCode.Location = New System.Drawing.Point(90, 69)
        Me.ContactCode.Name = "ContactCode"
        Me.ContactCode.ReadOnly = True
        Me.ContactCode.Size = New System.Drawing.Size(188, 20)
        Me.ContactCode.TabIndex = 3
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.Label11)
        Me.GroupBox1.Controls.Add(Me.TAMTaxExempt)
        Me.GroupBox1.Controls.Add(Me.TAMTaxable)
        Me.GroupBox1.Controls.Add(Me.PeroniMethod)
        Me.GroupBox1.Controls.Add(Me.High50Dividend)
        Me.GroupBox1.Controls.Add(Me.CoreTaxExempt)
        Me.GroupBox1.Controls.Add(Me.ShortTermTaxExempt)
        Me.GroupBox1.Controls.Add(Me.CreditOpportunities)
        Me.GroupBox1.Controls.Add(Me.MortgageInvestmentPortfolio)
        Me.GroupBox1.Controls.Add(Me.CorePlus)
        Me.GroupBox1.Controls.Add(Me.ConservativeTaxable)
        Me.GroupBox1.Location = New System.Drawing.Point(21, 111)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(414, 207)
        Me.GroupBox1.TabIndex = 5
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "Approvals"
        '
        'GroupBox2
        '
        Me.GroupBox2.Controls.Add(Me.Button1)
        Me.GroupBox2.Controls.Add(Me.CheckBox2)
        Me.GroupBox2.Controls.Add(Me.OK)
        Me.GroupBox2.Controls.Add(Me.DataGridView1)
        Me.GroupBox2.Controls.Add(Me.Label13)
        Me.GroupBox2.Location = New System.Drawing.Point(441, 111)
        Me.GroupBox2.Name = "GroupBox2"
        Me.GroupBox2.Size = New System.Drawing.Size(687, 207)
        Me.GroupBox2.TabIndex = 6
        Me.GroupBox2.TabStop = False
        Me.GroupBox2.Text = "Fee Info"
        '
        'GroupBox3
        '
        Me.GroupBox3.Controls.Add(Me.Label12)
        Me.GroupBox3.Controls.Add(Me.Label10)
        Me.GroupBox3.Controls.Add(Me.Fax)
        Me.GroupBox3.Controls.Add(Me.Label9)
        Me.GroupBox3.Controls.Add(Me.Phone)
        Me.GroupBox3.Controls.Add(Me.Label8)
        Me.GroupBox3.Controls.Add(Me.URL)
        Me.GroupBox3.Controls.Add(Me.Address)
        Me.GroupBox3.Controls.Add(Me.Label7)
        Me.GroupBox3.Controls.Add(Me.Label6)
        Me.GroupBox3.Controls.Add(Me.PlatformsWhiteLabel)
        Me.GroupBox3.Controls.Add(Me.Label5)
        Me.GroupBox3.Controls.Add(Me.Platforms)
        Me.GroupBox3.Controls.Add(Me.Label4)
        Me.GroupBox3.Controls.Add(Me.DateExecuted)
        Me.GroupBox3.Controls.Add(Me.Label1)
        Me.GroupBox3.Controls.Add(Me.AgreementVersion)
        Me.GroupBox3.Controls.Add(Me.ERISA)
        Me.GroupBox3.Controls.Add(Me.EPaperwork)
        Me.GroupBox3.Controls.Add(Me.PlatformOnly)
        Me.GroupBox3.Controls.Add(Me.UMA)
        Me.GroupBox3.Controls.Add(Me.AgreementSigned)
        Me.GroupBox3.Controls.Add(Me.AdditionalPlatformApproval)
        Me.GroupBox3.Location = New System.Drawing.Point(21, 324)
        Me.GroupBox3.Name = "GroupBox3"
        Me.GroupBox3.Size = New System.Drawing.Size(388, 271)
        Me.GroupBox3.TabIndex = 7
        Me.GroupBox3.TabStop = False
        Me.GroupBox3.Text = "Firm Info"
        '
        'GroupBox4
        '
        Me.GroupBox4.Controls.Add(Me.Label15)
        Me.GroupBox4.Location = New System.Drawing.Point(417, 324)
        Me.GroupBox4.Name = "GroupBox4"
        Me.GroupBox4.Size = New System.Drawing.Size(475, 271)
        Me.GroupBox4.TabIndex = 8
        Me.GroupBox4.TabStop = False
        Me.GroupBox4.Text = "Marketing Support"
        '
        'FirmName
        '
        Me.FirmName.BackColor = System.Drawing.Color.White
        Me.FirmName.BorderStyle = System.Windows.Forms.BorderStyle.None
        Me.FirmName.Font = New System.Drawing.Font("Microsoft Sans Serif", 21.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.FirmName.Location = New System.Drawing.Point(12, 3)
        Me.FirmName.Name = "FirmName"
        Me.FirmName.ReadOnly = True
        Me.FirmName.Size = New System.Drawing.Size(1122, 33)
        Me.FirmName.TabIndex = 9
        Me.FirmName.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'PlatformOnly
        '
        Me.PlatformOnly.AutoSize = True
        Me.PlatformOnly.Location = New System.Drawing.Point(8, 19)
        Me.PlatformOnly.Name = "PlatformOnly"
        Me.PlatformOnly.Size = New System.Drawing.Size(88, 17)
        Me.PlatformOnly.TabIndex = 0
        Me.PlatformOnly.Text = "Platform Only"
        Me.PlatformOnly.UseVisualStyleBackColor = True
        '
        'AdditionalPlatformApproval
        '
        Me.AdditionalPlatformApproval.AutoSize = True
        Me.AdditionalPlatformApproval.Location = New System.Drawing.Point(102, 19)
        Me.AdditionalPlatformApproval.Name = "AdditionalPlatformApproval"
        Me.AdditionalPlatformApproval.Size = New System.Drawing.Size(158, 17)
        Me.AdditionalPlatformApproval.TabIndex = 1
        Me.AdditionalPlatformApproval.Text = "Additional Platform Approval"
        Me.AdditionalPlatformApproval.UseVisualStyleBackColor = True
        '
        'UMA
        '
        Me.UMA.AutoSize = True
        Me.UMA.Location = New System.Drawing.Point(8, 42)
        Me.UMA.Name = "UMA"
        Me.UMA.Size = New System.Drawing.Size(50, 17)
        Me.UMA.TabIndex = 2
        Me.UMA.Text = "UMA"
        Me.UMA.UseVisualStyleBackColor = True
        '
        'AgreementSigned
        '
        Me.AgreementSigned.AutoSize = True
        Me.AgreementSigned.Location = New System.Drawing.Point(269, 19)
        Me.AgreementSigned.Name = "AgreementSigned"
        Me.AgreementSigned.Size = New System.Drawing.Size(113, 17)
        Me.AgreementSigned.TabIndex = 3
        Me.AgreementSigned.Text = "Agreement Signed"
        Me.AgreementSigned.UseVisualStyleBackColor = True
        '
        'EPaperwork
        '
        Me.EPaperwork.AutoSize = True
        Me.EPaperwork.Location = New System.Drawing.Point(102, 42)
        Me.EPaperwork.Name = "EPaperwork"
        Me.EPaperwork.Size = New System.Drawing.Size(121, 17)
        Me.EPaperwork.TabIndex = 4
        Me.EPaperwork.Text = "E-Paperwork Ready"
        Me.EPaperwork.UseVisualStyleBackColor = True
        '
        'SubAdvised
        '
        Me.SubAdvised.AutoSize = True
        Me.SubAdvised.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.SubAdvised.Location = New System.Drawing.Point(322, 51)
        Me.SubAdvised.Name = "SubAdvised"
        Me.SubAdvised.Size = New System.Drawing.Size(129, 24)
        Me.SubAdvised.TabIndex = 5
        Me.SubAdvised.Text = "Sub-Advised"
        Me.SubAdvised.UseVisualStyleBackColor = True
        '
        'Solicited
        '
        Me.Solicited.AutoSize = True
        Me.Solicited.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Solicited.Location = New System.Drawing.Point(673, 51)
        Me.Solicited.Name = "Solicited"
        Me.Solicited.Size = New System.Drawing.Size(97, 24)
        Me.Solicited.TabIndex = 6
        Me.Solicited.Text = "Solicited"
        Me.Solicited.UseVisualStyleBackColor = True
        '
        'Advised
        '
        Me.Advised.AutoSize = True
        Me.Advised.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Advised.Location = New System.Drawing.Point(1021, 51)
        Me.Advised.Name = "Advised"
        Me.Advised.Size = New System.Drawing.Size(91, 24)
        Me.Advised.TabIndex = 7
        Me.Advised.Text = "Advised"
        Me.Advised.UseVisualStyleBackColor = True
        '
        'CustomFeeSchedule
        '
        Me.CustomFeeSchedule.BackColor = System.Drawing.Color.White
        Me.CustomFeeSchedule.BorderStyle = System.Windows.Forms.BorderStyle.None
        Me.CustomFeeSchedule.Font = New System.Drawing.Font("Microsoft Sans Serif", 18.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.CustomFeeSchedule.ForeColor = System.Drawing.Color.Red
        Me.CustomFeeSchedule.Location = New System.Drawing.Point(284, 77)
        Me.CustomFeeSchedule.Name = "CustomFeeSchedule"
        Me.CustomFeeSchedule.ReadOnly = True
        Me.CustomFeeSchedule.Size = New System.Drawing.Size(850, 28)
        Me.CustomFeeSchedule.TabIndex = 10
        Me.CustomFeeSchedule.Text = "*** CUSTOM FEE SCHEDULE ***"
        Me.CustomFeeSchedule.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        Me.CustomFeeSchedule.Visible = False
        '
        'Timer1
        '
        Me.Timer1.Enabled = True
        Me.Timer1.Interval = 500
        '
        'ConservativeTaxable
        '
        Me.ConservativeTaxable.AutoSize = True
        Me.ConservativeTaxable.Location = New System.Drawing.Point(8, 19)
        Me.ConservativeTaxable.Name = "ConservativeTaxable"
        Me.ConservativeTaxable.Size = New System.Drawing.Size(129, 17)
        Me.ConservativeTaxable.TabIndex = 0
        Me.ConservativeTaxable.Text = "Conservative Taxable"
        Me.ConservativeTaxable.UseVisualStyleBackColor = True
        '
        'CorePlus
        '
        Me.CorePlus.AutoSize = True
        Me.CorePlus.Location = New System.Drawing.Point(8, 42)
        Me.CorePlus.Name = "CorePlus"
        Me.CorePlus.Size = New System.Drawing.Size(71, 17)
        Me.CorePlus.TabIndex = 1
        Me.CorePlus.Text = "Core Plus"
        Me.CorePlus.UseVisualStyleBackColor = True
        '
        'MortgageInvestmentPortfolio
        '
        Me.MortgageInvestmentPortfolio.AutoSize = True
        Me.MortgageInvestmentPortfolio.Location = New System.Drawing.Point(8, 65)
        Me.MortgageInvestmentPortfolio.Name = "MortgageInvestmentPortfolio"
        Me.MortgageInvestmentPortfolio.Size = New System.Drawing.Size(167, 17)
        Me.MortgageInvestmentPortfolio.TabIndex = 2
        Me.MortgageInvestmentPortfolio.Text = "Mortgage Investment Portfolio"
        Me.MortgageInvestmentPortfolio.UseVisualStyleBackColor = True
        '
        'CreditOpportunities
        '
        Me.CreditOpportunities.AutoSize = True
        Me.CreditOpportunities.Location = New System.Drawing.Point(8, 88)
        Me.CreditOpportunities.Name = "CreditOpportunities"
        Me.CreditOpportunities.Size = New System.Drawing.Size(118, 17)
        Me.CreditOpportunities.TabIndex = 3
        Me.CreditOpportunities.Text = "Credit Opportunities"
        Me.CreditOpportunities.UseVisualStyleBackColor = True
        '
        'ShortTermTaxExempt
        '
        Me.ShortTermTaxExempt.AutoSize = True
        Me.ShortTermTaxExempt.Location = New System.Drawing.Point(8, 111)
        Me.ShortTermTaxExempt.Name = "ShortTermTaxExempt"
        Me.ShortTermTaxExempt.Size = New System.Drawing.Size(137, 17)
        Me.ShortTermTaxExempt.TabIndex = 4
        Me.ShortTermTaxExempt.Text = "Short Term Tax-Exempt"
        Me.ShortTermTaxExempt.UseVisualStyleBackColor = True
        '
        'CoreTaxExempt
        '
        Me.CoreTaxExempt.AutoSize = True
        Me.CoreTaxExempt.Location = New System.Drawing.Point(8, 134)
        Me.CoreTaxExempt.Name = "CoreTaxExempt"
        Me.CoreTaxExempt.Size = New System.Drawing.Size(107, 17)
        Me.CoreTaxExempt.TabIndex = 5
        Me.CoreTaxExempt.Text = "Core Tax-Exempt"
        Me.CoreTaxExempt.UseVisualStyleBackColor = True
        '
        'High50Dividend
        '
        Me.High50Dividend.AutoSize = True
        Me.High50Dividend.Location = New System.Drawing.Point(8, 157)
        Me.High50Dividend.Name = "High50Dividend"
        Me.High50Dividend.Size = New System.Drawing.Size(108, 17)
        Me.High50Dividend.TabIndex = 6
        Me.High50Dividend.Text = "High 50 Dividend"
        Me.High50Dividend.UseVisualStyleBackColor = True
        '
        'PeroniMethod
        '
        Me.PeroniMethod.AutoSize = True
        Me.PeroniMethod.Location = New System.Drawing.Point(8, 180)
        Me.PeroniMethod.Name = "PeroniMethod"
        Me.PeroniMethod.Size = New System.Drawing.Size(92, 17)
        Me.PeroniMethod.TabIndex = 7
        Me.PeroniMethod.Text = "PeroniMethod"
        Me.PeroniMethod.UseVisualStyleBackColor = True
        '
        'ERISA
        '
        Me.ERISA.AutoSize = True
        Me.ERISA.Location = New System.Drawing.Point(269, 42)
        Me.ERISA.Name = "ERISA"
        Me.ERISA.Size = New System.Drawing.Size(58, 17)
        Me.ERISA.TabIndex = 5
        Me.ERISA.Text = "ERISA"
        Me.ERISA.UseVisualStyleBackColor = True
        '
        'TAMTaxable
        '
        Me.TAMTaxable.AutoSize = True
        Me.TAMTaxable.Location = New System.Drawing.Point(190, 19)
        Me.TAMTaxable.Name = "TAMTaxable"
        Me.TAMTaxable.Size = New System.Drawing.Size(189, 17)
        Me.TAMTaxable.TabIndex = 8
        Me.TAMTaxable.Text = "Tactical Allocation Model: Taxable"
        Me.TAMTaxable.UseVisualStyleBackColor = True
        '
        'TAMTaxExempt
        '
        Me.TAMTaxExempt.AutoSize = True
        Me.TAMTaxExempt.Location = New System.Drawing.Point(190, 42)
        Me.TAMTaxExempt.Name = "TAMTaxExempt"
        Me.TAMTaxExempt.Size = New System.Drawing.Size(207, 17)
        Me.TAMTaxExempt.TabIndex = 9
        Me.TAMTaxExempt.Text = "Tactical Allocation Model: Tax-Exempt"
        Me.TAMTaxExempt.UseVisualStyleBackColor = True
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(7, 68)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(99, 13)
        Me.Label1.TabIndex = 7
        Me.Label1.Text = "Agreement Version:"
        '
        'AgreementVersion
        '
        Me.AgreementVersion.BackColor = System.Drawing.Color.White
        Me.AgreementVersion.Location = New System.Drawing.Point(8, 84)
        Me.AgreementVersion.Name = "AgreementVersion"
        Me.AgreementVersion.ReadOnly = True
        Me.AgreementVersion.Size = New System.Drawing.Size(126, 20)
        Me.AgreementVersion.TabIndex = 6
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(142, 68)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(81, 13)
        Me.Label4.TabIndex = 9
        Me.Label4.Text = "Date Executed:"
        '
        'DateExecuted
        '
        Me.DateExecuted.BackColor = System.Drawing.Color.White
        Me.DateExecuted.Location = New System.Drawing.Point(145, 84)
        Me.DateExecuted.Name = "DateExecuted"
        Me.DateExecuted.ReadOnly = True
        Me.DateExecuted.Size = New System.Drawing.Size(78, 20)
        Me.DateExecuted.TabIndex = 8
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Location = New System.Drawing.Point(226, 68)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(59, 13)
        Me.Label5.TabIndex = 11
        Me.Label5.Text = "Platform(s):"
        '
        'Platforms
        '
        Me.Platforms.BackColor = System.Drawing.Color.White
        Me.Platforms.Location = New System.Drawing.Point(229, 84)
        Me.Platforms.Name = "Platforms"
        Me.Platforms.ReadOnly = True
        Me.Platforms.Size = New System.Drawing.Size(150, 20)
        Me.Platforms.TabIndex = 10
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Location = New System.Drawing.Point(7, 106)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(119, 13)
        Me.Label6.TabIndex = 13
        Me.Label6.Text = "Platform(s) White Label:"
        '
        'PlatformsWhiteLabel
        '
        Me.PlatformsWhiteLabel.BackColor = System.Drawing.Color.White
        Me.PlatformsWhiteLabel.Location = New System.Drawing.Point(10, 122)
        Me.PlatformsWhiteLabel.Name = "PlatformsWhiteLabel"
        Me.PlatformsWhiteLabel.ReadOnly = True
        Me.PlatformsWhiteLabel.Size = New System.Drawing.Size(175, 20)
        Me.PlatformsWhiteLabel.TabIndex = 12
        '
        'Label7
        '
        Me.Label7.AutoSize = True
        Me.Label7.Location = New System.Drawing.Point(7, 146)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(70, 13)
        Me.Label7.TabIndex = 15
        Me.Label7.Text = "Firm Address:"
        '
        'Address
        '
        Me.Address.Location = New System.Drawing.Point(10, 162)
        Me.Address.Name = "Address"
        Me.Address.ReadOnly = True
        Me.Address.Size = New System.Drawing.Size(238, 81)
        Me.Address.TabIndex = 16
        Me.Address.Text = ""
        '
        'Label8
        '
        Me.Label8.AutoSize = True
        Me.Label8.Location = New System.Drawing.Point(187, 106)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(71, 13)
        Me.Label8.TabIndex = 18
        Me.Label8.Text = "Firm Website:"
        '
        'URL
        '
        Me.URL.BackColor = System.Drawing.Color.White
        Me.URL.Location = New System.Drawing.Point(190, 122)
        Me.URL.Name = "URL"
        Me.URL.ReadOnly = True
        Me.URL.Size = New System.Drawing.Size(189, 20)
        Me.URL.TabIndex = 17
        '
        'Label9
        '
        Me.Label9.AutoSize = True
        Me.Label9.Location = New System.Drawing.Point(254, 145)
        Me.Label9.Name = "Label9"
        Me.Label9.Size = New System.Drawing.Size(86, 13)
        Me.Label9.TabIndex = 20
        Me.Label9.Text = "Business Phone:"
        '
        'Phone
        '
        Me.Phone.BackColor = System.Drawing.Color.White
        Me.Phone.Location = New System.Drawing.Point(257, 161)
        Me.Phone.Name = "Phone"
        Me.Phone.ReadOnly = True
        Me.Phone.Size = New System.Drawing.Size(122, 20)
        Me.Phone.TabIndex = 19
        '
        'Label10
        '
        Me.Label10.AutoSize = True
        Me.Label10.Location = New System.Drawing.Point(254, 191)
        Me.Label10.Name = "Label10"
        Me.Label10.Size = New System.Drawing.Size(70, 13)
        Me.Label10.TabIndex = 22
        Me.Label10.Text = "Other Phone:"
        '
        'Fax
        '
        Me.Fax.BackColor = System.Drawing.Color.White
        Me.Fax.Location = New System.Drawing.Point(257, 207)
        Me.Fax.Name = "Fax"
        Me.Fax.ReadOnly = True
        Me.Fax.Size = New System.Drawing.Size(122, 20)
        Me.Fax.TabIndex = 21
        '
        'GroupBox6
        '
        Me.GroupBox6.Controls.Add(Me.Label14)
        Me.GroupBox6.Controls.Add(Me.Notes)
        Me.GroupBox6.Location = New System.Drawing.Point(902, 324)
        Me.GroupBox6.Name = "GroupBox6"
        Me.GroupBox6.Size = New System.Drawing.Size(232, 271)
        Me.GroupBox6.TabIndex = 11
        Me.GroupBox6.TabStop = False
        Me.GroupBox6.Text = "Advent Notes"
        '
        'Notes
        '
        Me.Notes.Location = New System.Drawing.Point(6, 19)
        Me.Notes.Name = "Notes"
        Me.Notes.ReadOnly = True
        Me.Notes.Size = New System.Drawing.Size(220, 246)
        Me.Notes.TabIndex = 17
        Me.Notes.Text = ""
        '
        'Label11
        '
        Me.Label11.AutoSize = True
        Me.Label11.Font = New System.Drawing.Font("Microsoft Sans Serif", 14.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label11.Location = New System.Drawing.Point(156, 91)
        Me.Label11.Name = "Label11"
        Me.Label11.Size = New System.Drawing.Size(103, 24)
        Me.Label11.TabIndex = 16
        Me.Label11.Text = "Loading..."
        '
        'Label12
        '
        Me.Label12.AutoSize = True
        Me.Label12.Font = New System.Drawing.Font("Microsoft Sans Serif", 14.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label12.Location = New System.Drawing.Point(135, 122)
        Me.Label12.Name = "Label12"
        Me.Label12.Size = New System.Drawing.Size(103, 24)
        Me.Label12.TabIndex = 17
        Me.Label12.Text = "Loading..."
        '
        'Label13
        '
        Me.Label13.AutoSize = True
        Me.Label13.Font = New System.Drawing.Font("Microsoft Sans Serif", 14.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label13.Location = New System.Drawing.Point(185, 81)
        Me.Label13.Name = "Label13"
        Me.Label13.Size = New System.Drawing.Size(103, 24)
        Me.Label13.TabIndex = 23
        Me.Label13.Text = "Loading..."
        '
        'Label14
        '
        Me.Label14.AutoSize = True
        Me.Label14.Font = New System.Drawing.Font("Microsoft Sans Serif", 14.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label14.Location = New System.Drawing.Point(72, 122)
        Me.Label14.Name = "Label14"
        Me.Label14.Size = New System.Drawing.Size(103, 24)
        Me.Label14.TabIndex = 24
        Me.Label14.Text = "Loading..."
        '
        'Label15
        '
        Me.Label15.AutoSize = True
        Me.Label15.Font = New System.Drawing.Font("Microsoft Sans Serif", 14.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label15.Location = New System.Drawing.Point(179, 117)
        Me.Label15.Name = "Label15"
        Me.Label15.Size = New System.Drawing.Size(103, 24)
        Me.Label15.TabIndex = 25
        Me.Label15.Text = "Loading..."
        Me.Label15.Visible = False
        '
        'CheckBox1
        '
        Me.CheckBox1.AutoSize = True
        Me.CheckBox1.Location = New System.Drawing.Point(322, 82)
        Me.CheckBox1.Name = "CheckBox1"
        Me.CheckBox1.Size = New System.Drawing.Size(152, 17)
        Me.CheckBox1.TabIndex = 12
        Me.CheckBox1.Text = "Fee Schedule Label Driver"
        Me.CheckBox1.UseVisualStyleBackColor = True
        Me.CheckBox1.Visible = False
        '
        'DataGridView1
        '
        Me.DataGridView1.AllowUserToAddRows = False
        Me.DataGridView1.AllowUserToDeleteRows = False
        DataGridViewCellStyle1.BackColor = System.Drawing.Color.FromArgb(CType(CType(224, Byte), Integer), CType(CType(224, Byte), Integer), CType(CType(224, Byte), Integer))
        Me.DataGridView1.AlternatingRowsDefaultCellStyle = DataGridViewCellStyle1
        Me.DataGridView1.BackgroundColor = System.Drawing.Color.White
        Me.DataGridView1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.DataGridView1.ContextMenuStrip = Me.ContextMenuStrip1
        Me.DataGridView1.Location = New System.Drawing.Point(6, 20)
        Me.DataGridView1.Name = "DataGridView1"
        Me.DataGridView1.ReadOnly = True
        Me.DataGridView1.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect
        Me.DataGridView1.Size = New System.Drawing.Size(578, 177)
        Me.DataGridView1.TabIndex = 24
        '
        'OK
        '
        Me.OK.Location = New System.Drawing.Point(590, 111)
        Me.OK.Name = "OK"
        Me.OK.Size = New System.Drawing.Size(81, 86)
        Me.OK.TabIndex = 25
        Me.OK.Text = "Add"
        '
        'ContextMenuStrip1
        '
        Me.ContextMenuStrip1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.EditToolStripMenuItem, Me.DeleteToolStripMenuItem})
        Me.ContextMenuStrip1.Name = "ContextMenuStrip1"
        Me.ContextMenuStrip1.Size = New System.Drawing.Size(153, 70)
        '
        'EditToolStripMenuItem
        '
        Me.EditToolStripMenuItem.Name = "EditToolStripMenuItem"
        Me.EditToolStripMenuItem.Size = New System.Drawing.Size(152, 22)
        Me.EditToolStripMenuItem.Text = "&Edit"
        '
        'DeleteToolStripMenuItem
        '
        Me.DeleteToolStripMenuItem.Name = "DeleteToolStripMenuItem"
        Me.DeleteToolStripMenuItem.Size = New System.Drawing.Size(152, 22)
        Me.DeleteToolStripMenuItem.Text = "&Delete"
        '
        'CheckBox2
        '
        Me.CheckBox2.AutoSize = True
        Me.CheckBox2.Location = New System.Drawing.Point(6, 22)
        Me.CheckBox2.Name = "CheckBox2"
        Me.CheckBox2.Size = New System.Drawing.Size(114, 17)
        Me.CheckBox2.TabIndex = 27
        Me.CheckBox2.Text = "Loaded Fee Driver"
        Me.CheckBox2.UseVisualStyleBackColor = True
        Me.CheckBox2.Visible = False
        '
        'ViewAPX
        '
        Me.ViewAPX.Location = New System.Drawing.Point(946, 77)
        Me.ViewAPX.Name = "ViewAPX"
        Me.ViewAPX.Size = New System.Drawing.Size(182, 28)
        Me.ViewAPX.TabIndex = 27
        Me.ViewAPX.Text = "View Firm in APX"
        '
        'Button1
        '
        Me.Button1.Location = New System.Drawing.Point(590, 19)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(81, 86)
        Me.Button1.TabIndex = 28
        Me.Button1.Text = "Reload"
        '
        'FirmOverview
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.Color.White
        Me.ClientSize = New System.Drawing.Size(1146, 607)
        Me.Controls.Add(Me.ViewAPX)
        Me.Controls.Add(Me.CheckBox1)
        Me.Controls.Add(Me.GroupBox6)
        Me.Controls.Add(Me.CustomFeeSchedule)
        Me.Controls.Add(Me.Advised)
        Me.Controls.Add(Me.FirmName)
        Me.Controls.Add(Me.Solicited)
        Me.Controls.Add(Me.GroupBox4)
        Me.Controls.Add(Me.SubAdvised)
        Me.Controls.Add(Me.GroupBox3)
        Me.Controls.Add(Me.GroupBox2)
        Me.Controls.Add(Me.GroupBox1)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.ContactCode)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.ContactID)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "FirmOverview"
        Me.Text = "FirmOverview"
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox1.PerformLayout()
        Me.GroupBox2.ResumeLayout(False)
        Me.GroupBox2.PerformLayout()
        Me.GroupBox3.ResumeLayout(False)
        Me.GroupBox3.PerformLayout()
        Me.GroupBox4.ResumeLayout(False)
        Me.GroupBox4.PerformLayout()
        Me.GroupBox6.ResumeLayout(False)
        Me.GroupBox6.PerformLayout()
        CType(Me.DataGridView1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ContextMenuStrip1.ResumeLayout(False)
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents ContactID As System.Windows.Forms.TextBox
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents ContactCode As System.Windows.Forms.TextBox
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents GroupBox2 As System.Windows.Forms.GroupBox
    Friend WithEvents GroupBox3 As System.Windows.Forms.GroupBox
    Friend WithEvents GroupBox4 As System.Windows.Forms.GroupBox
    Friend WithEvents FirmName As System.Windows.Forms.TextBox
    Friend WithEvents AgreementSigned As System.Windows.Forms.CheckBox
    Friend WithEvents UMA As System.Windows.Forms.CheckBox
    Friend WithEvents AdditionalPlatformApproval As System.Windows.Forms.CheckBox
    Friend WithEvents PlatformOnly As System.Windows.Forms.CheckBox
    Friend WithEvents EPaperwork As System.Windows.Forms.CheckBox
    Friend WithEvents SubAdvised As System.Windows.Forms.CheckBox
    Friend WithEvents Solicited As System.Windows.Forms.CheckBox
    Friend WithEvents Advised As System.Windows.Forms.CheckBox
    Friend WithEvents CustomFeeSchedule As System.Windows.Forms.TextBox
    Friend WithEvents Timer1 As System.Windows.Forms.Timer
    Friend WithEvents PeroniMethod As System.Windows.Forms.CheckBox
    Friend WithEvents High50Dividend As System.Windows.Forms.CheckBox
    Friend WithEvents CoreTaxExempt As System.Windows.Forms.CheckBox
    Friend WithEvents ShortTermTaxExempt As System.Windows.Forms.CheckBox
    Friend WithEvents CreditOpportunities As System.Windows.Forms.CheckBox
    Friend WithEvents MortgageInvestmentPortfolio As System.Windows.Forms.CheckBox
    Friend WithEvents CorePlus As System.Windows.Forms.CheckBox
    Friend WithEvents ConservativeTaxable As System.Windows.Forms.CheckBox
    Friend WithEvents ERISA As System.Windows.Forms.CheckBox
    Friend WithEvents TAMTaxExempt As System.Windows.Forms.CheckBox
    Friend WithEvents TAMTaxable As System.Windows.Forms.CheckBox
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents Platforms As System.Windows.Forms.TextBox
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents DateExecuted As System.Windows.Forms.TextBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents AgreementVersion As System.Windows.Forms.TextBox
    Friend WithEvents Label10 As System.Windows.Forms.Label
    Friend WithEvents Fax As System.Windows.Forms.TextBox
    Friend WithEvents Label9 As System.Windows.Forms.Label
    Friend WithEvents Phone As System.Windows.Forms.TextBox
    Friend WithEvents Label8 As System.Windows.Forms.Label
    Friend WithEvents URL As System.Windows.Forms.TextBox
    Friend WithEvents Address As System.Windows.Forms.RichTextBox
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents PlatformsWhiteLabel As System.Windows.Forms.TextBox
    Friend WithEvents GroupBox6 As System.Windows.Forms.GroupBox
    Friend WithEvents Notes As System.Windows.Forms.RichTextBox
    Friend WithEvents Label11 As System.Windows.Forms.Label
    Friend WithEvents Label13 As System.Windows.Forms.Label
    Friend WithEvents Label12 As System.Windows.Forms.Label
    Friend WithEvents Label15 As System.Windows.Forms.Label
    Friend WithEvents Label14 As System.Windows.Forms.Label
    Friend WithEvents CheckBox1 As System.Windows.Forms.CheckBox
    Friend WithEvents DataGridView1 As System.Windows.Forms.DataGridView
    Friend WithEvents OK As System.Windows.Forms.Button
    Friend WithEvents ContextMenuStrip1 As System.Windows.Forms.ContextMenuStrip
    Friend WithEvents EditToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents DeleteToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents CheckBox2 As System.Windows.Forms.CheckBox
    Friend WithEvents ViewAPX As System.Windows.Forms.Button
    Friend WithEvents Button1 As System.Windows.Forms.Button
End Class
