<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class ETF_PriceImport_SecData
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(ETF_PriceImport_SecData))
        Me.Label1 = New System.Windows.Forms.Label
        Me.ID = New System.Windows.Forms.TextBox
        Me.Label2 = New System.Windows.Forms.Label
        Me.ComboBox1 = New System.Windows.Forms.ComboBox
        Me.lblSecName = New System.Windows.Forms.Label
        Me.MaturityDate = New System.Windows.Forms.DateTimePicker
        Me.Label4 = New System.Windows.Forms.Label
        Me.Label5 = New System.Windows.Forms.Label
        Me.InterestRate = New System.Windows.Forms.TextBox
        Me.PaymentFreq = New System.Windows.Forms.TextBox
        Me.Label6 = New System.Windows.Forms.Label
        Me.Rating = New System.Windows.Forms.TextBox
        Me.Label7 = New System.Windows.Forms.Label
        Me.AvgLife = New System.Windows.Forms.TextBox
        Me.Label8 = New System.Windows.Forms.Label
        Me.YTW = New System.Windows.Forms.TextBox
        Me.Label9 = New System.Windows.Forms.Label
        Me.Duration = New System.Windows.Forms.TextBox
        Me.Label10 = New System.Windows.Forms.Label
        Me.CMOResetRule = New System.Windows.Forms.TextBox
        Me.Label11 = New System.Windows.Forms.Label
        Me.PrimarySymbolType = New System.Windows.Forms.TextBox
        Me.Label12 = New System.Windows.Forms.Label
        Me.Button1 = New System.Windows.Forms.Button
        Me.SecType = New System.Windows.Forms.ComboBox
        Me.Label3 = New System.Windows.Forms.Label
        Me.NewSymbol = New System.Windows.Forms.TextBox
        Me.Label13 = New System.Windows.Forms.Label
        Me.AssetClass = New System.Windows.Forms.ComboBox
        Me.Label14 = New System.Windows.Forms.Label
        Me.Button2 = New System.Windows.Forms.Button
        Me.Label15 = New System.Windows.Forms.Label
        Me.Rating2 = New System.Windows.Forms.TextBox
        Me.Label16 = New System.Windows.Forms.Label
        Me.TextBox1 = New System.Windows.Forms.TextBox
        Me.TextBox2 = New System.Windows.Forms.TextBox
        Me.TextBox3 = New System.Windows.Forms.TextBox
        Me.TextBox4 = New System.Windows.Forms.TextBox
        Me.TextBox5 = New System.Windows.Forms.TextBox
        Me.TextBox6 = New System.Windows.Forms.TextBox
        Me.TextBox7 = New System.Windows.Forms.TextBox
        Me.SuspendLayout()
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(10, 15)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(21, 13)
        Me.Label1.TabIndex = 0
        Me.Label1.Text = "ID:"
        '
        'ID
        '
        Me.ID.Location = New System.Drawing.Point(37, 12)
        Me.ID.Name = "ID"
        Me.ID.ReadOnly = True
        Me.ID.Size = New System.Drawing.Size(100, 20)
        Me.ID.TabIndex = 1
        Me.ID.TabStop = False
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(10, 41)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(93, 13)
        Me.Label2.TabIndex = 2
        Me.Label2.Text = "Selected Security:"
        '
        'ComboBox1
        '
        Me.ComboBox1.AutoCompleteMode = System.Windows.Forms.AutoCompleteMode.SuggestAppend
        Me.ComboBox1.FormattingEnabled = True
        Me.ComboBox1.Location = New System.Drawing.Point(109, 38)
        Me.ComboBox1.Name = "ComboBox1"
        Me.ComboBox1.Size = New System.Drawing.Size(225, 21)
        Me.ComboBox1.TabIndex = 1
        '
        'lblSecName
        '
        Me.lblSecName.AutoSize = True
        Me.lblSecName.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblSecName.Location = New System.Drawing.Point(12, 71)
        Me.lblSecName.Name = "lblSecName"
        Me.lblSecName.Size = New System.Drawing.Size(157, 20)
        Me.lblSecName.TabIndex = 4
        Me.lblSecName.Text = "SELECT SECURITY"
        '
        'MaturityDate
        '
        Me.MaturityDate.Format = System.Windows.Forms.DateTimePickerFormat.[Short]
        Me.MaturityDate.Location = New System.Drawing.Point(93, 186)
        Me.MaturityDate.Name = "MaturityDate"
        Me.MaturityDate.Size = New System.Drawing.Size(112, 20)
        Me.MaturityDate.TabIndex = 4
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(14, 192)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(73, 13)
        Me.Label4.TabIndex = 6
        Me.Label4.Text = "Maturity Date:"
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Location = New System.Drawing.Point(14, 216)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(71, 13)
        Me.Label5.TabIndex = 7
        Me.Label5.Text = "Interest Rate:"
        '
        'InterestRate
        '
        Me.InterestRate.Location = New System.Drawing.Point(91, 213)
        Me.InterestRate.Name = "InterestRate"
        Me.InterestRate.Size = New System.Drawing.Size(156, 20)
        Me.InterestRate.TabIndex = 8
        Me.InterestRate.TabStop = False
        '
        'PaymentFreq
        '
        Me.PaymentFreq.Location = New System.Drawing.Point(124, 239)
        Me.PaymentFreq.Name = "PaymentFreq"
        Me.PaymentFreq.Size = New System.Drawing.Size(156, 20)
        Me.PaymentFreq.TabIndex = 10
        Me.PaymentFreq.TabStop = False
        Me.PaymentFreq.Text = "12"
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Location = New System.Drawing.Point(14, 242)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(104, 13)
        Me.Label6.TabIndex = 9
        Me.Label6.Text = "Payment Frequency:"
        '
        'Rating
        '
        Me.Rating.Location = New System.Drawing.Point(78, 265)
        Me.Rating.Name = "Rating"
        Me.Rating.Size = New System.Drawing.Size(156, 20)
        Me.Rating.TabIndex = 5
        '
        'Label7
        '
        Me.Label7.AutoSize = True
        Me.Label7.Location = New System.Drawing.Point(14, 268)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(58, 13)
        Me.Label7.TabIndex = 11
        Me.Label7.Text = "SP Rating:"
        '
        'AvgLife
        '
        Me.AvgLife.Location = New System.Drawing.Point(68, 316)
        Me.AvgLife.Name = "AvgLife"
        Me.AvgLife.Size = New System.Drawing.Size(156, 20)
        Me.AvgLife.TabIndex = 7
        '
        'Label8
        '
        Me.Label8.AutoSize = True
        Me.Label8.Location = New System.Drawing.Point(13, 319)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(49, 13)
        Me.Label8.TabIndex = 13
        Me.Label8.Text = "Avg Life:"
        '
        'YTW
        '
        Me.YTW.Location = New System.Drawing.Point(51, 342)
        Me.YTW.Name = "YTW"
        Me.YTW.Size = New System.Drawing.Size(156, 20)
        Me.YTW.TabIndex = 8
        '
        'Label9
        '
        Me.Label9.AutoSize = True
        Me.Label9.Location = New System.Drawing.Point(13, 345)
        Me.Label9.Name = "Label9"
        Me.Label9.Size = New System.Drawing.Size(32, 13)
        Me.Label9.TabIndex = 15
        Me.Label9.Text = "YTW"
        '
        'Duration
        '
        Me.Duration.Location = New System.Drawing.Point(69, 368)
        Me.Duration.Name = "Duration"
        Me.Duration.Size = New System.Drawing.Size(156, 20)
        Me.Duration.TabIndex = 9
        '
        'Label10
        '
        Me.Label10.AutoSize = True
        Me.Label10.Location = New System.Drawing.Point(13, 371)
        Me.Label10.Name = "Label10"
        Me.Label10.Size = New System.Drawing.Size(50, 13)
        Me.Label10.TabIndex = 17
        Me.Label10.Text = "Duration:"
        '
        'CMOResetRule
        '
        Me.CMOResetRule.Location = New System.Drawing.Point(152, 394)
        Me.CMOResetRule.Name = "CMOResetRule"
        Me.CMOResetRule.Size = New System.Drawing.Size(156, 20)
        Me.CMOResetRule.TabIndex = 20
        Me.CMOResetRule.TabStop = False
        Me.CMOResetRule.Text = "i"
        '
        'Label11
        '
        Me.Label11.AutoSize = True
        Me.Label11.Location = New System.Drawing.Point(13, 397)
        Me.Label11.Name = "Label11"
        Me.Label11.Size = New System.Drawing.Size(133, 13)
        Me.Label11.TabIndex = 19
        Me.Label11.Text = "CMO Payment Type Code:"
        '
        'PrimarySymbolType
        '
        Me.PrimarySymbolType.Location = New System.Drawing.Point(127, 420)
        Me.PrimarySymbolType.Name = "PrimarySymbolType"
        Me.PrimarySymbolType.Size = New System.Drawing.Size(156, 20)
        Me.PrimarySymbolType.TabIndex = 22
        Me.PrimarySymbolType.TabStop = False
        Me.PrimarySymbolType.Text = "T"
        '
        'Label12
        '
        Me.Label12.AutoSize = True
        Me.Label12.Location = New System.Drawing.Point(13, 423)
        Me.Label12.Name = "Label12"
        Me.Label12.Size = New System.Drawing.Size(108, 13)
        Me.Label12.TabIndex = 21
        Me.Label12.Text = "Primary Symbol Type:"
        '
        'Button1
        '
        Me.Button1.Location = New System.Drawing.Point(618, 15)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(75, 23)
        Me.Button1.TabIndex = 23
        Me.Button1.Text = "Save"
        Me.Button1.UseVisualStyleBackColor = True
        '
        'SecType
        '
        Me.SecType.AutoCompleteMode = System.Windows.Forms.AutoCompleteMode.SuggestAppend
        Me.SecType.FormattingEnabled = True
        Me.SecType.Location = New System.Drawing.Point(93, 103)
        Me.SecType.Name = "SecType"
        Me.SecType.Size = New System.Drawing.Size(225, 21)
        Me.SecType.TabIndex = 2
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(10, 106)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(78, 13)
        Me.Label3.TabIndex = 24
        Me.Label3.Text = "New SecType:"
        '
        'NewSymbol
        '
        Me.NewSymbol.Location = New System.Drawing.Point(78, 130)
        Me.NewSymbol.Name = "NewSymbol"
        Me.NewSymbol.Size = New System.Drawing.Size(156, 20)
        Me.NewSymbol.TabIndex = 27
        Me.NewSymbol.TabStop = False
        '
        'Label13
        '
        Me.Label13.AutoSize = True
        Me.Label13.Location = New System.Drawing.Point(10, 133)
        Me.Label13.Name = "Label13"
        Me.Label13.Size = New System.Drawing.Size(62, 13)
        Me.Label13.TabIndex = 26
        Me.Label13.Text = "New Ticker"
        '
        'AssetClass
        '
        Me.AssetClass.AutoCompleteMode = System.Windows.Forms.AutoCompleteMode.SuggestAppend
        Me.AssetClass.FormattingEnabled = True
        Me.AssetClass.Location = New System.Drawing.Point(80, 159)
        Me.AssetClass.Name = "AssetClass"
        Me.AssetClass.Size = New System.Drawing.Size(225, 21)
        Me.AssetClass.TabIndex = 3
        '
        'Label14
        '
        Me.Label14.AutoSize = True
        Me.Label14.Location = New System.Drawing.Point(10, 162)
        Me.Label14.Name = "Label14"
        Me.Label14.Size = New System.Drawing.Size(64, 13)
        Me.Label14.TabIndex = 28
        Me.Label14.Text = "Asset Class:"
        '
        'Button2
        '
        Me.Button2.Location = New System.Drawing.Point(233, 314)
        Me.Button2.Name = "Button2"
        Me.Button2.Size = New System.Drawing.Size(75, 23)
        Me.Button2.TabIndex = 29
        Me.Button2.Text = "What Date?"
        Me.Button2.UseVisualStyleBackColor = True
        '
        'Label15
        '
        Me.Label15.AutoSize = True
        Me.Label15.Location = New System.Drawing.Point(314, 319)
        Me.Label15.Name = "Label15"
        Me.Label15.Size = New System.Drawing.Size(45, 13)
        Me.Label15.TabIndex = 30
        Me.Label15.Text = "Label15"
        '
        'Rating2
        '
        Me.Rating2.Location = New System.Drawing.Point(91, 290)
        Me.Rating2.Name = "Rating2"
        Me.Rating2.Size = New System.Drawing.Size(156, 20)
        Me.Rating2.TabIndex = 6
        '
        'Label16
        '
        Me.Label16.AutoSize = True
        Me.Label16.Location = New System.Drawing.Point(14, 293)
        Me.Label16.Name = "Label16"
        Me.Label16.Size = New System.Drawing.Size(76, 13)
        Me.Label16.TabIndex = 32
        Me.Label16.Text = "Moody Rating:"
        '
        'TextBox1
        '
        Me.TextBox1.Location = New System.Drawing.Point(560, 99)
        Me.TextBox1.Name = "TextBox1"
        Me.TextBox1.Size = New System.Drawing.Size(57, 20)
        Me.TextBox1.TabIndex = 33
        '
        'TextBox2
        '
        Me.TextBox2.Location = New System.Drawing.Point(560, 125)
        Me.TextBox2.Name = "TextBox2"
        Me.TextBox2.Size = New System.Drawing.Size(57, 20)
        Me.TextBox2.TabIndex = 35
        '
        'TextBox3
        '
        Me.TextBox3.Location = New System.Drawing.Point(560, 151)
        Me.TextBox3.Name = "TextBox3"
        Me.TextBox3.Size = New System.Drawing.Size(57, 20)
        Me.TextBox3.TabIndex = 37
        '
        'TextBox4
        '
        Me.TextBox4.Location = New System.Drawing.Point(560, 177)
        Me.TextBox4.Name = "TextBox4"
        Me.TextBox4.Size = New System.Drawing.Size(57, 20)
        Me.TextBox4.TabIndex = 39
        '
        'TextBox5
        '
        Me.TextBox5.Location = New System.Drawing.Point(560, 203)
        Me.TextBox5.Name = "TextBox5"
        Me.TextBox5.Size = New System.Drawing.Size(57, 20)
        Me.TextBox5.TabIndex = 40
        '
        'TextBox6
        '
        Me.TextBox6.Location = New System.Drawing.Point(560, 229)
        Me.TextBox6.Name = "TextBox6"
        Me.TextBox6.Size = New System.Drawing.Size(57, 20)
        Me.TextBox6.TabIndex = 41
        '
        'TextBox7
        '
        Me.TextBox7.Location = New System.Drawing.Point(560, 255)
        Me.TextBox7.Name = "TextBox7"
        Me.TextBox7.Size = New System.Drawing.Size(57, 20)
        Me.TextBox7.TabIndex = 42
        '
        'ETF_PriceImport_SecData
        '
        Me.AcceptButton = Me.Button1
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.Color.White
        Me.ClientSize = New System.Drawing.Size(705, 450)
        Me.Controls.Add(Me.TextBox7)
        Me.Controls.Add(Me.TextBox6)
        Me.Controls.Add(Me.TextBox5)
        Me.Controls.Add(Me.TextBox4)
        Me.Controls.Add(Me.TextBox3)
        Me.Controls.Add(Me.TextBox2)
        Me.Controls.Add(Me.TextBox1)
        Me.Controls.Add(Me.Rating2)
        Me.Controls.Add(Me.Label16)
        Me.Controls.Add(Me.Label15)
        Me.Controls.Add(Me.Button2)
        Me.Controls.Add(Me.AssetClass)
        Me.Controls.Add(Me.Label14)
        Me.Controls.Add(Me.NewSymbol)
        Me.Controls.Add(Me.Label13)
        Me.Controls.Add(Me.SecType)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.Button1)
        Me.Controls.Add(Me.PrimarySymbolType)
        Me.Controls.Add(Me.Label12)
        Me.Controls.Add(Me.CMOResetRule)
        Me.Controls.Add(Me.Label11)
        Me.Controls.Add(Me.Duration)
        Me.Controls.Add(Me.Label10)
        Me.Controls.Add(Me.YTW)
        Me.Controls.Add(Me.Label9)
        Me.Controls.Add(Me.AvgLife)
        Me.Controls.Add(Me.Label8)
        Me.Controls.Add(Me.Rating)
        Me.Controls.Add(Me.Label7)
        Me.Controls.Add(Me.PaymentFreq)
        Me.Controls.Add(Me.Label6)
        Me.Controls.Add(Me.InterestRate)
        Me.Controls.Add(Me.Label5)
        Me.Controls.Add(Me.Label4)
        Me.Controls.Add(Me.MaturityDate)
        Me.Controls.Add(Me.lblSecName)
        Me.Controls.Add(Me.ComboBox1)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.ID)
        Me.Controls.Add(Me.Label1)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Name = "ETF_PriceImport_SecData"
        Me.Text = "ETF Security Data"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents ID As System.Windows.Forms.TextBox
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents ComboBox1 As System.Windows.Forms.ComboBox
    Friend WithEvents lblSecName As System.Windows.Forms.Label
    Friend WithEvents MaturityDate As System.Windows.Forms.DateTimePicker
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents InterestRate As System.Windows.Forms.TextBox
    Friend WithEvents PaymentFreq As System.Windows.Forms.TextBox
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents Rating As System.Windows.Forms.TextBox
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents AvgLife As System.Windows.Forms.TextBox
    Friend WithEvents Label8 As System.Windows.Forms.Label
    Friend WithEvents YTW As System.Windows.Forms.TextBox
    Friend WithEvents Label9 As System.Windows.Forms.Label
    Friend WithEvents Duration As System.Windows.Forms.TextBox
    Friend WithEvents Label10 As System.Windows.Forms.Label
    Friend WithEvents CMOResetRule As System.Windows.Forms.TextBox
    Friend WithEvents Label11 As System.Windows.Forms.Label
    Friend WithEvents PrimarySymbolType As System.Windows.Forms.TextBox
    Friend WithEvents Label12 As System.Windows.Forms.Label
    Friend WithEvents Button1 As System.Windows.Forms.Button
    Friend WithEvents SecType As System.Windows.Forms.ComboBox
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents NewSymbol As System.Windows.Forms.TextBox
    Friend WithEvents Label13 As System.Windows.Forms.Label
    Friend WithEvents AssetClass As System.Windows.Forms.ComboBox
    Friend WithEvents Label14 As System.Windows.Forms.Label
    Friend WithEvents Button2 As System.Windows.Forms.Button
    Friend WithEvents Label15 As System.Windows.Forms.Label
    Friend WithEvents Rating2 As System.Windows.Forms.TextBox
    Friend WithEvents Label16 As System.Windows.Forms.Label
    Friend WithEvents TextBox1 As System.Windows.Forms.TextBox
    Friend WithEvents TextBox2 As System.Windows.Forms.TextBox
    Friend WithEvents TextBox3 As System.Windows.Forms.TextBox
    Friend WithEvents TextBox4 As System.Windows.Forms.TextBox
    Friend WithEvents TextBox5 As System.Windows.Forms.TextBox
    Friend WithEvents TextBox6 As System.Windows.Forms.TextBox
    Friend WithEvents TextBox7 As System.Windows.Forms.TextBox
End Class
