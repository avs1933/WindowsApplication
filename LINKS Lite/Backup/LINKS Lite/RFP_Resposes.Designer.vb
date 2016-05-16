<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class RFP_Resposes
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(RFP_Resposes))
        Me.txtID = New System.Windows.Forms.TextBox
        Me.Label1 = New System.Windows.Forms.Label
        Me.Label2 = New System.Windows.Forms.Label
        Me.rtbQuestion = New System.Windows.Forms.RichTextBox
        Me.ContextMenuStrip1 = New System.Windows.Forms.ContextMenuStrip(Me.components)
        Me.CopyToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem
        Me.PasteToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem
        Me.rtbAnswer = New System.Windows.Forms.RichTextBox
        Me.ContextMenuStrip2 = New System.Windows.Forms.ContextMenuStrip(Me.components)
        Me.ToolStripMenuItem1 = New System.Windows.Forms.ToolStripMenuItem
        Me.ToolStripMenuItem2 = New System.Windows.Forms.ToolStripMenuItem
        Me.Label3 = New System.Windows.Forms.Label
        Me.cbPeople = New System.Windows.Forms.CheckBox
        Me.cbProduct = New System.Windows.Forms.CheckBox
        Me.cbComposite = New System.Windows.Forms.CheckBox
        Me.cbStale = New System.Windows.Forms.CheckBox
        Me.cboPeople = New System.Windows.Forms.ComboBox
        Me.cboProduct = New System.Windows.Forms.ComboBox
        Me.cboComposite = New System.Windows.Forms.ComboBox
        Me.DTPStale = New System.Windows.Forms.DateTimePicker
        Me.lblStale = New System.Windows.Forms.Label
        Me.cmdSave = New System.Windows.Forms.Button
        Me.cmdCancel = New System.Windows.Forms.Button
        Me.Label5 = New System.Windows.Forms.Label
        Me.Label4 = New System.Windows.Forms.Label
        Me.Label6 = New System.Windows.Forms.Label
        Me.Label7 = New System.Windows.Forms.Label
        Me.lblLevel = New System.Windows.Forms.Label
        Me.cboLevel = New System.Windows.Forms.ComboBox
        Me.txtDiscipline = New System.Windows.Forms.TextBox
        Me.txtContactID = New System.Windows.Forms.TextBox
        Me.txtCompositeID = New System.Windows.Forms.TextBox
        Me.txtLevelID = New System.Windows.Forms.TextBox
        Me.RichTextBox1 = New System.Windows.Forms.RichTextBox
        Me.RichTextBox2 = New System.Windows.Forms.RichTextBox
        Me.txtStaleDate = New System.Windows.Forms.TextBox
        Me.ckbPeople = New System.Windows.Forms.CheckBox
        Me.ckbProduct = New System.Windows.Forms.CheckBox
        Me.ckbComposite = New System.Windows.Forms.CheckBox
        Me.ckbDate1 = New System.Windows.Forms.CheckBox
        Me.Button1 = New System.Windows.Forms.Button
        Me.ckbActive = New System.Windows.Forms.CheckBox
        Me.Button2 = New System.Windows.Forms.Button
        Me.lblDeleted = New System.Windows.Forms.Label
        Me.Label8 = New System.Windows.Forms.Label
        Me.txtRptID = New System.Windows.Forms.TextBox
        Me.Label9 = New System.Windows.Forms.Label
        Me.cboStage = New System.Windows.Forms.ComboBox
        Me.Label10 = New System.Windows.Forms.Label
        Me.txtStageID = New System.Windows.Forms.TextBox
        Me.ContextMenuStrip1.SuspendLayout()
        Me.ContextMenuStrip2.SuspendLayout()
        Me.SuspendLayout()
        '
        'txtID
        '
        Me.txtID.Enabled = False
        Me.txtID.Location = New System.Drawing.Point(84, 12)
        Me.txtID.Name = "txtID"
        Me.txtID.Size = New System.Drawing.Size(100, 20)
        Me.txtID.TabIndex = 0
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(12, 15)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(66, 13)
        Me.Label1.TabIndex = 1
        Me.Label1.Text = "Question ID:"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(12, 97)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(52, 13)
        Me.Label2.TabIndex = 2
        Me.Label2.Text = "Question:"
        '
        'rtbQuestion
        '
        Me.rtbQuestion.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.rtbQuestion.ContextMenuStrip = Me.ContextMenuStrip1
        Me.rtbQuestion.Location = New System.Drawing.Point(15, 113)
        Me.rtbQuestion.Name = "rtbQuestion"
        Me.rtbQuestion.Size = New System.Drawing.Size(552, 86)
        Me.rtbQuestion.TabIndex = 3
        Me.rtbQuestion.Text = ""
        '
        'ContextMenuStrip1
        '
        Me.ContextMenuStrip1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.CopyToolStripMenuItem, Me.PasteToolStripMenuItem})
        Me.ContextMenuStrip1.Name = "ContextMenuStrip1"
        Me.ContextMenuStrip1.Size = New System.Drawing.Size(103, 48)
        '
        'CopyToolStripMenuItem
        '
        Me.CopyToolStripMenuItem.Name = "CopyToolStripMenuItem"
        Me.CopyToolStripMenuItem.Size = New System.Drawing.Size(102, 22)
        Me.CopyToolStripMenuItem.Text = "&Copy"
        '
        'PasteToolStripMenuItem
        '
        Me.PasteToolStripMenuItem.Name = "PasteToolStripMenuItem"
        Me.PasteToolStripMenuItem.Size = New System.Drawing.Size(102, 22)
        Me.PasteToolStripMenuItem.Text = "&Paste"
        '
        'rtbAnswer
        '
        Me.rtbAnswer.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.rtbAnswer.ContextMenuStrip = Me.ContextMenuStrip2
        Me.rtbAnswer.Location = New System.Drawing.Point(15, 226)
        Me.rtbAnswer.Name = "rtbAnswer"
        Me.rtbAnswer.Size = New System.Drawing.Size(552, 137)
        Me.rtbAnswer.TabIndex = 5
        Me.rtbAnswer.Text = ""
        '
        'ContextMenuStrip2
        '
        Me.ContextMenuStrip2.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.ToolStripMenuItem1, Me.ToolStripMenuItem2})
        Me.ContextMenuStrip2.Name = "ContextMenuStrip1"
        Me.ContextMenuStrip2.Size = New System.Drawing.Size(103, 48)
        '
        'ToolStripMenuItem1
        '
        Me.ToolStripMenuItem1.Name = "ToolStripMenuItem1"
        Me.ToolStripMenuItem1.Size = New System.Drawing.Size(102, 22)
        Me.ToolStripMenuItem1.Text = "&Copy"
        '
        'ToolStripMenuItem2
        '
        Me.ToolStripMenuItem2.Name = "ToolStripMenuItem2"
        Me.ToolStripMenuItem2.Size = New System.Drawing.Size(102, 22)
        Me.ToolStripMenuItem2.Text = "&Paste"
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(12, 210)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(45, 13)
        Me.Label3.TabIndex = 4
        Me.Label3.Text = "Answer:"
        '
        'cbPeople
        '
        Me.cbPeople.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.cbPeople.AutoSize = True
        Me.cbPeople.Location = New System.Drawing.Point(15, 376)
        Me.cbPeople.Name = "cbPeople"
        Me.cbPeople.Size = New System.Drawing.Size(111, 17)
        Me.cbPeople.TabIndex = 6
        Me.cbPeople.Text = "Related to People"
        Me.cbPeople.UseVisualStyleBackColor = True
        '
        'cbProduct
        '
        Me.cbProduct.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.cbProduct.AutoSize = True
        Me.cbProduct.Location = New System.Drawing.Point(14, 403)
        Me.cbProduct.Name = "cbProduct"
        Me.cbProduct.Size = New System.Drawing.Size(115, 17)
        Me.cbProduct.TabIndex = 7
        Me.cbProduct.Text = "Related to Product"
        Me.cbProduct.UseVisualStyleBackColor = True
        '
        'cbComposite
        '
        Me.cbComposite.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.cbComposite.AutoSize = True
        Me.cbComposite.Location = New System.Drawing.Point(14, 430)
        Me.cbComposite.Name = "cbComposite"
        Me.cbComposite.Size = New System.Drawing.Size(127, 17)
        Me.cbComposite.TabIndex = 8
        Me.cbComposite.Text = "Related to Composite"
        Me.cbComposite.UseVisualStyleBackColor = True
        '
        'cbStale
        '
        Me.cbStale.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.cbStale.AutoSize = True
        Me.cbStale.Location = New System.Drawing.Point(15, 459)
        Me.cbStale.Name = "cbStale"
        Me.cbStale.Size = New System.Drawing.Size(95, 17)
        Me.cbStale.TabIndex = 9
        Me.cbStale.Text = "Date Sensitive"
        Me.cbStale.UseVisualStyleBackColor = True
        '
        'cboPeople
        '
        Me.cboPeople.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.cboPeople.AutoCompleteMode = System.Windows.Forms.AutoCompleteMode.SuggestAppend
        Me.cboPeople.AutoCompleteSource = System.Windows.Forms.AutoCompleteSource.CustomSource
        Me.cboPeople.FormattingEnabled = True
        Me.cboPeople.Location = New System.Drawing.Point(147, 374)
        Me.cboPeople.Name = "cboPeople"
        Me.cboPeople.Size = New System.Drawing.Size(260, 21)
        Me.cboPeople.TabIndex = 10
        Me.cboPeople.Visible = False
        '
        'cboProduct
        '
        Me.cboProduct.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.cboProduct.AutoCompleteMode = System.Windows.Forms.AutoCompleteMode.SuggestAppend
        Me.cboProduct.FormattingEnabled = True
        Me.cboProduct.Location = New System.Drawing.Point(147, 401)
        Me.cboProduct.Name = "cboProduct"
        Me.cboProduct.Size = New System.Drawing.Size(260, 21)
        Me.cboProduct.TabIndex = 11
        Me.cboProduct.Visible = False
        '
        'cboComposite
        '
        Me.cboComposite.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.cboComposite.AutoCompleteMode = System.Windows.Forms.AutoCompleteMode.SuggestAppend
        Me.cboComposite.FormattingEnabled = True
        Me.cboComposite.Location = New System.Drawing.Point(147, 428)
        Me.cboComposite.Name = "cboComposite"
        Me.cboComposite.Size = New System.Drawing.Size(260, 21)
        Me.cboComposite.TabIndex = 12
        Me.cboComposite.Visible = False
        '
        'DTPStale
        '
        Me.DTPStale.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.DTPStale.Format = System.Windows.Forms.DateTimePickerFormat.[Short]
        Me.DTPStale.Location = New System.Drawing.Point(297, 456)
        Me.DTPStale.Name = "DTPStale"
        Me.DTPStale.Size = New System.Drawing.Size(110, 20)
        Me.DTPStale.TabIndex = 13
        Me.DTPStale.Visible = False
        '
        'lblStale
        '
        Me.lblStale.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.lblStale.AutoSize = True
        Me.lblStale.Location = New System.Drawing.Point(144, 460)
        Me.lblStale.Name = "lblStale"
        Me.lblStale.Size = New System.Drawing.Size(147, 13)
        Me.lblStale.TabIndex = 14
        Me.lblStale.Text = "Date question becomes stale:"
        Me.lblStale.Visible = False
        '
        'cmdSave
        '
        Me.cmdSave.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmdSave.Location = New System.Drawing.Point(473, 9)
        Me.cmdSave.Name = "cmdSave"
        Me.cmdSave.Size = New System.Drawing.Size(94, 23)
        Me.cmdSave.TabIndex = 21
        Me.cmdSave.Text = "&Save"
        '
        'cmdCancel
        '
        Me.cmdCancel.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmdCancel.Location = New System.Drawing.Point(367, 10)
        Me.cmdCancel.Name = "cmdCancel"
        Me.cmdCancel.Size = New System.Drawing.Size(94, 23)
        Me.cmdCancel.TabIndex = 20
        Me.cmdCancel.Text = "Delete"
        '
        'Label5
        '
        Me.Label5.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.Label5.AutoEllipsis = True
        Me.Label5.AutoSize = True
        Me.Label5.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label5.Location = New System.Drawing.Point(230, 375)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(73, 20)
        Me.Label5.TabIndex = 22
        Me.Label5.Text = "Loading"
        Me.Label5.Visible = False
        '
        'Label4
        '
        Me.Label4.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.Label4.AutoEllipsis = True
        Me.Label4.AutoSize = True
        Me.Label4.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label4.Location = New System.Drawing.Point(230, 402)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(73, 20)
        Me.Label4.TabIndex = 23
        Me.Label4.Text = "Loading"
        Me.Label4.Visible = False
        '
        'Label6
        '
        Me.Label6.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.Label6.AutoEllipsis = True
        Me.Label6.AutoSize = True
        Me.Label6.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label6.Location = New System.Drawing.Point(230, 430)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(73, 20)
        Me.Label6.TabIndex = 24
        Me.Label6.Text = "Loading"
        Me.Label6.Visible = False
        '
        'Label7
        '
        Me.Label7.AutoSize = True
        Me.Label7.Location = New System.Drawing.Point(12, 41)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(36, 13)
        Me.Label7.TabIndex = 25
        Me.Label7.Text = "Level:"
        '
        'lblLevel
        '
        Me.lblLevel.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.lblLevel.AutoEllipsis = True
        Me.lblLevel.AutoSize = True
        Me.lblLevel.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblLevel.Location = New System.Drawing.Point(167, 39)
        Me.lblLevel.Name = "lblLevel"
        Me.lblLevel.Size = New System.Drawing.Size(73, 20)
        Me.lblLevel.TabIndex = 27
        Me.lblLevel.Text = "Loading"
        Me.lblLevel.Visible = False
        '
        'cboLevel
        '
        Me.cboLevel.AutoCompleteMode = System.Windows.Forms.AutoCompleteMode.SuggestAppend
        Me.cboLevel.Enabled = False
        Me.cboLevel.FormattingEnabled = True
        Me.cboLevel.Location = New System.Drawing.Point(84, 38)
        Me.cboLevel.Name = "cboLevel"
        Me.cboLevel.Size = New System.Drawing.Size(260, 21)
        Me.cboLevel.TabIndex = 26
        '
        'txtDiscipline
        '
        Me.txtDiscipline.Location = New System.Drawing.Point(413, 398)
        Me.txtDiscipline.Name = "txtDiscipline"
        Me.txtDiscipline.Size = New System.Drawing.Size(47, 20)
        Me.txtDiscipline.TabIndex = 28
        Me.txtDiscipline.Visible = False
        '
        'txtContactID
        '
        Me.txtContactID.Location = New System.Drawing.Point(413, 372)
        Me.txtContactID.Name = "txtContactID"
        Me.txtContactID.Size = New System.Drawing.Size(47, 20)
        Me.txtContactID.TabIndex = 29
        Me.txtContactID.Visible = False
        '
        'txtCompositeID
        '
        Me.txtCompositeID.Location = New System.Drawing.Point(413, 422)
        Me.txtCompositeID.Name = "txtCompositeID"
        Me.txtCompositeID.Size = New System.Drawing.Size(47, 20)
        Me.txtCompositeID.TabIndex = 30
        Me.txtCompositeID.Visible = False
        '
        'txtLevelID
        '
        Me.txtLevelID.Location = New System.Drawing.Point(314, 10)
        Me.txtLevelID.Name = "txtLevelID"
        Me.txtLevelID.Size = New System.Drawing.Size(47, 20)
        Me.txtLevelID.TabIndex = 31
        Me.txtLevelID.Visible = False
        '
        'RichTextBox1
        '
        Me.RichTextBox1.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.RichTextBox1.Location = New System.Drawing.Point(499, 85)
        Me.RichTextBox1.Name = "RichTextBox1"
        Me.RichTextBox1.Size = New System.Drawing.Size(68, 98)
        Me.RichTextBox1.TabIndex = 32
        Me.RichTextBox1.Text = ""
        Me.RichTextBox1.Visible = False
        '
        'RichTextBox2
        '
        Me.RichTextBox2.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.RichTextBox2.Location = New System.Drawing.Point(499, 198)
        Me.RichTextBox2.Name = "RichTextBox2"
        Me.RichTextBox2.Size = New System.Drawing.Size(68, 98)
        Me.RichTextBox2.TabIndex = 33
        Me.RichTextBox2.Text = ""
        Me.RichTextBox2.Visible = False
        '
        'txtStaleDate
        '
        Me.txtStaleDate.Location = New System.Drawing.Point(413, 453)
        Me.txtStaleDate.Name = "txtStaleDate"
        Me.txtStaleDate.Size = New System.Drawing.Size(47, 20)
        Me.txtStaleDate.TabIndex = 34
        Me.txtStaleDate.Visible = False
        '
        'ckbPeople
        '
        Me.ckbPeople.AutoSize = True
        Me.ckbPeople.Location = New System.Drawing.Point(466, 378)
        Me.ckbPeople.Name = "ckbPeople"
        Me.ckbPeople.Size = New System.Drawing.Size(15, 14)
        Me.ckbPeople.TabIndex = 35
        Me.ckbPeople.UseVisualStyleBackColor = True
        Me.ckbPeople.Visible = False
        '
        'ckbProduct
        '
        Me.ckbProduct.AutoSize = True
        Me.ckbProduct.Location = New System.Drawing.Point(466, 400)
        Me.ckbProduct.Name = "ckbProduct"
        Me.ckbProduct.Size = New System.Drawing.Size(15, 14)
        Me.ckbProduct.TabIndex = 36
        Me.ckbProduct.UseVisualStyleBackColor = True
        Me.ckbProduct.Visible = False
        '
        'ckbComposite
        '
        Me.ckbComposite.AutoSize = True
        Me.ckbComposite.Location = New System.Drawing.Point(466, 424)
        Me.ckbComposite.Name = "ckbComposite"
        Me.ckbComposite.Size = New System.Drawing.Size(15, 14)
        Me.ckbComposite.TabIndex = 37
        Me.ckbComposite.UseVisualStyleBackColor = True
        Me.ckbComposite.Visible = False
        '
        'ckbDate1
        '
        Me.ckbDate1.AutoSize = True
        Me.ckbDate1.Location = New System.Drawing.Point(466, 456)
        Me.ckbDate1.Name = "ckbDate1"
        Me.ckbDate1.Size = New System.Drawing.Size(15, 14)
        Me.ckbDate1.TabIndex = 38
        Me.ckbDate1.UseVisualStyleBackColor = True
        Me.ckbDate1.Visible = False
        '
        'Button1
        '
        Me.Button1.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Button1.Location = New System.Drawing.Point(473, 36)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(94, 23)
        Me.Button1.TabIndex = 39
        Me.Button1.Text = "&Cancel"
        '
        'ckbActive
        '
        Me.ckbActive.AutoSize = True
        Me.ckbActive.Location = New System.Drawing.Point(535, 454)
        Me.ckbActive.Name = "ckbActive"
        Me.ckbActive.Size = New System.Drawing.Size(15, 14)
        Me.ckbActive.TabIndex = 40
        Me.ckbActive.UseVisualStyleBackColor = True
        Me.ckbActive.Visible = False
        '
        'Button2
        '
        Me.Button2.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Button2.Location = New System.Drawing.Point(367, 36)
        Me.Button2.Name = "Button2"
        Me.Button2.Size = New System.Drawing.Size(94, 23)
        Me.Button2.TabIndex = 41
        Me.Button2.Text = "Change Log"
        '
        'lblDeleted
        '
        Me.lblDeleted.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.lblDeleted.AutoEllipsis = True
        Me.lblDeleted.AutoSize = True
        Me.lblDeleted.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblDeleted.ForeColor = System.Drawing.Color.Red
        Me.lblDeleted.Location = New System.Drawing.Point(230, 9)
        Me.lblDeleted.Name = "lblDeleted"
        Me.lblDeleted.Size = New System.Drawing.Size(91, 20)
        Me.lblDeleted.TabIndex = 42
        Me.lblDeleted.Text = "DELETED"
        Me.lblDeleted.Visible = False
        '
        'Label8
        '
        Me.Label8.AutoSize = True
        Me.Label8.Location = New System.Drawing.Point(12, 68)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(56, 13)
        Me.Label8.TabIndex = 44
        Me.Label8.Text = "Report ID:"
        '
        'txtRptID
        '
        Me.txtRptID.Enabled = False
        Me.txtRptID.Location = New System.Drawing.Point(84, 65)
        Me.txtRptID.Name = "txtRptID"
        Me.txtRptID.Size = New System.Drawing.Size(100, 20)
        Me.txtRptID.TabIndex = 43
        '
        'Label9
        '
        Me.Label9.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Label9.AutoEllipsis = True
        Me.Label9.AutoSize = True
        Me.Label9.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label9.Location = New System.Drawing.Point(344, 69)
        Me.Label9.Name = "Label9"
        Me.Label9.Size = New System.Drawing.Size(73, 20)
        Me.Label9.TabIndex = 47
        Me.Label9.Text = "Loading"
        Me.Label9.Visible = False
        '
        'cboStage
        '
        Me.cboStage.AutoCompleteMode = System.Windows.Forms.AutoCompleteMode.SuggestAppend
        Me.cboStage.Enabled = False
        Me.cboStage.FormattingEnabled = True
        Me.cboStage.Location = New System.Drawing.Point(276, 68)
        Me.cboStage.Name = "cboStage"
        Me.cboStage.Size = New System.Drawing.Size(217, 21)
        Me.cboStage.TabIndex = 46
        '
        'Label10
        '
        Me.Label10.AutoSize = True
        Me.Label10.Location = New System.Drawing.Point(231, 71)
        Me.Label10.Name = "Label10"
        Me.Label10.Size = New System.Drawing.Size(38, 13)
        Me.Label10.TabIndex = 45
        Me.Label10.Text = "Stage:"
        '
        'txtStageID
        '
        Me.txtStageID.Location = New System.Drawing.Point(274, 85)
        Me.txtStageID.Name = "txtStageID"
        Me.txtStageID.Size = New System.Drawing.Size(47, 20)
        Me.txtStageID.TabIndex = 48
        Me.txtStageID.Visible = False
        '
        'RFP_Resposes
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.Color.White
        Me.ClientSize = New System.Drawing.Size(579, 494)
        Me.Controls.Add(Me.txtStageID)
        Me.Controls.Add(Me.Label9)
        Me.Controls.Add(Me.cboStage)
        Me.Controls.Add(Me.Label10)
        Me.Controls.Add(Me.Label8)
        Me.Controls.Add(Me.txtRptID)
        Me.Controls.Add(Me.lblDeleted)
        Me.Controls.Add(Me.Button2)
        Me.Controls.Add(Me.Button1)
        Me.Controls.Add(Me.RichTextBox2)
        Me.Controls.Add(Me.ckbActive)
        Me.Controls.Add(Me.RichTextBox1)
        Me.Controls.Add(Me.ckbDate1)
        Me.Controls.Add(Me.ckbComposite)
        Me.Controls.Add(Me.ckbProduct)
        Me.Controls.Add(Me.ckbPeople)
        Me.Controls.Add(Me.txtStaleDate)
        Me.Controls.Add(Me.txtLevelID)
        Me.Controls.Add(Me.lblLevel)
        Me.Controls.Add(Me.cboLevel)
        Me.Controls.Add(Me.txtCompositeID)
        Me.Controls.Add(Me.txtContactID)
        Me.Controls.Add(Me.txtDiscipline)
        Me.Controls.Add(Me.Label7)
        Me.Controls.Add(Me.Label6)
        Me.Controls.Add(Me.Label4)
        Me.Controls.Add(Me.Label5)
        Me.Controls.Add(Me.cmdSave)
        Me.Controls.Add(Me.cmdCancel)
        Me.Controls.Add(Me.lblStale)
        Me.Controls.Add(Me.DTPStale)
        Me.Controls.Add(Me.cboComposite)
        Me.Controls.Add(Me.cboProduct)
        Me.Controls.Add(Me.cboPeople)
        Me.Controls.Add(Me.cbStale)
        Me.Controls.Add(Me.cbComposite)
        Me.Controls.Add(Me.cbProduct)
        Me.Controls.Add(Me.cbPeople)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.txtID)
        Me.Controls.Add(Me.rtbAnswer)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.rtbQuestion)
        Me.Controls.Add(Me.Label2)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Name = "RFP_Resposes"
        Me.Text = "RFP Central - Question Engine"
        Me.ContextMenuStrip1.ResumeLayout(False)
        Me.ContextMenuStrip2.ResumeLayout(False)
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents txtID As System.Windows.Forms.TextBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents rtbQuestion As System.Windows.Forms.RichTextBox
    Friend WithEvents rtbAnswer As System.Windows.Forms.RichTextBox
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents cbPeople As System.Windows.Forms.CheckBox
    Friend WithEvents cbProduct As System.Windows.Forms.CheckBox
    Friend WithEvents cbComposite As System.Windows.Forms.CheckBox
    Friend WithEvents cbStale As System.Windows.Forms.CheckBox
    Friend WithEvents cboPeople As System.Windows.Forms.ComboBox
    Friend WithEvents cboProduct As System.Windows.Forms.ComboBox
    Friend WithEvents cboComposite As System.Windows.Forms.ComboBox
    Friend WithEvents DTPStale As System.Windows.Forms.DateTimePicker
    Friend WithEvents lblStale As System.Windows.Forms.Label
    Friend WithEvents cmdSave As System.Windows.Forms.Button
    Friend WithEvents cmdCancel As System.Windows.Forms.Button
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents lblLevel As System.Windows.Forms.Label
    Friend WithEvents cboLevel As System.Windows.Forms.ComboBox
    Friend WithEvents txtDiscipline As System.Windows.Forms.TextBox
    Friend WithEvents txtContactID As System.Windows.Forms.TextBox
    Friend WithEvents txtCompositeID As System.Windows.Forms.TextBox
    Friend WithEvents txtLevelID As System.Windows.Forms.TextBox
    Friend WithEvents RichTextBox1 As System.Windows.Forms.RichTextBox
    Friend WithEvents RichTextBox2 As System.Windows.Forms.RichTextBox
    Friend WithEvents txtStaleDate As System.Windows.Forms.TextBox
    Friend WithEvents ckbPeople As System.Windows.Forms.CheckBox
    Friend WithEvents ckbProduct As System.Windows.Forms.CheckBox
    Friend WithEvents ckbComposite As System.Windows.Forms.CheckBox
    Friend WithEvents ckbDate1 As System.Windows.Forms.CheckBox
    Friend WithEvents Button1 As System.Windows.Forms.Button
    Friend WithEvents ckbActive As System.Windows.Forms.CheckBox
    Friend WithEvents Button2 As System.Windows.Forms.Button
    Friend WithEvents lblDeleted As System.Windows.Forms.Label
    Friend WithEvents ContextMenuStrip1 As System.Windows.Forms.ContextMenuStrip
    Friend WithEvents CopyToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents PasteToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents ContextMenuStrip2 As System.Windows.Forms.ContextMenuStrip
    Friend WithEvents ToolStripMenuItem1 As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents ToolStripMenuItem2 As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents Label8 As System.Windows.Forms.Label
    Friend WithEvents txtRptID As System.Windows.Forms.TextBox
    Friend WithEvents Label9 As System.Windows.Forms.Label
    Friend WithEvents cboStage As System.Windows.Forms.ComboBox
    Friend WithEvents Label10 As System.Windows.Forms.Label
    Friend WithEvents txtStageID As System.Windows.Forms.TextBox
End Class
