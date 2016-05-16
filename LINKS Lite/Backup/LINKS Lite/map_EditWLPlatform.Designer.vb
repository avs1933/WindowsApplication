<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class map_EditWLPlatform
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
        Dim DataGridViewCellStyle2 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(map_EditWLPlatform))
        Me.Label1 = New System.Windows.Forms.Label
        Me.ID = New System.Windows.Forms.TextBox
        Me.WLName = New System.Windows.Forms.TextBox
        Me.Label2 = New System.Windows.Forms.Label
        Me.cboPlatform = New System.Windows.Forms.ComboBox
        Me.Label3 = New System.Windows.Forms.Label
        Me.GroupBox1 = New System.Windows.Forms.GroupBox
        Me.PlatformIDCopy = New System.Windows.Forms.TextBox
        Me.WLNameCopy = New System.Windows.Forms.TextBox
        Me.RichTextBox2 = New System.Windows.Forms.RichTextBox
        Me.cbcfcopy = New System.Windows.Forms.CheckBox
        Me.cb7copy = New System.Windows.Forms.CheckBox
        Me.cb6copy = New System.Windows.Forms.CheckBox
        Me.cb5copy = New System.Windows.Forms.CheckBox
        Me.cb4copy = New System.Windows.Forms.CheckBox
        Me.cb3copy = New System.Windows.Forms.CheckBox
        Me.cb2copy = New System.Windows.Forms.CheckBox
        Me.cb1copy = New System.Windows.Forms.CheckBox
        Me.ckbCustomFee = New System.Windows.Forms.CheckBox
        Me.CheckBox7 = New System.Windows.Forms.CheckBox
        Me.CheckBox6 = New System.Windows.Forms.CheckBox
        Me.CheckBox5 = New System.Windows.Forms.CheckBox
        Me.CheckBox4 = New System.Windows.Forms.CheckBox
        Me.CheckBox3 = New System.Windows.Forms.CheckBox
        Me.CheckBox2 = New System.Windows.Forms.CheckBox
        Me.CheckBox1 = New System.Windows.Forms.CheckBox
        Me.GroupBox2 = New System.Windows.Forms.GroupBox
        Me.RichTextBox1 = New System.Windows.Forms.RichTextBox
        Me.GroupBox3 = New System.Windows.Forms.GroupBox
        Me.Button2 = New System.Windows.Forms.Button
        Me.Button1 = New System.Windows.Forms.Button
        Me.DataGridView1 = New System.Windows.Forms.DataGridView
        Me.ContextMenuStrip1 = New System.Windows.Forms.ContextMenuStrip(Me.components)
        Me.AddFeeToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem
        Me.GroupBox4 = New System.Windows.Forms.GroupBox
        Me.Button6 = New System.Windows.Forms.Button
        Me.DataGridView2 = New System.Windows.Forms.DataGridView
        Me.ContextMenuStrip2 = New System.Windows.Forms.ContextMenuStrip(Me.components)
        Me.EditFeeToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem
        Me.Label7 = New System.Windows.Forms.Label
        Me.Button4 = New System.Windows.Forms.Button
        Me.Button3 = New System.Windows.Forms.Button
        Me.ckbActive = New System.Windows.Forms.CheckBox
        Me.Timer1 = New System.Windows.Forms.Timer(Me.components)
        Me.ckbChanges = New System.Windows.Forms.CheckBox
        Me.GroupBox1.SuspendLayout()
        Me.GroupBox2.SuspendLayout()
        Me.GroupBox3.SuspendLayout()
        CType(Me.DataGridView1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.ContextMenuStrip1.SuspendLayout()
        Me.GroupBox4.SuspendLayout()
        CType(Me.DataGridView2, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.ContextMenuStrip2.SuspendLayout()
        Me.SuspendLayout()
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(12, 15)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(18, 13)
        Me.Label1.TabIndex = 0
        Me.Label1.Text = "ID"
        '
        'ID
        '
        Me.ID.Location = New System.Drawing.Point(36, 12)
        Me.ID.Name = "ID"
        Me.ID.ReadOnly = True
        Me.ID.Size = New System.Drawing.Size(100, 20)
        Me.ID.TabIndex = 1
        '
        'WLName
        '
        Me.WLName.Location = New System.Drawing.Point(85, 38)
        Me.WLName.Name = "WLName"
        Me.WLName.Size = New System.Drawing.Size(360, 20)
        Me.WLName.TabIndex = 3
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(12, 41)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(67, 13)
        Me.Label2.TabIndex = 2
        Me.Label2.Text = "White Label:"
        '
        'cboPlatform
        '
        Me.cboPlatform.FormattingEnabled = True
        Me.cboPlatform.Location = New System.Drawing.Point(97, 64)
        Me.cboPlatform.Name = "cboPlatform"
        Me.cboPlatform.Size = New System.Drawing.Size(348, 21)
        Me.cboPlatform.TabIndex = 4
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(12, 67)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(79, 13)
        Me.Label3.TabIndex = 5
        Me.Label3.Text = "Platform Driver:"
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.PlatformIDCopy)
        Me.GroupBox1.Controls.Add(Me.WLNameCopy)
        Me.GroupBox1.Controls.Add(Me.RichTextBox2)
        Me.GroupBox1.Controls.Add(Me.cbcfcopy)
        Me.GroupBox1.Controls.Add(Me.cb7copy)
        Me.GroupBox1.Controls.Add(Me.cb6copy)
        Me.GroupBox1.Controls.Add(Me.cb5copy)
        Me.GroupBox1.Controls.Add(Me.cb4copy)
        Me.GroupBox1.Controls.Add(Me.cb3copy)
        Me.GroupBox1.Controls.Add(Me.cb2copy)
        Me.GroupBox1.Controls.Add(Me.cb1copy)
        Me.GroupBox1.Controls.Add(Me.ckbCustomFee)
        Me.GroupBox1.Controls.Add(Me.CheckBox7)
        Me.GroupBox1.Controls.Add(Me.CheckBox6)
        Me.GroupBox1.Controls.Add(Me.CheckBox5)
        Me.GroupBox1.Controls.Add(Me.CheckBox4)
        Me.GroupBox1.Controls.Add(Me.CheckBox3)
        Me.GroupBox1.Controls.Add(Me.CheckBox2)
        Me.GroupBox1.Controls.Add(Me.CheckBox1)
        Me.GroupBox1.Location = New System.Drawing.Point(12, 91)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(433, 187)
        Me.GroupBox1.TabIndex = 7
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "Attributes"
        '
        'PlatformIDCopy
        '
        Me.PlatformIDCopy.Location = New System.Drawing.Point(182, 157)
        Me.PlatformIDCopy.Name = "PlatformIDCopy"
        Me.PlatformIDCopy.Size = New System.Drawing.Size(123, 20)
        Me.PlatformIDCopy.TabIndex = 36
        Me.PlatformIDCopy.Visible = False
        '
        'WLNameCopy
        '
        Me.WLNameCopy.Location = New System.Drawing.Point(182, 127)
        Me.WLNameCopy.Name = "WLNameCopy"
        Me.WLNameCopy.Size = New System.Drawing.Size(123, 20)
        Me.WLNameCopy.TabIndex = 25
        Me.WLNameCopy.Visible = False
        '
        'RichTextBox2
        '
        Me.RichTextBox2.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.RichTextBox2.Location = New System.Drawing.Point(182, 62)
        Me.RichTextBox2.Name = "RichTextBox2"
        Me.RichTextBox2.Size = New System.Drawing.Size(123, 59)
        Me.RichTextBox2.TabIndex = 1
        Me.RichTextBox2.Text = ""
        Me.RichTextBox2.Visible = False
        '
        'cbcfcopy
        '
        Me.cbcfcopy.AutoSize = True
        Me.cbcfcopy.Location = New System.Drawing.Point(306, 42)
        Me.cbcfcopy.Name = "cbcfcopy"
        Me.cbcfcopy.Size = New System.Drawing.Size(15, 14)
        Me.cbcfcopy.TabIndex = 35
        Me.cbcfcopy.UseVisualStyleBackColor = True
        Me.cbcfcopy.Visible = False
        '
        'cb7copy
        '
        Me.cb7copy.AutoSize = True
        Me.cb7copy.Location = New System.Drawing.Point(161, 19)
        Me.cb7copy.Name = "cb7copy"
        Me.cb7copy.Size = New System.Drawing.Size(15, 14)
        Me.cb7copy.TabIndex = 34
        Me.cb7copy.UseVisualStyleBackColor = True
        Me.cb7copy.Visible = False
        '
        'cb6copy
        '
        Me.cb6copy.AutoSize = True
        Me.cb6copy.Location = New System.Drawing.Point(161, 157)
        Me.cb6copy.Name = "cb6copy"
        Me.cb6copy.Size = New System.Drawing.Size(15, 14)
        Me.cb6copy.TabIndex = 33
        Me.cb6copy.UseVisualStyleBackColor = True
        Me.cb6copy.Visible = False
        '
        'cb5copy
        '
        Me.cb5copy.AutoSize = True
        Me.cb5copy.Location = New System.Drawing.Point(161, 134)
        Me.cb5copy.Name = "cb5copy"
        Me.cb5copy.Size = New System.Drawing.Size(15, 14)
        Me.cb5copy.TabIndex = 32
        Me.cb5copy.UseVisualStyleBackColor = True
        Me.cb5copy.Visible = False
        '
        'cb4copy
        '
        Me.cb4copy.AutoSize = True
        Me.cb4copy.Location = New System.Drawing.Point(161, 111)
        Me.cb4copy.Name = "cb4copy"
        Me.cb4copy.Size = New System.Drawing.Size(15, 14)
        Me.cb4copy.TabIndex = 31
        Me.cb4copy.UseVisualStyleBackColor = True
        Me.cb4copy.Visible = False
        '
        'cb3copy
        '
        Me.cb3copy.AutoSize = True
        Me.cb3copy.Location = New System.Drawing.Point(161, 88)
        Me.cb3copy.Name = "cb3copy"
        Me.cb3copy.Size = New System.Drawing.Size(15, 14)
        Me.cb3copy.TabIndex = 30
        Me.cb3copy.UseVisualStyleBackColor = True
        Me.cb3copy.Visible = False
        '
        'cb2copy
        '
        Me.cb2copy.AutoSize = True
        Me.cb2copy.Location = New System.Drawing.Point(161, 65)
        Me.cb2copy.Name = "cb2copy"
        Me.cb2copy.Size = New System.Drawing.Size(15, 14)
        Me.cb2copy.TabIndex = 29
        Me.cb2copy.UseVisualStyleBackColor = True
        Me.cb2copy.Visible = False
        '
        'cb1copy
        '
        Me.cb1copy.AutoSize = True
        Me.cb1copy.Location = New System.Drawing.Point(161, 42)
        Me.cb1copy.Name = "cb1copy"
        Me.cb1copy.Size = New System.Drawing.Size(15, 14)
        Me.cb1copy.TabIndex = 28
        Me.cb1copy.UseVisualStyleBackColor = True
        Me.cb1copy.Visible = False
        '
        'ckbCustomFee
        '
        Me.ckbCustomFee.AutoSize = True
        Me.ckbCustomFee.Location = New System.Drawing.Point(306, 19)
        Me.ckbCustomFee.Name = "ckbCustomFee"
        Me.ckbCustomFee.Size = New System.Drawing.Size(104, 17)
        Me.ckbCustomFee.TabIndex = 27
        Me.ckbCustomFee.Text = "Has Custom Fee"
        Me.ckbCustomFee.UseVisualStyleBackColor = True
        '
        'CheckBox7
        '
        Me.CheckBox7.AutoSize = True
        Me.CheckBox7.Location = New System.Drawing.Point(6, 19)
        Me.CheckBox7.Name = "CheckBox7"
        Me.CheckBox7.Size = New System.Drawing.Size(98, 17)
        Me.CheckBox7.TabIndex = 13
        Me.CheckBox7.Text = "Mirrors Platform"
        Me.CheckBox7.UseVisualStyleBackColor = True
        '
        'CheckBox6
        '
        Me.CheckBox6.AutoSize = True
        Me.CheckBox6.Location = New System.Drawing.Point(6, 157)
        Me.CheckBox6.Name = "CheckBox6"
        Me.CheckBox6.Size = New System.Drawing.Size(82, 17)
        Me.CheckBox6.TabIndex = 12
        Me.CheckBox6.Text = "Offers UIT's"
        Me.CheckBox6.UseVisualStyleBackColor = True
        '
        'CheckBox5
        '
        Me.CheckBox5.AutoSize = True
        Me.CheckBox5.Location = New System.Drawing.Point(6, 134)
        Me.CheckBox5.Name = "CheckBox5"
        Me.CheckBox5.Size = New System.Drawing.Size(121, 17)
        Me.CheckBox5.TabIndex = 11
        Me.CheckBox5.Text = "Offers Mutual Funds"
        Me.CheckBox5.UseVisualStyleBackColor = True
        '
        'CheckBox4
        '
        Me.CheckBox4.AutoSize = True
        Me.CheckBox4.Location = New System.Drawing.Point(6, 111)
        Me.CheckBox4.Name = "CheckBox4"
        Me.CheckBox4.Size = New System.Drawing.Size(81, 17)
        Me.CheckBox4.TabIndex = 10
        Me.CheckBox4.Text = "Offers UMA"
        Me.CheckBox4.UseVisualStyleBackColor = True
        '
        'CheckBox3
        '
        Me.CheckBox3.AutoSize = True
        Me.CheckBox3.Location = New System.Drawing.Point(6, 88)
        Me.CheckBox3.Name = "CheckBox3"
        Me.CheckBox3.Size = New System.Drawing.Size(80, 17)
        Me.CheckBox3.TabIndex = 9
        Me.CheckBox3.Text = "Offers SMA"
        Me.CheckBox3.UseVisualStyleBackColor = True
        '
        'CheckBox2
        '
        Me.CheckBox2.AutoSize = True
        Me.CheckBox2.Location = New System.Drawing.Point(6, 65)
        Me.CheckBox2.Name = "CheckBox2"
        Me.CheckBox2.Size = New System.Drawing.Size(157, 17)
        Me.CheckBox2.TabIndex = 8
        Me.CheckBox2.Text = "Require Additional Approval"
        Me.CheckBox2.UseVisualStyleBackColor = True
        '
        'CheckBox1
        '
        Me.CheckBox1.AutoSize = True
        Me.CheckBox1.Location = New System.Drawing.Point(6, 42)
        Me.CheckBox1.Name = "CheckBox1"
        Me.CheckBox1.Size = New System.Drawing.Size(149, 17)
        Me.CheckBox1.TabIndex = 7
        Me.CheckBox1.Text = "Require Platform Approval"
        Me.CheckBox1.UseVisualStyleBackColor = True
        '
        'GroupBox2
        '
        Me.GroupBox2.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.GroupBox2.Controls.Add(Me.RichTextBox1)
        Me.GroupBox2.Enabled = False
        Me.GroupBox2.Location = New System.Drawing.Point(12, 284)
        Me.GroupBox2.Name = "GroupBox2"
        Me.GroupBox2.Size = New System.Drawing.Size(433, 428)
        Me.GroupBox2.TabIndex = 8
        Me.GroupBox2.TabStop = False
        Me.GroupBox2.Text = "Additional Approval Process"
        '
        'RichTextBox1
        '
        Me.RichTextBox1.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.RichTextBox1.Location = New System.Drawing.Point(3, 19)
        Me.RichTextBox1.Name = "RichTextBox1"
        Me.RichTextBox1.Size = New System.Drawing.Size(424, 403)
        Me.RichTextBox1.TabIndex = 0
        Me.RichTextBox1.Text = ""
        '
        'GroupBox3
        '
        Me.GroupBox3.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.GroupBox3.Controls.Add(Me.Button2)
        Me.GroupBox3.Controls.Add(Me.Button1)
        Me.GroupBox3.Controls.Add(Me.DataGridView1)
        Me.GroupBox3.Location = New System.Drawing.Point(451, 32)
        Me.GroupBox3.Name = "GroupBox3"
        Me.GroupBox3.Size = New System.Drawing.Size(746, 346)
        Me.GroupBox3.TabIndex = 9
        Me.GroupBox3.TabStop = False
        Me.GroupBox3.Text = "Product Worker"
        '
        'Button2
        '
        Me.Button2.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Button2.Location = New System.Drawing.Point(665, 19)
        Me.Button2.Name = "Button2"
        Me.Button2.Size = New System.Drawing.Size(75, 23)
        Me.Button2.TabIndex = 24
        Me.Button2.Text = "Edit"
        Me.Button2.UseVisualStyleBackColor = True
        '
        'Button1
        '
        Me.Button1.Location = New System.Drawing.Point(6, 19)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(75, 23)
        Me.Button1.TabIndex = 23
        Me.Button1.Text = "Reload"
        Me.Button1.UseVisualStyleBackColor = True
        '
        'DataGridView1
        '
        Me.DataGridView1.AllowUserToAddRows = False
        Me.DataGridView1.AllowUserToDeleteRows = False
        DataGridViewCellStyle1.BackColor = System.Drawing.Color.FromArgb(CType(CType(224, Byte), Integer), CType(CType(224, Byte), Integer), CType(CType(224, Byte), Integer))
        Me.DataGridView1.AlternatingRowsDefaultCellStyle = DataGridViewCellStyle1
        Me.DataGridView1.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.DataGridView1.BackgroundColor = System.Drawing.Color.White
        Me.DataGridView1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.DataGridView1.ContextMenuStrip = Me.ContextMenuStrip1
        Me.DataGridView1.Location = New System.Drawing.Point(6, 49)
        Me.DataGridView1.Name = "DataGridView1"
        Me.DataGridView1.ReadOnly = True
        Me.DataGridView1.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect
        Me.DataGridView1.Size = New System.Drawing.Size(734, 291)
        Me.DataGridView1.TabIndex = 0
        '
        'ContextMenuStrip1
        '
        Me.ContextMenuStrip1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.AddFeeToolStripMenuItem})
        Me.ContextMenuStrip1.Name = "ContextMenuStrip1"
        Me.ContextMenuStrip1.Size = New System.Drawing.Size(118, 26)
        '
        'AddFeeToolStripMenuItem
        '
        Me.AddFeeToolStripMenuItem.Name = "AddFeeToolStripMenuItem"
        Me.AddFeeToolStripMenuItem.Size = New System.Drawing.Size(117, 22)
        Me.AddFeeToolStripMenuItem.Text = "Add Fee"
        Me.AddFeeToolStripMenuItem.Visible = False
        '
        'GroupBox4
        '
        Me.GroupBox4.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.GroupBox4.Controls.Add(Me.Button6)
        Me.GroupBox4.Controls.Add(Me.DataGridView2)
        Me.GroupBox4.Location = New System.Drawing.Point(451, 384)
        Me.GroupBox4.Name = "GroupBox4"
        Me.GroupBox4.Size = New System.Drawing.Size(746, 322)
        Me.GroupBox4.TabIndex = 10
        Me.GroupBox4.TabStop = False
        Me.GroupBox4.Text = "Fee Worker"
        '
        'Button6
        '
        Me.Button6.Location = New System.Drawing.Point(6, 19)
        Me.Button6.Name = "Button6"
        Me.Button6.Size = New System.Drawing.Size(75, 23)
        Me.Button6.TabIndex = 26
        Me.Button6.Text = "Reload"
        Me.Button6.UseVisualStyleBackColor = True
        '
        'DataGridView2
        '
        Me.DataGridView2.AllowUserToAddRows = False
        Me.DataGridView2.AllowUserToDeleteRows = False
        DataGridViewCellStyle2.BackColor = System.Drawing.Color.FromArgb(CType(CType(224, Byte), Integer), CType(CType(224, Byte), Integer), CType(CType(224, Byte), Integer))
        Me.DataGridView2.AlternatingRowsDefaultCellStyle = DataGridViewCellStyle2
        Me.DataGridView2.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.DataGridView2.BackgroundColor = System.Drawing.Color.White
        Me.DataGridView2.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.DataGridView2.ContextMenuStrip = Me.ContextMenuStrip2
        Me.DataGridView2.Location = New System.Drawing.Point(6, 48)
        Me.DataGridView2.Name = "DataGridView2"
        Me.DataGridView2.ReadOnly = True
        Me.DataGridView2.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect
        Me.DataGridView2.Size = New System.Drawing.Size(734, 274)
        Me.DataGridView2.TabIndex = 25
        '
        'ContextMenuStrip2
        '
        Me.ContextMenuStrip2.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.EditFeeToolStripMenuItem})
        Me.ContextMenuStrip2.Name = "ContextMenuStrip2"
        Me.ContextMenuStrip2.Size = New System.Drawing.Size(116, 26)
        '
        'EditFeeToolStripMenuItem
        '
        Me.EditFeeToolStripMenuItem.Name = "EditFeeToolStripMenuItem"
        Me.EditFeeToolStripMenuItem.Size = New System.Drawing.Size(115, 22)
        Me.EditFeeToolStripMenuItem.Text = "Edit Fee"
        Me.EditFeeToolStripMenuItem.Visible = False
        '
        'Label7
        '
        Me.Label7.AutoSize = True
        Me.Label7.Font = New System.Drawing.Font("Microsoft Sans Serif", 14.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label7.ForeColor = System.Drawing.Color.Red
        Me.Label7.Location = New System.Drawing.Point(201, 7)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(104, 24)
        Me.Label7.TabIndex = 24
        Me.Label7.Text = "DELETED"
        Me.Label7.Visible = False
        '
        'Button4
        '
        Me.Button4.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Button4.Location = New System.Drawing.Point(1041, 10)
        Me.Button4.Name = "Button4"
        Me.Button4.Size = New System.Drawing.Size(75, 23)
        Me.Button4.TabIndex = 23
        Me.Button4.Text = "Delete"
        Me.Button4.UseVisualStyleBackColor = True
        '
        'Button3
        '
        Me.Button3.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Button3.Location = New System.Drawing.Point(1122, 10)
        Me.Button3.Name = "Button3"
        Me.Button3.Size = New System.Drawing.Size(75, 23)
        Me.Button3.TabIndex = 22
        Me.Button3.Text = "Save"
        Me.Button3.UseVisualStyleBackColor = True
        '
        'ckbActive
        '
        Me.ckbActive.AutoSize = True
        Me.ckbActive.Checked = True
        Me.ckbActive.CheckState = System.Windows.Forms.CheckState.Checked
        Me.ckbActive.Location = New System.Drawing.Point(569, 15)
        Me.ckbActive.Name = "ckbActive"
        Me.ckbActive.Size = New System.Drawing.Size(56, 17)
        Me.ckbActive.TabIndex = 21
        Me.ckbActive.Text = "Active"
        Me.ckbActive.UseVisualStyleBackColor = True
        Me.ckbActive.Visible = False
        '
        'Timer1
        '
        Me.Timer1.Enabled = True
        '
        'ckbChanges
        '
        Me.ckbChanges.AutoSize = True
        Me.ckbChanges.Location = New System.Drawing.Point(689, 14)
        Me.ckbChanges.Name = "ckbChanges"
        Me.ckbChanges.Size = New System.Drawing.Size(115, 17)
        Me.ckbChanges.TabIndex = 25
        Me.ckbChanges.Text = "Changes Detected"
        Me.ckbChanges.UseVisualStyleBackColor = True
        Me.ckbChanges.Visible = False
        '
        'map_EditWLPlatform
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.Color.White
        Me.ClientSize = New System.Drawing.Size(1209, 724)
        Me.Controls.Add(Me.ckbChanges)
        Me.Controls.Add(Me.Label7)
        Me.Controls.Add(Me.Button4)
        Me.Controls.Add(Me.Button3)
        Me.Controls.Add(Me.ckbActive)
        Me.Controls.Add(Me.GroupBox4)
        Me.Controls.Add(Me.GroupBox3)
        Me.Controls.Add(Me.GroupBox2)
        Me.Controls.Add(Me.GroupBox1)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.cboPlatform)
        Me.Controls.Add(Me.WLName)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.ID)
        Me.Controls.Add(Me.Label1)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Name = "map_EditWLPlatform"
        Me.Text = "Edit White Label Platforms"
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox1.PerformLayout()
        Me.GroupBox2.ResumeLayout(False)
        Me.GroupBox3.ResumeLayout(False)
        CType(Me.DataGridView1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ContextMenuStrip1.ResumeLayout(False)
        Me.GroupBox4.ResumeLayout(False)
        CType(Me.DataGridView2, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ContextMenuStrip2.ResumeLayout(False)
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents ID As System.Windows.Forms.TextBox
    Friend WithEvents WLName As System.Windows.Forms.TextBox
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents cboPlatform As System.Windows.Forms.ComboBox
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents CheckBox1 As System.Windows.Forms.CheckBox
    Friend WithEvents CheckBox7 As System.Windows.Forms.CheckBox
    Friend WithEvents CheckBox6 As System.Windows.Forms.CheckBox
    Friend WithEvents CheckBox5 As System.Windows.Forms.CheckBox
    Friend WithEvents CheckBox4 As System.Windows.Forms.CheckBox
    Friend WithEvents CheckBox3 As System.Windows.Forms.CheckBox
    Friend WithEvents CheckBox2 As System.Windows.Forms.CheckBox
    Friend WithEvents GroupBox2 As System.Windows.Forms.GroupBox
    Friend WithEvents RichTextBox1 As System.Windows.Forms.RichTextBox
    Friend WithEvents GroupBox3 As System.Windows.Forms.GroupBox
    Friend WithEvents GroupBox4 As System.Windows.Forms.GroupBox
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents Button4 As System.Windows.Forms.Button
    Friend WithEvents Button3 As System.Windows.Forms.Button
    Friend WithEvents ckbActive As System.Windows.Forms.CheckBox
    Friend WithEvents DataGridView1 As System.Windows.Forms.DataGridView
    Friend WithEvents Button2 As System.Windows.Forms.Button
    Friend WithEvents Button1 As System.Windows.Forms.Button
    Friend WithEvents Button6 As System.Windows.Forms.Button
    Friend WithEvents DataGridView2 As System.Windows.Forms.DataGridView
    Friend WithEvents Timer1 As System.Windows.Forms.Timer
    Friend WithEvents ckbCustomFee As System.Windows.Forms.CheckBox
    Friend WithEvents ContextMenuStrip1 As System.Windows.Forms.ContextMenuStrip
    Friend WithEvents AddFeeToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents ContextMenuStrip2 As System.Windows.Forms.ContextMenuStrip
    Friend WithEvents EditFeeToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents cbcfcopy As System.Windows.Forms.CheckBox
    Friend WithEvents cb7copy As System.Windows.Forms.CheckBox
    Friend WithEvents cb6copy As System.Windows.Forms.CheckBox
    Friend WithEvents cb5copy As System.Windows.Forms.CheckBox
    Friend WithEvents cb4copy As System.Windows.Forms.CheckBox
    Friend WithEvents cb3copy As System.Windows.Forms.CheckBox
    Friend WithEvents cb2copy As System.Windows.Forms.CheckBox
    Friend WithEvents cb1copy As System.Windows.Forms.CheckBox
    Friend WithEvents PlatformIDCopy As System.Windows.Forms.TextBox
    Friend WithEvents WLNameCopy As System.Windows.Forms.TextBox
    Friend WithEvents RichTextBox2 As System.Windows.Forms.RichTextBox
    Friend WithEvents ckbChanges As System.Windows.Forms.CheckBox
End Class
