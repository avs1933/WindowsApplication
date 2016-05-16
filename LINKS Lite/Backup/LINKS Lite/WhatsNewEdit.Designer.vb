<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class WhatsNewEdit
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
        Dim DataGridViewCellStyle1 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(WhatsNewEdit))
        Me.GroupBox1 = New System.Windows.Forms.GroupBox
        Me.Button1 = New System.Windows.Forms.Button
        Me.ckbFunction = New System.Windows.Forms.CheckBox
        Me.txtSFunction = New System.Windows.Forms.TextBox
        Me.Label4 = New System.Windows.Forms.Label
        Me.ckbTable = New System.Windows.Forms.CheckBox
        Me.txtSTable = New System.Windows.Forms.TextBox
        Me.Label3 = New System.Windows.Forms.Label
        Me.ckbFormName = New System.Windows.Forms.CheckBox
        Me.ckbVersion = New System.Windows.Forms.CheckBox
        Me.txtSForm = New System.Windows.Forms.TextBox
        Me.Label2 = New System.Windows.Forms.Label
        Me.cboSVersion = New System.Windows.Forms.ComboBox
        Me.Label1 = New System.Windows.Forms.Label
        Me.DataGridView1 = New System.Windows.Forms.DataGridView
        Me.GroupBox2 = New System.Windows.Forms.GroupBox
        Me.btnDelete = New System.Windows.Forms.Button
        Me.btnReset = New System.Windows.Forms.Button
        Me.btnSave = New System.Windows.Forms.Button
        Me.Label12 = New System.Windows.Forms.Label
        Me.dtpRelease = New System.Windows.Forms.DateTimePicker
        Me.Label11 = New System.Windows.Forms.Label
        Me.rtbDevNotes = New System.Windows.Forms.RichTextBox
        Me.Label10 = New System.Windows.Forms.Label
        Me.rtbChanges = New System.Windows.Forms.RichTextBox
        Me.txtFunction = New System.Windows.Forms.TextBox
        Me.Label8 = New System.Windows.Forms.Label
        Me.txtTable = New System.Windows.Forms.TextBox
        Me.Label9 = New System.Windows.Forms.Label
        Me.txtForm = New System.Windows.Forms.TextBox
        Me.Label7 = New System.Windows.Forms.Label
        Me.txtVersion = New System.Windows.Forms.TextBox
        Me.Label6 = New System.Windows.Forms.Label
        Me.txtID = New System.Windows.Forms.TextBox
        Me.Label5 = New System.Windows.Forms.Label
        Me.txtHeader = New System.Windows.Forms.TextBox
        Me.Button5 = New System.Windows.Forms.Button
        Me.Button2 = New System.Windows.Forms.Button
        Me.GroupBox1.SuspendLayout()
        CType(Me.DataGridView1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.GroupBox2.SuspendLayout()
        Me.SuspendLayout()
        '
        'GroupBox1
        '
        Me.GroupBox1.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.GroupBox1.Controls.Add(Me.Button1)
        Me.GroupBox1.Controls.Add(Me.ckbFunction)
        Me.GroupBox1.Controls.Add(Me.txtSFunction)
        Me.GroupBox1.Controls.Add(Me.Label4)
        Me.GroupBox1.Controls.Add(Me.ckbTable)
        Me.GroupBox1.Controls.Add(Me.txtSTable)
        Me.GroupBox1.Controls.Add(Me.Label3)
        Me.GroupBox1.Controls.Add(Me.ckbFormName)
        Me.GroupBox1.Controls.Add(Me.ckbVersion)
        Me.GroupBox1.Controls.Add(Me.txtSForm)
        Me.GroupBox1.Controls.Add(Me.Label2)
        Me.GroupBox1.Controls.Add(Me.cboSVersion)
        Me.GroupBox1.Controls.Add(Me.Label1)
        Me.GroupBox1.Controls.Add(Me.DataGridView1)
        Me.GroupBox1.Location = New System.Drawing.Point(12, 61)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(721, 658)
        Me.GroupBox1.TabIndex = 0
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "Search Events"
        '
        'Button1
        '
        Me.Button1.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Button1.Location = New System.Drawing.Point(640, 15)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(75, 23)
        Me.Button1.TabIndex = 13
        Me.Button1.Text = "Search"
        Me.Button1.UseVisualStyleBackColor = True
        '
        'ckbFunction
        '
        Me.ckbFunction.AutoSize = True
        Me.ckbFunction.Location = New System.Drawing.Point(532, 49)
        Me.ckbFunction.Name = "ckbFunction"
        Me.ckbFunction.Size = New System.Drawing.Size(15, 14)
        Me.ckbFunction.TabIndex = 12
        Me.ckbFunction.UseVisualStyleBackColor = True
        '
        'txtSFunction
        '
        Me.txtSFunction.Enabled = False
        Me.txtSFunction.Location = New System.Drawing.Point(361, 46)
        Me.txtSFunction.Name = "txtSFunction"
        Me.txtSFunction.Size = New System.Drawing.Size(165, 20)
        Me.txtSFunction.TabIndex = 11
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(304, 49)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(51, 13)
        Me.Label4.TabIndex = 10
        Me.Label4.Text = "Function:"
        '
        'ckbTable
        '
        Me.ckbTable.AutoSize = True
        Me.ckbTable.Location = New System.Drawing.Point(532, 20)
        Me.ckbTable.Name = "ckbTable"
        Me.ckbTable.Size = New System.Drawing.Size(15, 14)
        Me.ckbTable.TabIndex = 9
        Me.ckbTable.UseVisualStyleBackColor = True
        '
        'txtSTable
        '
        Me.txtSTable.Enabled = False
        Me.txtSTable.Location = New System.Drawing.Point(361, 17)
        Me.txtSTable.Name = "txtSTable"
        Me.txtSTable.Size = New System.Drawing.Size(165, 20)
        Me.txtSTable.TabIndex = 8
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(287, 20)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(68, 13)
        Me.Label3.TabIndex = 7
        Me.Label3.Text = "Table Name:"
        '
        'ckbFormName
        '
        Me.ckbFormName.AutoSize = True
        Me.ckbFormName.Location = New System.Drawing.Point(248, 49)
        Me.ckbFormName.Name = "ckbFormName"
        Me.ckbFormName.Size = New System.Drawing.Size(15, 14)
        Me.ckbFormName.TabIndex = 6
        Me.ckbFormName.UseVisualStyleBackColor = True
        '
        'ckbVersion
        '
        Me.ckbVersion.AutoSize = True
        Me.ckbVersion.Location = New System.Drawing.Point(248, 20)
        Me.ckbVersion.Name = "ckbVersion"
        Me.ckbVersion.Size = New System.Drawing.Size(15, 14)
        Me.ckbVersion.TabIndex = 5
        Me.ckbVersion.UseVisualStyleBackColor = True
        '
        'txtSForm
        '
        Me.txtSForm.Enabled = False
        Me.txtSForm.Location = New System.Drawing.Point(77, 46)
        Me.txtSForm.Name = "txtSForm"
        Me.txtSForm.Size = New System.Drawing.Size(165, 20)
        Me.txtSForm.TabIndex = 4
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(7, 49)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(64, 13)
        Me.Label2.TabIndex = 3
        Me.Label2.Text = "Form Name:"
        '
        'cboSVersion
        '
        Me.cboSVersion.Enabled = False
        Me.cboSVersion.FormattingEnabled = True
        Me.cboSVersion.Location = New System.Drawing.Point(53, 17)
        Me.cboSVersion.Name = "cboSVersion"
        Me.cboSVersion.Size = New System.Drawing.Size(189, 21)
        Me.cboSVersion.TabIndex = 2
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(7, 20)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(40, 13)
        Me.Label1.TabIndex = 1
        Me.Label1.Text = "Verion:"
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
        Me.DataGridView1.Location = New System.Drawing.Point(6, 75)
        Me.DataGridView1.Name = "DataGridView1"
        Me.DataGridView1.ReadOnly = True
        Me.DataGridView1.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect
        Me.DataGridView1.Size = New System.Drawing.Size(709, 577)
        Me.DataGridView1.TabIndex = 0
        '
        'GroupBox2
        '
        Me.GroupBox2.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.GroupBox2.Controls.Add(Me.Button2)
        Me.GroupBox2.Controls.Add(Me.Button5)
        Me.GroupBox2.Controls.Add(Me.btnDelete)
        Me.GroupBox2.Controls.Add(Me.btnReset)
        Me.GroupBox2.Controls.Add(Me.btnSave)
        Me.GroupBox2.Controls.Add(Me.Label12)
        Me.GroupBox2.Controls.Add(Me.dtpRelease)
        Me.GroupBox2.Controls.Add(Me.Label11)
        Me.GroupBox2.Controls.Add(Me.rtbDevNotes)
        Me.GroupBox2.Controls.Add(Me.Label10)
        Me.GroupBox2.Controls.Add(Me.rtbChanges)
        Me.GroupBox2.Controls.Add(Me.txtFunction)
        Me.GroupBox2.Controls.Add(Me.Label8)
        Me.GroupBox2.Controls.Add(Me.txtTable)
        Me.GroupBox2.Controls.Add(Me.Label9)
        Me.GroupBox2.Controls.Add(Me.txtForm)
        Me.GroupBox2.Controls.Add(Me.Label7)
        Me.GroupBox2.Controls.Add(Me.txtVersion)
        Me.GroupBox2.Controls.Add(Me.Label6)
        Me.GroupBox2.Controls.Add(Me.txtID)
        Me.GroupBox2.Controls.Add(Me.Label5)
        Me.GroupBox2.Location = New System.Drawing.Point(739, 61)
        Me.GroupBox2.Name = "GroupBox2"
        Me.GroupBox2.Size = New System.Drawing.Size(469, 652)
        Me.GroupBox2.TabIndex = 1
        Me.GroupBox2.TabStop = False
        Me.GroupBox2.Text = "Edit Events"
        '
        'btnDelete
        '
        Me.btnDelete.Location = New System.Drawing.Point(194, 522)
        Me.btnDelete.Name = "btnDelete"
        Me.btnDelete.Size = New System.Drawing.Size(75, 23)
        Me.btnDelete.TabIndex = 30
        Me.btnDelete.Text = "Delete"
        Me.btnDelete.UseVisualStyleBackColor = True
        '
        'btnReset
        '
        Me.btnReset.Location = New System.Drawing.Point(14, 522)
        Me.btnReset.Name = "btnReset"
        Me.btnReset.Size = New System.Drawing.Size(75, 23)
        Me.btnReset.TabIndex = 29
        Me.btnReset.Text = "Reset"
        Me.btnReset.UseVisualStyleBackColor = True
        '
        'btnSave
        '
        Me.btnSave.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnSave.Location = New System.Drawing.Point(378, 522)
        Me.btnSave.Name = "btnSave"
        Me.btnSave.Size = New System.Drawing.Size(75, 23)
        Me.btnSave.TabIndex = 28
        Me.btnSave.Text = "Save"
        Me.btnSave.UseVisualStyleBackColor = True
        '
        'Label12
        '
        Me.Label12.AutoSize = True
        Me.Label12.Location = New System.Drawing.Point(11, 488)
        Me.Label12.Name = "Label12"
        Me.Label12.Size = New System.Drawing.Size(87, 13)
        Me.Label12.TabIndex = 27
        Me.Label12.Text = "Date of Release:"
        '
        'dtpRelease
        '
        Me.dtpRelease.Format = System.Windows.Forms.DateTimePickerFormat.[Short]
        Me.dtpRelease.Location = New System.Drawing.Point(104, 485)
        Me.dtpRelease.Name = "dtpRelease"
        Me.dtpRelease.Size = New System.Drawing.Size(114, 20)
        Me.dtpRelease.TabIndex = 26
        '
        'Label11
        '
        Me.Label11.AutoSize = True
        Me.Label11.Location = New System.Drawing.Point(11, 321)
        Me.Label11.Name = "Label11"
        Me.Label11.Size = New System.Drawing.Size(163, 13)
        Me.Label11.TabIndex = 25
        Me.Label11.Text = "Dev Notes (Not visable to users):"
        '
        'rtbDevNotes
        '
        Me.rtbDevNotes.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.rtbDevNotes.Location = New System.Drawing.Point(14, 337)
        Me.rtbDevNotes.Name = "rtbDevNotes"
        Me.rtbDevNotes.Size = New System.Drawing.Size(414, 142)
        Me.rtbDevNotes.TabIndex = 24
        Me.rtbDevNotes.Text = ""
        '
        'Label10
        '
        Me.Label10.AutoSize = True
        Me.Label10.Location = New System.Drawing.Point(11, 157)
        Me.Label10.Name = "Label10"
        Me.Label10.Size = New System.Drawing.Size(167, 13)
        Me.Label10.TabIndex = 23
        Me.Label10.Text = "Changes Made (Visable to Users):"
        '
        'rtbChanges
        '
        Me.rtbChanges.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.rtbChanges.Location = New System.Drawing.Point(14, 173)
        Me.rtbChanges.Name = "rtbChanges"
        Me.rtbChanges.Size = New System.Drawing.Size(414, 142)
        Me.rtbChanges.TabIndex = 22
        Me.rtbChanges.Text = ""
        '
        'txtFunction
        '
        Me.txtFunction.Location = New System.Drawing.Point(81, 127)
        Me.txtFunction.Name = "txtFunction"
        Me.txtFunction.Size = New System.Drawing.Size(165, 20)
        Me.txtFunction.TabIndex = 21
        '
        'Label8
        '
        Me.Label8.AutoSize = True
        Me.Label8.Location = New System.Drawing.Point(24, 130)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(51, 13)
        Me.Label8.TabIndex = 20
        Me.Label8.Text = "Function:"
        '
        'txtTable
        '
        Me.txtTable.Location = New System.Drawing.Point(81, 98)
        Me.txtTable.Name = "txtTable"
        Me.txtTable.Size = New System.Drawing.Size(165, 20)
        Me.txtTable.TabIndex = 19
        '
        'Label9
        '
        Me.Label9.AutoSize = True
        Me.Label9.Location = New System.Drawing.Point(7, 101)
        Me.Label9.Name = "Label9"
        Me.Label9.Size = New System.Drawing.Size(68, 13)
        Me.Label9.TabIndex = 18
        Me.Label9.Text = "Table Name:"
        '
        'txtForm
        '
        Me.txtForm.Location = New System.Drawing.Point(81, 72)
        Me.txtForm.Name = "txtForm"
        Me.txtForm.Size = New System.Drawing.Size(165, 20)
        Me.txtForm.TabIndex = 17
        '
        'Label7
        '
        Me.Label7.AutoSize = True
        Me.Label7.Location = New System.Drawing.Point(11, 75)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(64, 13)
        Me.Label7.TabIndex = 16
        Me.Label7.Text = "Form Name:"
        '
        'txtVersion
        '
        Me.txtVersion.Location = New System.Drawing.Point(60, 46)
        Me.txtVersion.Name = "txtVersion"
        Me.txtVersion.Size = New System.Drawing.Size(165, 20)
        Me.txtVersion.TabIndex = 15
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Location = New System.Drawing.Point(9, 52)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(45, 13)
        Me.Label6.TabIndex = 14
        Me.Label6.Text = "Version:"
        '
        'txtID
        '
        Me.txtID.Location = New System.Drawing.Point(38, 20)
        Me.txtID.Name = "txtID"
        Me.txtID.ReadOnly = True
        Me.txtID.Size = New System.Drawing.Size(74, 20)
        Me.txtID.TabIndex = 13
        Me.txtID.Text = "NEW"
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Location = New System.Drawing.Point(11, 23)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(21, 13)
        Me.Label5.TabIndex = 12
        Me.Label5.Text = "ID:"
        '
        'txtHeader
        '
        Me.txtHeader.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtHeader.BackColor = System.Drawing.Color.White
        Me.txtHeader.BorderStyle = System.Windows.Forms.BorderStyle.None
        Me.txtHeader.Enabled = False
        Me.txtHeader.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtHeader.Location = New System.Drawing.Point(12, 13)
        Me.txtHeader.Name = "txtHeader"
        Me.txtHeader.ReadOnly = True
        Me.txtHeader.Size = New System.Drawing.Size(1196, 19)
        Me.txtHeader.TabIndex = 2
        Me.txtHeader.Text = "Edit Whats New Events"
        Me.txtHeader.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'Button5
        '
        Me.Button5.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Button5.Image = CType(resources.GetObject("Button5.Image"), System.Drawing.Image)
        Me.Button5.Location = New System.Drawing.Point(434, 173)
        Me.Button5.Name = "Button5"
        Me.Button5.Size = New System.Drawing.Size(29, 30)
        Me.Button5.TabIndex = 71
        Me.Button5.UseVisualStyleBackColor = True
        '
        'Button2
        '
        Me.Button2.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Button2.Image = CType(resources.GetObject("Button2.Image"), System.Drawing.Image)
        Me.Button2.Location = New System.Drawing.Point(434, 337)
        Me.Button2.Name = "Button2"
        Me.Button2.Size = New System.Drawing.Size(29, 30)
        Me.Button2.TabIndex = 72
        Me.Button2.UseVisualStyleBackColor = True
        '
        'WhatsNewEdit
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.Color.White
        Me.ClientSize = New System.Drawing.Size(1220, 731)
        Me.Controls.Add(Me.txtHeader)
        Me.Controls.Add(Me.GroupBox2)
        Me.Controls.Add(Me.GroupBox1)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Name = "WhatsNewEdit"
        Me.Text = "Edit Whats New"
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox1.PerformLayout()
        CType(Me.DataGridView1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.GroupBox2.ResumeLayout(False)
        Me.GroupBox2.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents cboSVersion As System.Windows.Forms.ComboBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents DataGridView1 As System.Windows.Forms.DataGridView
    Friend WithEvents GroupBox2 As System.Windows.Forms.GroupBox
    Friend WithEvents Button1 As System.Windows.Forms.Button
    Friend WithEvents ckbFunction As System.Windows.Forms.CheckBox
    Friend WithEvents txtSFunction As System.Windows.Forms.TextBox
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents ckbTable As System.Windows.Forms.CheckBox
    Friend WithEvents txtSTable As System.Windows.Forms.TextBox
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents ckbFormName As System.Windows.Forms.CheckBox
    Friend WithEvents ckbVersion As System.Windows.Forms.CheckBox
    Friend WithEvents txtSForm As System.Windows.Forms.TextBox
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents txtHeader As System.Windows.Forms.TextBox
    Friend WithEvents Label12 As System.Windows.Forms.Label
    Friend WithEvents dtpRelease As System.Windows.Forms.DateTimePicker
    Friend WithEvents Label11 As System.Windows.Forms.Label
    Friend WithEvents rtbDevNotes As System.Windows.Forms.RichTextBox
    Friend WithEvents Label10 As System.Windows.Forms.Label
    Friend WithEvents rtbChanges As System.Windows.Forms.RichTextBox
    Friend WithEvents txtFunction As System.Windows.Forms.TextBox
    Friend WithEvents Label8 As System.Windows.Forms.Label
    Friend WithEvents txtTable As System.Windows.Forms.TextBox
    Friend WithEvents Label9 As System.Windows.Forms.Label
    Friend WithEvents txtForm As System.Windows.Forms.TextBox
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents txtVersion As System.Windows.Forms.TextBox
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents txtID As System.Windows.Forms.TextBox
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents btnDelete As System.Windows.Forms.Button
    Friend WithEvents btnReset As System.Windows.Forms.Button
    Friend WithEvents btnSave As System.Windows.Forms.Button
    Friend WithEvents Button5 As System.Windows.Forms.Button
    Friend WithEvents Button2 As System.Windows.Forms.Button
End Class
