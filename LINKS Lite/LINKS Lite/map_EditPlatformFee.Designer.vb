<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class map_EditPlatformFee
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(map_EditPlatformFee))
        Me.Label1 = New System.Windows.Forms.Label
        Me.ID = New System.Windows.Forms.TextBox
        Me.Label2 = New System.Windows.Forms.Label
        Me.cboPlatform = New System.Windows.Forms.ComboBox
        Me.cboWLPlatform = New System.Windows.Forms.ComboBox
        Me.Label3 = New System.Windows.Forms.Label
        Me.cboProduct = New System.Windows.Forms.ComboBox
        Me.Label4 = New System.Windows.Forms.Label
        Me.ckbWL = New System.Windows.Forms.CheckBox
        Me.GroupBox2 = New System.Windows.Forms.GroupBox
        Me.RepFee = New System.Windows.Forms.TextBox
        Me.Label8 = New System.Windows.Forms.Label
        Me.Fee = New System.Windows.Forms.TextBox
        Me.Label7 = New System.Windows.Forms.Label
        Me.BP2 = New System.Windows.Forms.TextBox
        Me.Label6 = New System.Windows.Forms.Label
        Me.BP1 = New System.Windows.Forms.TextBox
        Me.Label5 = New System.Windows.Forms.Label
        Me.Button1 = New System.Windows.Forms.Button
        Me.Button2 = New System.Windows.Forms.Button
        Me.Label9 = New System.Windows.Forms.Label
        Me.Timer1 = New System.Windows.Forms.Timer(Me.components)
        Me.ckbActive = New System.Windows.Forms.CheckBox
        Me.GroupBox2.SuspendLayout()
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
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(12, 41)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(45, 13)
        Me.Label2.TabIndex = 2
        Me.Label2.Text = "Platform"
        '
        'cboPlatform
        '
        Me.cboPlatform.Enabled = False
        Me.cboPlatform.FormattingEnabled = True
        Me.cboPlatform.Location = New System.Drawing.Point(63, 38)
        Me.cboPlatform.Name = "cboPlatform"
        Me.cboPlatform.Size = New System.Drawing.Size(289, 21)
        Me.cboPlatform.TabIndex = 3
        '
        'cboWLPlatform
        '
        Me.cboWLPlatform.Enabled = False
        Me.cboWLPlatform.FormattingEnabled = True
        Me.cboWLPlatform.Location = New System.Drawing.Point(123, 65)
        Me.cboWLPlatform.Name = "cboWLPlatform"
        Me.cboWLPlatform.Size = New System.Drawing.Size(303, 21)
        Me.cboWLPlatform.TabIndex = 5
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(12, 68)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(105, 13)
        Me.Label3.TabIndex = 4
        Me.Label3.Text = "Platform White Label"
        '
        'cboProduct
        '
        Me.cboProduct.Enabled = False
        Me.cboProduct.FormattingEnabled = True
        Me.cboProduct.Location = New System.Drawing.Point(62, 92)
        Me.cboProduct.Name = "cboProduct"
        Me.cboProduct.Size = New System.Drawing.Size(289, 21)
        Me.cboProduct.TabIndex = 7
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(12, 95)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(44, 13)
        Me.Label4.TabIndex = 6
        Me.Label4.Text = "Product"
        '
        'ckbWL
        '
        Me.ckbWL.AutoSize = True
        Me.ckbWL.Location = New System.Drawing.Point(479, 41)
        Me.ckbWL.Name = "ckbWL"
        Me.ckbWL.Size = New System.Drawing.Size(81, 17)
        Me.ckbWL.TabIndex = 8
        Me.ckbWL.Text = "WLPlatform"
        Me.ckbWL.UseVisualStyleBackColor = True
        Me.ckbWL.Visible = False
        '
        'GroupBox2
        '
        Me.GroupBox2.Controls.Add(Me.RepFee)
        Me.GroupBox2.Controls.Add(Me.Label8)
        Me.GroupBox2.Controls.Add(Me.Fee)
        Me.GroupBox2.Controls.Add(Me.Label7)
        Me.GroupBox2.Controls.Add(Me.BP2)
        Me.GroupBox2.Controls.Add(Me.Label6)
        Me.GroupBox2.Controls.Add(Me.BP1)
        Me.GroupBox2.Controls.Add(Me.Label5)
        Me.GroupBox2.Location = New System.Drawing.Point(12, 119)
        Me.GroupBox2.Name = "GroupBox2"
        Me.GroupBox2.Size = New System.Drawing.Size(506, 80)
        Me.GroupBox2.TabIndex = 13
        Me.GroupBox2.TabStop = False
        Me.GroupBox2.Text = "Non 40 Act Fee Structure"
        '
        'RepFee
        '
        Me.RepFee.Location = New System.Drawing.Point(86, 45)
        Me.RepFee.Name = "RepFee"
        Me.RepFee.Size = New System.Drawing.Size(100, 20)
        Me.RepFee.TabIndex = 7
        '
        'Label8
        '
        Me.Label8.AutoSize = True
        Me.Label8.Location = New System.Drawing.Point(6, 48)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(74, 13)
        Me.Label8.TabIndex = 6
        Me.Label8.Text = "Max Rep Fee:"
        '
        'Fee
        '
        Me.Fee.Location = New System.Drawing.Point(382, 19)
        Me.Fee.Name = "Fee"
        Me.Fee.Size = New System.Drawing.Size(100, 20)
        Me.Fee.TabIndex = 5
        '
        'Label7
        '
        Me.Label7.AutoSize = True
        Me.Label7.Location = New System.Drawing.Point(351, 22)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(25, 13)
        Me.Label7.TabIndex = 4
        Me.Label7.Text = "Fee"
        '
        'BP2
        '
        Me.BP2.Location = New System.Drawing.Point(228, 19)
        Me.BP2.Name = "BP2"
        Me.BP2.Size = New System.Drawing.Size(100, 20)
        Me.BP2.TabIndex = 3
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Location = New System.Drawing.Point(202, 22)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(20, 13)
        Me.Label6.TabIndex = 2
        Me.Label6.Text = "To"
        '
        'BP1
        '
        Me.BP1.Location = New System.Drawing.Point(96, 19)
        Me.BP1.Name = "BP1"
        Me.BP1.Size = New System.Drawing.Size(100, 20)
        Me.BP1.TabIndex = 1
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Location = New System.Drawing.Point(6, 22)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(84, 13)
        Me.Label5.TabIndex = 0
        Me.Label5.Text = "Breakpoint From"
        '
        'Button1
        '
        Me.Button1.Location = New System.Drawing.Point(515, 6)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(75, 23)
        Me.Button1.TabIndex = 14
        Me.Button1.Text = "Save"
        Me.Button1.UseVisualStyleBackColor = True
        '
        'Button2
        '
        Me.Button2.Location = New System.Drawing.Point(434, 6)
        Me.Button2.Name = "Button2"
        Me.Button2.Size = New System.Drawing.Size(75, 23)
        Me.Button2.TabIndex = 15
        Me.Button2.Text = "Delete"
        Me.Button2.UseVisualStyleBackColor = True
        '
        'Label9
        '
        Me.Label9.AutoSize = True
        Me.Label9.Font = New System.Drawing.Font("Microsoft Sans Serif", 14.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label9.ForeColor = System.Drawing.Color.Red
        Me.Label9.Location = New System.Drawing.Point(213, 8)
        Me.Label9.Name = "Label9"
        Me.Label9.Size = New System.Drawing.Size(104, 24)
        Me.Label9.TabIndex = 24
        Me.Label9.Text = "DELETED"
        Me.Label9.Visible = False
        '
        'Timer1
        '
        Me.Timer1.Enabled = True
        '
        'ckbActive
        '
        Me.ckbActive.AutoSize = True
        Me.ckbActive.Checked = True
        Me.ckbActive.CheckState = System.Windows.Forms.CheckState.Checked
        Me.ckbActive.Location = New System.Drawing.Point(479, 64)
        Me.ckbActive.Name = "ckbActive"
        Me.ckbActive.Size = New System.Drawing.Size(87, 17)
        Me.ckbActive.TabIndex = 25
        Me.ckbActive.Text = "Active Driver"
        Me.ckbActive.UseVisualStyleBackColor = True
        Me.ckbActive.Visible = False
        '
        'map_EditPlatformFee
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.Color.White
        Me.ClientSize = New System.Drawing.Size(604, 218)
        Me.Controls.Add(Me.ckbActive)
        Me.Controls.Add(Me.Label9)
        Me.Controls.Add(Me.Button2)
        Me.Controls.Add(Me.Button1)
        Me.Controls.Add(Me.GroupBox2)
        Me.Controls.Add(Me.ckbWL)
        Me.Controls.Add(Me.cboProduct)
        Me.Controls.Add(Me.Label4)
        Me.Controls.Add(Me.cboWLPlatform)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.cboPlatform)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.ID)
        Me.Controls.Add(Me.Label1)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Name = "map_EditPlatformFee"
        Me.Text = "Edit Platform Fee"
        Me.GroupBox2.ResumeLayout(False)
        Me.GroupBox2.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents ID As System.Windows.Forms.TextBox
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents cboPlatform As System.Windows.Forms.ComboBox
    Friend WithEvents cboWLPlatform As System.Windows.Forms.ComboBox
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents cboProduct As System.Windows.Forms.ComboBox
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents ckbWL As System.Windows.Forms.CheckBox
    Friend WithEvents GroupBox2 As System.Windows.Forms.GroupBox
    Friend WithEvents RepFee As System.Windows.Forms.TextBox
    Friend WithEvents Label8 As System.Windows.Forms.Label
    Friend WithEvents Fee As System.Windows.Forms.TextBox
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents BP2 As System.Windows.Forms.TextBox
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents BP1 As System.Windows.Forms.TextBox
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents Button1 As System.Windows.Forms.Button
    Friend WithEvents Button2 As System.Windows.Forms.Button
    Friend WithEvents Label9 As System.Windows.Forms.Label
    Friend WithEvents Timer1 As System.Windows.Forms.Timer
    Friend WithEvents ckbActive As System.Windows.Forms.CheckBox
End Class
