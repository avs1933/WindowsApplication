<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class map_PlatformAssignment
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(map_PlatformAssignment))
        Me.Label1 = New System.Windows.Forms.Label
        Me.ID = New System.Windows.Forms.TextBox
        Me.GroupBox1 = New System.Windows.Forms.GroupBox
        Me.Label2 = New System.Windows.Forms.Label
        Me.cboPlatform = New System.Windows.Forms.ComboBox
        Me.cboWLPlatform = New System.Windows.Forms.ComboBox
        Me.Label3 = New System.Windows.Forms.Label
        Me.cboFirm = New System.Windows.Forms.ComboBox
        Me.Label4 = New System.Windows.Forms.Label
        Me.ckbPlatform = New System.Windows.Forms.CheckBox
        Me.ckbWLPlatform = New System.Windows.Forms.CheckBox
        Me.Button1 = New System.Windows.Forms.Button
        Me.Button2 = New System.Windows.Forms.Button
        Me.Button3 = New System.Windows.Forms.Button
        Me.ckbActive = New System.Windows.Forms.CheckBox
        Me.Label7 = New System.Windows.Forms.Label
        Me.Timer1 = New System.Windows.Forms.Timer(Me.components)
        Me.GroupBox1.SuspendLayout()
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
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.ckbWLPlatform)
        Me.GroupBox1.Controls.Add(Me.ckbPlatform)
        Me.GroupBox1.Location = New System.Drawing.Point(15, 68)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(144, 72)
        Me.GroupBox1.TabIndex = 2
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "Assignment Type"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(12, 149)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(48, 13)
        Me.Label2.TabIndex = 3
        Me.Label2.Text = "Platform:"
        '
        'cboPlatform
        '
        Me.cboPlatform.Enabled = False
        Me.cboPlatform.FormattingEnabled = True
        Me.cboPlatform.Location = New System.Drawing.Point(66, 146)
        Me.cboPlatform.Name = "cboPlatform"
        Me.cboPlatform.Size = New System.Drawing.Size(259, 21)
        Me.cboPlatform.TabIndex = 4
        '
        'cboWLPlatform
        '
        Me.cboWLPlatform.Enabled = False
        Me.cboWLPlatform.FormattingEnabled = True
        Me.cboWLPlatform.Location = New System.Drawing.Point(126, 173)
        Me.cboWLPlatform.Name = "cboWLPlatform"
        Me.cboWLPlatform.Size = New System.Drawing.Size(259, 21)
        Me.cboWLPlatform.TabIndex = 6
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(12, 176)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(108, 13)
        Me.Label3.TabIndex = 5
        Me.Label3.Text = "White Label Platform:"
        '
        'cboFirm
        '
        Me.cboFirm.FormattingEnabled = True
        Me.cboFirm.Location = New System.Drawing.Point(129, 41)
        Me.cboFirm.Name = "cboFirm"
        Me.cboFirm.Size = New System.Drawing.Size(301, 21)
        Me.cboFirm.TabIndex = 8
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(12, 44)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(111, 13)
        Me.Label4.TabIndex = 7
        Me.Label4.Text = "Firm to associate with:"
        '
        'ckbPlatform
        '
        Me.ckbPlatform.AutoSize = True
        Me.ckbPlatform.Location = New System.Drawing.Point(7, 20)
        Me.ckbPlatform.Name = "ckbPlatform"
        Me.ckbPlatform.Size = New System.Drawing.Size(64, 17)
        Me.ckbPlatform.TabIndex = 0
        Me.ckbPlatform.Text = "Platform"
        Me.ckbPlatform.UseVisualStyleBackColor = True
        '
        'ckbWLPlatform
        '
        Me.ckbWLPlatform.AutoSize = True
        Me.ckbWLPlatform.Location = New System.Drawing.Point(7, 43)
        Me.ckbWLPlatform.Name = "ckbWLPlatform"
        Me.ckbWLPlatform.Size = New System.Drawing.Size(124, 17)
        Me.ckbWLPlatform.TabIndex = 1
        Me.ckbWLPlatform.Text = "White Label Platform"
        Me.ckbWLPlatform.UseVisualStyleBackColor = True
        '
        'Button1
        '
        Me.Button1.Location = New System.Drawing.Point(355, 217)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(75, 23)
        Me.Button1.TabIndex = 9
        Me.Button1.Text = "Save"
        Me.Button1.UseVisualStyleBackColor = True
        '
        'Button2
        '
        Me.Button2.Location = New System.Drawing.Point(178, 217)
        Me.Button2.Name = "Button2"
        Me.Button2.Size = New System.Drawing.Size(75, 23)
        Me.Button2.TabIndex = 10
        Me.Button2.Text = "Delete"
        Me.Button2.UseVisualStyleBackColor = True
        '
        'Button3
        '
        Me.Button3.Location = New System.Drawing.Point(15, 217)
        Me.Button3.Name = "Button3"
        Me.Button3.Size = New System.Drawing.Size(75, 23)
        Me.Button3.TabIndex = 11
        Me.Button3.Text = "Cancel"
        Me.Button3.UseVisualStyleBackColor = True
        '
        'ckbActive
        '
        Me.ckbActive.AutoSize = True
        Me.ckbActive.Checked = True
        Me.ckbActive.CheckState = System.Windows.Forms.CheckState.Checked
        Me.ckbActive.Location = New System.Drawing.Point(374, 11)
        Me.ckbActive.Name = "ckbActive"
        Me.ckbActive.Size = New System.Drawing.Size(56, 17)
        Me.ckbActive.TabIndex = 12
        Me.ckbActive.Text = "Active"
        Me.ckbActive.UseVisualStyleBackColor = True
        Me.ckbActive.Visible = False
        '
        'Label7
        '
        Me.Label7.AutoSize = True
        Me.Label7.Font = New System.Drawing.Font("Microsoft Sans Serif", 14.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label7.ForeColor = System.Drawing.Color.Red
        Me.Label7.Location = New System.Drawing.Point(197, 8)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(104, 24)
        Me.Label7.TabIndex = 25
        Me.Label7.Text = "DELETED"
        Me.Label7.Visible = False
        '
        'Timer1
        '
        Me.Timer1.Enabled = True
        '
        'map_PlatformAssignment
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.Color.White
        Me.ClientSize = New System.Drawing.Size(452, 252)
        Me.Controls.Add(Me.Label7)
        Me.Controls.Add(Me.ckbActive)
        Me.Controls.Add(Me.Button3)
        Me.Controls.Add(Me.Button2)
        Me.Controls.Add(Me.Button1)
        Me.Controls.Add(Me.cboFirm)
        Me.Controls.Add(Me.Label4)
        Me.Controls.Add(Me.cboWLPlatform)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.cboPlatform)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.GroupBox1)
        Me.Controls.Add(Me.ID)
        Me.Controls.Add(Me.Label1)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "map_PlatformAssignment"
        Me.Text = "Edit Platform Assignments"
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox1.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents ID As System.Windows.Forms.TextBox
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents cboPlatform As System.Windows.Forms.ComboBox
    Friend WithEvents cboWLPlatform As System.Windows.Forms.ComboBox
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents cboFirm As System.Windows.Forms.ComboBox
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents ckbWLPlatform As System.Windows.Forms.CheckBox
    Friend WithEvents ckbPlatform As System.Windows.Forms.CheckBox
    Friend WithEvents Button1 As System.Windows.Forms.Button
    Friend WithEvents Button2 As System.Windows.Forms.Button
    Friend WithEvents Button3 As System.Windows.Forms.Button
    Friend WithEvents ckbActive As System.Windows.Forms.CheckBox
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents Timer1 As System.Windows.Forms.Timer
End Class
