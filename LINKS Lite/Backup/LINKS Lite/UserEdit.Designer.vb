<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class UserEdit
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(UserEdit))
        Me.ID = New System.Windows.Forms.TextBox
        Me.Label1 = New System.Windows.Forms.Label
        Me.Label2 = New System.Windows.Forms.Label
        Me.FullName = New System.Windows.Forms.TextBox
        Me.Label3 = New System.Windows.Forms.Label
        Me.NetworkID = New System.Windows.Forms.TextBox
        Me.Label4 = New System.Windows.Forms.Label
        Me.Password = New System.Windows.Forms.TextBox
        Me.Label5 = New System.Windows.Forms.Label
        Me.Team = New System.Windows.Forms.TextBox
        Me.Label6 = New System.Windows.Forms.Label
        Me.APXName = New System.Windows.Forms.TextBox
        Me.Label7 = New System.Windows.Forms.Label
        Me.APXID = New System.Windows.Forms.TextBox
        Me.Active = New System.Windows.Forms.CheckBox
        Me.WI = New System.Windows.Forms.CheckBox
        Me.ComboBox1 = New System.Windows.Forms.ComboBox
        Me.Label8 = New System.Windows.Forms.Label
        Me.cmdSave = New System.Windows.Forms.Button
        Me.cmdCancel = New System.Windows.Forms.Button
        Me.SuspendLayout()
        '
        'ID
        '
        Me.ID.Location = New System.Drawing.Point(92, 12)
        Me.ID.Name = "ID"
        Me.ID.ReadOnly = True
        Me.ID.Size = New System.Drawing.Size(100, 20)
        Me.ID.TabIndex = 0
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(43, 15)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(43, 13)
        Me.Label1.TabIndex = 1
        Me.Label1.Text = "UserID:"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(32, 41)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(57, 13)
        Me.Label2.TabIndex = 3
        Me.Label2.Text = "Full Name:"
        '
        'FullName
        '
        Me.FullName.Location = New System.Drawing.Point(92, 38)
        Me.FullName.Name = "FullName"
        Me.FullName.Size = New System.Drawing.Size(173, 20)
        Me.FullName.TabIndex = 2
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(7, 67)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(79, 13)
        Me.Label3.TabIndex = 5
        Me.Label3.Text = "Network Login:"
        '
        'NetworkID
        '
        Me.NetworkID.Location = New System.Drawing.Point(92, 64)
        Me.NetworkID.Name = "NetworkID"
        Me.NetworkID.Size = New System.Drawing.Size(173, 20)
        Me.NetworkID.TabIndex = 4
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(30, 93)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(56, 13)
        Me.Label4.TabIndex = 7
        Me.Label4.Text = "Password:"
        '
        'Password
        '
        Me.Password.Location = New System.Drawing.Point(92, 90)
        Me.Password.Name = "Password"
        Me.Password.Size = New System.Drawing.Size(173, 20)
        Me.Password.TabIndex = 6
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Location = New System.Drawing.Point(18, 119)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(68, 13)
        Me.Label5.TabIndex = 9
        Me.Label5.Text = "Team Name:"
        '
        'Team
        '
        Me.Team.Location = New System.Drawing.Point(92, 116)
        Me.Team.Name = "Team"
        Me.Team.Size = New System.Drawing.Size(173, 20)
        Me.Team.TabIndex = 8
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Location = New System.Drawing.Point(24, 145)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(62, 13)
        Me.Label6.TabIndex = 11
        Me.Label6.Text = "APX Name:"
        '
        'APXName
        '
        Me.APXName.Location = New System.Drawing.Point(92, 142)
        Me.APXName.Name = "APXName"
        Me.APXName.Size = New System.Drawing.Size(173, 20)
        Me.APXName.TabIndex = 10
        '
        'Label7
        '
        Me.Label7.AutoSize = True
        Me.Label7.Location = New System.Drawing.Point(1, 171)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(85, 13)
        Me.Label7.TabIndex = 13
        Me.Label7.Text = "APX Contact ID:"
        '
        'APXID
        '
        Me.APXID.Location = New System.Drawing.Point(92, 168)
        Me.APXID.Name = "APXID"
        Me.APXID.Size = New System.Drawing.Size(173, 20)
        Me.APXID.TabIndex = 12
        '
        'Active
        '
        Me.Active.AutoSize = True
        Me.Active.Location = New System.Drawing.Point(27, 194)
        Me.Active.Name = "Active"
        Me.Active.Size = New System.Drawing.Size(56, 17)
        Me.Active.TabIndex = 14
        Me.Active.Text = "Active"
        Me.Active.UseVisualStyleBackColor = True
        '
        'WI
        '
        Me.WI.AutoSize = True
        Me.WI.Location = New System.Drawing.Point(115, 194)
        Me.WI.Name = "WI"
        Me.WI.Size = New System.Drawing.Size(150, 17)
        Me.WI.TabIndex = 15
        Me.WI.Text = "Windows Integrated Login"
        Me.WI.UseVisualStyleBackColor = True
        '
        'ComboBox1
        '
        Me.ComboBox1.FormattingEnabled = True
        Me.ComboBox1.Location = New System.Drawing.Point(12, 237)
        Me.ComboBox1.Name = "ComboBox1"
        Me.ComboBox1.Size = New System.Drawing.Size(252, 21)
        Me.ComboBox1.TabIndex = 16
        '
        'Label8
        '
        Me.Label8.AutoSize = True
        Me.Label8.Location = New System.Drawing.Point(12, 221)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(29, 13)
        Me.Label8.TabIndex = 17
        Me.Label8.Text = "Role"
        '
        'cmdSave
        '
        Me.cmdSave.Location = New System.Drawing.Point(170, 269)
        Me.cmdSave.Name = "cmdSave"
        Me.cmdSave.Size = New System.Drawing.Size(94, 23)
        Me.cmdSave.TabIndex = 19
        Me.cmdSave.Text = "&Save"
        '
        'cmdCancel
        '
        Me.cmdCancel.Location = New System.Drawing.Point(4, 269)
        Me.cmdCancel.Name = "cmdCancel"
        Me.cmdCancel.Size = New System.Drawing.Size(94, 23)
        Me.cmdCancel.TabIndex = 18
        Me.cmdCancel.Text = "&Cancel"
        '
        'UserEdit
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.Color.White
        Me.ClientSize = New System.Drawing.Size(278, 303)
        Me.Controls.Add(Me.cmdSave)
        Me.Controls.Add(Me.cmdCancel)
        Me.Controls.Add(Me.Label8)
        Me.Controls.Add(Me.ComboBox1)
        Me.Controls.Add(Me.WI)
        Me.Controls.Add(Me.Active)
        Me.Controls.Add(Me.Label7)
        Me.Controls.Add(Me.APXID)
        Me.Controls.Add(Me.Label6)
        Me.Controls.Add(Me.APXName)
        Me.Controls.Add(Me.Label5)
        Me.Controls.Add(Me.Team)
        Me.Controls.Add(Me.Label4)
        Me.Controls.Add(Me.Password)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.NetworkID)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.FullName)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.ID)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "UserEdit"
        Me.Text = "Add/Edit User"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents ID As System.Windows.Forms.TextBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents FullName As System.Windows.Forms.TextBox
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents NetworkID As System.Windows.Forms.TextBox
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents Password As System.Windows.Forms.TextBox
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents Team As System.Windows.Forms.TextBox
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents APXName As System.Windows.Forms.TextBox
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents APXID As System.Windows.Forms.TextBox
    Friend WithEvents Active As System.Windows.Forms.CheckBox
    Friend WithEvents WI As System.Windows.Forms.CheckBox
    Friend WithEvents ComboBox1 As System.Windows.Forms.ComboBox
    Friend WithEvents Label8 As System.Windows.Forms.Label
    Friend WithEvents cmdSave As System.Windows.Forms.Button
    Friend WithEvents cmdCancel As System.Windows.Forms.Button
End Class
