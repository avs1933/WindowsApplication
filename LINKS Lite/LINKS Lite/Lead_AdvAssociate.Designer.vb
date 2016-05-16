<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Lead_AdvAssociate
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(Lead_AdvAssociate))
        Me.Label1 = New System.Windows.Forms.Label
        Me.cboLead = New System.Windows.Forms.ComboBox
        Me.cboAdvisor = New System.Windows.Forms.ComboBox
        Me.Label2 = New System.Windows.Forms.Label
        Me.Button1 = New System.Windows.Forms.Button
        Me.SuspendLayout()
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(13, 13)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(98, 13)
        Me.Label1.TabIndex = 0
        Me.Label1.Text = "Select Lead Name:"
        '
        'cboLead
        '
        Me.cboLead.AutoCompleteMode = System.Windows.Forms.AutoCompleteMode.SuggestAppend
        Me.cboLead.FormattingEnabled = True
        Me.cboLead.Location = New System.Drawing.Point(117, 10)
        Me.cboLead.Name = "cboLead"
        Me.cboLead.Size = New System.Drawing.Size(401, 21)
        Me.cboLead.TabIndex = 1
        '
        'cboAdvisor
        '
        Me.cboAdvisor.AutoCompleteMode = System.Windows.Forms.AutoCompleteMode.SuggestAppend
        Me.cboAdvisor.FormattingEnabled = True
        Me.cboAdvisor.Location = New System.Drawing.Point(133, 37)
        Me.cboAdvisor.Name = "cboAdvisor"
        Me.cboAdvisor.Size = New System.Drawing.Size(385, 21)
        Me.cboAdvisor.TabIndex = 3
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(13, 40)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(114, 13)
        Me.Label2.TabIndex = 2
        Me.Label2.Text = "Select Known Advisor:"
        '
        'Button1
        '
        Me.Button1.Location = New System.Drawing.Point(443, 64)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(75, 23)
        Me.Button1.TabIndex = 4
        Me.Button1.Text = "Save"
        Me.Button1.UseVisualStyleBackColor = True
        '
        'Lead_AdvAssociate
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.Color.White
        Me.ClientSize = New System.Drawing.Size(534, 101)
        Me.Controls.Add(Me.Button1)
        Me.Controls.Add(Me.cboAdvisor)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.cboLead)
        Me.Controls.Add(Me.Label1)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Name = "Lead_AdvAssociate"
        Me.Text = "Associate Lead with Known Advisor"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents cboLead As System.Windows.Forms.ComboBox
    Friend WithEvents cboAdvisor As System.Windows.Forms.ComboBox
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Button1 As System.Windows.Forms.Button
End Class
