<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class rfp_ReportQuestion
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(rfp_ReportQuestion))
        Me.Label1 = New System.Windows.Forms.Label
        Me.txtID = New System.Windows.Forms.TextBox
        Me.Label2 = New System.Windows.Forms.Label
        Me.txtRFPID = New System.Windows.Forms.TextBox
        Me.Label3 = New System.Windows.Forms.Label
        Me.txtSortOrder = New System.Windows.Forms.TextBox
        Me.Label4 = New System.Windows.Forms.Label
        Me.rtbQuestion = New System.Windows.Forms.RichTextBox
        Me.Button1 = New System.Windows.Forms.Button
        Me.cmdNew = New System.Windows.Forms.Button
        Me.lblConfirm = New System.Windows.Forms.Label
        Me.Button2 = New System.Windows.Forms.Button
        Me.ckbAutoSpell = New System.Windows.Forms.CheckBox
        Me.SuspendLayout()
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(6, 15)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(18, 13)
        Me.Label1.TabIndex = 7
        Me.Label1.Text = "ID"
        '
        'txtID
        '
        Me.txtID.Location = New System.Drawing.Point(30, 12)
        Me.txtID.Name = "txtID"
        Me.txtID.ReadOnly = True
        Me.txtID.Size = New System.Drawing.Size(100, 20)
        Me.txtID.TabIndex = 6
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(6, 41)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(28, 13)
        Me.Label2.TabIndex = 9
        Me.Label2.Text = "RFP"
        '
        'txtRFPID
        '
        Me.txtRFPID.Location = New System.Drawing.Point(40, 38)
        Me.txtRFPID.Name = "txtRFPID"
        Me.txtRFPID.ReadOnly = True
        Me.txtRFPID.Size = New System.Drawing.Size(100, 20)
        Me.txtRFPID.TabIndex = 8
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(6, 67)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(58, 13)
        Me.Label3.TabIndex = 10
        Me.Label3.Text = "Sort Order:"
        '
        'txtSortOrder
        '
        Me.txtSortOrder.Location = New System.Drawing.Point(70, 64)
        Me.txtSortOrder.Name = "txtSortOrder"
        Me.txtSortOrder.Size = New System.Drawing.Size(70, 20)
        Me.txtSortOrder.TabIndex = 11
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(6, 93)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(52, 13)
        Me.Label4.TabIndex = 12
        Me.Label4.Text = "Question:"
        '
        'rtbQuestion
        '
        Me.rtbQuestion.Location = New System.Drawing.Point(9, 109)
        Me.rtbQuestion.Name = "rtbQuestion"
        Me.rtbQuestion.Size = New System.Drawing.Size(603, 212)
        Me.rtbQuestion.TabIndex = 13
        Me.rtbQuestion.Text = ""
        '
        'Button1
        '
        Me.Button1.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Button1.Location = New System.Drawing.Point(478, 38)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(134, 23)
        Me.Button1.TabIndex = 31
        Me.Button1.Text = "Save and Close"
        '
        'cmdNew
        '
        Me.cmdNew.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmdNew.Location = New System.Drawing.Point(478, 5)
        Me.cmdNew.Name = "cmdNew"
        Me.cmdNew.Size = New System.Drawing.Size(134, 23)
        Me.cmdNew.TabIndex = 32
        Me.cmdNew.Text = "Save and Add Another"
        '
        'lblConfirm
        '
        Me.lblConfirm.AutoSize = True
        Me.lblConfirm.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblConfirm.Location = New System.Drawing.Point(149, 93)
        Me.lblConfirm.Name = "lblConfirm"
        Me.lblConfirm.Size = New System.Drawing.Size(95, 13)
        Me.lblConfirm.TabIndex = 33
        Me.lblConfirm.Text = "Nothing Saved."
        Me.lblConfirm.Visible = False
        '
        'Button2
        '
        Me.Button2.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Button2.Image = CType(resources.GetObject("Button2.Image"), System.Drawing.Image)
        Me.Button2.Location = New System.Drawing.Point(572, 64)
        Me.Button2.Name = "Button2"
        Me.Button2.Size = New System.Drawing.Size(40, 42)
        Me.Button2.TabIndex = 34
        Me.Button2.UseVisualStyleBackColor = True
        '
        'ckbAutoSpell
        '
        Me.ckbAutoSpell.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.ckbAutoSpell.AutoSize = True
        Me.ckbAutoSpell.Checked = True
        Me.ckbAutoSpell.CheckState = System.Windows.Forms.CheckState.Checked
        Me.ckbAutoSpell.Location = New System.Drawing.Point(458, 86)
        Me.ckbAutoSpell.Name = "ckbAutoSpell"
        Me.ckbAutoSpell.Size = New System.Drawing.Size(108, 17)
        Me.ckbAutoSpell.TabIndex = 35
        Me.ckbAutoSpell.Text = "Auto Spell Check"
        Me.ckbAutoSpell.UseVisualStyleBackColor = True
        '
        'rfp_ReportQuestion
        '
        Me.AcceptButton = Me.cmdNew
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.Color.White
        Me.ClientSize = New System.Drawing.Size(624, 345)
        Me.Controls.Add(Me.ckbAutoSpell)
        Me.Controls.Add(Me.Button2)
        Me.Controls.Add(Me.lblConfirm)
        Me.Controls.Add(Me.cmdNew)
        Me.Controls.Add(Me.Button1)
        Me.Controls.Add(Me.rtbQuestion)
        Me.Controls.Add(Me.Label4)
        Me.Controls.Add(Me.txtSortOrder)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.txtRFPID)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.txtID)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Name = "rfp_ReportQuestion"
        Me.Text = "Add Question"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents txtID As System.Windows.Forms.TextBox
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents txtRFPID As System.Windows.Forms.TextBox
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents txtSortOrder As System.Windows.Forms.TextBox
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents rtbQuestion As System.Windows.Forms.RichTextBox
    Friend WithEvents Button1 As System.Windows.Forms.Button
    Friend WithEvents cmdNew As System.Windows.Forms.Button
    Friend WithEvents lblConfirm As System.Windows.Forms.Label
    Friend WithEvents Button2 As System.Windows.Forms.Button
    Friend WithEvents ckbAutoSpell As System.Windows.Forms.CheckBox
End Class
