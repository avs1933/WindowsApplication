<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FirmView
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(FirmView))
        Me.FirmName = New System.Windows.Forms.TextBox
        Me.SuspendLayout()
        '
        'FirmName
        '
        Me.FirmName.BackColor = System.Drawing.Color.White
        Me.FirmName.BorderStyle = System.Windows.Forms.BorderStyle.None
        Me.FirmName.Font = New System.Drawing.Font("Microsoft Sans Serif", 21.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.FirmName.Location = New System.Drawing.Point(12, 3)
        Me.FirmName.Name = "FirmName"
        Me.FirmName.ReadOnly = True
        Me.FirmName.Size = New System.Drawing.Size(1136, 33)
        Me.FirmName.TabIndex = 10
        Me.FirmName.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'FirmView
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.Color.White
        Me.ClientSize = New System.Drawing.Size(1160, 632)
        Me.Controls.Add(Me.FirmName)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Name = "FirmView"
        Me.Text = "FirmView"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents FirmName As System.Windows.Forms.TextBox
End Class
