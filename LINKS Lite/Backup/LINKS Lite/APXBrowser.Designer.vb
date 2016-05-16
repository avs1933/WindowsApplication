<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class APXBrowser
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(APXBrowser))
        Me.wbAdvent = New System.Windows.Forms.WebBrowser
        Me.SuspendLayout()
        '
        'wbAdvent
        '
        Me.wbAdvent.Dock = System.Windows.Forms.DockStyle.Fill
        Me.wbAdvent.Location = New System.Drawing.Point(0, 0)
        Me.wbAdvent.MinimumSize = New System.Drawing.Size(20, 20)
        Me.wbAdvent.Name = "wbAdvent"
        Me.wbAdvent.Size = New System.Drawing.Size(1038, 543)
        Me.wbAdvent.TabIndex = 0
        '
        'APXBrowser
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.Color.White
        Me.ClientSize = New System.Drawing.Size(1038, 543)
        Me.Controls.Add(Me.wbAdvent)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Name = "APXBrowser"
        Me.Text = "Advent APX"
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents wbAdvent As System.Windows.Forms.WebBrowser
End Class
