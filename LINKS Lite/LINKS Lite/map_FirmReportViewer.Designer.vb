<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class map_FirmReportViewer
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(map_FirmReportViewer))
        Me.map_ReportFirm1 = New SMA_Toolbox.map_ReportFirm
        Me.AMSCover1 = New SMA_Toolbox.AMSCover
        Me.CrystalReportViewer1 = New CrystalDecisions.Windows.Forms.CrystalReportViewer
        Me.map_Report_FirmFINAL2 = New SMA_Toolbox.map_Report_FirmFINAL
        Me.map_Report_FirmFINAL1 = New SMA_Toolbox.map_Report_FirmFINAL
        Me.map_Report_FirmTEST1 = New SMA_Toolbox.map_Report_FirmTEST
        Me.SuspendLayout()
        '
        'CrystalReportViewer1
        '
        Me.CrystalReportViewer1.ActiveViewIndex = 0
        Me.CrystalReportViewer1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.CrystalReportViewer1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.CrystalReportViewer1.Location = New System.Drawing.Point(0, 0)
        Me.CrystalReportViewer1.Name = "CrystalReportViewer1"
        Me.CrystalReportViewer1.ReportSource = Me.map_Report_FirmTEST1
        Me.CrystalReportViewer1.Size = New System.Drawing.Size(1203, 706)
        Me.CrystalReportViewer1.TabIndex = 0
        '
        'map_FirmReportViewer
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(1203, 706)
        Me.Controls.Add(Me.CrystalReportViewer1)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Name = "map_FirmReportViewer"
        Me.Text = "Report Viewer"
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents CrystalReportViewer1 As CrystalDecisions.Windows.Forms.CrystalReportViewer
    Friend WithEvents map_ReportFirm1 As SMA_Toolbox.map_ReportFirm
    Friend WithEvents map_Report_FirmFINAL1 As SMA_Toolbox.map_Report_FirmFINAL
    Friend WithEvents map_Report_FirmFINAL2 As SMA_Toolbox.map_Report_FirmFINAL
    Friend WithEvents AMSCover1 As SMA_Toolbox.AMSCover
    Friend WithEvents map_Report_FirmTEST1 As SMA_Toolbox.map_Report_FirmTEST
End Class
