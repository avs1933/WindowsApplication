Imports System.Windows.Forms

Public Class Home

    Private Sub ShowNewForm(ByVal sender As Object, ByVal e As EventArgs) Handles NewToolStripMenuItem.Click, NewWindowToolStripMenuItem.Click
        ' Create a new instance of the child form.
        Dim ChildForm As New System.Windows.Forms.Form
        ' Make it a child of this MDI form before showing it.
        ChildForm.MdiParent = Me

        m_ChildFormNumber += 1
        ChildForm.Text = "Window " & m_ChildFormNumber

        ChildForm.Show()
    End Sub

    Private Sub OpenFile(ByVal sender As Object, ByVal e As EventArgs) Handles OpenToolStripMenuItem.Click
        Dim OpenFileDialog As New OpenFileDialog
        OpenFileDialog.InitialDirectory = My.Computer.FileSystem.SpecialDirectories.MyDocuments
        OpenFileDialog.Filter = "Text Files (*.txt)|*.txt|All Files (*.*)|*.*"
        If (OpenFileDialog.ShowDialog(Me) = System.Windows.Forms.DialogResult.OK) Then
            Dim FileName As String = OpenFileDialog.FileName
            ' TODO: Add code here to open the file.
        End If
    End Sub

    Private Sub SaveAsToolStripMenuItem_Click(ByVal sender As Object, ByVal e As EventArgs) Handles SaveAsToolStripMenuItem.Click
        Dim SaveFileDialog As New SaveFileDialog
        SaveFileDialog.InitialDirectory = My.Computer.FileSystem.SpecialDirectories.MyDocuments
        SaveFileDialog.Filter = "Text Files (*.txt)|*.txt|All Files (*.*)|*.*" 

        If (SaveFileDialog.ShowDialog(Me) = System.Windows.Forms.DialogResult.OK) Then
            Dim FileName As String = SaveFileDialog.FileName
            ' TODO: Add code here to save the current contents of the form to a file.
        End If
    End Sub


    Private Sub ExitToolsStripMenuItem_Click(ByVal sender As Object, ByVal e As EventArgs) Handles ExitToolStripMenuItem.Click
        Me.Close()
    End Sub

    Private Sub CutToolStripMenuItem_Click(ByVal sender As Object, ByVal e As EventArgs) Handles CutToolStripMenuItem.Click
        ' Use My.Computer.Clipboard to insert the selected text or images into the clipboard
    End Sub

    Private Sub CopyToolStripMenuItem_Click(ByVal sender As Object, ByVal e As EventArgs) Handles CopyToolStripMenuItem.Click
        ' Use My.Computer.Clipboard to insert the selected text or images into the clipboard
    End Sub

    Private Sub PasteToolStripMenuItem_Click(ByVal sender As Object, ByVal e As EventArgs) Handles PasteToolStripMenuItem.Click
        'Use My.Computer.Clipboard.GetText() or My.Computer.Clipboard.GetData to retrieve information from the clipboard.
    End Sub

    Private Sub ToolBarToolStripMenuItem_Click(ByVal sender As Object, ByVal e As EventArgs) Handles ToolBarToolStripMenuItem.Click
        Me.ToolStrip.Visible = Me.ToolBarToolStripMenuItem.Checked
    End Sub

    Private Sub StatusBarToolStripMenuItem_Click(ByVal sender As Object, ByVal e As EventArgs) Handles StatusBarToolStripMenuItem.Click
        Me.StatusStrip.Visible = Me.StatusBarToolStripMenuItem.Checked
    End Sub

    Private Sub CascadeToolStripMenuItem_Click(ByVal sender As Object, ByVal e As EventArgs) Handles CascadeToolStripMenuItem.Click
        Me.LayoutMdi(MdiLayout.Cascade)
    End Sub

    Private Sub TileVerticalToolStripMenuItem_Click(ByVal sender As Object, ByVal e As EventArgs) Handles TileVerticalToolStripMenuItem.Click
        Me.LayoutMdi(MdiLayout.TileVertical)
    End Sub

    Private Sub TileHorizontalToolStripMenuItem_Click(ByVal sender As Object, ByVal e As EventArgs) Handles TileHorizontalToolStripMenuItem.Click
        Me.LayoutMdi(MdiLayout.TileHorizontal)
    End Sub

    Private Sub ArrangeIconsToolStripMenuItem_Click(ByVal sender As Object, ByVal e As EventArgs) Handles ArrangeIconsToolStripMenuItem.Click
        Me.LayoutMdi(MdiLayout.ArrangeIcons)
    End Sub

    Private Sub CloseAllToolStripMenuItem_Click(ByVal sender As Object, ByVal e As EventArgs) Handles CloseAllToolStripMenuItem.Click
        ' Close all child forms of the parent.
        For Each ChildForm As Form In Me.MdiChildren
            ChildForm.Close()
        Next
    End Sub

    Private m_ChildFormNumber As Integer

    Private Sub ViewMyTeamToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub

    Private Sub ToolStripButton1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton1.Click
        Application.Exit()
    End Sub

    Private Sub ToolStripDropDownButton2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripDropDownButton2.Click

    End Sub

    Private Sub ToolStripDropDownButton3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripDropDownButton3.Click

    End Sub

    Private Sub tsSpoof_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tsSpoof.Click
        Dim child As New SpoofUser
        child.MdiParent = Me
        child.Show()
    End Sub

    Private Sub UserControlToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles UserControlToolStripMenuItem.Click
        Dim child As New Users
        child.MdiParent = Me
        child.Show()
    End Sub

    Private Sub Home_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        Application.Exit()
    End Sub

    Private Sub Home_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        
    End Sub

    Private Sub AccountSearchToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles AccountSearchToolStripMenuItem.Click
        Dim child As New AccountSearch
        child.MdiParent = Me
        child.Show()
    End Sub

    Private Sub FirmLookupToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FirmLookupToolStripMenuItem.Click
        Dim child As New map_UserViewFirms
        child.MdiParent = Me
        child.Show()
    End Sub

    Private Sub ToolStripDropDownButton5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripDropDownButton5.Click

    End Sub

    Private Sub AddUserToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles AddUserToolStripMenuItem.Click
        Dim child As New UserEdit
        child.MdiParent = Me
        child.Show()
        child.ID.Text = "NEW"
        child.ID.Enabled = False
    End Sub

    Private Sub ToolStripButton2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton2.Click
        Call Permissions.Populate()
    End Sub

    Private Sub TSDashboard_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TSDashboard.Click
        Dim child As New Dashboard
        child.MdiParent = Me
        child.Show()
    End Sub

    Private Sub ToolStripButton3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton3.Click
        Dim child As New CommCalc
        child.MdiParent = Me
        child.Show()
    End Sub

    Private Sub ToolStripDropDownButton7_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripDropDownButton7.Click

    End Sub

    Private Sub RFPCenterToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RFPCenterToolStripMenuItem.Click
        Dim child As New RFPCenter
        child.MdiParent = Me
        child.Show()
    End Sub

    Private Sub MasterMappingCenterToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MasterMappingCenterToolStripMenuItem.Click
        Dim child As New MasterMapping
        child.MdiParent = Me
        child.Show()
    End Sub

    Private Sub UMAToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles UMAToolStripMenuItem.Click
        Dim child As New UMA
        child.MdiParent = Me
        child.Show()
    End Sub

    Private Sub ToolStripDropDownButton4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripDropDownButton4.Click

    End Sub

    Private Sub PositionByAccountToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PositionByAccountToolStripMenuItem.Click
        Dim child As New map_ReportViewer_FirmMaster
        child.MdiParent = Me
        child.Show()
    End Sub

    Private Sub TransactionReportToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TransactionReportToolStripMenuItem.Click
        Dim child As New map_ReportViewer_ProductMaster
        child.MdiParent = Me
        child.Show()
    End Sub

    Private Sub ToolStripMenuItem6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripMenuItem6.Click
        Dim child As New ETF_PriceImportMain
        child.MdiParent = Me
        child.Show()
    End Sub

    Private Sub CmdToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Dim child As New TestCmd
        child.MdiParent = Me
        child.Show()
    End Sub

    Private Sub RevenueCenterToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RevenueCenterToolStripMenuItem.Click
        Dim child As New RevenueCenter
        child.MdiParent = Me
        child.Show()
    End Sub

    Private Sub RepBusinessBreakdownToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RepBusinessBreakdownToolStripMenuItem.Click
        Dim child As New rpt_RepBookBusiness
        child.MdiParent = Me
        child.Show()
    End Sub

    Private Sub TeamManagerToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TeamManagerToolStripMenuItem.Click
        Dim child As New SalesTeams
        child.MdiParent = Me
        child.Show()
    End Sub

    Private Sub LeadManagerToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles LeadManagerToolStripMenuItem.Click
        Dim child As New Leads
        child.MdiParent = Me
        child.Show()
    End Sub

    Private Sub ToolStripDropDownButton9_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripDropDownButton9.Click

    End Sub

    Private Sub AMToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles AMToolStripMenuItem.Click
        Dim child As New AMPFExcelExport
        child.MdiParent = Me
        child.Show()
    End Sub

    Private Sub AccountMovesToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles AccountMovesToolStripMenuItem.Click
        Dim child As New AccountMoves
        child.MdiParent = Me
        child.Show()
    End Sub

    Private Sub EditWhatsNewToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles EditWhatsNewToolStripMenuItem.Click
        Dim child As New WhatsNewEdit
        child.MdiParent = Me
        child.Show()
    End Sub

    Private Sub ToolStripDropDownButton10_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripDropDownButton10.Click
        If Permissions.EditWhatsNew.Checked Then
            ViewWhatsNewToolStripMenuItem.Visible = True
        Else
            ViewWhatsNewToolStripMenuItem.Visible = False
            Dim child As New WhatsNew
            child.MdiParent = Me
            child.Show()
        End If
    End Sub

    Private Sub ViewWhatsNewToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ViewWhatsNewToolStripMenuItem.Click
        Dim child As New WhatsNew
        child.MdiParent = Me
        child.Show()
    End Sub
End Class
