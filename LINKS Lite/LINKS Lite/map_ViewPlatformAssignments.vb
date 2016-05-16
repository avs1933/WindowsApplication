Public Class map_ViewPlatformAssignments

    Private Sub map_ViewPlatformAssignments_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Call LoadPermissions()
        Call LoadPlatformAssigned()
        Call LoadWLAssigned()
    End Sub

    Public Sub LoadPermissions()
        If Permissions.MAPAssociatePlatform.Checked Then
            EditToolStripMenuItem.Visible = True
            DeleteToolStripMenuItem.Visible = True
            ToolStripMenuItem1.Visible = True
            ToolStripMenuItem2.Visible = True
        Else
            EditToolStripMenuItem.Visible = False
            DeleteToolStripMenuItem.Visible = False
            ToolStripMenuItem1.Visible = False
            ToolStripMenuItem2.Visible = False
        End If
    End Sub

    Public Sub LoadPlatformAssigned()
        Try
            Dim conn As New OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & My.Settings.dbloc & "';Jet OLEDB:Database Password='" & My.Settings.dbpass & "';")

            Dim SqlString As String

            SqlString = "SELECT map_PlatformAssignments.ID, Map_Firms.FirmName As [Firm Name], Map_Firms_1.FirmName As [Platform Name], map_Platforms.OfferSMA As [SMA Offered], map_Platforms.OfferUMA As [UMA Offered], map_Platforms.OfferMF As [Mutual Fund Offered], map_Platforms.OfferUIT As [UIT Offered]" & _
            " FROM ((map_PlatformAssignments INNER JOIN Map_Firms ON map_PlatformAssignments.FirmID = Map_Firms.ID) INNER JOIN map_Platforms ON map_PlatformAssignments.PlatformID = map_Platforms.ID) INNER JOIN Map_Firms AS Map_Firms_1 ON map_Platforms.PlatformID = Map_Firms_1.ID" & _
            " WHERE(((map_PlatformAssignments.Active) = True) And ((map_PlatformAssignments.UsePlatform) = True))" & _
            " GROUP BY map_PlatformAssignments.ID, Map_Firms.FirmName, Map_Firms_1.FirmName, map_Platforms.OfferSMA, map_Platforms.OfferUMA, map_Platforms.OfferMF, map_Platforms.OfferUIT;"

            Dim da As New OleDb.OleDbDataAdapter(SqlString, conn)
            Dim ds As New DataSet
            da.Fill(ds, "Users")

            With DataGridView1
                .DataSource = ds.Tables("Users")
                .Columns(0).Visible = False
            End With

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "ERROR")
        End Try
    End Sub

    Public Sub LoadWLAssigned()
        Try
            Dim conn As New OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & My.Settings.dbloc & "';Jet OLEDB:Database Password='" & My.Settings.dbpass & "';")

            Dim SqlString As String

            SqlString = "SELECT map_PlatformAssignments.ID, Map_Firms.FirmName As [Firm Name], map_PlatformsWL.WLName As [White Label Name], Map_Firms_1.FirmName As [Platform Name], map_Platforms.OfferSMA As [SMA Offered], map_Platforms.OfferUMA As [UMA Offered], map_Platforms.OfferMF As [Mutual Fund Offered], map_Platforms.OfferUIT As [UIT Offered]" & _
            " FROM (map_PlatformAssignments INNER JOIN Map_Firms ON map_PlatformAssignments.FirmID = Map_Firms.ID) INNER JOIN (map_PlatformsWL INNER JOIN (map_Platforms INNER JOIN Map_Firms AS Map_Firms_1 ON map_Platforms.PlatformID = Map_Firms_1.ID) ON map_PlatformsWL.PlatformDriverID = map_Platforms.ID) ON map_PlatformAssignments.WLID = map_PlatformsWL.ID" & _
            " WHERE(((map_PlatformAssignments.Active) = True) And ((map_PlatformAssignments.UseWL) = True))" & _
            " GROUP BY map_PlatformAssignments.ID, Map_Firms.FirmName, map_PlatformsWL.WLName, Map_Firms_1.FirmName, map_Platforms.OfferSMA, map_Platforms.OfferUMA, map_Platforms.OfferMF, map_Platforms.OfferUIT;"

            Dim da As New OleDb.OleDbDataAdapter(SqlString, conn)
            Dim ds As New DataSet
            da.Fill(ds, "Users")

            With DataGridView2
                .DataSource = ds.Tables("Users")
                .Columns(0).Visible = False
            End With

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "ERROR")
        End Try
    End Sub

    Private Sub RefreshToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RefreshToolStripMenuItem.Click
        Call LoadPlatformAssigned()
    End Sub

    Private Sub ToolStripMenuItem3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripMenuItem3.Click
        Call LoadWLAssigned()
    End Sub

    Private Sub ToolStripMenuItem1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripMenuItem1.Click
        Call LoadAssignmentEditWL()
    End Sub

    Public Sub LoadAssignmentEditWL()
        If DataGridView2.RowCount = "0" Then
            'do nothing
        Else
            Dim child As New map_PlatformAssignment
            child.MdiParent = Home
            child.Show()

            Try
                Dim conn As New OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & My.Settings.dbloc & "';Jet OLEDB:Database Password='" & My.Settings.dbpass & "';")

                Dim strSQL As String = "SELECT * FROM map_PlatformAssignments WHERE ID = " & DataGridView2.SelectedCells(0).Value
                Dim queryString As String = String.Format(strSQL)
                Dim cmd As New OleDb.OleDbCommand(queryString, conn)
                Dim da As New OleDb.OleDbDataAdapter(cmd)
                Dim ds As New DataSet

                da.Fill(ds, "User")
                Dim dt As DataTable = ds.Tables("User")
                Dim row1 As DataRow = dt.Rows(0)

                child.ID.Text = (row1("ID"))

                child.cboFirm.SelectedValue = (row1("FirmID"))
                child.cboWLPlatform.SelectedValue = (row1("WLID"))
                child.ckbWLPlatform.Checked = True
                child.cboPlatform.Enabled = False
                child.cboWLPlatform.Enabled = True
                child.ckbPlatform.Enabled = False
                child.ckbWLPlatform.Enabled = False
                child.ckbActive.Checked = (row1("Active"))

            Catch ex As Exception
                MsgBox(ex.Message, MsgBoxStyle.Critical, "ERROR")
            End Try
        End If
    End Sub

    Public Sub LoadAssignmentEditPlatform()
        If DataGridView1.RowCount = "0" Then
            'do nothing
        Else
            Dim child As New map_PlatformAssignment
            child.MdiParent = Home
            child.Show()

            Try
                Dim conn As New OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & My.Settings.dbloc & "';Jet OLEDB:Database Password='" & My.Settings.dbpass & "';")

                Dim strSQL As String = "SELECT * FROM map_PlatformAssignments WHERE ID = " & DataGridView1.SelectedCells(0).Value
                Dim queryString As String = String.Format(strSQL)
                Dim cmd As New OleDb.OleDbCommand(queryString, conn)
                Dim da As New OleDb.OleDbDataAdapter(cmd)
                Dim ds As New DataSet

                da.Fill(ds, "User")
                Dim dt As DataTable = ds.Tables("User")
                Dim row1 As DataRow = dt.Rows(0)

                child.ID.Text = (row1("ID"))

                child.cboFirm.SelectedValue = (row1("FirmID"))
                child.cboPlatform.SelectedValue = (row1("PlatformID"))
                child.ckbPlatform.Checked = True
                child.cboPlatform.Enabled = True
                child.cboWLPlatform.Enabled = False
                child.ckbPlatform.Enabled = False
                child.ckbWLPlatform.Enabled = False
                child.ckbActive.Checked = (row1("Active"))

            Catch ex As Exception
                MsgBox(ex.Message, MsgBoxStyle.Critical, "ERROR")
            End Try
        End If
    End Sub

    Private Sub EditToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles EditToolStripMenuItem.Click
        Call LoadAssignmentEditPlatform()
    End Sub

    Private Sub ToolStripMenuItem2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripMenuItem2.Click
        If DataGridView2.RowCount = 0 Then

        Else
            Dim ir As MsgBoxResult
            ir = MsgBox("Are you sure you want to delete this record?  This cannot be undone!", MsgBoxStyle.YesNo + MsgBoxStyle.Critical, "Confirm Delete")
            If ir = MsgBoxResult.Yes Then
                Call RemoveRecordWL()
            Else

            End If
        End If


    End Sub

    Public Sub RemoveRecordWL()
        Dim Mycn As OleDb.OleDbConnection
        Dim Command As OleDb.OleDbCommand
        Dim SQLstr As String
        Try
            Mycn = New OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & My.Settings.dbloc & "';Jet OLEDB:Database Password='" & My.Settings.dbpass & "';")
            Mycn.Open()

            SQLstr = "Update map_PlatformAssignments SET Active = False WHERE ID = " & DataGridView2.SelectedCells(0).Value
            
            Command = New OleDb.OleDbCommand(SQLstr, Mycn)
            Command.ExecuteNonQuery()

            'Call LoadChangesDetection()

            Mycn.Close()

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "ERROR")
        End Try
    End Sub

    Private Sub DeleteToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles DeleteToolStripMenuItem.Click
        If DataGridView1.RowCount = 0 Then

        Else
            Dim ir As MsgBoxResult
            ir = MsgBox("Are you sure you want to delete this record?  This cannot be undone!", MsgBoxStyle.YesNo + MsgBoxStyle.Critical, "Confirm Delete")
            If ir = MsgBoxResult.Yes Then
                Call RemoveRecordPlatform()
            Else

            End If
        End If
    End Sub

    Public Sub RemoveRecordPlatform()
        Dim Mycn As OleDb.OleDbConnection
        Dim Command As OleDb.OleDbCommand
        Dim SQLstr As String
        Try
            Mycn = New OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & My.Settings.dbloc & "';Jet OLEDB:Database Password='" & My.Settings.dbpass & "';")
            Mycn.Open()

            SQLstr = "Update map_PlatformAssignments SET Active = False WHERE ID = " & DataGridView1.SelectedCells(0).Value

            Command = New OleDb.OleDbCommand(SQLstr, Mycn)
            Command.ExecuteNonQuery()

            'Call LoadChangesDetection()

            Mycn.Close()

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "ERROR")
        End Try
    End Sub
End Class