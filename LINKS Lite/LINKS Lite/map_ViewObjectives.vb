Public Class map_ViewObjectives

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        Dim child As New map_EditObjective
        child.MdiParent = Home
        child.Show()
        child.ID.Text = "NEW"
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Me.Close()
    End Sub

    Private Sub map_ViewObjectives_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Call LoadTbl()
    End Sub

    Public Sub LoadPermissions()
        If Permissions.MAPAddObjective.Checked Then
            Button3.Enabled = True
        Else
            Button3.Enabled = False
        End If

        If Permissions.MAPEditObjective.Checked Then
            EditToolStripMenuItem.Visible = True
        Else
            EditToolStripMenuItem.Visible = False
        End If
    End Sub
    Public Sub LoadTbl()
        Try
            Dim conn As New OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & My.Settings.dbloc & "';Jet OLEDB:Database Password='" & My.Settings.dbpass & "';")

            Dim strSQL As String = "SELECT * FROM map_Objective"
            Dim da As New OleDb.OleDbDataAdapter(strSQL, conn)
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

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        Call LoadTbl()
    End Sub

    Private Sub EditToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles EditToolStripMenuItem.Click
        Call EditRow()
    End Sub

    Public Sub EditRow()
        If DataGridView1.RowCount = "0" Then
            'do nothing
        Else
            Dim child As New map_EditObjective
            child.MdiParent = Home
            child.Show()
            child.ID.Text = DataGridView1.SelectedCells(0).Value
            child.TextBox2.Text = DataGridView1.SelectedCells(1).Value
        End If
    End Sub

    Private Sub DataGridView1_CellContentDoubleClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView1.CellContentDoubleClick
        If Permissions.MAPEditObjective.Checked Then
            Call EditRow()
        Else
        End If
    End Sub

    Private Sub DataGridView1_CellContentClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick

    End Sub
End Class