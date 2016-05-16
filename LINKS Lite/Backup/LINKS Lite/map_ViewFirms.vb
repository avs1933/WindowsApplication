Public Class map_ViewFirms
    Dim tableloader As System.Threading.Thread

    Private Sub map_ViewFirms_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        'Control.CheckForIllegalCrossThreadCalls = False
        'tableloader = New System.Threading.Thread(AddressOf LoadTbl)
        'tableloader.Start()
        Call LoadPermissions()
        Call LoadTbl()
    End Sub

    Public Sub LoadPermissions()
        If Permissions.MAPEditFirms.Checked Then
            EditToolStripMenuItem.Visible = True
            RemoveAsActiveToolStripMenuItem.Visible = True
        Else
            EditToolStripMenuItem.Visible = False
            RemoveAsActiveToolStripMenuItem.Visible = False
        End If

        If Permissions.MAPAddFirm.Checked Then
            Button3.Enabled = True
        Else
            Button3.Enabled = False
        End If
    End Sub

    Public Sub LoadTbl()
        Try
            Dim conn As New OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & My.Settings.dbloc & "';Jet OLEDB:Database Password='" & My.Settings.dbpass & "';")

            Dim strSQL As String = "SELECT Map_Firms.ID, map_FirmType.ID, Map_Firms.AdventContactID, Map_Firms.AdventPortfolioFirm, Map_Firms.FirmName, Map_Firms.Address, Map_Firms.City, Map_Firms.State, Map_Firms.Zip, Map_Firms.Phone, map_FirmType.FirmType, Map_Firms.NotInAdvent, Map_Firms.Active" & _
            " FROM (Map_Firms INNER JOIN AdvApp_vContact ON Map_Firms.AdventContactID = AdvApp_vContact.ContactID) INNER JOIN map_FirmType ON Map_Firms.TypeID = map_FirmType.ID" & _
            " WHERE Map_Firms.Active = -1 AND map_FirmType.SystemValue = False"
            Dim da As New OleDb.OleDbDataAdapter(strSQL, conn)
            Dim ds As New DataSet
            da.Fill(ds, "Users")

            With DataGridView1
                .DataSource = ds.Tables("Users")
                .Columns(0).Visible = False
                .Columns(1).Visible = False
                .Columns(2).Visible = False
                .Columns(3).Visible = False
            End With

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "ERROR")
        End Try
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        'Control.CheckForIllegalCrossThreadCalls = False
        'tableloader = New System.Threading.Thread(AddressOf LoadTbl)
        'tableloader.Start()

        Call LoadTbl()
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Me.Close()
    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        Dim child As New map_EditMFirms
        child.MdiParent = Home
        child.Show()
        child.ID.Text = "NEW"
    End Sub

    Private Sub EditToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles EditToolStripMenuItem.Click
        If DataGridView1.RowCount = "0" Then
            'do nothing
        Else
            Dim child As New map_EditMFirms
            child.MdiParent = Home
            child.Show()
            child.ID.Text = DataGridView1.SelectedCells(0).Value
            child.TextBox1.Text = DataGridView1.SelectedCells(4).Value
            child.TextBox2.Text = DataGridView1.SelectedCells(5).Value
            child.TextBox3.Text = DataGridView1.SelectedCells(6).Value
            child.TextBox4.Text = DataGridView1.SelectedCells(7).Value
            child.TextBox5.Text = DataGridView1.SelectedCells(8).Value
            child.TextBox6.Text = DataGridView1.SelectedCells(9).Value
            child.CheckBox1.Checked = DataGridView1.SelectedCells(11).Value
            child.cboAPXFirms.SelectedValue = DataGridView1.SelectedCells(2).Value
            child.cboFirmType.SelectedValue = DataGridView1.SelectedCells(1).Value

            If IsDBNull(DataGridView1.SelectedCells(3).Value) Then
            Else
                child.cboPortfolioFirm.Text = DataGridView1.SelectedCells(3).Value
            End If
        End If
    End Sub

    Private Sub RemoveAsActiveToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RemoveAsActiveToolStripMenuItem.Click
        If DataGridView1.RowCount = "0" Then
            'do nothing
        Else
            Dim ir As MsgBoxResult
            ir = MsgBox("You are mark this firm as Inactive.  This cannot be undone!  Are you sure you want to continue?", MsgBoxStyle.YesNo + MsgBoxStyle.Critical, "Confirm Delete")

            If ir = MsgBoxResult.Yes Then
                If DataGridView1.SelectedCells(12).Value = -1 Then


                    Dim Mycn As OleDb.OleDbConnection
                    Dim Command1 As OleDb.OleDbCommand
                    Dim SQLstr As String
                    Try
                        Mycn = New OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & My.Settings.dbloc & "';Jet OLEDB:Database Password='" & My.Settings.dbpass & "';")
                        Mycn.Open()

                        SQLstr = "UPDATE Map_Firms SET Active = Null WHERE ID = " & DataGridView1.SelectedCells(0).Value

                        Command1 = New OleDb.OleDbCommand(SQLstr, Mycn)
                        Command1.ExecuteNonQuery()

                        Mycn.Close()

                        Call LoadTbl()

                    Catch ex As Exception
                        MsgBox(ex.Message, MsgBoxStyle.Critical, "ERROR")
                    End Try
                Else
                    Dim Mycn As OleDb.OleDbConnection
                    Dim Command1 As OleDb.OleDbCommand
                    Dim SQLstr As String
                    Try
                        Mycn = New OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & My.Settings.dbloc & "';Jet OLEDB:Database Password='" & My.Settings.dbpass & "';")
                        Mycn.Open()

                        SQLstr = "UPDATE Map_Firms SET Active = -1 WHERE ID = " & DataGridView1.SelectedCells(0).Value

                        Command1 = New OleDb.OleDbCommand(SQLstr, Mycn)
                        Command1.ExecuteNonQuery()

                        Mycn.Close()

                        Call LoadTbl()

                    Catch ex As Exception
                        MsgBox(ex.Message, MsgBoxStyle.Critical, "ERROR")
                    End Try
                End If
            Else
            End If
        End If
    End Sub

    Private Sub DataGridView1_CellContentClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick

    End Sub

    Private Sub DataGridView1_CellContentDoubleClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView1.CellContentDoubleClick

        If Permissions.MAPEditFirms.Checked Then
            If DataGridView1.RowCount = "0" Then
                'do nothing
            Else
                Dim child As New map_EditMFirms
                child.MdiParent = Home
                child.Show()
                child.ID.Text = DataGridView1.SelectedCells(0).Value
                child.TextBox1.Text = DataGridView1.SelectedCells(4).Value
                child.TextBox2.Text = DataGridView1.SelectedCells(5).Value
                child.TextBox3.Text = DataGridView1.SelectedCells(6).Value
                child.TextBox4.Text = DataGridView1.SelectedCells(7).Value
                child.TextBox5.Text = DataGridView1.SelectedCells(8).Value
                child.TextBox6.Text = DataGridView1.SelectedCells(9).Value
                child.CheckBox1.Checked = DataGridView1.SelectedCells(11).Value
                child.cboAPXFirms.SelectedValue = DataGridView1.SelectedCells(2).Value
                child.cboFirmType.SelectedValue = DataGridView1.SelectedCells(1).Value

                If IsDBNull(DataGridView1.SelectedCells(3).Value) Then
                Else
                    child.cboPortfolioFirm.Text = DataGridView1.SelectedCells(3).Value
                End If
            End If
        Else
        End If
    End Sub

    Private Sub ToolStripMenuItem1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripMenuItem1.Click
        Dim child As New map_Firm
        child.MdiParent = Home
        child.ID.Text = DataGridView1.SelectedCells(0).Value
        child.TextBox1.Text = DataGridView1.SelectedCells(4).Value
        child.Show()
    End Sub
End Class