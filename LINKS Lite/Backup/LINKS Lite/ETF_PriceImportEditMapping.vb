Public Class ETF_PriceImportEditMapping

    Private Sub ETF_PriceImportEditMapping_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Call LoadTbl()
    End Sub

    Public Sub LoadTbl()
        Try
            Dim conn As New OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & My.Settings.dbloc & "';Jet OLEDB:Database Password='" & My.Settings.dbpass & "';")

            Dim strSQL As String
            If CheckBox1.Checked Then
                strSQL = "SELECT mdb_SymbolMapping.ID, mdb_SymbolMapping.APXSymbol, mdb_SymbolMapping.NewAPXSymbol, mdb_SymbolMapping.APXSecType As [Mapped From Sec Type], AdvApp_vSecurity.Symbol As [Mapped From Symbol], mdb_SymbolMapping.NewAPXSecType As [Mapped To Sec Type], AdvApp_vSecurity_1.Symbol As [Mapped To Symbol], mdb_SymbolMapping.Active" & _
                " FROM (mdb_SymbolMapping INNER JOIN AdvApp_vSecurity ON mdb_SymbolMapping.APXSymbol = AdvApp_vSecurity.SecurityID) INNER JOIN AdvApp_vSecurity AS AdvApp_vSecurity_1 ON mdb_SymbolMapping.NewAPXSymbol = AdvApp_vSecurity_1.SecurityID" & _
                " WHERE mdb_SymbolMapping.Active = False"
            Else
                strSQL = "SELECT mdb_SymbolMapping.ID, mdb_SymbolMapping.APXSymbol, mdb_SymbolMapping.NewAPXSymbol, mdb_SymbolMapping.APXSecType As [Mapped From Sec Type], AdvApp_vSecurity.Symbol As [Mapped From Symbol], mdb_SymbolMapping.NewAPXSecType As [Mapped To Sec Type], AdvApp_vSecurity_1.Symbol As [Mapped To Symbol], mdb_SymbolMapping.Active" & _
                " FROM (mdb_SymbolMapping INNER JOIN AdvApp_vSecurity ON mdb_SymbolMapping.APXSymbol = AdvApp_vSecurity.SecurityID) INNER JOIN AdvApp_vSecurity AS AdvApp_vSecurity_1 ON mdb_SymbolMapping.NewAPXSymbol = AdvApp_vSecurity_1.SecurityID" & _
                " WHERE mdb_SymbolMapping.Active = True"
            End If

            Dim da As New OleDb.OleDbDataAdapter(strSQL, conn)
            Dim ds As New DataSet
            da.Fill(ds, "Users")

            With DataGridView1
                .DataSource = ds.Tables("Users")
                .Columns(0).Visible = False
                .Columns(1).Visible = False
                .Columns(2).Visible = False
            End With

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "ERROR")
        End Try
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Call LoadTbl()
    End Sub

    Private Sub EditToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles EditToolStripMenuItem.Click
        If DataGridView1.RowCount = 0 Then

        Else
            Dim child As New ETF_PriceImportMapping
            child.MdiParent = Home
            child.Show()
            child.ID.Text = DataGridView1.SelectedCells(0).Value
            child.SecTypeFrom.Text = DataGridView1.SelectedCells(3).Value
            child.cboSymbolFrom.SelectedValue = DataGridView1.SelectedCells(1).Value
            child.SecTypeTo.Text = DataGridView1.SelectedCells(5).Value
            child.cboSymbolTo.SelectedValue = DataGridView1.SelectedCells(2).Value
            child.Active.Checked = DataGridView1.SelectedCells(7).Value
        End If
    End Sub

    Private Sub DeleteToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles DeleteToolStripMenuItem.Click
        If DataGridView1.RowCount = 0 Then
            'do nothing
        Else
            Dim ir As MsgBoxResult
            ir = MsgBox("Are you sure you want to delete this record?", MsgBoxStyle.YesNo + MsgBoxStyle.Information, "Confirm Delete")
            If ir = MsgBoxResult.Yes Then
                Dim Mycn As OleDb.OleDbConnection
                Dim Command As OleDb.OleDbCommand
                Dim SQLstr As String
                Try
                    Mycn = New OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & My.Settings.dbloc & "';Jet OLEDB:Database Password='" & My.Settings.dbpass & "';")
                    Mycn.Open()


                    SQLstr = "Update Map_Firms SET Active = False WHERE ID = " & DataGridView1.SelectedCells(0).Value


                    Command = New OleDb.OleDbCommand(SQLstr, Mycn)
                    Command.ExecuteNonQuery()

                    Mycn.Close()
                    Call LoadTbl()

                Catch ex As Exception
                    MsgBox(ex.Message, MsgBoxStyle.Critical, "ERROR")
                End Try
            Else
                'do nothing
            End If
        End If
    End Sub

    
End Class