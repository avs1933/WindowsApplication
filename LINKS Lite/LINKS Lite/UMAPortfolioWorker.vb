Public Class UMAPortfolioWorker

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        If CheckBox1.Checked Then
            LoadIAtbl()
        Else
            Call LoadTbl()
        End If
    End Sub

    Public Sub LoadIAtbl()
        Try

            'DataGridView2.Enabled = False

            Dim conn As New OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & My.Settings.dbloc & "';Jet OLEDB:Database Password='" & My.Settings.dbpass & "';")

            Dim strSQL As String = "SELECT uma_portfolios_envestnet.ID, uma_portfolios_envestnet.PortfolioCode As [APX Portfolio Code], uma_products_envestnet.ProductID, uma_products_envestnet.Discipline, uma_portfolios_envestnet.PortfolioCodeReal As [Platform Portfolio Code], uma_portfolios_envestnet.Active" & _
            " FROM uma_portfolios_envestnet INNER JOIN uma_products_envestnet ON uma_portfolios_envestnet.ProductID = uma_products_envestnet.ID" & _
            " WHERE uma_portfolios_envestnet.Active <> -1 and uma_portfolios_envestnet.PortfolioCode Like '%" & TextBox1.Text & "%'"
            Dim da As New OleDb.OleDbDataAdapter(strSQL, conn)
            Dim ds As New DataSet
            da.Fill(ds, "Users")

            With DataGridView1
                .DataSource = ds.Tables("Users")
                .Columns(0).Visible = False
            End With

            'Label1.Text = "Loaded " & DataGridView2.RowCount & " ready trades."

            'DataGridView2.Enabled = True

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "ERROR")
        End Try
    End Sub

    Public Sub LoadTbl()
        Try

            'DataGridView2.Enabled = False

            Dim conn As New OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & My.Settings.dbloc & "';Jet OLEDB:Database Password='" & My.Settings.dbpass & "';")

            Dim strSQL As String = "SELECT uma_portfolios_envestnet.ID, uma_products_envestnet.ID, uma_portfolios_envestnet.PortfolioCode As [APX Portfolio Code], uma_products_envestnet.ProductID, uma_products_envestnet.Discipline, uma_portfolios_envestnet.PortfolioCodeReal As [Platform Portfolio Code], uma_portfolios_envestnet.Active" & _
            " FROM uma_portfolios_envestnet INNER JOIN uma_products_envestnet ON uma_portfolios_envestnet.ProductID = uma_products_envestnet.ID" & _
            " WHERE uma_portfolios_envestnet.PortfolioCode Like '%" & TextBox1.Text & "%'"
            Dim da As New OleDb.OleDbDataAdapter(strSQL, conn)
            Dim ds As New DataSet
            da.Fill(ds, "Users")

            With DataGridView1
                .DataSource = ds.Tables("Users")
                .Columns(0).Visible = False
                .Columns(1).Visible = False
            End With

            'Label1.Text = "Loaded " & DataGridView2.RowCount & " ready trades."

            'DataGridView2.Enabled = True

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "ERROR")
        End Try
    End Sub

    Private Sub DeleteToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles DeleteToolStripMenuItem.Click
        Dim ir As MsgBoxResult
        If DataGridView1.RowCount = "0" Then
            'do nothing
        Else
            ir = MsgBox("You are about to delete this portfolio from the UMA table.  This cannot be undone!  Are you sure you want to continue?", MsgBoxStyle.YesNo + MsgBoxStyle.Critical, "Confirm Delete")

            If ir = MsgBoxResult.Yes Then
                Dim Mycn As OleDb.OleDbConnection
                Dim Command1 As OleDb.OleDbCommand
                Dim SQLstr As String
                Try
                    Mycn = New OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & My.Settings.dbloc & "';Jet OLEDB:Database Password='" & My.Settings.dbpass & "';")
                    Mycn.Open()

                    SQLstr = "DELETE * FROM uma_portfolios_envestnet WHERE ID = " & DataGridView1.SelectedCells(0).Value

                    Command1 = New OleDb.OleDbCommand(SQLstr, Mycn)
                    Command1.ExecuteNonQuery()

                    Mycn.Close()

                    Call LoadTbl()

                Catch ex As Exception
                    MsgBox(ex.Message, MsgBoxStyle.Critical, "ERROR")
                End Try
            Else

            End If
        End If
    End Sub

    Private Sub InActiveToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles InActiveToolStripMenuItem.Click
        If DataGridView1.RowCount = "0" Then
            'do nothing
        Else
            'ir = MsgBox("You are about to delete this portfolio from the UMA table.  This cannot be undone!  Are you sure you want to continue?", MsgBoxStyle.YesNo + MsgBoxStyle.Critical, "Confirm Delete")

            If DataGridView1.SelectedCells(5).Value = -1 Then


                Dim Mycn As OleDb.OleDbConnection
                Dim Command1 As OleDb.OleDbCommand
                Dim SQLstr As String
                Try
                    Mycn = New OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & My.Settings.dbloc & "';Jet OLEDB:Database Password='" & My.Settings.dbpass & "';")
                    Mycn.Open()

                    SQLstr = "UPDATE uma_portfolios_envestnet SET Active = Null WHERE ID = " & DataGridView1.SelectedCells(0).Value

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

                    SQLstr = "UPDATE uma_portfolios_envestnet SET Active = -1 WHERE ID = " & DataGridView1.SelectedCells(0).Value

                    Command1 = New OleDb.OleDbCommand(SQLstr, Mycn)
                    Command1.ExecuteNonQuery()

                    Mycn.Close()

                    Call LoadTbl()

                Catch ex As Exception
                    MsgBox(ex.Message, MsgBoxStyle.Critical, "ERROR")
                End Try
            End If
        End If
    End Sub

    Private Sub EditToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles EditToolStripMenuItem.Click
        If DataGridView1.RowCount = "0" Then
            'do nothing
        Else
            Dim child As New UMAPortfolio
            child.MdiParent = Home
            child.Show()

            child.ComboBox1.SelectedValue = DataGridView1.SelectedCells(2).Value
            child.ComboBox2.SelectedValue = DataGridView1.SelectedCells(1).Value

            If IsDBNull(DataGridView1.SelectedCells(5).Value) Then
            Else
                child.TextBox1.Text = DataGridView1.SelectedCells(5).Value
            End If

            child.ID.Text = DataGridView1.SelectedCells(0).Value
        End If
    End Sub

    Private Sub DataGridView1_CellContentClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick

    End Sub

    Private Sub DataGridView1_CellContentDoubleClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView1.CellContentDoubleClick
        If DataGridView1.RowCount = "0" Then
            'do nothing
        Else
            Dim child As New UMAPortfolio
            child.MdiParent = Home
            child.Show()

            child.ComboBox1.SelectedValue = DataGridView1.SelectedCells(2).Value
            child.ComboBox2.SelectedValue = DataGridView1.SelectedCells(1).Value
            If IsDBNull(DataGridView1.SelectedCells(5).Value) Then
            Else
                child.TextBox1.Text = DataGridView1.SelectedCells(5).Value
            End If
            child.ID.Text = DataGridView1.SelectedCells(0).Value
        End If
    End Sub

    Private Sub EditToolStripMenuItem1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub

    Private Sub UMAPortfolioWorker_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

    End Sub
End Class