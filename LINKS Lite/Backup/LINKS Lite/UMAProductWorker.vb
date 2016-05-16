Public Class UMAProductWorker

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        Me.Close()
    End Sub

    Private Sub UMAProductWorker_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Call loadproducts()
    End Sub

    Public Sub loadproducts()
        Try

            'DataGridView2.Enabled = False

            Dim conn As New OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & My.Settings.dbloc & "';Jet OLEDB:Database Password='" & My.Settings.dbpass & "';")

            Dim strSQL As String = "SELECT * FROM uma_products_envestnet"
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

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        Call loadproducts()
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Dim child As New UMAProduct
        child.MdiParent = Home
        child.Show()

        child.ID.Text = "NEW"
    End Sub

    Private Sub EditToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles EditToolStripMenuItem.Click
        If DataGridView1.RowCount = "0" Then
            'do nothing
        Else
            Dim child As New UMAProduct
            child.MdiParent = Home
            child.Show()

            child.ID.Text = DataGridView1.SelectedCells(0).Value
            child.TextBox1.Text = DataGridView1.SelectedCells(1).Value
            child.TextBox2.Text = DataGridView1.SelectedCells(2).Value
        End If
    End Sub

    Private Sub DataGridView1_CellContentClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick

    End Sub

    Private Sub DataGridView1_CellContentDoubleClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView1.CellContentDoubleClick
        If DataGridView1.RowCount = "0" Then
            'do nothing
        Else
            Dim child As New UMAProduct
            child.MdiParent = Home
            child.Show()

            child.ID.Text = DataGridView1.SelectedCells(0).Value
            child.TextBox1.Text = DataGridView1.SelectedCells(1).Value
            child.TextBox2.Text = DataGridView1.SelectedCells(2).Value
        End If
    End Sub

    Private Sub DeleteToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles DeleteToolStripMenuItem.Click
        If DataGridView1.RowCount = "0" Then
            'do nothing
        Else
            Dim ir As MsgBoxResult

            ir = MsgBox("Are you sure you want to delete this record?  This cannot be undone!", MsgBoxStyle.YesNo + MsgBoxStyle.Critical, "Confirm Delete")

            If ir = MsgBoxResult.Yes Then
                Dim Mycn As OleDb.OleDbConnection
                Dim Command As OleDb.OleDbCommand
                Dim SQLstr As String
                Try
                    Mycn = New OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & My.Settings.dbloc & "';Jet OLEDB:Database Password='" & My.Settings.dbpass & "';")
                    Mycn.Open()

                    SQLstr = "DELETE * FROM uma_products_envestnet WHERE ID = " & DataGridView1.SelectedCells(0).Value

                    Command = New OleDb.OleDbCommand(SQLstr, Mycn)
                    Command.ExecuteNonQuery()

                    Mycn.Close()
                    Call loadproducts()

                Catch ex As Exception
                    MsgBox(ex.Message, MsgBoxStyle.Critical, "ERROR")
                End Try
            Else

            End If
        End If
    End Sub
End Class