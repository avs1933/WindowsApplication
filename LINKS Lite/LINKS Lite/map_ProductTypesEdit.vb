Public Class map_ProductTypesEdit

    Private Sub map_ProductTypesEdit_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Call loadproducttypes()
    End Sub

    Public Sub loadproducttypes()
        Try

            'DataGridView2.Enabled = False

            Dim conn As New OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & My.Settings.dbloc & "';Jet OLEDB:Database Password='" & My.Settings.dbpass & "';")

            Dim strSQL As String = "SELECT * FROM map_ProductType"
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
        Dim child As New map_EditProductType
        child.MdiParent = Home
        child.Show()
        child.TypeID.Text = "NEW"
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        Call loadproducttypes()
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Me.Close()
    End Sub

    Private Sub EditToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles EditToolStripMenuItem.Click
        If DataGridView1.RowCount = "0" Then
            'do nothing
        Else
            Dim child As New map_EditProductType
            child.MdiParent = Home
            child.Show()

            child.TypeID.Text = DataGridView1.SelectedCells(0).Value
            child.TextBox2.Text = DataGridView1.SelectedCells(1).Value
            child.CheckBox1.Checked = DataGridView1.SelectedCells(2).Value
            child.CheckBox2.Checked = DataGridView1.SelectedCells(3).Value
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

                    SQLstr = "DELETE * FROM map_ProductType WHERE ID = " & DataGridView1.SelectedCells(0).Value

                    Command = New OleDb.OleDbCommand(SQLstr, Mycn)
                    Command.ExecuteNonQuery()

                    Mycn.Close()
                    Call loadproducttypes()

                Catch ex As Exception
                    MsgBox(ex.Message, MsgBoxStyle.Critical, "ERROR")
                End Try
            Else

            End If
        End If
    End Sub
End Class