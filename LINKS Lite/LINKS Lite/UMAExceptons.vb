Public Class UMAExceptons

    Private Sub UMAExceptons_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Call LoadExceptions()
    End Sub

    Public Sub LoadExceptions()
        Try
            'DataGridView1.Enabled = False

            Dim conn As New OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & My.Settings.dbloc & "';Jet OLEDB:Database Password='" & My.Settings.dbpass & "';")

            Dim strSQL As String = "SELECT * FROM uma_exceptions"
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
End Class