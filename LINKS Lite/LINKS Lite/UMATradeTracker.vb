Public Class UMATradeTracker

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Call LoadField()
    End Sub


    Public Sub LoadField()
        Try

            'DataGridView2.Enabled = False

            Dim conn As New OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & My.Settings.dbloc & "';Jet OLEDB:Database Password='" & My.Settings.dbpass & "';")

            Dim strSQL As String = "SELECT * FROM uma_tradetracker"

            'If IsDBNull(TextBox1.Text) And IsDBNull(TextBox2.Text) And CheckBox1.Checked = False And CheckBox2.Checked = False Then
            'GoTo line1
            'Else
            'strSQL = strSQL & " WHERE"
            'End If

            'If CheckBox1.Checked Then
            'strSQL = strSQL & " and TradeUploaded = -1"
            'Else

            'End If
            'If CheckBox2.Checked Then
            'strSQL = strSQL & " and TradeRemoved = -1"
            'Else

            'End If
            'If IsDBNull(TextBox1.Text) Then

            'Else
            'strSQL = strSQL & " MoxyOrderID Like '%" & TextBox1.Text & "%'"
            'End If

            'If IsDBNull(TextBox2.Text) Then

            'Else
            'strSQL = strSQL & " ReadyBy Like '%" & TextBox2.Text & "%'"
            'End If

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


    Private Sub DeleteToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles DeleteToolStripMenuItem.Click
        Dim ir As MsgBoxResult

        ir = MsgBox("This will delete this trade from the table and could result in a duplicate file being sent to platform.  This cannot be undone!  Are you sure you want to continue?", MsgBoxStyle.YesNo + MsgBoxStyle.Critical, "Confirm Delete")

        If ir = MsgBoxResult.Yes Then
            Dim Mycn As OleDb.OleDbConnection
            Dim Command As OleDb.OleDbCommand
            Dim SQLstr As String
            Try
                Mycn = New OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & My.Settings.dbloc & "';Jet OLEDB:Database Password='" & My.Settings.dbpass & "';")
                Mycn.Open()

                SQLstr = "DELETE * FROM uma_tradetracker WHERE ID = " & DataGridView1.SelectedCells(0).Value

                Command = New OleDb.OleDbCommand(SQLstr, Mycn)
                Command.ExecuteNonQuery()

                Mycn.Close()
                Call LoadField()

            Catch ex As Exception
                MsgBox(ex.Message, MsgBoxStyle.Critical, "ERROR")
            End Try
        Else

        End If
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        
    End Sub

    Private Sub Button2_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        Call LoadField()
    End Sub

    Private Sub Button1_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Me.Close()
    End Sub

    Private Sub UMATradeTracker_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Call LoadField()
    End Sub
End Class