Public Class AMP_SecTypes

    Private Sub AMP_SecTypes_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Call LoadGrid()
        Call loadsectypes()
    End Sub
    Public Sub LoadGrid()
        Try

            Dim conn As New OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & My.Settings.dbloc & "';Jet OLEDB:Database Password='" & My.Settings.dbpass & "';")

            Dim strSQL As String = "SELECT AMP_SecTypes.ID, AMP_SecTypes.APXSecType As [Sec Type], AdvApp_vSecType.SecTypeNameLong As [Type Name]" & _
            " FROM AMP_SecTypes INNER JOIN AdvApp_vSecType ON AMP_SecTypes.APXSecType = AdvApp_vSecType.SecTypeCode;"

            Dim da As New OleDb.OleDbDataAdapter(strSQL, conn)
            Dim ds As New DataSet
            da.Fill(ds, "Users")

            With DataGridView1
                .DataSource = ds.Tables("Users")
                .Columns(0).Visible = False
            End With

            'TextBox6.Text = DataGridView4.RowCount

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "ERROR")
        End Try
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        Call LoadGrid()
        Call loadsectypes()
    End Sub

    Public Sub loadsectypes()
        Try

            Dim conn As New OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & My.Settings.dbloc & "';Jet OLEDB:Database Password='" & My.Settings.dbpass & "';")

            Dim strSQL As String = "SELECT AdvApp_vSecType.SecTypeCode" & _
            " FROM(AdvApp_vSecType)" & _
            " WHERE (((AdvApp_vSecType.SecTypeCode) Not In (Select APXSecType from AMP_SecTypes)))" & _
            " GROUP BY AdvApp_vSecType.SecTypeCode;"


            Dim da As New OleDb.OleDbDataAdapter(strSQL, conn)
            Dim ds As New DataSet
            da.Fill(ds, "Users")

            With ComboBox1
                .DataSource = ds.Tables("Users")
                .DisplayMember = "SecTypeCode"
                .ValueMember = "SecTypeCode"
                .SelectedIndex = 0
            End With

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "ERROR")
        End Try
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Dim Mycn As OleDb.OleDbConnection
        Dim Command As OleDb.OleDbCommand
        Dim SQLstr As String

        If ComboBox1.Text = "" Or IsDBNull(ComboBox1.Text) Then
            ComboBox1.BackColor = Color.Red
            ComboBox1.ForeColor = Color.White
        Else
            ComboBox1.BackColor = Color.White
            ComboBox1.ForeColor = Color.Black
            Try
                Mycn = New OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & My.Settings.dbloc & "';Jet OLEDB:Database Password='" & My.Settings.dbpass & "';")
                Dim ds1 As New DataSet
                Dim eds1 As New DataGridView
                Dim dv1 As New DataView
                Mycn.Open()

                SQLstr = "INSERT INTO AMP_SecTypes (APXSecType, DateAdded, AddedByID)" & _
                "VALUES('" & ComboBox1.Text & "',#" & Format(Now()) & "#," & My.Settings.userid & ")"
                Command = New OleDb.OleDbCommand(SQLstr, Mycn)
                Command.ExecuteNonQuery()

                Call LoadGrid()
                Call loadsectypes()

            Catch ex As Exception
                MsgBox(ex.Message, MsgBoxStyle.Critical, "ERROR")
            End Try
        End If
    End Sub

    Private Sub RemoveToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RemoveToolStripMenuItem.Click
        Dim ir As MsgBoxResult
        Dim Mycn As OleDb.OleDbConnection
        Dim Command As OleDb.OleDbCommand
        Dim SQLstr As String

        If DataGridView1.RowCount = 0 Then

        Else

            ir = MsgBox("Are you sure you want to delete this security type?", MsgBoxStyle.YesNo + MsgBoxStyle.Critical, "Confirm Delete")

            If ir = MsgBoxResult.Yes Then
                Try
                    Mycn = New OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & My.Settings.dbloc & "';Jet OLEDB:Database Password='" & My.Settings.dbpass & "';")
                    Dim ds1 As New DataSet
                    Dim eds1 As New DataGridView
                    Dim dv1 As New DataView
                    Mycn.Open()

                    SQLstr = "DELETE * FROM AMP_SecTypes WHERE ID = " & DataGridView1.SelectedCells(0).Value
                    Command = New OleDb.OleDbCommand(SQLstr, Mycn)
                    Command.ExecuteNonQuery()

                    Call LoadGrid()
                    Call loadsectypes()

                Catch ex As Exception
                    MsgBox(ex.Message, MsgBoxStyle.Critical, "ERROR")
                End Try
            Else
            End If
        End If
    End Sub
End Class