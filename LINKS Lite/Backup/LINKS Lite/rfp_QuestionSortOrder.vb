Public Class rfp_QuestionSortOrder

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        Me.Close()
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Dim Mycn As OleDb.OleDbConnection
        Dim Command As OleDb.OleDbCommand
        Dim SQLstr As String

        Try
            Mycn = New OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & My.Settings.roledbloc & "';Jet OLEDB:Database Password='" & My.Settings.dbpass & "';")

            Mycn.Open()
            SQLstr = "Update rfp_ReportQuestions SET [SortOrder] = " & txtNewSort.Text & " WHERE ID = " & txtID.Text
            Command = New OleDb.OleDbCommand(SQLstr, Mycn)
            Command.ExecuteNonQuery()
            Mycn.Close()

            Mycn.Open()
            SQLstr = "INSERT INTO rfp_ResponseTracking([QuestionID], [Field], [OldValue], [NewValue], [ChangedBy], [DateStamp])" & _
            "VALUES(" & txtID.Text & ",'Sort Order','" & txtOldSort.Text & "','" & txtNewSort.Text & "','" & Environ("USERNAME") & "',#" & Format(Now()) & "#)"
            Command = New OleDb.OleDbCommand(SQLstr, Mycn)
            Command.ExecuteNonQuery()
            Mycn.Close()

            Me.Close()

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "ERROR")
        End Try
    End Sub
End Class