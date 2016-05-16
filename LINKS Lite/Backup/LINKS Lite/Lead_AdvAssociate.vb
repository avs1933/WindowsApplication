Public Class Lead_AdvAssociate

    Private Sub Lead_AdvAssociate_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Try

            Dim conn As New OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & My.Settings.roledbloc & "';Jet OLEDB:Database Password='" & My.Settings.dbpass & "';")

            Dim strSQL As String = "SELECT DISTINCT leads_Reps.Advisor" & _
            " FROM(leads_Reps)" & _
            " WHERE (((leads_Reps.Advisor) Not In (Select LeadName FROM [leads_advisors])))" & _
            " GROUP BY leads_Reps.Advisor" & _
            " ORDER BY leads_Reps.Advisor;"
            Dim da As New OleDb.OleDbDataAdapter(strSQL, conn)
            Dim ds As New DataSet
            da.Fill(ds, "Users")

            With cboLead
                .DataSource = ds.Tables("Users")
                .DisplayMember = "Advisor"
                '.ValueMember = "ID"
                .SelectedIndex = 0
            End With

            Dim strSQL1 As String = "SELECT DISTINCT dbo_vQbRowDefPortfolio.ExternalAdvisor" & _
            " FROM(dbo_vQbRowDefPortfolio)" & _
            " WHERE dbo_vQbRowDefPortfolio.ExternalAdvisor Not In (SELECT APXName FROM leads_advisors)" & _
            " GROUP BY dbo_vQbRowDefPortfolio.ExternalAdvisor;"
            Dim da1 As New OleDb.OleDbDataAdapter(strSQL1, conn)
            Dim ds1 As New DataSet
            da1.Fill(ds1, "Users")

            With cboAdvisor
                .DataSource = ds1.Tables("Users")
                .DisplayMember = "ExternalAdvisor"
                .SelectedIndex = 0
            End With

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "ERROR")
        End Try
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Dim Mycn As OleDb.OleDbConnection
        Dim Command1 As OleDb.OleDbCommand
        'Dim Command2 As OleDb.OleDbCommand
        Dim SQLstr As String
        Try
            Mycn = New OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & My.Settings.dbloc & "';Jet OLEDB:Database Password='" & My.Settings.dbpass & "';")
            Dim ds1 As New DataSet
            Dim eds1 As New DataGridView
            Dim dv1 As New DataView
            Mycn.Open()

            Dim advisor1 As String
            advisor1 = cboLead.Text
            advisor1 = Replace(advisor1, "'", "''")

            SQLstr = "Insert INTO leads_advisors(LeadName, APXName)" & _
            "Values('" & advisor1 & "','" & cboAdvisor.Text & "')"
            Command1 = New OleDb.OleDbCommand(SQLstr, Mycn)
            Command1.ExecuteNonQuery()

            Mycn.Close()
            MsgBox("Record Saved")
            Me.Close()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "ERROR")
        End Try
    End Sub
End Class