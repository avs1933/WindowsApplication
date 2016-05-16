Public Class mdb_ETF_Pricing_ChangeHoldingWeight

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Dim Mycn As OleDb.OleDbConnection
        Dim Command As OleDb.OleDbCommand
        Dim SQLstr As String

        Try
            Mycn = New OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & My.Settings.dbloc & "';Jet OLEDB:Database Password='" & My.Settings.dbpass & "';")
            Mycn.Open()

            Dim wgt As Double
            wgt = (TextBox1.Text / 100)

            'If wgt = 0 Then
            'SQLstr = "DELETE * FROM mdb_ETFPricing_ModelHoldings WHERE ID = " & ID.Text
            'Else
            SQLstr = "UPDATE mdb_ETFPricing_ModelHoldings SET Weight = " & wgt & " WHERE ID = " & ID.Text
            'End If


            Command = New OleDb.OleDbCommand(SQLstr, Mycn)
            Command.ExecuteNonQuery()

            Mycn.Close()

            Me.Close()

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "ERROR")
        End Try
    End Sub

    Private Sub mdb_ETF_Pricing_ChangeHoldingWeight_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        TextBox1.Focus()
    End Sub
End Class