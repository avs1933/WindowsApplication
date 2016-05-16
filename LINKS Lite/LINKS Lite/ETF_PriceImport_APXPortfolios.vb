Public Class ETF_PriceImport_APXPortfolios

    Private Sub ETF_PriceImport_APXPortfolios_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Call LoadApprovedPortfolios()
        Call LoadAPXPortfolio()
    End Sub

    Public Sub LoadAPXPortfolio()
        Try

            Dim conn As New OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & My.Settings.dbloc & "';Jet OLEDB:Database Password='" & My.Settings.dbpass & "';")

            Dim strSQL As String = "SELECT PortfolioID, PortfolioCode FROM AdvApp_vPortfolio"
            Dim da As New OleDb.OleDbDataAdapter(strSQL, conn)
            Dim ds As New DataSet
            da.Fill(ds, "Users")

            With ComboBox1
                .DataSource = ds.Tables("Users")
                .DisplayMember = "PortfolioCode"
                .ValueMember = "PortfolioID"
                .SelectedIndex = 0
            End With

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "ERROR")
        End Try
    End Sub

    Public Sub LoadApprovedPortfolios()
        Try
            Dim conn As New OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & My.Settings.dbloc & "';Jet OLEDB:Database Password='" & My.Settings.dbpass & "';")

            Dim strSQL As String

            strSQL = "SELECT mdb_ETFPrice_APXPortfolios.ID, AdvApp_vPortfolio.PortfolioCode" & _
            " FROM AdvApp_vPortfolio INNER JOIN mdb_ETFPrice_APXPortfolios ON AdvApp_vPortfolio.PortfolioID = mdb_ETFPrice_APXPortfolios.PortfolioID" & _
            " WHERE(((mdb_ETFPrice_APXPortfolios.Active) = True))" & _
            " GROUP BY mdb_ETFPrice_APXPortfolios.ID, AdvApp_vPortfolio.PortfolioCode;"

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

    Public Sub CheckForDupe()
        Dim Mycn As OleDb.OleDbConnection
        Dim SQLstring As String
        Try
            Mycn = New OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & My.Settings.dbloc & "';Jet OLEDB:Database Password='" & My.Settings.dbpass & "';")
            Mycn.Open()

            SQLstring = "SELECT * FROM mdb_ETFPrice_APXPortfolios WHERE PortfolioID = " & ComboBox1.SelectedValue & " AND Active = True"

            Dim queryString As String = String.Format(SQLstring)
            Dim cmd As New OleDb.OleDbCommand(queryString, Mycn)
            Dim da As New OleDb.OleDbDataAdapter(cmd)
            Dim ds As New DataSet

            da.Fill(ds, "User")
            Dim dt As DataTable = ds.Tables("User")
            If dt.Rows.Count = 0 Then
                Call ApprovePortfolio()
            Else
                MsgBox("That portfolio is already approved for TAM modeling.", MsgBoxStyle.Information, "DUPLICATE RECORD DETECTED")
            End If

        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub

    Public Sub ApprovePortfolio()
        Dim Mycn As OleDb.OleDbConnection
        Dim Command As OleDb.OleDbCommand
        Dim SQLstr As String
        Try
            Mycn = New OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & My.Settings.dbloc & "';Jet OLEDB:Database Password='" & My.Settings.dbpass & "';")
            Mycn.Open()

            SQLstr = "INSERT INTO mdb_ETFPrice_APXPortfolios(PortfolioID, PortfolioCode, Active)" & _
            "VALUES(" & ComboBox1.SelectedValue & ",'" & ComboBox1.SelectedText & "', -1)"

            Command = New OleDb.OleDbCommand(SQLstr, Mycn)
            Command.ExecuteNonQuery()

            Mycn.Close()

            Call LoadApprovedPortfolios()

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "ERROR")
        End Try
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Call CheckForDupe()
    End Sub

    Private Sub DeleteToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles DeleteToolStripMenuItem.Click
        If DataGridView1.RowCount = 0 Then

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

                    SQLstr = "UPDATE mdb_ETFPrice_APXPortfolios SET Active = False WHERE ID = " & DataGridView1.SelectedCells(0).Value

                    Command = New OleDb.OleDbCommand(SQLstr, Mycn)
                    Command.ExecuteNonQuery()

                    Mycn.Close()

                    Call LoadApprovedPortfolios()

                Catch ex As Exception
                    MsgBox(ex.Message, MsgBoxStyle.Critical, "ERROR")
                End Try
            Else
            End If
        End If
    End Sub
End Class