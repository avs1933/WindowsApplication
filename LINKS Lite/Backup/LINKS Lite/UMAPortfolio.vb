Public Class UMAPortfolio

    Private Sub UMAPortfolio_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Call LoadPortfolios()
        Call LoadProducts()
    End Sub

    Public Sub LoadPortfolios()
        Try

            Dim conn As New OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & My.Settings.dbloc & "';Jet OLEDB:Database Password='" & My.Settings.dbpass & "';")

            Dim strSQL As String = "SELECT PortfolioID, PortfolioCode FROM dbo_vQbRowDefPortfolio"
            Dim da As New OleDb.OleDbDataAdapter(strSQL, conn)
            Dim ds As New DataSet
            da.Fill(ds, "Users")

            With ComboBox1
                .DataSource = ds.Tables("Users")
                .DisplayMember = "PortfolioCode"
                .ValueMember = "PortfolioCode"
                .SelectedIndex = 0
            End With

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "ERROR")
            'Application.Exit()
        End Try
    End Sub

    Public Sub LoadProducts()
        Try

            Dim conn As New OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & My.Settings.dbloc & "';Jet OLEDB:Database Password='" & My.Settings.dbpass & "';")

            Dim strSQL As String = "SELECT * FROM uma_products_envestnet"
            Dim da As New OleDb.OleDbDataAdapter(strSQL, conn)
            Dim ds As New DataSet
            da.Fill(ds, "Users")

            With ComboBox2
                .DataSource = ds.Tables("Users")
                .DisplayMember = "ProductID"
                .ValueMember = "ID"
                .SelectedIndex = 0
            End With

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "ERROR")
            'Application.Exit()
        End Try
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        If IsDBNull(ComboBox1.Text) Or ComboBox1.Text = "" Then
            ComboBox1.BackColor = Color.Red
            ComboBox1.ForeColor = Color.White
        Else
            ComboBox1.BackColor = Color.White
            ComboBox1.ForeColor = Color.Black
            If IsDBNull(TextBox1.Text) Or TextBox1.Text = "" Then
                TextBox1.BackColor = Color.Red
                TextBox1.ForeColor = Color.White
            Else
                TextBox1.BackColor = Color.White
                TextBox1.ForeColor = Color.Black
                If ID.Text = "NEW" Then
                    Call SaveNew()
                Else
                    Call SaveOld()
                End If
            End If
        End If
    End Sub

    Public Sub SaveOld()
        Dim Mycn As OleDb.OleDbConnection
        Dim Command As OleDb.OleDbCommand
        Dim SQLstr As String
        Try
            Mycn = New OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & My.Settings.dbloc & "';Jet OLEDB:Database Password='" & My.Settings.dbpass & "';")
            Mycn.Open()

            SQLstr = "Update uma_portfolios_envestnet SET PortfolioCode = '" & ComboBox1.SelectedValue & "', ProductID = " & ComboBox2.SelectedValue & ", Active = " & CheckBox1.CheckState & ", PortfolioCodeReal = '" & TextBox1.Text & "' WHERE ID = " & ID.Text

            Command = New OleDb.OleDbCommand(SQLstr, Mycn)
            Command.ExecuteNonQuery()

            Mycn.Close()

            MsgBox("Record Updated.")
            Me.Close()

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "ERROR")
        End Try
    End Sub

    Public Sub SaveNew()
        Dim Mycn As OleDb.OleDbConnection
        Dim Command As OleDb.OleDbCommand
        Dim SQLstr As String
        Try
            Mycn = New OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & My.Settings.dbloc & "';Jet OLEDB:Database Password='" & My.Settings.dbpass & "';")
            Mycn.Open()

            SQLstr = "INSERT INTO uma_portfolios_envestnet(PortfolioCode, ProductID, Active, PortfolioCodeReal)" & _
            "VALUES('" & ComboBox1.SelectedValue & "'," & ComboBox2.SelectedValue & ", -1, '" & TextBox1.Text & "')"

            Command = New OleDb.OleDbCommand(SQLstr, Mycn)
            Command.ExecuteNonQuery()

            Mycn.Close()

            MsgBox("Record Saved.  Portfolio " & ComboBox1.SelectedValue & " is now coded as a UMA Account.")
            Me.Close()

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "ERROR")
        End Try
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        Me.Close()
    End Sub

    Private Sub ComboBox1_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles ComboBox1.LostFocus
        If IsDBNull(TextBox1.Text) Or TextBox1.Text = "" Then
            TextBox1.Text = ComboBox1.SelectedValue
        Else
        End If
    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ComboBox1.SelectedIndexChanged

    End Sub
End Class