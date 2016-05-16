Public Class ETF_PriceImport_SecData

    Private Sub ETF_PriceImport_SecData_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Call LoadSecurities()
        Call LoadAssetClass()
        Call LoadSecType()
    End Sub

    Public Sub LoadSecurities()
        Try

            Dim conn As New OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & My.Settings.dbloc & "';Jet OLEDB:Database Password='" & My.Settings.dbpass & "';")

            Dim strSQL As String = "SELECT * FROM mdb_SecurityWithSecType"
            Dim da As New OleDb.OleDbDataAdapter(strSQL, conn)
            Dim ds As New DataSet
            da.Fill(ds, "Users")

            With ComboBox1
                .DataSource = ds.Tables("Users")
                .DisplayMember = "Symbol"
                .ValueMember = "SecurityID"
                .SelectedIndex = 0
            End With

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "ERROR")
        End Try
    End Sub

    Public Sub LoadSecType()
        Try

            Dim conn As New OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & My.Settings.dbloc & "';Jet OLEDB:Database Password='" & My.Settings.dbpass & "';")

            Dim strSQL As String = "SELECT (AdvApp_vSecType.SecTypeCode & AdvApp_vSecType.PrincipalCurrencyCode) As SecType FROM AdvApp_vSecType;"
            Dim da As New OleDb.OleDbDataAdapter(strSQL, conn)
            Dim ds As New DataSet
            da.Fill(ds, "Users")

            With SecType
                .DataSource = ds.Tables("Users")
                .DisplayMember = "SecType"
                .ValueMember = "SecType"
                .SelectedIndex = 0
            End With

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "ERROR")
        End Try
    End Sub

    Public Sub LoadAssetClass()
        Try

            Dim conn As New OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & My.Settings.dbloc & "';Jet OLEDB:Database Password='" & My.Settings.dbpass & "';")

            Dim strSQL As String = "SELECT * FROM AdvApp_vAssetClass"
            Dim da As New OleDb.OleDbDataAdapter(strSQL, conn)
            Dim ds As New DataSet
            da.Fill(ds, "Users")

            With AssetClass
                .DataSource = ds.Tables("Users")
                .DisplayMember = "AssetClassName"
                .ValueMember = "AssetClassCode"
                .SelectedIndex = 0
            End With

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "ERROR")
        End Try
    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ComboBox1.SelectedIndexChanged
        If ComboBox1.ValueMember.Count = 0 Then

        Else
            Try
                Dim conn As New OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & My.Settings.dbloc & "';Jet OLEDB:Database Password='" & My.Settings.dbpass & "';")

                Dim strSQL As String = "SELECT * FROM mdb_SecurityWithSecType WHERE SecurityID = " & ComboBox1.SelectedValue
                Dim queryString As String = String.Format(strSQL)
                Dim cmd As New OleDb.OleDbCommand(queryString, conn)
                Dim da As New OleDb.OleDbDataAdapter(cmd)
                Dim ds As New DataSet

                da.Fill(ds, "User")
                Dim dt As DataTable = ds.Tables("User")

                Dim row1 As DataRow = dt.Rows(0)

                lblSecName.Text = (row1("FullName"))
                NewSymbol.Text = "Z" & (row1("Symbol"))
                InterestRate.Text = (row1("InterestOrDividendRate"))

            Catch ex As Exception
                MsgBox(ex.Message, MsgBoxStyle.Critical, "ERROR")
            End Try

        End If
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        If ID.Text = "NEW" Then
            Call CheckForDupe()
        Else
            Call SaveOld()
        End If
    End Sub

    Public Sub SaveNew()
        Dim Mycn As OleDb.OleDbConnection
        Dim Command As OleDb.OleDbCommand
        Dim SQLstr As String
        Dim dte1 As Date
        dte1 = MaturityDate.Text
        Dim dte2 As String
        dte2 = Format(dte1, "MMddyyyy").ToString
        Try
            Mycn = New OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & My.Settings.dbloc & "';Jet OLEDB:Database Password='" & My.Settings.dbpass & "';")
            Mycn.Open()

            SQLstr = "INSERT INTO mdb_ETFPrice_SecurityData(APXSecurityID, MaturityDate, InterestRate, PaymentFreq, Rating, Rating2, AvgLife, YTW, Duration, CMOResetRule, PrimarySymbolType, AssetClassLongCode, IssueCountry, NewSecType, NewSymbol, [Desc], TrueDate)" & _
            "VALUES(" & ComboBox1.SelectedValue & ", '" & dte2 & "','" & InterestRate.Text & "','" & PaymentFreq.Text & "','" & Rating.Text & "','" & Rating2.Text & "','" & AvgLife.Text & "','" & YTW.Text & "','" & Duration.Text & "','" & CMOResetRule.Text & "','" & PrimarySymbolType.Text & "', '" & AssetClass.SelectedValue & "','us','" & SecType.SelectedValue & "','" & NewSymbol.Text & "','" & lblSecName.Text & "', #" & MaturityDate.Text & "#)"

            Command = New OleDb.OleDbCommand(SQLstr, Mycn)
            Command.ExecuteNonQuery()

            Mycn.Close()

            MsgBox("Record Saved.")
            Me.Close()

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "ERROR")
        End Try
    End Sub

    Public Sub SaveOld()
        Dim Mycn As OleDb.OleDbConnection
        Dim Command As OleDb.OleDbCommand
        Dim SQLstr As String

        Dim dte1 As Date
        dte1 = MaturityDate.Text
        Dim dte2 As String
        dte2 = Format(dte1, "MMddyyyy").ToString
        Try
            Mycn = New OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & My.Settings.dbloc & "';Jet OLEDB:Database Password='" & My.Settings.dbpass & "';")
            Mycn.Open()

            SQLstr = "Update mdb_ETFPrice_SecurityData SET APXSecurityID = " & ComboBox1.SelectedValue & ", MaturityDate = '" & dte2 & "', InterestRate = '" & InterestRate.Text & "', PaymentFreq = '" & PaymentFreq.Text & "', Rating = '" & Rating.Text & "', Rating2 = '" & Rating2.Text & "', AvgLife = '" & AvgLife.Text & "', YTW = '" & YTW.Text & "', Duration = '" & Duration.Text & "', CMOResetRule = '" & CMOResetRule.Text & "', PrimarySymbolType = '" & PrimarySymbolType.Text & "', AssetClassLongCode = '" & AssetClass.SelectedValue & "', IssueCountry = 'us', NewSecType = '" & SecType.SelectedValue & "', NewSymbol = '" & NewSymbol.Text & "', [Desc] = '" & lblSecName.Text & "', TrueDate = #" & MaturityDate.Text & "# WHERE ID = " & ID.Text

            Command = New OleDb.OleDbCommand(SQLstr, Mycn)
            Command.ExecuteNonQuery()

            Mycn.Close()

            MsgBox("Record Updated.")
            Me.Close()

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

            SQLstring = "SELECT * FROM mdb_ETFPrice_SecurityData WHERE APXSecurityID = " & ComboBox1.SelectedValue

            Dim queryString As String = String.Format(SQLstring)
            Dim cmd As New OleDb.OleDbCommand(queryString, Mycn)
            Dim da As New OleDb.OleDbDataAdapter(cmd)
            Dim ds As New DataSet

            da.Fill(ds, "User")
            Dim dt As DataTable = ds.Tables("User")
            If dt.Rows.Count = 0 Then
                Call SaveNew()
            Else
                MsgBox("A record already exsists for that Security.", MsgBoxStyle.Information, "DUPLICATE RECORD DETECTED")
            End If

        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        Call AddDates()

    End Sub

    Public Sub AddDates()
        Dim daysadd As Integer
        daysadd = AvgLife.Text
        'Label15.Text = DateAdd(DateInterval.Year, daysadd, Now())
        Dim figure1 As Double
        Dim figure2 As Double
        Dim figure3 As Double
        Dim figure4 As Double
        Dim figure5 As Double
        Dim figure6 As Double
        Dim figure7 As Double

        figure1 = Math.Floor(daysadd)
        figure2 = AvgLife.Text - figure1
        figure3 = 12 * figure2
        figure4 = Math.Floor(figure3)
        figure5 = figure3 - figure4
        figure6 = 30 * figure5
        figure7 = Math.Floor(figure6)

        Dim dte As Date

        TextBox1.Text = figure1
        TextBox2.Text = figure2
        TextBox3.Text = figure3
        TextBox4.Text = figure4
        TextBox5.Text = figure5
        TextBox6.Text = figure6
        TextBox7.Text = figure7

        dte = DateAdd(DateInterval.Year, figure1, Now())
        dte = DateAdd(DateInterval.Month, figure4, dte)
        dte = DateAdd(DateInterval.Day, figure7, dte)

        MaturityDate.Text = dte
        Label15.Text = dte
    End Sub

    Private Sub AvgLife_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles AvgLife.LostFocus
        Call AddDates()
    End Sub

    Private Sub AvgLife_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles AvgLife.TextChanged

    End Sub
End Class