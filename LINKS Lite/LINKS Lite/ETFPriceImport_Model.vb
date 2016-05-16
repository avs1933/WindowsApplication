Public Class ETFPriceImport_Model

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Dim child As New ETF_PriceImport_APXPortfolios
        child.MdiParent = Home
        child.Show()
    End Sub

    Private Sub ETFPriceImport_Model_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Call LoadPortfolioCodes1()
        Call LoadPortfolioCodes2()
        Call LoadAssetClass()
        Call LoadAssignedAssetClasses()
    End Sub

    Public Sub LoadPortfolioCodes1()
        Try

            Dim conn As New OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & My.Settings.dbloc & "';Jet OLEDB:Database Password='" & My.Settings.dbpass & "';")

            Dim strSQL As String = "SELECT mdb_ETFPrice_APXPortfolios.PortfolioID, AdvApp_vPortfolio.PortfolioCode" & _
            " FROM AdvApp_vPortfolio INNER JOIN mdb_ETFPrice_APXPortfolios ON AdvApp_vPortfolio.PortfolioID = mdb_ETFPrice_APXPortfolios.PortfolioID" & _
            " WHERE(((mdb_ETFPrice_APXPortfolios.Active) = True))" & _
            " GROUP BY mdb_ETFPrice_APXPortfolios.PortfolioID, AdvApp_vPortfolio.PortfolioCode;"

            Dim da As New OleDb.OleDbDataAdapter(strSQL, conn)
            Dim ds As New DataSet
            da.Fill(ds, "Users")

            With ETFPortfolioID
                .DataSource = ds.Tables("Users")
                .DisplayMember = "PortfolioCode"
                .ValueMember = "PortfolioID"
                .SelectedIndex = 0
            End With

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "ERROR")
        End Try
    End Sub

    Public Sub LoadPortfolioCodes2()
        Try

            Dim conn As New OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & My.Settings.dbloc & "';Jet OLEDB:Database Password='" & My.Settings.dbpass & "';")

            Dim strSQL As String = "SELECT mdb_ETFPrice_APXPortfolios.PortfolioID, AdvApp_vPortfolio.PortfolioCode" & _
            " FROM AdvApp_vPortfolio INNER JOIN mdb_ETFPrice_APXPortfolios ON AdvApp_vPortfolio.PortfolioID = mdb_ETFPrice_APXPortfolios.PortfolioID" & _
            " WHERE(((mdb_ETFPrice_APXPortfolios.Active) = True))" & _
            " GROUP BY mdb_ETFPrice_APXPortfolios.PortfolioID, AdvApp_vPortfolio.PortfolioCode;"

            Dim da As New OleDb.OleDbDataAdapter(strSQL, conn)
            Dim ds As New DataSet
            da.Fill(ds, "Users")

            With TAMPortfolioID
                .DataSource = ds.Tables("Users")
                .DisplayMember = "PortfolioCode"
                .ValueMember = "PortfolioID"
                .SelectedIndex = 0
            End With

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "ERROR")
        End Try
    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        If ID.Text = "NEW" Then
            Call SaveNew()
        Else
            Call SaveOld()
        End If
    End Sub

    Public Sub SaveNew()
        Dim Mycn As OleDb.OleDbConnection
        Dim Command As OleDb.OleDbCommand
        Dim SQLstr As String

        Try
            Mycn = New OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & My.Settings.dbloc & "';Jet OLEDB:Database Password='" & My.Settings.dbpass & "';")
            Mycn.Open()

            SQLstr = "INSERT INTO mdb_ETFPrice_ModelMain(ModelName, CashBal, ETFPortfolioID, TAMPortfolioID, InceptionDate, Active)" & _
            "VALUES('" & ModelName.Text & "', " & Cash.Text & ", " & ETFPortfolioID.SelectedValue & "," & TAMPortfolioID.SelectedValue & ", #" & InceptionDate.Text & "#, -1)"

            Command = New OleDb.OleDbCommand(SQLstr, Mycn)
            Command.ExecuteNonQuery()

            Mycn.Close()

            Call PullID()
            Call LoadAssetClass()
            Call LoadManualAddSymbol()

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "ERROR")
        End Try
    End Sub

    Public Sub SaveOld()
        Dim Mycn As OleDb.OleDbConnection
        Dim Command As OleDb.OleDbCommand
        Dim SQLstr As String

        Try
            Mycn = New OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & My.Settings.dbloc & "';Jet OLEDB:Database Password='" & My.Settings.dbpass & "';")
            Mycn.Open()

            SQLstr = "Update mdb_ETFPrice_ModelMain SET ModelName = '" & ModelName.Text & "', CashBal = " & Cash.Text & ", ETFPortfolioID = " & ETFPortfolioID.SelectedValue & ", TAMPortfolioID = " & TAMPortfolioID.SelectedValue & ", InceptionDate = #" & InceptionDate.Text & "# WHERE ID = " & ID.Text

            Command = New OleDb.OleDbCommand(SQLstr, Mycn)
            Command.ExecuteNonQuery()

            Mycn.Close()

            Call LoadAssetClass()
            MsgBox("Record Updated.")

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "ERROR")
        End Try
    End Sub

    Public Sub PullID()
        Try
            Dim conn As New OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & My.Settings.dbloc & "';Jet OLEDB:Database Password='" & My.Settings.dbpass & "';")

            Dim strSQL As String = "SELECT Top 1 ID FROM mdb_ETFPrice_ModelMain WHERE ModelName = '" & ModelName.Text & "' AND ETFPortfolioID = " & ETFPortfolioID.SelectedValue & " AND TAMPortfolioID = " & TAMPortfolioID.SelectedValue & " AND Active = -1 ORDER BY ID DESC"
            Dim queryString As String = String.Format(strSQL)
            Dim cmd As New OleDb.OleDbCommand(queryString, conn)
            Dim da As New OleDb.OleDbDataAdapter(cmd)
            Dim ds As New DataSet

            da.Fill(ds, "User")
            Dim dt As DataTable = ds.Tables("User")

            Dim row1 As DataRow = dt.Rows(0)

            If IsDBNull(row1("ID")) Then
                ID.Text = "NEW"
            Else
                ID.Text = (row1("ID"))
            End If

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "ERROR")
        End Try
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        If ID.Text = "NEW" Then
            Dim ir As MsgBoxResult
            ir = MsgBox("You must save the record before adding Asset Classes to the model.  Would you like to save now?", MsgBoxStyle.YesNo + MsgBoxStyle.Information, "Save Record?")
            If ir = MsgBoxResult.Yes Then

                If IsDBNull(AssetClass.SelectedValue) Then
                    MsgBox("No Asset Class selected.", MsgBoxStyle.Critical, "Cant Save")
                Else
                    Call SaveNew()
                    Call MoveAssetClass()
                End If

            Else
                'do nothing
            End If
        Else
            Call MoveAssetClass()
        End If
    End Sub

    Public Sub LoadAssetClass()
        If ID.Text = "NEW" Then
        Else

            Try

                Dim conn As New OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & My.Settings.dbloc & "';Jet OLEDB:Database Password='" & My.Settings.dbpass & "';")

                Dim sqlstring As String

                sqlstring = "SELECT dbo_AdvSecurity.[LongAssetClassCode], AdvApp_vAssetClass.[AssetClassName]" & _
                    " FROM (dbo_AdvSecurity INNER JOIN mdb_ETFPricing_ApprovedList ON dbo_AdvSecurity.[Symbol] = mdb_ETFPricing_ApprovedList.[APXSymbol]) INNER JOIN AdvApp_vAssetClass ON dbo_AdvSecurity.[LongAssetClassCode] = AdvApp_vAssetClass.[AssetClassCode]" & _
                    " WHERE(((mdb_ETFPricing_ApprovedList.[Current]) = True) AND dbo_AdvSecurity.[LongAssetClassCode] Not In (SELECT [AssetClassCode] FROM mdb_ETFPrice_ModelAssetClass WHERE [ModelID] = " & ID.Text & " AND [Active] = True))" & _
                    " GROUP BY dbo_AdvSecurity.[LongAssetClassCode], AdvApp_vAssetClass.[AssetClassName]"

                Dim queryString As String = String.Format(sqlstring)
                Dim cmd As New OleDb.OleDbCommand(queryString, conn)
                Dim da1 As New OleDb.OleDbDataAdapter(cmd)
                Dim ds1 As New DataSet

                da1.Fill(ds1, "User")
                Dim dt1 As DataTable = ds1.Tables("User")
                If dt1.Rows.Count > 0 Then

                    Dim strSQL As String = "SELECT dbo_AdvSecurity.[LongAssetClassCode], AdvApp_vAssetClass.[AssetClassName]" & _
                    " FROM (dbo_AdvSecurity INNER JOIN mdb_ETFPricing_ApprovedList ON dbo_AdvSecurity.[Symbol] = mdb_ETFPricing_ApprovedList.[APXSymbol]) INNER JOIN AdvApp_vAssetClass ON dbo_AdvSecurity.[LongAssetClassCode] = AdvApp_vAssetClass.[AssetClassCode]" & _
                    " WHERE(((mdb_ETFPricing_ApprovedList.[Current]) = True) AND dbo_AdvSecurity.[LongAssetClassCode] Not In (SELECT [AssetClassCode] FROM mdb_ETFPrice_ModelAssetClass WHERE [ModelID] = " & ID.Text & " AND [Active] = True))" & _
                    " GROUP BY dbo_AdvSecurity.[LongAssetClassCode], AdvApp_vAssetClass.[AssetClassName]"

                    Dim da As New OleDb.OleDbDataAdapter(strSQL, conn)
                    Dim ds As New DataSet
                    da.Fill(ds, "Users")

                    With AssetClass
                        .DataSource = ds.Tables("Users")
                        .DisplayMember = "AssetClassName"
                        .ValueMember = "LongAssetClassCode"
                        .SelectedIndex = 0
                    End With

                Else
                    AssetClass.Enabled = False
                    Button2.Enabled = False
                End If
            Catch ex As Exception
                MsgBox(ex.Message, MsgBoxStyle.Critical, "ERROR")
            End Try
        End If
    End Sub

    Public Sub LoadAssignedAssetClasses()
        If ID.Text = "NEW" Then

        Else

            Try
                Dim conn As New OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & My.Settings.dbloc & "';Jet OLEDB:Database Password='" & My.Settings.dbpass & "';")

                Dim SqlString As String


                SqlString = "SELECT mdb_ETFPrice_ModelAssetClass.[AssetClassCode], AdvApp_vAssetClass.AssetClassName As [Asset Class]" & _
                " FROM AdvApp_vAssetClass INNER JOIN mdb_ETFPrice_ModelAssetClass ON AdvApp_vAssetClass.[AssetClassCode] = mdb_ETFPrice_ModelAssetClass.[AssetClassCode]" & _
                " WHERE(((mdb_ETFPrice_ModelAssetClass.[Active]) = True) And mdb_ETFPrice_ModelAssetClass.[ModelID] = " & ID.Text & ")" & _
                " GROUP BY mdb_ETFPrice_ModelAssetClass.[AssetClassCode], AdvApp_vAssetClass.[AssetClassName]"

                Dim da As New OleDb.OleDbDataAdapter(SqlString, conn)
                Dim ds As New DataSet
                da.Fill(ds, "Users")

                With DataGridView1
                    .DataSource = ds.Tables("Users")
                End With

            Catch ex As Exception
                MsgBox(ex.Message, MsgBoxStyle.Critical, "ERROR")
            End Try
        End If
    End Sub

    Public Sub MoveAssetClass()
        Dim Mycn As OleDb.OleDbConnection
        Dim Command As OleDb.OleDbCommand
        Dim SQLstr As String

        Try
            Mycn = New OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & My.Settings.dbloc & "';Jet OLEDB:Database Password='" & My.Settings.dbpass & "';")
            Mycn.Open()

            SQLstr = "INSERT INTO mdb_ETFPrice_ModelAssetClass([ModelID], [AssetClassCode], [Active])" & _
            "VALUES(" & ID.Text & ",'" & AssetClass.SelectedValue & "', -1)"

            Command = New OleDb.OleDbCommand(SQLstr, Mycn)
            Command.ExecuteNonQuery()

            Mycn.Close()

            Call LoadAssetClass()
            Call LoadAssignedAssetClasses()

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "ERROR")
        End Try
    End Sub

    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        If ID.Text = "NEW" Then
            MsgBox("You must save this record and add Asset Classes before auto adding securities.", MsgBoxStyle.Exclamation, "Cant Update")
        Else
            Dim ir As MsgBoxResult
            ir = MsgBox("You are about to add all approved securities within the asset classes you selected.  This will refresh all securities inside this model, and any current weightings will be lost.  Are you sure you want to continue?", MsgBoxStyle.YesNo + MsgBoxStyle.Information, "Confirm Selection")
            If ir = MsgBoxResult.Yes Then
                Call InsertNewSecurities()
            Else
                'do nothing
            End If
        End If
    End Sub

    Public Sub InsertNewSecurities()
        Dim Mycn As OleDb.OleDbConnection
        Dim Command As OleDb.OleDbCommand
        Dim SQLstr As String

        Dim Command2 As OleDb.OleDbCommand
        Dim SQLstr2 As String

        Dim Command3 As OleDb.OleDbCommand
        Dim SQLstr3 As String

        Try
            Mycn = New OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & My.Settings.dbloc & "';Jet OLEDB:Database Password='" & My.Settings.dbpass & "';")
            Mycn.Open()

            SQLstr2 = "DELETE * FROM mdb_ETFPricing_ModelHoldings WHERE ModelID = " & ID.Text


            Command2 = New OleDb.OleDbCommand(SQLstr2, Mycn)
            Command2.ExecuteNonQuery()

            SQLstr = "INSERT INTO mdb_ETFPricing_ModelHoldings(ModelID, SecurityID, InterestRate, YTW, Duration, AvgLife, Rating)" & _
            "SELECT " & ID.Text & ", AdvApp_vSecurity.SecurityID, AdvApp_vSecurity.InterestOrDividendRate, AdvApp_vSecurity.YTMOnMarket, AdvApp_vSecurity.DurationToMaturity, AdvApp_vSecurity.AverageLife, AdvApp_vSecurity.SPRating" & _
            " FROM mdb_ETFPricing_ApprovedList INNER JOIN AdvApp_vSecurity ON mdb_ETFPricing_ApprovedList.APXSymbol = AdvApp_vSecurity.Symbol" & _
            " WHERE(((mdb_ETFPricing_ApprovedList.Current) = True) AND (AdvApp_vSecurity.SecurityID Not In (SELECT APXSymbol FROM mdb_SymbolMapping Where Active = True)) And ((AdvApp_vSecurity.LongAssetClassCode) In (SELECT AssetClassCode FROM mdb_ETFPrice_ModelAssetClass WHERE ModelID = " & ID.Text & " AND Active = -1)))" & _
            " GROUP BY AdvApp_vSecurity.SecurityID, AdvApp_vSecurity.InterestOrDividendRate, AdvApp_vSecurity.YTMOnMarket, AdvApp_vSecurity.DurationToMaturity, AdvApp_vSecurity.AverageLife, AdvApp_vSecurity.SPRating;"

            Command = New OleDb.OleDbCommand(SQLstr, Mycn)
            Command.ExecuteNonQuery()

            SQLstr3 = "INSERT INTO mdb_ETFPricing_ModelHoldings(ModelID, SecurityID, InterestRate, YTW, Duration, AvgLife, Rating)" & _
            "SELECT " & ID.Text & ",dbo_AdvSecurity.SecurityID, dbo_AdvSecurity.InterestOrDividendRate, (AdvApp_vSecurity.YTMOnMarket / 100), AdvApp_vSecurity.DurationToMaturity, AdvApp_vSecurity.AverageLife, AdvApp_vSecurity.SPRating" & _
            " FROM (mdb_ETFPricing_ApprovedList INNER JOIN dbo_AdvSecurity ON mdb_ETFPricing_ApprovedList.APXSymbol = dbo_AdvSecurity.Symbol) INNER JOIN (mdb_SymbolMapping INNER JOIN AdvApp_vSecurity ON mdb_SymbolMapping.NewAPXSymbol = AdvApp_vSecurity.SecurityID) ON dbo_AdvSecurity.SecurityID = mdb_SymbolMapping.APXSymbol" & _
            " WHERE (((mdb_ETFPricing_ApprovedList.Current)=True) AND ((dbo_AdvSecurity.SecurityID) IN (SELECT APXSymbol FROM mdb_SymbolMapping WHERE Active = True)) AND ((AdvApp_vSecurity.LongAssetClassCode) In (SELECT AssetClassCode FROM mdb_ETFPrice_ModelAssetClass WHERE ModelID = " & ID.Text & " AND Active = -1)))" & _
            " GROUP BY dbo_AdvSecurity.SecurityID, dbo_AdvSecurity.InterestOrDividendRate, AdvApp_vSecurity.YTMOnMarket, AdvApp_vSecurity.DurationToMaturity, AdvApp_vSecurity.AverageLife, AdvApp_vSecurity.SPRating;"

            Command3 = New OleDb.OleDbCommand(SQLstr3, Mycn)
            Command3.ExecuteNonQuery()

            Mycn.Close()

            Call LoadModelHoldings()

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "ERROR")
        End Try
    End Sub

    Public Sub LoadModelHoldings()
        Call UpdateInceptionDate()
        Call UpdatePrice()
        Call GetAvgHoldingWeight()
        Call UpdatePositionAmtShares()
        Call LoadHoldingsGrid()
        Call AddPerCash()
        Call UpdateAssetClassWeightings()
        Call UpdateTMVAssetClassWeightings()
        Call LoadAssetClassGrid()
        Call LoadCurrentYieldPortfolio()
        Call LoadCurrentYTWPortfolio()
        Call LoadManualAddSymbol()
        'update inception date 1st
        'pull price
        'calculate CY (((Rate*12)/price)*10)
    End Sub

    Public Sub UpdateInceptionDate()
        Try
            Dim conn As New OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & My.Settings.dbloc & "';Jet OLEDB:Database Password='" & My.Settings.dbpass & "';")

            Dim strSQL As String = "SELECT Top 1 mdb_ETFPricing_ApprovedList.Inception" & _
            " FROM mdb_ETFPricing_ApprovedList INNER JOIN (mdb_ETFPricing_ModelHoldings INNER JOIN dbo_AdvSecurity ON mdb_ETFPricing_ModelHoldings.SecurityID = dbo_AdvSecurity.SecurityID) ON mdb_ETFPricing_ApprovedList.APXSymbol = dbo_AdvSecurity.Symbol" & _
            " WHERE(((mdb_ETFPricing_ModelHoldings.ModelID) = " & ID.Text & "))" & _
            " GROUP BY mdb_ETFPricing_ApprovedList.Inception" & _
            " ORDER BY mdb_ETFPricing_ApprovedList.Inception DESC;"

            Dim queryString As String = String.Format(strSQL)
            Dim cmd As New OleDb.OleDbCommand(queryString, conn)
            Dim da As New OleDb.OleDbDataAdapter(cmd)
            Dim ds As New DataSet

            da.Fill(ds, "User")
            Dim dt As DataTable = ds.Tables("User")

            Dim row1 As DataRow = dt.Rows(0)

            Dim dte1 As Date
            dte1 = (row1("Inception"))

            InceptionDate.Text = dte1

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "ERROR")
        End Try
    End Sub

    Public Sub UpdatePrice()
        Dim Mycn As OleDb.OleDbConnection
        Dim Command As OleDb.OleDbCommand
        Dim SQLstr As String

        Dim dte1 As Date
        dte1 = InceptionDate.Text

        Dim dte2 As String
        dte2 = Format(dte1, "yyyy-MM-dd").ToString


        Try
            Mycn = New OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & My.Settings.dbloc & "';Jet OLEDB:Database Password='" & My.Settings.dbpass & "';")
            Mycn.Open()

            SQLstr = "UPDATE mdb_ETFPricing_ModelHoldings INNER JOIN dbo_AdvPriceHistory ON mdb_ETFPricing_ModelHoldings.SecurityID = dbo_AdvPriceHistory.SecurityID SET mdb_ETFPricing_ModelHoldings.Price = [dbo_AdvPriceHistory].[PriceValue]" & _
            " WHERE (((mdb_ETFPricing_ModelHoldings.[ModelID])=" & ID.Text & ") AND ((dbo_AdvPriceHistory.FromDate)='" & dte2 & "'))"

            Command = New OleDb.OleDbCommand(SQLstr, Mycn)
            Command.ExecuteNonQuery()

            Mycn.Close()


        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "ERROR")
        End Try
    End Sub

    Public Sub GetAvgHoldingWeight()
        Dim Mycn As OleDb.OleDbConnection

        Try

            Mycn = New OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & My.Settings.roledbloc & "';Jet OLEDB:Database Password='" & My.Settings.dbpass & "';")

            Mycn.Open()

            Dim sqlstring As String

            sqlstring = "SELECT * FROM mdb_ETFPricing_ModelHoldings WHERE ModelID = " & ID.Text


            Dim queryString As String = String.Format(sqlstring)
            Dim cmd As New OleDb.OleDbCommand(queryString, Mycn)
            Dim da As New OleDb.OleDbDataAdapter(cmd)
            Dim ds As New DataSet

            da.Fill(ds, "User")
            Dim dt As DataTable = ds.Tables("User")

            Dim cnt As Double
            Dim rcount As Double

            rcount = dt.Rows.Count

            cnt = ((100 / rcount) / 100)

            Dim csh As Double
            csh = Cash.Text

            'Dim posamt As Double
            'posamt = (csh * cnt)

            Dim Command As OleDb.OleDbCommand
            Dim SQLstr As String

            Mycn = New OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & My.Settings.dbloc & "';Jet OLEDB:Database Password='" & My.Settings.dbpass & "';")
            Mycn.Open()

            SQLstr = "UPDATE mdb_ETFPricing_ModelHoldings SET [Weight] = " & cnt & " WHERE ModelID = " & ID.Text

            Command = New OleDb.OleDbCommand(SQLstr, Mycn)
            Command.ExecuteNonQuery()

            Mycn.Close()

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "ERROR")
        End Try
    End Sub

    Public Sub UpdatePositionAmtShares()
        Dim Mycn As OleDb.OleDbConnection
        Dim Command As OleDb.OleDbCommand
        Dim SQLstr As String

        Try
            Mycn = New OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & My.Settings.dbloc & "';Jet OLEDB:Database Password='" & My.Settings.dbpass & "';")
            Mycn.Open()

            Dim csh As Double
            csh = Cash.Text

            SQLstr = "UPDATE mdb_ETFPricing_ModelHoldings SET PositionAmt = (" & csh & " * mdb_ETFPricing_ModelHoldings.Weight), Shares = ((" & csh & " * mdb_ETFPricing_ModelHoldings.Weight) / mdb_ETFPricing_ModelHoldings.Price), CurrentYield = (((mdb_ETFPricing_ModelHoldings.InterestRate * 12) / mdb_ETFPricing_ModelHoldings.Price)/10) WHERE ModelID = " & ID.Text

            Command = New OleDb.OleDbCommand(SQLstr, Mycn)
            Command.ExecuteNonQuery()

            Mycn.Close()


        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "ERROR")
        End Try
    End Sub

    Public Sub UpdateTMVAssetClassWeightings()
        Dim Mycn As OleDb.OleDbConnection
        Dim Command As OleDb.OleDbCommand
        Dim SQLstr As String

        Dim Command2 As OleDb.OleDbCommand
        Dim SQLstr2 As String

        Try
            Mycn = New OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & My.Settings.dbloc & "';Jet OLEDB:Database Password='" & My.Settings.dbpass & "';")
            Mycn.Open()

            Dim csh As Double
            csh = Cash.Text

            SQLstr2 = "DELETE * FROM mdb_ETFPrice_AssetClassWeighting WHERE ModelID = " & ID.Text

            Command2 = New OleDb.OleDbCommand(SQLstr2, Mycn)
            Command2.ExecuteNonQuery()

            SQLstr = "Insert Into mdb_ETFPrice_AssetClassWeighting(ModelID, AssetClassCode, TMV)" & _
            "SELECT " & ID.Text & ",mdb_ETFPrice_ModelAssetClass.AssetClassCode, Sum(mdb_ETFPricing_ModelHoldings.PositionAmt) AS SumOfPositionAmt" & _
            " FROM (mdb_ETFPricing_ModelHoldings INNER JOIN dbo_AdvSecurity ON mdb_ETFPricing_ModelHoldings.SecurityID = dbo_AdvSecurity.SecurityID) INNER JOIN mdb_ETFPrice_ModelAssetClass ON (dbo_AdvSecurity.LongAssetClassCode = mdb_ETFPrice_ModelAssetClass.AssetClassCode) AND (mdb_ETFPricing_ModelHoldings.ModelID = mdb_ETFPrice_ModelAssetClass.ModelID)" & _
            " WHERE(((mdb_ETFPrice_ModelAssetClass.ModelID) = " & ID.Text & "))" & _
            " GROUP BY mdb_ETFPrice_ModelAssetClass.AssetClassCode;"


            Command = New OleDb.OleDbCommand(SQLstr, Mycn)
            Command.ExecuteNonQuery()

            Mycn.Close()


        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "ERROR")
        End Try
    End Sub

    Public Sub Remove0WeightedHoldings()
        Dim Mycn As OleDb.OleDbConnection
        Dim Command As OleDb.OleDbCommand
        Dim SQLstr As String

        Try
            Mycn = New OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & My.Settings.dbloc & "';Jet OLEDB:Database Password='" & My.Settings.dbpass & "';")
            Mycn.Open()


            SQLstr = "DELETE * FROM mdb_ETFPricing_ModelHoldings WHERE Weight = 0 AND ModelID = " & ID.Text


            Command = New OleDb.OleDbCommand(SQLstr, Mycn)
            Command.ExecuteNonQuery()

            Mycn.Close()


        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "ERROR")
        End Try
    End Sub

    Public Sub UpdateAssetClassWeightings()
        Dim Mycn As OleDb.OleDbConnection

        Try

            Mycn = New OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & My.Settings.roledbloc & "';Jet OLEDB:Database Password='" & My.Settings.dbpass & "';")

            Mycn.Open()

            Dim sqlstring As String

            sqlstring = "SELECT * FROM mdb_ETFPrice_ModelAssetClass WHERE ModelID = " & ID.Text


            Dim queryString As String = String.Format(sqlstring)
            Dim cmd As New OleDb.OleDbCommand(queryString, Mycn)
            Dim da As New OleDb.OleDbDataAdapter(cmd)
            Dim ds As New DataSet

            da.Fill(ds, "User")
            Dim dt As DataTable = ds.Tables("User")

            Dim cnt As Double
            Dim rcount As Double

            rcount = dt.Rows.Count

            cnt = ((100 / rcount) / 100)

            Dim csh As Double
            csh = Cash.Text

            Dim Command As OleDb.OleDbCommand
            Dim SQLstr As String

            Mycn = New OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & My.Settings.dbloc & "';Jet OLEDB:Database Password='" & My.Settings.dbpass & "';")
            Mycn.Open()

            SQLstr = "UPDATE mdb_ETFPrice_ModelAssetClass SET [TargetWeight] = " & cnt & " WHERE ModelID = " & ID.Text

            Command = New OleDb.OleDbCommand(SQLstr, Mycn)
            Command.ExecuteNonQuery()

            Mycn.Close()

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "ERROR")
        End Try
    End Sub

    Public Sub LoadHoldingsGrid()
        Try
            Dim conn As New OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & My.Settings.dbloc & "';Jet OLEDB:Database Password='" & My.Settings.dbpass & "';")

            Dim SqlString As String


            SqlString = "SELECT mdb_ETFPricing_ModelHoldings.ID, AdvApp_vAssetClass.AssetClassName As [Asset Class], dbo_AdvSecurity.Symbol, dbo_AdvSecurity.FullName As [Security Desc], Format((mdb_ETFPricing_ModelHoldings.Weight * 100), '#.##') As [Weighting], mdb_ETFPricing_ModelHoldings.Price, mdb_ETFPricing_ModelHoldings.PositionAmt, mdb_ETFPricing_ModelHoldings.Shares, Format((mdb_ETFPricing_ModelHoldings.CurrentYield * 100), '#.##') As [CY], mdb_ETFPricing_ModelHoldings.YTW, mdb_ETFPricing_ModelHoldings.Duration, mdb_ETFPricing_ModelHoldings.AvgLife, mdb_ETFPricing_ModelHoldings.Rating" & _
            " FROM AdvApp_vAssetClass INNER JOIN (dbo_AdvSecurity INNER JOIN mdb_ETFPricing_ModelHoldings ON dbo_AdvSecurity.SecurityID = mdb_ETFPricing_ModelHoldings.SecurityID) ON AdvApp_vAssetClass.AssetClassCode = dbo_AdvSecurity.LongAssetClassCode" & _
            " WHERE(((mdb_ETFPricing_ModelHoldings.ModelID) = " & ID.Text & "))" & _
            " GROUP BY mdb_ETFPricing_ModelHoldings.ID, AdvApp_vAssetClass.AssetClassName, dbo_AdvSecurity.Symbol, dbo_AdvSecurity.FullName, mdb_ETFPricing_ModelHoldings.Weight, mdb_ETFPricing_ModelHoldings.Price, mdb_ETFPricing_ModelHoldings.PositionAmt, mdb_ETFPricing_ModelHoldings.Shares, mdb_ETFPricing_ModelHoldings.CurrentYield, mdb_ETFPricing_ModelHoldings.YTW, mdb_ETFPricing_ModelHoldings.Duration, mdb_ETFPricing_ModelHoldings.AvgLife, mdb_ETFPricing_ModelHoldings.Rating;"


            Dim da As New OleDb.OleDbDataAdapter(SqlString, conn)
            Dim ds As New DataSet
            da.Fill(ds, "Users")

            With DataGridView2
                .DataSource = ds.Tables("Users")
                .Columns(0).Visible = False
                .Columns(4).DefaultCellStyle.Format = "p"
                .Columns(5).DefaultCellStyle.Format = "c"
                .Columns(8).DefaultCellStyle.Format = "p"
                .Columns(6).DefaultCellStyle.Format = "c"
                .Columns(9).DefaultCellStyle.Format = "p"
            End With

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "ERROR")
        End Try
    End Sub

    Public Sub LoadAssetClassGrid()
        Try
            Dim conn As New OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & My.Settings.dbloc & "';Jet OLEDB:Database Password='" & My.Settings.dbpass & "';")

            Dim SqlString As String


            SqlString = "SELECT mdb_ETFPrice_ModelAssetClass.ID, AdvApp_vAssetClass.AssetClassName, mdb_ETFPrice_ModelAssetClass.TargetWeight As [Target Weight], (mdb_ETFPrice_AssetClassWeighting.TMV/" & Cash.Text & ") As [Current Weight]" & _
            " FROM (mdb_ETFPrice_ModelAssetClass INNER JOIN mdb_ETFPrice_AssetClassWeighting ON (mdb_ETFPrice_ModelAssetClass.ModelID = mdb_ETFPrice_AssetClassWeighting.ModelID) AND (mdb_ETFPrice_ModelAssetClass.AssetClassCode = mdb_ETFPrice_AssetClassWeighting.AssetClassCode)) INNER JOIN AdvApp_vAssetClass ON mdb_ETFPrice_AssetClassWeighting.AssetClassCode = AdvApp_vAssetClass.AssetClassCode" & _
            " WHERE mdb_ETFPrice_ModelAssetClass.ModelID = " & ID.Text


            Dim da As New OleDb.OleDbDataAdapter(SqlString, conn)
            Dim ds As New DataSet
            da.Fill(ds, "Users")

            With DataGridView3
                .DataSource = ds.Tables("Users")
                .Columns(0).Visible = False
                .Columns(2).DefaultCellStyle.Format = "p"
                .Columns(3).DefaultCellStyle.Format = "p"
            End With

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "ERROR")
        End Try
    End Sub

    Private Sub DataGridView2_CellContentDoubleClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView2.CellContentDoubleClick
        Dim child As New mdb_ETF_Pricing_ChangeHoldingWeight
        child.MdiParent = Home
        child.Show()
        child.ID.Text = DataGridView2.SelectedCells(0).Value
        child.TextBox1.Text = DataGridView2.SelectedCells(4).Value
    End Sub

    Private Sub Button5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button5.Click
        Call LoadAfterInitialUpdate()
    End Sub

    Public Sub LoadAfterInitialUpdate()
        Call Remove0WeightedHoldings()
        Call UpdateInceptionDate()
        Call UpdatePrice()
        Call UpdatePositionAmtShares()
        Call LoadHoldingsGrid()
        Call AddPerCash()
        Call UpdateTMVAssetClassWeightings()
        Call LoadAssetClassGrid()
        Call LoadCurrentYieldPortfolio()
        Call LoadCurrentYTWPortfolio()
        Call LoadManualAddSymbol()
    End Sub

    Public Sub AddPerCash()
        Try
            Dim conn As New OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & My.Settings.dbloc & "';Jet OLEDB:Database Password='" & My.Settings.dbpass & "';")

            Dim strSQL As String = "SELECT Sum(mdb_ETFPricing_ModelHoldings.Weight) AS SumOfWeight, Sum(mdb_ETFPricing_ModelHoldings.PositionAmt) AS SumOfPositionAmt" & _
            " FROM mdb_ETFPricing_ModelHoldings" & _
            " WHERE ModelID = " & ID.Text

            Dim queryString As String = String.Format(strSQL)
            Dim cmd As New OleDb.OleDbCommand(queryString, conn)
            Dim da As New OleDb.OleDbDataAdapter(cmd)
            Dim ds As New DataSet

            da.Fill(ds, "User")
            Dim dt As DataTable = ds.Tables("User")

            Dim row1 As DataRow = dt.Rows(0)

            TotalMoney.Text = Format((row1("SumOfPositionAmt")), "$#,###.##")
            TotalWeighting.Text = Format((row1("SumOfWeight")), "#.##%")

            If TotalWeighting.Text <> "100%" Then
                TotalWeighting.ForeColor = Color.Red
            Else
                TotalWeighting.ForeColor = Color.Black
            End If

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "ERROR")
        End Try
    End Sub

    Public Sub LoadCurrentYieldPortfolio()
        Try
            Dim conn As New OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & My.Settings.dbloc & "';Jet OLEDB:Database Password='" & My.Settings.dbpass & "';")

            Dim strSQL As String = "SELECT ModelID, (SUM(CurrentYield * Weight) / Sum(Weight)) As [CurrentYield1]" & _
            " FROM(mdb_ETFPricing_ModelHoldings)" & _
            " WHERE ModelID = " & ID.Text & _
            " GROUP BY ModelID"

            Dim queryString As String = String.Format(strSQL)
            Dim cmd As New OleDb.OleDbCommand(queryString, conn)
            Dim da As New OleDb.OleDbDataAdapter(cmd)
            Dim ds As New DataSet

            da.Fill(ds, "User")
            Dim dt As DataTable = ds.Tables("User")

            Dim row1 As DataRow = dt.Rows(0)

            lblCurrentYield.Text = Format((row1("CurrentYield1")), "#.##%")

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "ERROR")
        End Try
    End Sub

    Public Sub LoadCurrentYTWPortfolio()
        Try
            Dim conn As New OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & My.Settings.dbloc & "';Jet OLEDB:Database Password='" & My.Settings.dbpass & "';")

            Dim strSQL As String = "SELECT ModelID, (SUM(YTW * Weight) / Sum(Weight)) As [YTW1], (SUM(Duration * Weight) / Sum(Weight)) As [Duration1], (SUM(AvgLife * Weight) / Sum(Weight)) As [AvgLife1]" & _
            " FROM(mdb_ETFPricing_ModelHoldings)" & _
            " WHERE YTW Is Not NULL AND ModelID = " & ID.Text & _
            " GROUP BY ModelID"

            Dim queryString As String = String.Format(strSQL)
            Dim cmd As New OleDb.OleDbCommand(queryString, conn)
            Dim da As New OleDb.OleDbDataAdapter(cmd)
            Dim ds As New DataSet

            da.Fill(ds, "User")
            Dim dt As DataTable = ds.Tables("User")

            Dim row1 As DataRow = dt.Rows(0)

            lblYTW.Text = Format((row1("YTW1")), "#.##%")
            lblDuration.Text = Format((row1("Duration1")), "#.##")
            lblAvgLife.Text = Format((row1("AvgLife1")), "#.##")

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "ERROR")
        End Try
    End Sub

    Public Sub LoadManualAddSymbol()
        If ID.Text = "NEW" Then

        Else

            Try

                Dim conn As New OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & My.Settings.dbloc & "';Jet OLEDB:Database Password='" & My.Settings.dbpass & "';")

                Dim sqlstring As String

                SqlString = "SELECT dbo_AdvSecurity.SecurityID, mdb_ETFPricing_ApprovedList.APXSymbol" & _
                " FROM mdb_ETFPricing_ApprovedList INNER JOIN dbo_AdvSecurity ON mdb_ETFPricing_ApprovedList.APXSymbol = dbo_AdvSecurity.Symbol" & _
                " WHERE dbo_AdvSecurity.SecurityID Not In (SELECT SecurityID FROM mdb_ETFPricing_ModelHoldings WHERE ModelID = " & ID.Text & ")" & _
                " ORDER BY mdb_ETFPricing_ApprovedList.APXSymbol;"

                Dim queryString As String = String.Format(sqlstring)
                Dim cmd As New OleDb.OleDbCommand(queryString, conn)
                Dim da1 As New OleDb.OleDbDataAdapter(cmd)
                Dim ds1 As New DataSet

                da1.Fill(ds1, "User")
                Dim dt1 As DataTable = ds1.Tables("User")
                If dt1.Rows.Count > 0 Then

                    Dim strSQL As String = "SELECT dbo_AdvSecurity.SecurityID, mdb_ETFPricing_ApprovedList.APXSymbol" & _
                    " FROM mdb_ETFPricing_ApprovedList INNER JOIN dbo_AdvSecurity ON mdb_ETFPricing_ApprovedList.APXSymbol = dbo_AdvSecurity.Symbol" & _
                    " WHERE dbo_AdvSecurity.SecurityID Not In (SELECT SecurityID FROM mdb_ETFPricing_ModelHoldings WHERE ModelID = " & ID.Text & ")" & _
                    " ORDER BY mdb_ETFPricing_ApprovedList.APXSymbol;"

                    Dim da As New OleDb.OleDbDataAdapter(strSQL, conn)
                    Dim ds As New DataSet
                    da.Fill(ds, "Users")

                    lblSecName.Text = "Select a security to add."
                    cboAddSecurity.Enabled = True
                    txtWeight.Enabled = True
                    Button6.Enabled = True

                    With cboAddSecurity
                        .DataSource = ds.Tables("Users")
                        .DisplayMember = "APXSymbol"
                        .ValueMember = "SecurityID"
                        .SelectedIndex = 0
                    End With

                Else
                    cboAddSecurity.Enabled = False
                    txtWeight.Enabled = False
                    Button6.Enabled = False
                    lblSecName.Text = "No securities to add."
                End If

            Catch ex As Exception
                MsgBox(ex.Message, MsgBoxStyle.Critical, "ERROR")
            End Try
        End If
    End Sub

    Public Sub InsertAddSecurities()
        Dim Mycn As OleDb.OleDbConnection
        Dim Command As OleDb.OleDbCommand
        Dim SQLstr As String

        Dim Command3 As OleDb.OleDbCommand
        Dim SQLstr3 As String

        Try
            Mycn = New OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & My.Settings.dbloc & "';Jet OLEDB:Database Password='" & My.Settings.dbpass & "';")
            Mycn.Open()

            Dim wgt As Double
            wgt = (txtWeight.Text / 100)

            SQLstr = "INSERT INTO mdb_ETFPricing_ModelHoldings(ModelID, SecurityID, Weight, InterestRate, YTW, Duration, AvgLife, Rating)" & _
            "SELECT " & ID.Text & ", AdvApp_vSecurity.SecurityID, " & wgt & ", AdvApp_vSecurity.InterestOrDividendRate, AdvApp_vSecurity.YTMOnMarket, AdvApp_vSecurity.DurationToMaturity, AdvApp_vSecurity.AverageLife, AdvApp_vSecurity.SPRating" & _
            " FROM mdb_ETFPricing_ApprovedList INNER JOIN AdvApp_vSecurity ON mdb_ETFPricing_ApprovedList.APXSymbol = AdvApp_vSecurity.Symbol" & _
            " WHERE(((mdb_ETFPricing_ApprovedList.Current) = True) AND AdvApp_vSecurity.SecurityID = " & cboAddSecurity.SelectedValue & " AND (AdvApp_vSecurity.SecurityID Not In (SELECT APXSymbol FROM mdb_SymbolMapping Where Active = True)) And ((AdvApp_vSecurity.LongAssetClassCode) In (SELECT AssetClassCode FROM mdb_ETFPrice_ModelAssetClass WHERE ModelID = " & ID.Text & " AND Active = -1)))" & _
            " GROUP BY AdvApp_vSecurity.SecurityID, AdvApp_vSecurity.InterestOrDividendRate, AdvApp_vSecurity.YTMOnMarket, AdvApp_vSecurity.DurationToMaturity, AdvApp_vSecurity.AverageLife, AdvApp_vSecurity.SPRating;"

            Command = New OleDb.OleDbCommand(SQLstr, Mycn)
            Command.ExecuteNonQuery()

            SQLstr3 = "INSERT INTO mdb_ETFPricing_ModelHoldings(ModelID, SecurityID, Weight, InterestRate, YTW, Duration, AvgLife, Rating)" & _
            "SELECT " & ID.Text & ",dbo_AdvSecurity.SecurityID, " & wgt & ", dbo_AdvSecurity.InterestOrDividendRate, (AdvApp_vSecurity.YTMOnMarket / 100), AdvApp_vSecurity.DurationToMaturity, AdvApp_vSecurity.AverageLife, AdvApp_vSecurity.SPRating" & _
            " FROM (mdb_ETFPricing_ApprovedList INNER JOIN dbo_AdvSecurity ON mdb_ETFPricing_ApprovedList.APXSymbol = dbo_AdvSecurity.Symbol) INNER JOIN (mdb_SymbolMapping INNER JOIN AdvApp_vSecurity ON mdb_SymbolMapping.NewAPXSymbol = AdvApp_vSecurity.SecurityID) ON dbo_AdvSecurity.SecurityID = mdb_SymbolMapping.APXSymbol" & _
            " WHERE (((mdb_ETFPricing_ApprovedList.Current)=True) AND dbo_AdvSecurity.SecurityID = " & cboAddSecurity.SelectedValue & " AND ((AdvApp_vSecurity.LongAssetClassCode) In (SELECT AssetClassCode FROM mdb_ETFPrice_ModelAssetClass WHERE ModelID = " & ID.Text & " AND Active = -1)))" & _
            " GROUP BY dbo_AdvSecurity.SecurityID, dbo_AdvSecurity.InterestOrDividendRate, AdvApp_vSecurity.YTMOnMarket, AdvApp_vSecurity.DurationToMaturity, AdvApp_vSecurity.AverageLife, AdvApp_vSecurity.SPRating;"

            Command3 = New OleDb.OleDbCommand(SQLstr3, Mycn)
            Command3.ExecuteNonQuery()

            Mycn.Close()

            Call LoadAfterInitialUpdate()

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "ERROR")
        End Try
    End Sub

    Private Sub Button6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button6.Click
        If ID.Text = "NEW" Then
        Else
            Dim wgt As Double
            wgt = txtWeight.Text
            If wgt > 100 Then
                MsgBox("Weighting must be 100% or less", MsgBoxStyle.Critical, "Invalid Weight")
            Else
                Call InsertAddSecurities()
            End If

        End If
    End Sub

    Private Sub cboAddSecurity_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles cboAddSecurity.LostFocus
        Try
            Dim conn As New OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & My.Settings.dbloc & "';Jet OLEDB:Database Password='" & My.Settings.dbpass & "';")

            Dim strSQL As String = "SELECT FullName FROM dbo_AdvSecurity WHERE SecurityID = " & cboAddSecurity.SelectedValue

            Dim queryString As String = String.Format(strSQL)
            Dim cmd As New OleDb.OleDbCommand(queryString, conn)
            Dim da As New OleDb.OleDbDataAdapter(cmd)
            Dim ds As New DataSet

            da.Fill(ds, "User")
            Dim dt As DataTable = ds.Tables("User")

            Dim row1 As DataRow = dt.Rows(0)

            lblSecName.Text = (row1("FullName"))

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "ERROR")
        End Try
    End Sub

    Private Sub Button7_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button7.Click
        Call LoadModelHoldings()
    End Sub

    Private Sub DataGridView3_CellContentDoubleClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView3.CellContentDoubleClick
        Dim child As New ETF_PriceImport_AssetClassWeighting
        child.MdiParent = Home
        child.Show()
        child.ID.Text = DataGridView3.SelectedCells(0).Value
        child.TextBox1.Text = DataGridView3.SelectedCells(2).Value
    End Sub

    Public Sub LoadPerfDates()
        Dim dte As Date 'StartDate
        Dim dte1 As String 'Start Date Text
        Dim dte2 As String 'Date 3 to end of month - find week day
        'Dim dte3 As Date 'End of Month Date
        Dim dte4 As Date 'End of Month Date
        Dim dte5 As String 'End of date to Text
        Dim stdate As Date
        Dim stdatetxt As String
        Dim Mycn As OleDb.OleDbConnection
        Dim Command As OleDb.OleDbCommand
        Dim SQLstr As String

        Try
            Mycn = New OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & My.Settings.dbloc & "';Jet OLEDB:Database Password='" & My.Settings.dbpass & "';")
            Mycn.Open()

            '/////////////////////////////////PUT INITIAL DATE IN DB\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
            dte = InceptionDate.Text
            dte1 = Format(dte, "yyyy-MM-dd")

            dte4 = DateSerial(Year(dte), Month(dte) + 1, 1 - 1)
            dte2 = Format(dte4, "dddd")
            '/////////////////////////////////FIND LAST WEEKDAY OF MONTH\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
            Do Until dte2 = "Monday" Or dte2 = "Tuesday" Or dte2 = "Wednesday" Or dte2 = "Thursday" Or dte2 = "Friday"
                dte4 = DateAdd("d", -1, dte4)
                dte2 = Format(dte4, "dddd")
            Loop

            dte5 = Format(dte4, "yyyy-MM-dd")
            '/////////////////////////////////INSERT VALUES INTO DB\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
            SQLstr = "INSERT INTO mdb_ETFPrice_ModelPerfDates(ModelID, Date1, Date1Text, Date2, Date2Text)" & _
            " VALUES(" & ID.Text & ",#" & dte & "#, '" & dte1 & "',#" & dte4 & "#, '" & dte5 & "')"

            stdate = dte4
            stdatetxt = dte5

            Command = New OleDb.OleDbCommand(SQLstr, Mycn)
            Command.ExecuteNonQuery()

            '/////////////////////////////////START FINDING EVERY MONTH VALUE\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
            Do
                dte = DateAdd(DateInterval.Month, 1, dte4)
                dte1 = Format(dte, "yyyy-MM-dd")

                dte4 = DateSerial(Year(dte), Month(dte) + 1, 1 - 1)
                dte2 = Format(dte4, "dddd")
                '/////////////////////////////////FIND LAST WEEKDAY OF MONTH\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                Do Until dte2 = "Monday" Or dte2 = "Tuesday" Or dte2 = "Wednesday" Or dte2 = "Thursday" Or dte2 = "Friday"
                    dte4 = DateAdd("d", -1, dte4)
                    dte2 = Format(dte4, "dddd")
                Loop

                dte5 = Format(dte4, "yyyy-MM-dd")
                '/////////////////////////////////INSERT VALUES INTO DB\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                SQLstr = "INSERT INTO mdb_ETFPrice_ModelPerfDates(ModelID, Date1, Date1Text, Date2, Date2Text)" & _
                " VALUES(" & ID.Text & ",#" & stdate & "#, '" & stdatetxt & "',#" & dte4 & "#, '" & dte5 & "')"

                stdate = dte4
                stdatetxt = dte5

                Command = New OleDb.OleDbCommand(SQLstr, Mycn)
                Command.ExecuteNonQuery()

            Loop Until dte4 = DateSerial(Year(Now()), Month(Now()), 1 - 1)

                Mycn.Close()


        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "ERROR")
        End Try
    End Sub

    Private Sub Button8_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button8.Click
        Call LoadPerfDates()
    End Sub

    Private Sub txtWeight_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtWeight.KeyPress
        If e.KeyChar = Microsoft.VisualBasic.ChrW(Keys.Return) Then
            If ID.Text = "NEW" Then
            Else
                Dim wgt As Double
                wgt = txtWeight.Text
                If wgt > 100 Then
                    MsgBox("Weighting must be 100% or less", MsgBoxStyle.Critical, "Invalid Weight")
                Else
                    Call InsertAddSecurities()
                End If

            End If
            e.Handled = True
        End If
    End Sub

    Private Sub txtWeight_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtWeight.TextChanged

    End Sub

    Private Sub Button9_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button9.Click
        If ID.Text = "NEW" Then
            Dim ir As MsgBoxResult
            ir = MsgBox("You must save the record before adding Asset Classes to the model.  Would you like to save now?", MsgBoxStyle.YesNo + MsgBoxStyle.Information, "Save Record?")
            If ir = MsgBoxResult.Yes Then
                Call SaveNew()
                Call MoveAllAssetClasses()
            Else

            End If

        Else
            Call MoveAllAssetClasses()
        End If
    End Sub

    Public Sub MoveAllAssetClasses()

        Dim Mycn As OleDb.OleDbConnection
        Dim Command As OleDb.OleDbCommand
        Dim SQLstr As String

        Try
            Mycn = New OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & My.Settings.dbloc & "';Jet OLEDB:Database Password='" & My.Settings.dbpass & "';")
            Mycn.Open()

            SQLstr = "INSERT INTO mdb_ETFPrice_ModelAssetClass([ModelID], [AssetClassCode], [Active])" & _
            "SELECT " & ID.Text & ",dbo_AdvSecurity.[LongAssetClassCode], -1" & _
            " FROM (dbo_AdvSecurity INNER JOIN mdb_ETFPricing_ApprovedList ON dbo_AdvSecurity.[Symbol] = mdb_ETFPricing_ApprovedList.[APXSymbol]) INNER JOIN AdvApp_vAssetClass ON dbo_AdvSecurity.[LongAssetClassCode] = AdvApp_vAssetClass.[AssetClassCode]" & _
            " WHERE(((mdb_ETFPricing_ApprovedList.[Current]) = True) AND dbo_AdvSecurity.[LongAssetClassCode] Not In (SELECT [AssetClassCode] FROM mdb_ETFPrice_ModelAssetClass WHERE [ModelID] = " & ID.Text & " AND [Active] = True))" & _
            " GROUP BY dbo_AdvSecurity.[LongAssetClassCode], AdvApp_vAssetClass.[AssetClassName]"

            Command = New OleDb.OleDbCommand(SQLstr, Mycn)
            Command.ExecuteNonQuery()

            Mycn.Close()

            Call LoadAssetClass()
            Call LoadAssignedAssetClasses()

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "ERROR")
        End Try

    End Sub

    Private Sub Button10_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button10.Click
        Dim ir As MsgBoxResult
        ir = MsgBox("You are about to copy and create a new model.  Are you sure you want to continue?", MsgBoxStyle.YesNo + MsgBoxStyle.Information, "Save and create new?")
        If ir = MsgBoxResult.Yes Then
            If ID.Text = "NEW" Then
                MsgBox("Cannot create new model.  Current Model has not been saved.", MsgBoxStyle.Critical, "Cant create model")
            Else
                Call SaveOld()
                Call CopyCreate()
            End If
        End If
    End Sub

    Public Sub CopyCreate()
        Dim Mycn As OleDb.OleDbConnection
        Dim Command As OleDb.OleDbCommand
        Dim SQLstr As String

        Dim Command1 As OleDb.OleDbCommand
        Dim SQLstr1 As String

        Dim Command2 As OleDb.OleDbCommand
        Dim SQLstr2 As String

        Dim Command3 As OleDb.OleDbCommand
        Dim SQLstr3 As String

        Try
            Mycn = New OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & My.Settings.dbloc & "';Jet OLEDB:Database Password='" & My.Settings.dbpass & "';")
            Mycn.Open()

            Dim modelname1 As String
            modelname1 = "Copy of " & ModelName.Text

            SQLstr = "INSERT INTO mdb_ETFPrice_ModelMain(ModelName, CashBal, ETFPortfolioID, TAMPortfolioID, InceptionDate, Active)" & _
            "VALUES('" & modelname1 & "', " & Cash.Text & ", " & ETFPortfolioID.SelectedValue & "," & TAMPortfolioID.SelectedValue & ", #" & InceptionDate.Text & "#, -1)"

            Command = New OleDb.OleDbCommand(SQLstr, Mycn)
            Command.ExecuteNonQuery()

            Dim conn As New OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & My.Settings.dbloc & "';Jet OLEDB:Database Password='" & My.Settings.dbpass & "';")

            Dim strSQL As String = "SELECT Top 1 ID FROM mdb_ETFPrice_ModelMain WHERE ModelName = '" & modelname1 & "' AND ETFPortfolioID = " & ETFPortfolioID.SelectedValue & " AND TAMPortfolioID = " & TAMPortfolioID.SelectedValue & " AND Active = -1 ORDER BY ID DESC"
            Dim queryString As String = String.Format(strSQL)
            Dim cmd As New OleDb.OleDbCommand(queryString, conn)
            Dim da As New OleDb.OleDbDataAdapter(cmd)
            Dim ds As New DataSet

            da.Fill(ds, "User")
            Dim dt As DataTable = ds.Tables("User")

            Dim row1 As DataRow = dt.Rows(0)

            Dim nid As Integer
            nid = (row1("ID"))
            Pause(0.05)

            SQLstr1 = "INSERT INTO mdb_ETFPrice_AssetClassWeighting(ModelID, AssetClassCode, TMV)" & _
            "SELECT " & nid & ", AssetClassCode, TMV FROM mdb_ETFPrice_AssetClassWeighting WHERE ModelID = " & ID.Text

            Command1 = New OleDb.OleDbCommand(SQLstr1, Mycn)
            Command1.ExecuteNonQuery()
            'Pause(0.02)

            SQLstr2 = "INSERT INTO mdb_ETFPrice_ModelAssetClass(ModelID, AssetClassCode, Active, TargetWeight)" & _
            "SELECT " & nid & ", AssetClassCode, Active, TargetWeight FROM mdb_ETFPrice_ModelAssetClass WHERE ModelID = " & ID.Text

            Command2 = New OleDb.OleDbCommand(SQLstr2, Mycn)
            Command2.ExecuteNonQuery()
            'Pause(0.02)

            SQLstr3 = "INSERT INTO mdb_ETFPricing_ModelHoldings([ModelID], [SecurityID], [Weight], [Price], [PositionAmt], [Shares], [Updated], [InterestRate], [CurrentYield], [YTW], [Duration], [AvgLife], [Rating])" & _
            "Select " & nid & ", [SecurityID], [Weight], [Price], [PositionAmt], [Shares], [Updated], [InterestRate], [CurrentYield], [YTW], [Duration], [AvgLife], [Rating] FROM mdb_ETFPricing_ModelHoldings WHERE ModelID = " & ID.Text


            Command3 = New OleDb.OleDbCommand(SQLstr3, Mycn)
            Command3.ExecuteNonQuery()
            'Pause(0.02)

            Mycn.Close()

            Dim child As New ETFPriceImport_Model
            child.MdiParent = Home
            child.ID.Text = nid
            child.ModelName.Text = modelname1
            child.ETFPortfolioID.SelectedValue = ETFPortfolioID.SelectedValue
            child.TAMPortfolioID.SelectedValue = TAMPortfolioID.SelectedValue
            child.Cash.Text = Cash.Text
            child.Show()
            Call child.LoadAfterInitialUpdate()

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "ERROR")
        End Try
    End Sub

    Public Sub Pause(ByVal dblSecs As Double)
        Const OneSec As Double = 1.0# / (1440.0# * 60.0#)
        Dim dblWaitTil As Date
        Now.AddSeconds(OneSec)
        dblWaitTil = Now.AddSeconds(OneSec).AddSeconds(dblSecs)
        Do Until Now > dblWaitTil
            Application.DoEvents() ' Allow windows messages to be processed
        Loop
    End Sub

    Public Sub LoadInitialPerformance()


    End Sub

    Private Sub Button11_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button11.Click
        Dim ir As MsgBoxResult
        ir = MsgBox("You are about to transfer this model to an APX blotter.  Please be sure to save the record and your portfolio mappings are correct." & vbNewLine & "Do you wish to continue?", MsgBoxStyle.YesNo + MsgBoxStyle.Information, "Have you saved the record?")
        If ir = MsgBoxResult.Yes Then
            Call ClearTradeBlotter()
        Else

        End If


    End Sub

    Public Sub ClearTradeBlotter()
        Dim Mycn As OleDb.OleDbConnection
        Dim Command As OleDb.OleDbCommand
        Dim SQLstr As String

        Try
            Mycn = New OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & My.Settings.dbloc & "';Jet OLEDB:Database Password='" & My.Settings.dbpass & "';")
            Mycn.Open()

            SQLstr = "DELETE * FROM uma_tradeblotter"

            Command = New OleDb.OleDbCommand(SQLstr, Mycn)
            Command.ExecuteNonQuery()

            Mycn.Close()

            Call LoadETFBlotter()

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "ERROR")
        End Try
    End Sub

    Public Sub LoadETFBlotter()
        Dim Mycn As OleDb.OleDbConnection
        Dim Command As OleDb.OleDbCommand
        Dim SQLstr As String

        Try
            Mycn = New OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & My.Settings.dbloc & "';Jet OLEDB:Database Password='" & My.Settings.dbpass & "';")
            Mycn.Open()

            SQLstr = "INSERT INTO uma_tradeblotter ( Field1, Field2, Field4, Field5, Field6, Field8, Field9, Field17, Field18, Field19, Field29, Field30, Field42, Field45, Field46, Field59 )" & _
            " SELECT dbo_vQbRowDefPortfolio.PortfolioCode, 'li' AS Expr1, dbo_vQbRowDefSecurity.SecType, dbo_vQbRowDefSecurity.Symbol, Format(mdb_ETFPrice_ModelMain.InceptionDate,'mmddyyyy') AS Expr2, Format(mdb_ETFPrice_ModelMain.InceptionDate,'mmddyyyy') AS Expr3, mdb_ETFPricing_ModelHoldings.Shares, 'y' AS Expr4, (mdb_ETFPricing_ModelHoldings.Shares*mdb_ETFPricing_ModelHoldings.Price) AS Expr5, (mdb_ETFPricing_ModelHoldings.Shares*mdb_ETFPricing_ModelHoldings.Price) AS Expr6, 'n' AS Expr7, '65533' AS Expr8, '1' AS Expr9, 'n' AS Expr10, 'y' AS Expr11, 'y' AS Expr12" & _
            " FROM ((mdb_ETFPricing_ModelHoldings INNER JOIN mdb_ETFPrice_ModelMain ON mdb_ETFPricing_ModelHoldings.ModelID = mdb_ETFPrice_ModelMain.ID) INNER JOIN dbo_vQbRowDefSecurity ON mdb_ETFPricing_ModelHoldings.SecurityID = dbo_vQbRowDefSecurity.SecurityID) INNER JOIN dbo_vQbRowDefPortfolio ON mdb_ETFPrice_ModelMain.ETFPortfolioID = dbo_vQbRowDefPortfolio.PortfolioID" & _
            " WHERE mdb_ETFPrice_ModelMain.ID = " & ID.Text

            Command = New OleDb.OleDbCommand(SQLstr, Mycn)
            Command.ExecuteNonQuery()

            Mycn.Close()

            Mycn.Open()

            SQLstr = "INSERT INTO uma_tradeblotter ( Field1, Field2, Field4, Field5, Field6, Field8, Field9, Field17, Field18, Field19, Field29, Field30, Field42, Field45, Field46, Field59)" & _
            " SELECT dbo_vQbRowDefPortfolio.PortfolioCode, 'li' AS Expr1, dbo_vQbRowDefSecurity.SecType, dbo_vQbRowDefSecurity.Symbol, Format(mdb_ETFPrice_ModelMain.InceptionDate,'mmddyyyy') AS Expr2, Format(mdb_ETFPrice_ModelMain.InceptionDate,'mmddyyyy') AS Expr3, mdb_ETFPricing_ModelHoldings.Shares, 'y' AS Expr4, (mdb_ETFPricing_ModelHoldings.Shares*mdb_ETFPricing_ModelHoldings.Price) AS Expr5, (mdb_ETFPricing_ModelHoldings.Shares*mdb_ETFPricing_ModelHoldings.Price) AS Expr6, 'n' AS Expr7, '65533' AS Expr8, '1' AS Expr9, 'n' AS Expr10, 'y' AS Expr11, 'y' AS Expr12" & _
            " FROM ((mdb_ETFPricing_ModelHoldings INNER JOIN mdb_ETFPrice_ModelMain ON mdb_ETFPricing_ModelHoldings.ModelID = mdb_ETFPrice_ModelMain.ID) INNER JOIN dbo_vQbRowDefSecurity ON mdb_ETFPricing_ModelHoldings.SecurityID = dbo_vQbRowDefSecurity.SecurityID) INNER JOIN dbo_vQbRowDefPortfolio ON mdb_ETFPrice_ModelMain.TAMPortfolioID = dbo_vQbRowDefPortfolio.PortfolioID" & _
            " WHERE mdb_ETFPricing_ModelHoldings.SecurityID Not In (Select APXSymbol FROM mdb_SymbolMapping WHERE Active = True) AND mdb_ETFPrice_ModelMain.ID = " & ID.Text

            Command = New OleDb.OleDbCommand(SQLstr, Mycn)
            Command.ExecuteNonQuery()

            Mycn.Close()

            Mycn.Open()

            SQLstr = "INSERT INTO uma_tradeblotter ( Field1, Field2, Field4, Field5, Field6, Field8, Field9, Field17, Field18, Field19, Field29, Field30, Field33, Field42, Field45, Field46, Field59)" & _
            " SELECT dbo_vQbRowDefPortfolio.PortfolioCode, 'li' AS Expr1, dbo_vQbRowDefSecurity.SecType, dbo_vQbRowDefSecurity.Symbol, Format(mdb_ETFPrice_ModelMain.InceptionDate,'mmddyyyy') AS Expr2, Format(mdb_ETFPrice_ModelMain.InceptionDate,'mmddyyyy') AS Expr3, (mdb_ETFPricing_ModelHoldings.Shares * mdb_SymbolMapping.QntyMultiplier), 'y' AS Expr4, (mdb_ETFPricing_ModelHoldings.Shares*mdb_ETFPricing_ModelHoldings.Price) AS Expr5, (mdb_ETFPricing_ModelHoldings.Shares*mdb_ETFPricing_ModelHoldings.Price) AS Expr6, 'n' AS Expr7, '65533' AS Expr8,(mdb_ETFPricing_ModelHoldings.Shares * mdb_SymbolMapping.QntyMultiplier), '1' AS Expr9, 'n' AS Expr10, 'y' AS Expr11, 'y' AS Expr12" & _
            " FROM (((mdb_ETFPricing_ModelHoldings INNER JOIN mdb_ETFPrice_ModelMain ON mdb_ETFPricing_ModelHoldings.ModelID = mdb_ETFPrice_ModelMain.ID) INNER JOIN dbo_vQbRowDefPortfolio ON mdb_ETFPrice_ModelMain.TAMPortfolioID = dbo_vQbRowDefPortfolio.PortfolioID) INNER JOIN mdb_SymbolMapping ON mdb_ETFPricing_ModelHoldings.SecurityID = mdb_SymbolMapping.APXSymbol) INNER JOIN dbo_vQbRowDefSecurity ON mdb_SymbolMapping.NewAPXSymbol = dbo_vQbRowDefSecurity.SecurityID" & _
            " WHERE mdb_SymbolMapping.Active=True AND mdb_ETFPrice_ModelMain.ID = " & ID.Text

            Command = New OleDb.OleDbCommand(SQLstr, Mycn)
            Command.ExecuteNonQuery()

            Mycn.Close()

            Call SendETFBlotter()

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "ERROR")
        End Try
    End Sub

    Public Sub SendETFBlotter()
        Dim path As String
        path = "C:\_Blotters"

        If (System.IO.Directory.Exists(path)) Then
            System.IO.Directory.Delete(path, True)
        End If

        Dim AccessConn As New System.Data.OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & My.Settings.dbloc & "';Jet OLEDB:Database Password='" & My.Settings.dbpass & "';")

        AccessConn.Open()

        Dim AccessCommand As New OleDb.OleDbCommand("SELECT Field1,Field2,Field3,Field4,Field5,Field6,Field7,Field8,Field9,Field10,Field11,Field12,Field13,Field14,Field15,Field16,Field17,Field18,Field19,Field20,Field21,Field22,Field23,Field24,Field25,Field26,Field27,Field28,Field29,Field30,Field31,Field32,Field33,Field34,Field35,Field36,Field37,Field38,Field39,Field40,Field41,Field42,Field43,Field44,Field45,Field46,Field47,Field48,Field49,Field50,Field51,Field52,Field53,Field54,Field55,Field56,Field57,Field58,Field59,Field60,Field61,Field62,Field63,Field64,Field65,Field66,Field67,Field68,Field69,Field70,Field71,Field72,Field73,Field74,Field75,Field76,Field77,Field78 INTO [Text;HDR=NO;DATABASE=" & path & "].[Envestnet_UMA_Trades.csv] FROM uma_tradeblotter", AccessConn)
        System.IO.Directory.CreateDirectory(path)
        AccessCommand.ExecuteNonQuery()
        AccessConn.Close()

        Call cleantradefile()

        System.Diagnostics.Process.Start("\\aamapxapps01\apx$\Automation\UMATradeBlotter.BAT")
        System.Diagnostics.Process.Start("\\aamapxapps01\apx$\Automation\UMATradeBlotter64.BAT")

        MsgBox("Trades have been loaded into the APX Trade Blotter.", MsgBoxStyle.Information, "Import Successful.")
    End Sub

    Public Sub cleantradefile()
        Dim myFiles As String()
        Dim path As String
        path = "C:\_Blotters"
        myFiles = IO.Directory.GetFiles(path, "*.csv")

        Dim newFilePath As String

        For Each filepath As String In myFiles

            newFilePath = filepath.Replace(".csv", ".trn")

            System.IO.File.Move(filepath, newFilePath)

        Next

        Dim FileToDelete As String

        FileToDelete = "\\aamapxapps01\apx$\imp\Envestnet_UMA_Trades.trn"

        If System.IO.File.Exists(FileToDelete) = True Then

            System.IO.File.Delete(FileToDelete)
            'MsgBox("File Deleted")
        End If

        My.Computer.FileSystem.CopyFile("C:\_Blotters\Envestnet_UMA_Trades.trn", "\\aamapxapps01\apx$\imp\Envestnet_UMA_Trades.trn")
    End Sub

End Class