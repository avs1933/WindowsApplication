Public Class map_Firm
    Dim loadinhouseSMA As System.Threading.Thread
    Dim loadoutsideSMA As System.Threading.Thread
    Dim loadtopviews As System.Threading.Thread
    Dim loadad As System.Threading.Thread

    Private Sub SplitContainer2_Panel1_Paint(ByVal sender As System.Object, ByVal e As System.Windows.Forms.PaintEventArgs) Handles SplitContainer2.Panel1.Paint

    End Sub

    Private Sub map_Firm_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Control.CheckForIllegalCrossThreadCalls = False
        loadinhouseSMA = New System.Threading.Thread(AddressOf LoadInHouseSMAData)
        loadinhouseSMA.Start()

        loadoutsideSMA = New System.Threading.Thread(AddressOf LoadOutsideSMAData)
        loadoutsideSMA.Start()

        loadtopviews = New System.Threading.Thread(AddressOf LoadTopInhouseAdvisors)
        loadtopviews.Start()

        Timer1.Enabled = True
        TextBox4.Visible = True
        TextBox5.Visible = True

        Timer2.Enabled = True
        TextBox3.Visible = True
        TextBox2.Visible = True

        Timer3.Enabled = True
        TextBox7.Visible = True
        TextBox6.Visible = True

        Timer4.Enabled = True
        TextBox9.Visible = True
        TextBox8.Visible = True

        Timer5.Enabled = True
        TextBox11.Visible = True
        TextBox10.Visible = True

        'Call CleanTempTables()
        Call LoadADVISEDTable()
        Call LoadSOLICITEDTable()
        Call LoadSubAdvisedTable()
        Call LoadPlatformTable()
        Call LoadWLPlatformTable()
        'Call LoadTopInhouseAdvisors()
    End Sub

    Public Sub LoadTopInhouseAdvisors()
        Try
            Dim conn As New OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & My.Settings.dbloc & "';Jet OLEDB:Database Password='" & My.Settings.dbpass & "';")

            Dim strSQL As String = "SELECT ExternalAdvisor As Advisor, Count(PortfolioCode) AS [# of accounts], Sum(TMV) AS [Market Value], Sum(AAM1) AS BSC" & _
            " FROM(env_BSC_map_Advisor)" & _
            " WHERE ExternalFirm = (SELECT Map_Firms.AdventPortfolioFirm FROM (Map_Firms) WHERE(((Map_Firms.ID) = " & ID.Text & ")) GROUP BY Map_Firms.AdventPortfolioFirm)" & _
            " GROUP BY ExternalAdvisor"

            Dim da As New OleDb.OleDbDataAdapter(strSQL, conn)
            Dim ds As New DataSet
            da.Fill(ds, "Users")

            With DataGridView8
                .DataSource = ds.Tables("Users")
                .Columns(2).DefaultCellStyle.Format = "c"
                .Columns(3).DefaultCellStyle.Format = "c"
                .Sort(DataGridView8.Columns(2), System.ComponentModel.ListSortDirection.Descending)
            End With

            Label28.Text = DataGridView8.RowCount & " Records Found"
            Call LoadTopInhouseProducts()

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "ERROR")
        End Try
    End Sub

    Public Sub LoadTopInhouseProducts()
        Try
            Dim conn As New OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & My.Settings.dbloc & "';Jet OLEDB:Database Password='" & My.Settings.dbpass & "';")

            Dim strSQL As String = "SELECT Discipline6 As Discipline, Count(PortfolioCode) AS [# of accounts], Sum(TMV) AS [Market Value], Sum(AAM1) AS BSC" & _
            " FROM(env_BSC_map_Advisor)" & _
            " WHERE ExternalFirm = (SELECT Map_Firms.AdventPortfolioFirm FROM (Map_Firms) WHERE(((Map_Firms.ID) = " & ID.Text & ")) GROUP BY Map_Firms.AdventPortfolioFirm)" & _
            " GROUP BY Discipline6"

            Dim da As New OleDb.OleDbDataAdapter(strSQL, conn)
            Dim ds As New DataSet
            da.Fill(ds, "Users")

            With DataGridView10
                .DataSource = ds.Tables("Users")
                .Columns(2).DefaultCellStyle.Format = "c"
                .Columns(3).DefaultCellStyle.Format = "c"
                .Sort(DataGridView10.Columns(2), System.ComponentModel.ListSortDirection.Descending)
            End With

            Label29.Text = DataGridView10.RowCount & " Records Found"
            Call LoadTopOutsideAdvisors()

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "ERROR")
        End Try
    End Sub

    Public Sub LoadTopOutsideAdvisors()
        Try
            Dim conn As New OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & My.Settings.dbloc & "';Jet OLEDB:Database Password='" & My.Settings.dbpass & "';")

            Dim strSQL As String = "SELECT ExternalAdvisor As Advisor, Count(PortfolioCode) AS [# of accounts], Sum(TMV) AS [Market Value], Sum(BSC1) AS BSC" & _
            " FROM(env_BSC_External_map)" & _
            " WHERE ExternalFirm = (SELECT Map_Firms.AdventPortfolioFirm FROM (Map_Firms) WHERE(((Map_Firms.ID) = " & ID.Text & ")) GROUP BY Map_Firms.AdventPortfolioFirm)" & _
            " GROUP BY ExternalAdvisor"

            Dim da As New OleDb.OleDbDataAdapter(strSQL, conn)
            Dim ds As New DataSet
            da.Fill(ds, "Users")

            With DataGridView9
                .DataSource = ds.Tables("Users")
                .Columns(2).DefaultCellStyle.Format = "c"
                .Columns(3).DefaultCellStyle.Format = "c"
                .Sort(DataGridView9.Columns(2), System.ComponentModel.ListSortDirection.Descending)
            End With
            Label30.Text = DataGridView9.RowCount & " Records Found"
            Call LoadTopOutsideProducts()

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "ERROR")
        End Try
    End Sub

    Public Sub LoadTopOutsideProducts()
        Try
            Dim conn As New OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & My.Settings.dbloc & "';Jet OLEDB:Database Password='" & My.Settings.dbpass & "';")

            Dim strSQL As String = "SELECT ExternalInformation As Discipline, Count(PortfolioCode) AS [# of accounts], Sum(TMV) AS [Market Value], Sum(BSC1) AS BSC" & _
            " FROM(env_BSC_External_map)" & _
            " WHERE ExternalFirm = (SELECT Map_Firms.AdventPortfolioFirm FROM (Map_Firms) WHERE(((Map_Firms.ID) = " & ID.Text & ")) GROUP BY Map_Firms.AdventPortfolioFirm)" & _
            " GROUP BY ExternalInformation"

            Dim da As New OleDb.OleDbDataAdapter(strSQL, conn)
            Dim ds As New DataSet
            da.Fill(ds, "Users")

            With DataGridView11
                .DataSource = ds.Tables("Users")
                .Columns(2).DefaultCellStyle.Format = "c"
                .Columns(3).DefaultCellStyle.Format = "c"
                .Sort(DataGridView11.Columns(2), System.ComponentModel.ListSortDirection.Descending)
            End With

            Label31.Text = DataGridView11.RowCount & " Records Found"

            Call LoadTopOutsideReps()

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "ERROR")
        End Try
    End Sub

    Public Sub LoadTopOutsideReps()
        Try
            Dim conn As New OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & My.Settings.dbloc & "';Jet OLEDB:Database Password='" & My.Settings.dbpass & "';")

            Dim strSQL As String = "SELECT AAMRepName As [AAM Rep], Count(PortfolioCode) AS [# of accounts], Sum(TMV) AS [Market Value], Sum(BSC1) AS BSC" & _
            " FROM(env_BSC_External_map)" & _
            " WHERE ExternalFirm = (SELECT Map_Firms.AdventPortfolioFirm FROM (Map_Firms) WHERE(((Map_Firms.ID) = " & ID.Text & ")) GROUP BY Map_Firms.AdventPortfolioFirm)" & _
            " GROUP BY AAMRepName"

            Dim da As New OleDb.OleDbDataAdapter(strSQL, conn)
            Dim ds As New DataSet
            da.Fill(ds, "Users")

            With DataGridView13
                .DataSource = ds.Tables("Users")
                .Columns(2).DefaultCellStyle.Format = "c"
                .Columns(3).DefaultCellStyle.Format = "c"
                .Sort(DataGridView13.Columns(2), System.ComponentModel.ListSortDirection.Descending)
            End With

            Label33.Text = DataGridView13.RowCount & " Records Found"

            Call LoadTopInhouseReps()

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "ERROR")
        End Try
    End Sub

    Public Sub LoadTopInhouseReps()
        Try
            Dim conn As New OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & My.Settings.dbloc & "';Jet OLEDB:Database Password='" & My.Settings.dbpass & "';")

            Dim strSQL As String = "SELECT AAMRepName As [AAM Rep], Count(PortfolioCode) AS [# of accounts], Sum(TMV) AS [Market Value], Sum(AAM1) AS BSC" & _
            " FROM(env_BSC_map)" & _
            " WHERE ExternalFirm = (SELECT Map_Firms.AdventPortfolioFirm FROM (Map_Firms) WHERE(((Map_Firms.ID) = " & ID.Text & ")) GROUP BY Map_Firms.AdventPortfolioFirm)" & _
            " GROUP BY AAMRepName"

            Dim da As New OleDb.OleDbDataAdapter(strSQL, conn)
            Dim ds As New DataSet
            da.Fill(ds, "Users")

            With DataGridView12
                .DataSource = ds.Tables("Users")
                .Columns(2).DefaultCellStyle.Format = "c"
                .Columns(3).DefaultCellStyle.Format = "c"
                .Sort(DataGridView12.Columns(2), System.ComponentModel.ListSortDirection.Descending)
            End With

            Label32.Text = DataGridView12.RowCount & " Records Found"

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "ERROR")
        End Try
    End Sub
    Public Sub LoadInHouseSMAData()
        Dim Mycn As OleDb.OleDbConnection

        Try

            Mycn = New OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & My.Settings.dbloc & "';Jet OLEDB:Database Password='" & My.Settings.dbpass & "';")

            Mycn.Open()

            Dim sqlstring As String
            sqlstring = "SELECT Count(env_BSC_map.PortfolioCode) AS Accounts1, Sum(TMV) AS TMV1, Sum(AAM1) AS AAM1" & _
            " FROM(env_BSC_map)" & _
            " WHERE env_BSC_map.ExternalFirm = (SELECT Map_Firms.AdventPortfolioFirm FROM (Map_Firms) WHERE(((Map_Firms.ID) = " & ID.Text & ")) GROUP BY Map_Firms.AdventPortfolioFirm)"

            Dim queryString As String = String.Format(sqlstring)
            Dim cmd As New OleDb.OleDbCommand(queryString, Mycn)
            Dim da As New OleDb.OleDbDataAdapter(cmd)
            Dim ds As New DataSet

            da.Fill(ds, "User")
            Dim dt As DataTable = ds.Tables("User")
            If dt.Rows.Count > 0 Then

                Dim row As DataRow = dt.Rows(0)


                Dim Billing As Double
                If IsDBNull((row("AAM1"))) Then
                    Billing = 0.0
                Else
                    Billing = (row("AAM1"))
                End If

                Dim TMV As Integer
                If IsDBNull((row("TMV1"))) Then
                    TMV = 0.0
                Else
                    TMV = (row("TMV1"))
                End If

                Dim velocity As Double
                If TMV = 0.0 Or Billing = 0.0 Then
                    velocity = 0
                Else
                    velocity = ((Billing / TMV) * 100)
                End If

                Dim accts As Integer
                If IsDBNull((row("Accounts1"))) Then
                    accts = 0
                Else
                    accts = (row("Accounts1"))
                End If

                txtAccounts.Text = accts
                txtAUM.Text = Format(TMV, "$#,###.00")
                txtBSC.Text = Format(Billing, "$#,###.00")
                txtVelocity.Text = Format(velocity, "#.00")
            Else

            End If
            Mycn.Close()
            Mycn.Dispose()

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "ERROR")

        End Try
    End Sub

    Public Sub LoadOutsideSMAData()
        Dim Mycn As OleDb.OleDbConnection

        Try

            Mycn = New OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & My.Settings.dbloc & "';Jet OLEDB:Database Password='" & My.Settings.dbpass & "';")

            Mycn.Open()

            Dim sqlstring As String
            sqlstring = "SELECT Count(env_BSC_External_map.PortfolioCode) AS Accounts1, Sum(TMV) AS TMV1, Sum(BSC1) AS AAM1" & _
            " FROM(env_BSC_External_map)" & _
            " WHERE env_BSC_External_map.ExternalFirm = (SELECT Map_Firms.AdventPortfolioFirm FROM (Map_Firms) WHERE(((Map_Firms.ID) = " & ID.Text & ")) GROUP BY Map_Firms.AdventPortfolioFirm)"

            Dim queryString As String = String.Format(sqlstring)
            Dim cmd As New OleDb.OleDbCommand(queryString, Mycn)
            Dim da As New OleDb.OleDbDataAdapter(cmd)
            Dim ds As New DataSet

            da.Fill(ds, "User")
            Dim dt As DataTable = ds.Tables("User")
            If dt.Rows.Count > 0 Then

                Dim row As DataRow = dt.Rows(0)


                Dim Billing As Double
                If IsDBNull((row("AAM1"))) Then
                    Billing = 0.0
                Else
                    Billing = (row("AAM1"))
                End If

                Dim TMV As Integer
                If IsDBNull((row("TMV1"))) Then
                    TMV = 0.0
                Else
                    TMV = (row("TMV1"))
                End If

                Dim velocity As Double
                If TMV = 0.0 Or Billing = 0.0 Then
                    velocity = 0
                Else
                    velocity = ((Billing / TMV) * 100)
                End If

                Dim accts As Integer
                If IsDBNull((row("Accounts1"))) Then
                    accts = 0
                Else
                    accts = (row("Accounts1"))
                End If

                txtAUSAccts.Text = accts
                txtAUS.Text = Format(TMV, "$#,###.00")
                txtAUSBSC.Text = Format(Billing, "$#,###.00")
                txtAUSVelocity.Text = Format(velocity, "#.00")
            Else

            End If
            Mycn.Close()
            Mycn.Dispose()

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "ERROR")

        End Try
    End Sub

    Public Sub LoadSubAdvisedTable()
        TextBox5.Text = "Drawing Grid..."

        Try
            Dim conn As New OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & My.Settings.dbloc & "';Jet OLEDB:Database Password='" & My.Settings.dbpass & "';")

            Dim SqlString As String

            SqlString = "SELECT map_TempAssignments.ID, map_Products.ProductName, Map_Firms.FirmName, map_TempAssignments.BreakpointFrom, map_TempAssignments.BreakpointTo, map_TempAssignments.Fee, map_TempAssignments.MaxRepFee" & _
            " FROM (map_TempAssignments INNER JOIN map_Products ON map_TempAssignments.ProductID = map_Products.ID) INNER JOIN Map_Firms ON map_TempAssignments.ManagingFirmID = Map_Firms.ID" & _
            " WHERE map_TempAssignments.FirmID = " & ID.Text & " AND map_TempAssignments.TypeID = 1"

            Dim da As New OleDb.OleDbDataAdapter(SqlString, conn)
            Dim ds As New DataSet
            da.Fill(ds, "Users")

            With DataGridView4
                .DataSource = ds.Tables("Users")
                .Columns(0).Visible = False
                .Columns(3).DefaultCellStyle.Format = "c"
                .Columns(4).DefaultCellStyle.Format = "c"
                .Columns(3).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                .Columns(4).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
            End With

            Timer1.Enabled = False
            TextBox4.Visible = False
            TextBox5.Visible = False

            DataGridView4.Refresh()
            Call CheckSMAStatus()
            If DataGridView4.RowCount = 0 Then
                Label4.Text = "NO SUB-ADVISED AGREEMENTS"
            Else
                Label4.Text = "SUB-ADVISED AGREEMENTS IN PLACE"
            End If

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "ERROR")
        End Try
    End Sub

    Public Sub LoadSOLICITEDTable()
        TextBox2.Text = "Drawing Grid..."

        Try
            Dim conn As New OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & My.Settings.dbloc & "';Jet OLEDB:Database Password='" & My.Settings.dbpass & "';")

            Dim SqlString As String

            SqlString = "SELECT map_TempAssignments.ID, map_Products.ProductName, Map_Firms.FirmName, map_TempAssignments.BreakpointFrom, map_TempAssignments.BreakpointTo, map_TempAssignments.Fee, map_TempAssignments.MaxRepFee" & _
            " FROM (map_TempAssignments INNER JOIN map_Products ON map_TempAssignments.ProductID = map_Products.ID) INNER JOIN Map_Firms ON map_TempAssignments.ManagingFirmID = Map_Firms.ID" & _
            " WHERE map_TempAssignments.FirmID = " & ID.Text & " AND ((map_TempAssignments.TypeID = 2))"

            Dim da As New OleDb.OleDbDataAdapter(SqlString, conn)
            Dim ds As New DataSet
            da.Fill(ds, "Users")

            With DataGridView5
                .DataSource = ds.Tables("Users")
                .Columns(0).Visible = False
                .Columns(3).DefaultCellStyle.Format = "c"
                .Columns(4).DefaultCellStyle.Format = "c"
                .Columns(3).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                .Columns(4).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
            End With

            Timer2.Enabled = False
            TextBox3.Visible = False
            TextBox2.Visible = False

            DataGridView5.Refresh()
            Call CheckSMAStatus()
            If DataGridView5.RowCount = 0 Then
                Label5.Text = "NO SOLICITOR AGREEMENTS"
            Else
                Label5.Text = "SOLICITOR AGREEMENTS IN PLACE"
            End If

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "ERROR")
        End Try
    End Sub

    Public Sub LoadADVISEDTable()
        TextBox6.Text = "Drawing Grid..."

        Try
            Dim conn As New OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & My.Settings.dbloc & "';Jet OLEDB:Database Password='" & My.Settings.dbpass & "';")

            Dim SqlString As String

            SqlString = "SELECT map_TempAssignments.ID, map_Products.ProductName, Map_Firms.FirmName, map_TempAssignments.BreakpointFrom, map_TempAssignments.BreakpointTo, map_TempAssignments.Fee, map_TempAssignments.MaxRepFee" & _
            " FROM (map_TempAssignments INNER JOIN map_Products ON map_TempAssignments.ProductID = map_Products.ID) INNER JOIN Map_Firms ON map_TempAssignments.ManagingFirmID = Map_Firms.ID" & _
            " WHERE map_TempAssignments.FirmID = " & ID.Text & " AND ((map_TempAssignments.TypeID = 3))"

            Dim da As New OleDb.OleDbDataAdapter(SqlString, conn)
            Dim ds As New DataSet
            da.Fill(ds, "Users")

            With DataGridView6
                .DataSource = ds.Tables("Users")
                .Columns(0).Visible = False
                .Columns(3).DefaultCellStyle.Format = "c"
                .Columns(4).DefaultCellStyle.Format = "c"
                .Columns(3).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                .Columns(4).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
            End With

            Timer3.Enabled = False
            TextBox7.Visible = False
            TextBox6.Visible = False

            DataGridView6.Refresh()
            Call CheckSMAStatus()
            If DataGridView6.RowCount = 0 Then
                Label6.Text = "NO ADVISED AGREEMENTS"
            Else
                Label6.Text = "ADVISED AGREEMENTS IN PLACE"
            End If

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "ERROR")
        End Try
    End Sub

    Public Sub LoadPlatformTable()
        TextBox8.Text = "Drawing Grid..."

        Try
            Dim conn As New OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & My.Settings.dbloc & "';Jet OLEDB:Database Password='" & My.Settings.dbpass & "';")

            Dim SqlString As String

            SqlString = "SELECT map_TempAssignments.ID, Map_Firms.FirmName AS [Platform], map_Products.ProductName AS [Product Name], map_ProductType.TypeName AS [Type of Product], Map_Firms_1.FirmName AS [Managing Firm], map_TempAssignments.BreakpointFrom As [Starting Breakpoint], map_TempAssignments.BreakpointTo As [Ending Breakpoint], map_TempAssignments.Fee, map_TempAssignments.MaxRepFee As [Max Rep Fee], map_TempAssignments.PlatformApproved As [Platform Approved], map_TempAssignments.SMAOffered As [SMA], map_TempAssignments.UMAOffered As [UMA]" & _
            " FROM (((map_TempAssignments INNER JOIN (map_Platforms INNER JOIN Map_Firms ON map_Platforms.PlatformID = Map_Firms.ID) ON map_TempAssignments.PlatformID = map_Platforms.ID) INNER JOIN map_Products ON map_TempAssignments.ProductID = map_Products.ID) INNER JOIN map_ProductType ON map_TempAssignments.TypeOfProductID = map_ProductType.ID) INNER JOIN Map_Firms AS Map_Firms_1 ON map_TempAssignments.ManagingFirmID = Map_Firms_1.ID" & _
            " WHERE(((map_TempAssignments.FirmID) = " & ID.Text & ") AND map_TempAssignments.TypeID = 4)" & _
            " GROUP BY map_TempAssignments.ID, Map_Firms.FirmName, map_Products.ProductName, map_ProductType.TypeName, Map_Firms_1.FirmName, map_TempAssignments.BreakpointFrom, map_TempAssignments.BreakpointTo, map_TempAssignments.Fee, map_TempAssignments.MaxRepFee, map_TempAssignments.SMAOffered, map_TempAssignments.UMAOffered, map_TempAssignments.PlatformApproved;"


            Dim da As New OleDb.OleDbDataAdapter(SqlString, conn)
            Dim ds As New DataSet
            da.Fill(ds, "Users")

            With DataGridView1
                .DataSource = ds.Tables("Users")
                .Columns(0).Visible = False
                .Columns(5).DefaultCellStyle.Format = "c"
                .Columns(6).DefaultCellStyle.Format = "c"
                .Columns(5).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                .Columns(6).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
            End With

            Timer4.Enabled = False
            TextBox9.Visible = False
            TextBox8.Visible = False

            DataGridView6.Refresh()
            Call CheckSMAStatus()
            If DataGridView1.RowCount = 0 Then
                Label1.Text = "NO PLATFORMS"
            Else
                Label1.Text = "FIRM USES PLATFORMS"
            End If

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "ERROR")
        End Try
    End Sub

    Public Sub LoadWLPlatformTable()
        TextBox10.Text = "Drawing Grid..."

        Try
            Dim conn As New OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & My.Settings.dbloc & "';Jet OLEDB:Database Password='" & My.Settings.dbpass & "';")

            Dim SqlString As String

            SqlString = "SELECT map_TempAssignments.ID, map_PlatformsWL.WLName As [WL Name], Map_Firms.FirmName AS [Platform], map_Products.ProductName AS [Product Name], map_ProductType.TypeName AS [Type of Product], Map_Firms_1.FirmName AS [Managing Firm], map_TempAssignments.BreakpointFrom As [Starting Breakpoint], map_TempAssignments.BreakpointTo As [Ending Breakpoint], map_TempAssignments.Fee, map_TempAssignments.MaxRepFee As [Max Rep Fee], map_TempAssignments.PlatformApproved As [Platform Approved], map_TempAssignments.SMAOffered As [SMA], map_TempAssignments.UMAOffered As [UMA]" & _
            " FROM (((((map_TempAssignments INNER JOIN map_PlatformsWL ON map_TempAssignments.WLID = map_PlatformsWL.ID) INNER JOIN map_Platforms ON map_PlatformsWL.PlatformDriverID = map_Platforms.ID) INNER JOIN Map_Firms ON map_Platforms.PlatformID = Map_Firms.ID) INNER JOIN map_Products ON map_TempAssignments.ProductID = map_Products.ID) INNER JOIN map_ProductType ON map_TempAssignments.TypeOfProductID = map_ProductType.ID) INNER JOIN Map_Firms AS Map_Firms_1 ON map_TempAssignments.ManagingFirmID = Map_Firms_1.ID" & _
            " WHERE(((map_TempAssignments.FirmID) = " & ID.Text & ") AND map_TempAssignments.TypeID = 5)" & _
            " GROUP BY map_TempAssignments.ID, map_PlatformsWL.WLName, Map_Firms.FirmName, map_Products.ProductName, map_ProductType.TypeName, Map_Firms_1.FirmName, map_TempAssignments.BreakpointFrom, map_TempAssignments.BreakpointTo, map_TempAssignments.Fee, map_TempAssignments.MaxRepFee, map_TempAssignments.SMAOffered, map_TempAssignments.UMAOffered, map_TempAssignments.PlatformApproved;"


            Dim da As New OleDb.OleDbDataAdapter(SqlString, conn)
            Dim ds As New DataSet
            da.Fill(ds, "Users")

            With DataGridView2
                .DataSource = ds.Tables("Users")
                .Columns(0).Visible = False
                .Columns(7).DefaultCellStyle.Format = "c"
                .Columns(6).DefaultCellStyle.Format = "c"
                .Columns(7).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                .Columns(6).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
            End With

            Timer5.Enabled = False
            TextBox11.Visible = False
            TextBox10.Visible = False

            DataGridView6.Refresh()
            Call CheckSMAStatus()
            If DataGridView2.RowCount = 0 Then
                Label2.Text = "NO WHITE LABEL PLATFORMS"
            Else
                Label2.Text = "FIRM USES WHITE LABEL PLATFORMS"
            End If

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "ERROR")
        End Try
    End Sub

    Private Sub Timer1_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer1.Tick
        If TextBox4.Text = "Working" Then
            TextBox4.Text = "Working."
        Else
            If TextBox4.Text = "Working." Then
                TextBox4.Text = "Working.."
            Else
                If TextBox4.Text = "Working.." Then
                    TextBox4.Text = "Working..."
                Else
                    If TextBox4.Text = "Working..." Then
                        TextBox4.Text = "Working"
                    End If
                End If
            End If
        End If
    End Sub

    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        Timer1.Enabled = True
        TextBox4.Visible = True
        TextBox5.Visible = True
        Call LoadSubAdvisedTable()
        Call CheckSMAStatus()
    End Sub

    Private Sub Timer2_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer2.Tick
        If TextBox3.Text = "Working" Then
            TextBox3.Text = "Working."
        Else
            If TextBox3.Text = "Working." Then
                TextBox3.Text = "Working.."
            Else
                If TextBox3.Text = "Working.." Then
                    TextBox3.Text = "Working..."
                Else
                    If TextBox3.Text = "Working..." Then
                        TextBox3.Text = "Working"
                    End If
                End If
            End If
        End If
    End Sub

    Private Sub Button5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button5.Click
        Timer2.Enabled = True
        TextBox3.Visible = True
        TextBox2.Visible = True
        Call LoadSOLICITEDTable()
        Call CheckSMAStatus()
    End Sub

    Private Sub Timer3_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer3.Tick
        If TextBox7.Text = "Working" Then
            TextBox7.Text = "Working."
        Else
            If TextBox7.Text = "Working." Then
                TextBox7.Text = "Working.."
            Else
                If TextBox7.Text = "Working.." Then
                    TextBox7.Text = "Working..."
                Else
                    If TextBox7.Text = "Working..." Then
                        TextBox7.Text = "Working"
                    End If
                End If
            End If
        End If
    End Sub

    Private Sub Button6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button6.Click
        Timer3.Enabled = True
        TextBox7.Visible = True
        TextBox6.Visible = True
        Call LoadADVISEDTable()
        Call CheckSMAStatus()
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        Timer1.Enabled = True
        TextBox4.Visible = True
        TextBox5.Visible = True
        Call LoadSubAdvisedTable()

        Timer2.Enabled = True
        TextBox3.Visible = True
        TextBox2.Visible = True
        Call LoadSOLICITEDTable()

        Timer3.Enabled = True
        TextBox7.Visible = True
        TextBox6.Visible = True
        Call LoadADVISEDTable()

        Timer4.Enabled = True
        TextBox9.Visible = True
        TextBox8.Visible = True
        Call LoadPlatformTable()

        Timer5.Enabled = True
        TextBox11.Visible = True
        TextBox10.Visible = True
        Call LoadWLPlatformTable()
        Call CheckSMAStatus()

    End Sub

    Public Sub CheckSMAStatus()
        If ((DataGridView4.RowCount >= 1) Or (DataGridView5.RowCount >= 1) Or (DataGridView6.RowCount >= 1)) Then
            Label3.Visible = True
            Label3.Text = "APPROVED"
            Label3.BackColor = Color.Green
            Label3.ForeColor = Color.White
        Else
            Label3.Visible = True
            Label3.Text = "NOT APPROVED"
            Label3.BackColor = Color.Red
            Label3.ForeColor = Color.White
        End If

        If ((DataGridView1.RowCount >= 1) Or (DataGridView2.RowCount >= 1)) Then
            Label7.Visible = True
            Label7.Text = "APPROVED"
            Label7.BackColor = Color.Green
            Label7.ForeColor = Color.White
        Else
            Label7.Visible = True
            Label7.Text = "NOT APPROVED"
            Label7.BackColor = Color.Red
            Label7.ForeColor = Color.White
        End If
    End Sub

    Private Sub Timer4_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer4.Tick
        If TextBox8.Text = "Working" Then
            TextBox8.Text = "Working."
        Else
            If TextBox8.Text = "Working." Then
                TextBox8.Text = "Working.."
            Else
                If TextBox8.Text = "Working.." Then
                    TextBox8.Text = "Working..."
                Else
                    If TextBox8.Text = "Working..." Then
                        TextBox8.Text = "Working"
                    End If
                End If
            End If
        End If
    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        Timer4.Enabled = True
        TextBox9.Visible = True
        TextBox8.Visible = True
        Call LoadPlatformTable()
        Call CheckSMAStatus()
    End Sub

    Private Sub Timer5_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer5.Tick
        If TextBox11.Text = "Working" Then
            TextBox11.Text = "Working."
        Else
            If TextBox11.Text = "Working." Then
                TextBox11.Text = "Working.."
            Else
                If TextBox11.Text = "Working.." Then
                    TextBox11.Text = "Working..."
                Else
                    If TextBox11.Text = "Working..." Then
                        TextBox11.Text = "Working"
                    End If
                End If
            End If
        End If
    End Sub

    Private Sub Button7_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button7.Click
        Timer5.Enabled = True
        TextBox11.Visible = True
        TextBox10.Visible = True
        Call LoadWLPlatformTable()
        Call CheckSMAStatus()
    End Sub

    Private Sub Timer6_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer6.Tick
        If txtAccounts.Text = "Loading..." Or txtAUSAccts.Text = "Loading..." Then

        Else
            Dim accts As Integer
            accts = txtAccounts.Text
            Dim ausaccts As Integer
            ausaccts = txtAUSAccts.Text
            txtTAccts.Text = accts + ausaccts

            Dim aus As Double
            aus = txtAUS.Text
            Dim aum As Double
            aum = txtAUM.Text
            Dim taum As Double
            taum = aus + aum
            TxtTAssets.Text = Format(taum, "$#,###.00")

            Dim bsc As Double
            bsc = txtBSC.Text
            Dim ausbsc As Double
            ausbsc = txtAUSBSC.Text
            Dim tbsc As Double
            tbsc = bsc + ausbsc
            TxtTBSC.Text = Format(tbsc, "$#,###.00")

            Dim tvel As Double
            tvel = ((tbsc / taum) * 100)
            TxtTVel.Text = Format(tvel, "#.00")

            Timer1.Enabled = False

        End If
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click

        If Permissions.MAPGenerateReports.Checked Then
            Call CheckReportSystemStatus()
        Else
            MsgBox("You do not have permission to perform this task.", MsgBoxStyle.Critical, "Insufficient Permissions")
        End If
    End Sub

    Public Sub CheckReportSystemStatus()
        Dim Mycn As OleDb.OleDbConnection

        Try

            Mycn = New OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & My.Settings.dbloc & "';Jet OLEDB:Database Password='" & My.Settings.dbpass & "';")

            Mycn.Open()

            Dim sqlstring As String
            sqlstring = "SELECT map_ReportRequest.ID, sys_Users.FullName, map_ReportRequest.Finished" & _
            " FROM map_ReportRequest INNER JOIN sys_Users ON map_ReportRequest.EmployeeID = sys_Users.ID" & _
            " WHERE map_ReportRequest.Finished = False"

            Dim queryString As String = String.Format(sqlstring)
            Dim cmd As New OleDb.OleDbCommand(queryString, Mycn)
            Dim da As New OleDb.OleDbDataAdapter(cmd)
            Dim ds As New DataSet

            da.Fill(ds, "User")
            Dim dt As DataTable = ds.Tables("User")
            If dt.Rows.Count > 0 Then

                Dim row As DataRow = dt.Rows(0)

                Dim usr As String
                usr = (row("FullName"))
                MsgBox("The report has been placed in a state of 'locked' by user: " & usr & ".  Please try your request again", MsgBoxStyle.Information, "Report Locked")

            Else
                Call ReportPlaceLock()
            End If

            Mycn.Close()
            Mycn.Dispose()

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "ERROR")

        End Try
    End Sub

    Public Sub ReportPlaceLock()
        Dim Mycn As OleDb.OleDbConnection
        Dim Command As OleDb.OleDbCommand
        Dim SQLstr As String
        Try
            Mycn = New OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & My.Settings.dbloc & "';Jet OLEDB:Database Password='" & My.Settings.dbpass & "';")
            Mycn.Open()

            SQLstr = "INSERT INTO map_ReportRequest(EmployeeID, FirmID)" & _
            "VALUES(" & My.Settings.userid & ", " & ID.Text & ")"

            Command = New OleDb.OleDbCommand(SQLstr, Mycn)
            Command.ExecuteNonQuery()

            Mycn.Close()

            Call LoadProductReport()

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "ERROR")
        End Try
    End Sub

    Public Sub LoadProductReport()
        Dim child As New map_FirmReportViewer
        child.MdiParent = Home
        child.Show()
        Timer7.Enabled = True
    End Sub

    Public Sub ReportReleaseLock()
        Dim Mycn As OleDb.OleDbConnection
        Dim Command As OleDb.OleDbCommand
        Dim SQLstr As String
        Try
            Mycn = New OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & My.Settings.dbloc & "';Jet OLEDB:Database Password='" & My.Settings.dbpass & "';")
            Mycn.Open()

            SQLstr = "UPDATE map_ReportRequest SET Finished = True WHERE EmployeeID = " & My.Settings.userid & " AND FirmID = " & ID.Text

            Command = New OleDb.OleDbCommand(SQLstr, Mycn)
            Command.ExecuteNonQuery()

            Mycn.Close()

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "ERROR")
        End Try
    End Sub

    Private Sub Timer7_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer7.Tick
        Call ReportReleaseLock()
        Timer7.Enabled = False
    End Sub
End Class