Public Class ETF_PriceImportMain
    Dim qdate As System.Threading.Thread


    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        Dim child As New ETF_PriceImportMapping
        child.MdiParent = Home
        child.Show()
        child.ID.Text = "NEW"
    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        Dim child As New ETF_PriceImportEditMapping
        child.MdiParent = Home
        child.Show()
    End Sub

    Private Sub ETF_PriceImportMain_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        'Call LoadRateCheck()
    End Sub

    Public Sub LoadRateCheck()
        Try
            Dim conn As New OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & My.Settings.dbloc & "';Jet OLEDB:Database Password='" & My.Settings.dbpass & "';")

            Dim strSQL As String

            strSQL = "SELECT AdvApp_vSecurity.Symbol As [Actual Symbol], AdvApp_vSecurity_1.Symbol As [Mapped Symbol], AdvApp_vSecurity.InterestOrDividendRate As [Actual Rate], AdvApp_vSecurity_1.InterestOrDividendRate As [Mapped Rate]" & _
            " FROM (AdvApp_vSecurity INNER JOIN mdb_SymbolMapping ON AdvApp_vSecurity.SecurityID = mdb_SymbolMapping.APXSymbol) INNER JOIN AdvApp_vSecurity AS AdvApp_vSecurity_1 ON mdb_SymbolMapping.NewAPXSymbol = AdvApp_vSecurity_1.SecurityID" & _
            " WHERE(AdvApp_vSecurity.InterestOrDividendRate <> AdvApp_vSecurity_1.InterestOrDividendRate)"

            Dim da As New OleDb.OleDbDataAdapter(strSQL, conn)
            Dim ds As New DataSet
            da.Fill(ds, "Users")

            With DataGridView3
                .DataSource = ds.Tables("Users")
            End With

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "ERROR")
        End Try
    End Sub

    Public Sub QueueDates()
        Dim dte As Date
        Dim dte1 As String
        'Dim dte3 As String
        dte = DateTimePicker1.Text
        dte1 = Format(dte, "yyyy-MM-dd")
        'dte3 = dte1.ToString
        Dim Mycn As OleDb.OleDbConnection
        Dim Command As OleDb.OleDbCommand
        Dim SQLstr As String
        Try
            Mycn = New OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & My.Settings.dbloc & "';Jet OLEDB:Database Password='" & My.Settings.dbpass & "';")
            Mycn.Open()

            SQLstr = "INSERT INTO mdb_ETFPricing_DateQueue(QueueDate, PostDate)" & _
            " VALUES('" & dte1 & "', #" & dte & "#)"

            Command = New OleDb.OleDbCommand(SQLstr, Mycn)
            Command.ExecuteNonQuery()

            If ((DateTimePicker1.Text = DateTimePicker2.Text)) Then
                
            Else
                If (DateAdd(DateInterval.Day, 1, dte) = DateTimePicker2.Text) Then
                    dte = DateAdd(DateInterval.Day, 1, dte)
                    dte1 = Format(dte, "yyyy-MM-dd")
                    'dte3 = dte1.ToString
                    SQLstr = "INSERT INTO mdb_ETFPricing_DateQueue(QueueDate, PostDate)" & _
                    " VALUES('" & dte1 & "', #" & dte & "#)"

                    Command = New OleDb.OleDbCommand(SQLstr, Mycn)
                    Command.ExecuteNonQuery()
                Else
                    Do
                        dte = DateAdd(DateInterval.Day, 1, dte)
                        dte1 = Format(dte, "yyyy-MM-dd")
                        'dte3 = dte1.ToString
                        SQLstr = "INSERT INTO mdb_ETFPricing_DateQueue(QueueDate, PostDate)" & _
                        " VALUES('" & dte1 & "', #" & dte & "#)"

                        Command = New OleDb.OleDbCommand(SQLstr, Mycn)
                        Command.ExecuteNonQuery()

                    Loop Until dte = DateTimePicker2.Text
                End If
            End If


            RichTextBox1.Text = RichTextBox1.Text & vbNewLine & "Dates Loaded..."
            Mycn.Close()

            Call LoadDateQueue()

            Call LoadTempPriceFiles()

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "ERROR")
        End Try

    End Sub

    Public Sub LoadDateQueue()
        Try
            Dim conn As New OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & My.Settings.dbloc & "';Jet OLEDB:Database Password='" & My.Settings.dbpass & "';")

            Dim strSQL As String

            strSQL = "SELECT mdb_ETFPricing_DateQueue.ID, mdb_ETFPricing_DateQueue.QueueDate As [Date], mdb_ETFPricing_DateQueue.Working" & _
            " FROM(mdb_ETFPricing_DateQueue)" & _
            " WHERE(((mdb_ETFPricing_DateQueue.Done) = False))" & _
            " GROUP BY mdb_ETFPricing_DateQueue.ID, mdb_ETFPricing_DateQueue.QueueDate, mdb_ETFPricing_DateQueue.Working"
            
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

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        RichTextBox1.Text = "Loading Dates to get Pricing..."
        'Control.CheckForIllegalCrossThreadCalls = False
        'qdate = New System.Threading.Thread(AddressOf QueueDates)
        'qdate.Start()
        Call QueueDates()
        'RichTextBox1.Text = RichTextBox1.Text & vbNewLine & "Dates Loaded..."
        'Call LoadDateQueue()
    End Sub

    Public Sub LoadTempPriceFiles()
        Dim Mycn As OleDb.OleDbConnection
        Dim Command As OleDb.OleDbCommand
        Dim SQLstr As String

        Dim command2 As OleDb.OleDbCommand
        Dim sqlstr2 As String

        Dim command3 As OleDb.OleDbCommand
        Dim sqlstr3 As String

        Dim command4 As OleDb.OleDbCommand
        Dim sqlstr4 As String

        Dim command5 As OleDb.OleDbCommand
        Dim sqlstr5 As String

        Dim command6 As OleDb.OleDbCommand
        Dim sqlstr6 As String

        Mycn = New OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & My.Settings.dbloc & "';Jet OLEDB:Database Password='" & My.Settings.dbpass & "';")
        Mycn.Open()

        'Update Working = False Done = True where ID = ID
        sqlstr6 = "DELETE * FROM mdb_ETFPrice_Rejects"

        command6 = New OleDb.OleDbCommand(sqlstr6, Mycn)
        command6.ExecuteNonQuery()
        Do
            'Select top 1 from mdb_ETFPricing_DateQueue WHERE Done = False
            Try
                Dim conn As New OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & My.Settings.dbloc & "';Jet OLEDB:Database Password='" & My.Settings.dbpass & "';")

                Dim strSQL As String = "SELECT Top 1 * FROM mdb_ETFPricing_DateQueue WHERE Done = False"
                Dim queryString As String = String.Format(strSQL)
                Dim cmd As New OleDb.OleDbCommand(queryString, conn)
                Dim da As New OleDb.OleDbDataAdapter(cmd)
                Dim ds As New DataSet

                da.Fill(ds, "User")
                Dim dt As DataTable = ds.Tables("User")

                Dim row1 As DataRow = dt.Rows(0)

                'Pull ID, PostDate, QueueDate
                Dim fileid As Integer = (row1("ID"))
                Dim postdate As Date = (row1("PostDate"))
                Dim queuedate As String = (row1("QueueDate"))

                Dim datefile As String = Format(postdate, "MMddyy").ToString

                RichTextBox1.Text = RichTextBox1.Text & vbNewLine & "Finding Prices for " & postdate & "..."
                RichTextBox1.SelectionStart = RichTextBox1.TextLength
                Pause(0.01)
                'Update Working = True where ID = ID
                SQLstr = "Update mdb_ETFPricing_DateQueue SET Working = True WHERE ID = " & fileid

                Command = New OleDb.OleDbCommand(SQLstr, Mycn)
                Command.ExecuteNonQuery()
                Pause(0.01)
                'Call LoadDateQueue()

                'Insert Into mdb_ETFPrice_TempLoad query
                sqlstr2 = "INSERT INTO mdb_ETFPrice_TempLoad(FileDate, SecType, Symbol, Price, Source)" & _
                "SELECT mdb_ETFPricing_DateQueue.PostDate, mdb_SymbolMapping.NewAPXSecType, dbo_AdvSecurity.Symbol, dbo_AdvPriceHistory.PriceValue, dbo_AdvPriceHistory.SourceID" & _
                " FROM mdb_ETF_Price_TEMPQuery"

                command2 = New OleDb.OleDbCommand(sqlstr2, Mycn)
                command2.ExecuteNonQuery()
                RichTextBox1.Text = RichTextBox1.Text & vbNewLine & "Prices Loaded..."
                RichTextBox1.SelectionStart = RichTextBox1.TextLength
                Pause(0.5)

                RichTextBox1.Text = RichTextBox1.Text & vbNewLine & "Looking for prices not loaded..."
                RichTextBox1.SelectionStart = RichTextBox1.TextLength

                'Look for prices not loaded due to missing data
                sqlstr5 = "INSERT INTO mdb_ETFPrice_Rejects(NewSecType, NewSymbol, OldSymbol, DateNeeded)" & _
                "SELECT mdb_SymbolMapping.NewAPXSecType, dbo_AdvSecurity.Symbol, dbo_AdvSecurity_1.Symbol, #" & postdate & "#" & _
                " FROM (mdb_SymbolMapping INNER JOIN dbo_AdvSecurity ON mdb_SymbolMapping.NewAPXSymbol = dbo_AdvSecurity.SecurityID) INNER JOIN dbo_AdvSecurity AS dbo_AdvSecurity_1 ON mdb_SymbolMapping.APXSymbol = dbo_AdvSecurity_1.SecurityID" & _
                " WHERE (((dbo_AdvSecurity.[Symbol]) Not In (Select Symbol FROM mdb_ETFPrice_TempLoad)));"


                command5 = New OleDb.OleDbCommand(sqlstr5, Mycn)
                command5.ExecuteNonQuery()

                Dim sqlstring As String
                sqlstring = "SELECT * FROM mdb_ETFPrice_Rejects WHERE DateNeeded = #" & postdate & "#"

                Dim queryString1 As String = String.Format(sqlstring)
                Dim cmd1 As New OleDb.OleDbCommand(queryString1, Mycn)
                Dim da1 As New OleDb.OleDbDataAdapter(cmd1)
                Dim ds1 As New DataSet

                da1.Fill(ds1, "User")
                Dim dt1 As DataTable = ds1.Tables("User")
                If dt1.Rows.Count > 0 Then
                    RichTextBox1.Text = RichTextBox1.Text & vbNewLine & "Found " & dt1.Rows.Count.ToString & " Rejected Prices..."
                    RichTextBox1.SelectionStart = RichTextBox1.TextLength
                    Call LoadRejectQueue()
                    Pause(0.1)
                Else
                    RichTextBox1.Text = RichTextBox1.Text & vbNewLine & "No Rejected Prices Found..."
                    RichTextBox1.SelectionStart = RichTextBox1.TextLength
                    Pause(0.1)
                End If

                Dim sqlstring1 As String
                sqlstring1 = "SELECT * FROM mdb_ETFPrice_TempLoad"

                Dim queryString2 As String = String.Format(sqlstring1)
                Dim cmd2 As New OleDb.OleDbCommand(queryString2, Mycn)
                Dim da2 As New OleDb.OleDbDataAdapter(cmd2)
                Dim ds2 As New DataSet

                da2.Fill(ds2, "User")
                Dim dt2 As DataTable = ds2.Tables("User")
                If dt2.Rows.Count = 0 Then
                    RichTextBox1.Text = RichTextBox1.Text & vbNewLine & "No Price Files Returned.  Moving to next date..."
                    RichTextBox1.SelectionStart = RichTextBox1.TextLength
                    Pause(0.1)
                    GoTo Line1
                Else
                    
                End If
                

                Dim path As String
                path = "C:\_PriceFiles"

                If (System.IO.Directory.Exists(path)) Then
                    System.IO.Directory.Delete(path, True)
                End If

                Dim AccessConn As New System.Data.OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & My.Settings.dbloc & "';Jet OLEDB:Database Password='" & My.Settings.dbpass & "';")

                AccessConn.Open()
                RichTextBox1.Text = RichTextBox1.Text & vbNewLine & "Creating APX File Nammed '" & datefile & ".pri'..."
                RichTextBox1.SelectionStart = RichTextBox1.TextLength
                Pause(0.01)
                Dim AccessCommand As New OleDb.OleDbCommand("SELECT SecType, Symbol, Price, Message, Source INTO [Text;HDR=NO;DATABASE=" & path & "].[" & datefile & ".csv] FROM mdb_ETFPrice_TempLoad", AccessConn)
                System.IO.Directory.CreateDirectory(path)
                AccessCommand.ExecuteNonQuery()
                AccessConn.Close()

                Dim myFiles As String()

                myFiles = IO.Directory.GetFiles(path, "*.csv")
                Pause(0.5)
                Dim newFilePath As String

                For Each filepath As String In myFiles

                    newFilePath = filepath.Replace(".csv", ".pri")

                    System.IO.File.Move(filepath, newFilePath)

                Next

                Dim FileToDelete As String

                FileToDelete = "\\aamapxapps01\apx$\imp\" & datefile & ".pri"
                Pause(0.5)

                If System.IO.File.Exists(FileToDelete) = True Then

                    System.IO.File.Delete(FileToDelete)
                    'MsgBox("File Deleted")
                End If
                RichTextBox1.Text = RichTextBox1.Text & vbNewLine & "File Created..."
                RichTextBox1.SelectionStart = RichTextBox1.TextLength
                RichTextBox1.Text = RichTextBox1.Text & vbNewLine & "Moving File to APX Server..."
                RichTextBox1.SelectionStart = RichTextBox1.TextLength
                Pause(0.01)
                My.Computer.FileSystem.CopyFile("C:\_PriceFiles\" & datefile & ".pri", "\\aamapxapps01\apx$\imp\" & datefile & ".pri")
                'RichTextBox1.Text = RichTextBox1.Text & vbNewLine & "MOVE DENIED - CODE BLOCKED..."
                RichTextBox1.Text = RichTextBox1.Text & vbNewLine & "File moved to APX Server..."
                Pause(0.5)
                '***********************************User Update***********************************
                RichTextBox1.Text = RichTextBox1.Text & vbNewLine & "Creating Import Script..."
                Pause(0.1)
                '***********************************User Update***********************************
                '////////////////////////////////////////////////////////////////////////////////////////////
                '***********************************CREATE 62 Bit BAT File***********************************
                RichTextBox2.Text = "c:" & vbNewLine & "cd C:\Program Files (x86)\Advent\ApxClient\4.0" & vbNewLine & "AdvScriptRunner Imex -i -s\\aamapxapps01\APX$\imp\ -Aua -f\\aamapxapps01\APX$\imp\" & datefile & ".pri -tcsv4 -u -N_268"
                RichTextBox1.Text = RichTextBox1.Text & vbNewLine & "64 Bit Script Created..."
                Pause(0.1)

                RichTextBox1.Text = RichTextBox1.Text & vbNewLine & "Creating BAT File..."
                Pause(0.1)

                RichTextBox2.SaveFile("\\aamapxapps01\apx$\automation\_AutomatedBATFiles\" & datefile & "64.bat", _
                    RichTextBoxStreamType.PlainText)

                RichTextBox1.Text = RichTextBox1.Text & vbNewLine & "64 Bit BAT file created..."
                Pause(0.5)
                '***********************************CREATE 64 Bit BAT File***********************************
                '////////////////////////////////////////////////////////////////////////////////////////////
                '***********************************CREATE 32 Bit BAT File***********************************
                RichTextBox2.Text = "c:" & vbNewLine & "cd C:\Program Files\Advent\ApxClient\4.0" & vbNewLine & "AdvScriptRunner Imex -i -s\\aamapxapps01\APX$\imp\ -Aua -f\\aamapxapps01\APX$\imp\" & datefile & ".pri -tcsv4 -u -N_268"
                RichTextBox1.Text = RichTextBox1.Text & vbNewLine & "32 Bit Script Created..."
                Pause(0.1)

                RichTextBox1.Text = RichTextBox1.Text & vbNewLine & "Creating 64 Bit BAT File..."
                Pause(0.1)

                RichTextBox2.SaveFile("\\aamapxapps01\apx$\automation\_AutomatedBATFiles\" & datefile & ".bat", _
                    RichTextBoxStreamType.PlainText)

                RichTextBox1.Text = RichTextBox1.Text & vbNewLine & "32 Bit BAT file created..."
                Pause(0.5)
                '***********************************CREATE 32 Bit BAT File***********************************
                '////////////////////////////////////////////////////////////////////////////////////////////
                '***********************************Send BAT files to APX***********************************
                RichTextBox1.Text = RichTextBox1.Text & vbNewLine & "Sending BAT files to RepRunner for Processing..."
                Pause(0.1)
                System.Diagnostics.Process.Start("\\aamapxapps01\apx$\Automation\_AutomatedBATFiles\" & datefile & ".BAT")
                System.Diagnostics.Process.Start("\\aamapxapps01\apx$\Automation\_AutomatedBATFiles\" & datefile & "64.BAT")
                RichTextBox1.Text = RichTextBox1.Text & vbNewLine & "Sent files to RepRunner for Processing..."
                Pause(0.5)

Line1:
                'Update Working = False Done = True where ID = ID
                sqlstr3 = "Update mdb_ETFPricing_DateQueue SET Working = False, Done = True WHERE ID = " & fileid

                command3 = New OleDb.OleDbCommand(sqlstr3, Mycn)
                command3.ExecuteNonQuery()
                Pause(0.01)
                'Update Working = False Done = True where ID = ID
                sqlstr4 = "DELETE * FROM mdb_ETFPrice_TempLoad"

                command4 = New OleDb.OleDbCommand(sqlstr4, Mycn)
                command4.ExecuteNonQuery()
                Pause(0.01)
                Call LoadDateQueue()
                Pause(0.01)
                RichTextBox1.Text = RichTextBox1.Text & vbNewLine & "Waiting on Next Command..."
                RichTextBox1.SelectionStart = RichTextBox1.TextLength
                Pause(0.01)
            Catch ex As Exception
                MsgBox(ex.Message, MsgBoxStyle.Critical, "ERROR")
            End Try






            'Update working = false and done = true

            'call loaddatequeue

        Loop Until DataGridView1.RowCount = 0

        Mycn.Close()

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

    Public Sub LoadRejectQueue()
        Try
            Dim conn As New OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & My.Settings.dbloc & "';Jet OLEDB:Database Password='" & My.Settings.dbpass & "';")

            Dim strSQL As String

            strSQL = "SELECT ID, NewSecType As [Sec Type], NewSymbol As [Security Symbol], OldSymbol As [Symbol Needed], DateNeeded As [Date Needed] FROM mdb_ETFPrice_Rejects"

            Dim da As New OleDb.OleDbDataAdapter(strSQL, conn)
            Dim ds As New DataSet
            da.Fill(ds, "Users")

            With DataGridView2
                .DataSource = ds.Tables("Users")
                .Columns(0).Visible = False
            End With

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "ERROR")
        End Try
    End Sub

    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        Call LoadRateCheck()
    End Sub

    Private Sub Button5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button5.Click
        Dim child As New ETF_PriceImport_SecData
        child.MdiParent = Home
        child.Show()
        child.ID.Text = "NEW"
    End Sub

    Private Sub Button6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button6.Click
        Dim child As New ETF_PriceImport_EditImportSecurity
        child.MdiParent = Home
        child.Show()
    End Sub

    Private Sub Button7_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button7.Click
        Dim ir As MsgBoxResult
        ir = MsgBox("Auto mapping the securities will delete the current records and replace them with securities entered in the Security Data screen.  Are you sure you want to continue?", MsgBoxStyle.YesNo + MsgBoxStyle.Information, "Confirm Auto Map")
        If ir = MsgBoxResult.Yes Then
            Call ClearSymbolMappingTable()
        Else
            'do nothing
        End If
    End Sub

    Public Sub ClearSymbolMappingTable()
        Dim Mycn As OleDb.OleDbConnection
        Dim Command As OleDb.OleDbCommand
        Dim SQLstr As String
        'Try
        Mycn = New OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & My.Settings.dbloc & "';Jet OLEDB:Database Password='" & My.Settings.dbpass & "';")
        Mycn.Open()

        SQLstr = "DELETE * FROM mdb_SymbolMapping"

        Command = New OleDb.OleDbCommand(SQLstr, Mycn)
        Command.ExecuteNonQuery()

        Mycn.Close()

        Call AutoMapSymbols()

        'Catch ex As Exception
        'MsgBox(ex.Message, MsgBoxStyle.Critical, "ERROR")
        'End Try
    End Sub

    Public Sub AutoMapSymbols()
        Dim Mycn As OleDb.OleDbConnection
        Dim Command As OleDb.OleDbCommand
        Dim SQLstr As String
        'Try
        Mycn = New OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & My.Settings.dbloc & "';Jet OLEDB:Database Password='" & My.Settings.dbpass & "';")
        Mycn.Open()

        SQLstr = "INSERT INTO mdb_SymbolMapping(APXSecType, APXSymbol, NewAPXSecType, NewAPXSymbol, Active, QntyMultiplier)" & _
        " SELECT (dbo_AdvSecurity.SecTypeBaseCode & dbo_AdvSecurity.PrincipalCurrencyCode), dbo_AdvSecurity.SecurityID, mdb_ETFPrice_SecurityData.NewSecType, dbo_AdvSecurity_1.SecurityID,-1,100" & _
        " FROM (mdb_ETFPrice_SecurityData INNER JOIN dbo_AdvSecurity ON mdb_ETFPrice_SecurityData.APXSecurityID = dbo_AdvSecurity.SecurityID) INNER JOIN dbo_AdvSecurity AS dbo_AdvSecurity_1 ON mdb_ETFPrice_SecurityData.NewSymbol = dbo_AdvSecurity_1.Symbol;"
        Command = New OleDb.OleDbCommand(SQLstr, Mycn)
        Command.ExecuteNonQuery()

        Mycn.Close()

        'Catch ex As Exception
        'MsgBox(ex.Message, MsgBoxStyle.Critical, "ERROR")
        'End Try
    End Sub

    Private Sub Button8_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button8.Click
        RichTextBox4.Text = "Loading Dates to get Pricing..."
        'Control.CheckForIllegalCrossThreadCalls = False
        'qdate = New System.Threading.Thread(AddressOf QueueDates)
        'qdate.Start()
        Call QueueDates1()
    End Sub

    Public Sub QueueDates1()
        Dim dte As Date
        Dim dte1 As String
        'Dim dte3 As String
        dte = DateTimePicker4.Text
        dte1 = Format(dte, "yyyy-MM-dd")
        'dte3 = dte1.ToString
        Dim Mycn As OleDb.OleDbConnection
        Dim Command As OleDb.OleDbCommand
        Dim SQLstr As String
        Try
            Mycn = New OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & My.Settings.dbloc & "';Jet OLEDB:Database Password='" & My.Settings.dbpass & "';")
            Mycn.Open()

            SQLstr = "INSERT INTO mdb_ETFPricing_DateQueue(QueueDate, PostDate)" & _
            " VALUES('" & dte1 & "', #" & dte & "#)"

            Command = New OleDb.OleDbCommand(SQLstr, Mycn)
            Command.ExecuteNonQuery()

            Do
                dte = DateAdd(DateInterval.Day, 1, dte)
                dte1 = Format(dte, "yyyy-MM-dd")
                'dte3 = dte1.ToString
                SQLstr = "INSERT INTO mdb_ETFPricing_DateQueue(QueueDate, PostDate)" & _
                " VALUES('" & dte1 & "', #" & dte & "#)"

                Command = New OleDb.OleDbCommand(SQLstr, Mycn)
                Command.ExecuteNonQuery()

            Loop Until dte = DateTimePicker3.Text
            RichTextBox4.Text = RichTextBox4.Text & vbNewLine & "Dates Loaded..."
            Mycn.Close()

            Call LoadDateQueue1()

            Call LoadTempPriceFiles1()

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "ERROR")
        End Try

    End Sub

    Public Sub LoadDateQueue1()
        Try
            Dim conn As New OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & My.Settings.dbloc & "';Jet OLEDB:Database Password='" & My.Settings.dbpass & "';")

            Dim strSQL As String

            strSQL = "SELECT mdb_ETFPricing_DateQueue.ID, mdb_ETFPricing_DateQueue.QueueDate As [Date], mdb_ETFPricing_DateQueue.Working" & _
            " FROM(mdb_ETFPricing_DateQueue)" & _
            " WHERE(((mdb_ETFPricing_DateQueue.Done) = False))" & _
            " GROUP BY mdb_ETFPricing_DateQueue.ID, mdb_ETFPricing_DateQueue.QueueDate, mdb_ETFPricing_DateQueue.Working"

            Dim da As New OleDb.OleDbDataAdapter(strSQL, conn)
            Dim ds As New DataSet
            da.Fill(ds, "Users")

            With DataGridView5
                .DataSource = ds.Tables("Users")
                .Columns(0).Visible = False
            End With

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "ERROR")
        End Try
    End Sub

    Public Sub LoadTempPriceFiles1()
        Dim Mycn As OleDb.OleDbConnection
        Dim Command As OleDb.OleDbCommand
        Dim SQLstr As String

        Dim command2 As OleDb.OleDbCommand
        Dim sqlstr2 As String

        Dim command3 As OleDb.OleDbCommand
        Dim sqlstr3 As String

        Dim command4 As OleDb.OleDbCommand
        Dim sqlstr4 As String

        Dim command5 As OleDb.OleDbCommand
        Dim sqlstr5 As String

        Dim command6 As OleDb.OleDbCommand
        Dim sqlstr6 As String

        Dim command7 As OleDb.OleDbCommand
        Dim sqlstr7 As String

        Dim command8 As OleDb.OleDbCommand
        Dim sqlstr8 As String

        Mycn = New OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & My.Settings.dbloc & "';Jet OLEDB:Database Password='" & My.Settings.dbpass & "';")
        Mycn.Open()

        'Update Working = False Done = True where ID = ID
        sqlstr6 = "DELETE * FROM mdb_ETFPrice_Rejects"

        command6 = New OleDb.OleDbCommand(sqlstr6, Mycn)
        command6.ExecuteNonQuery()
        Do
            'Select top 1 from mdb_ETFPricing_DateQueue WHERE Done = False
            Try
                Dim conn As New OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & My.Settings.dbloc & "';Jet OLEDB:Database Password='" & My.Settings.dbpass & "';")

                Dim strSQL As String = "SELECT Top 1 * FROM mdb_ETFPricing_DateQueue WHERE Done = False"
                Dim queryString As String = String.Format(strSQL)
                Dim cmd As New OleDb.OleDbCommand(queryString, conn)
                Dim da As New OleDb.OleDbDataAdapter(cmd)
                Dim ds As New DataSet

                da.Fill(ds, "User")
                Dim dt As DataTable = ds.Tables("User")

                Dim row1 As DataRow = dt.Rows(0)

                'Pull ID, PostDate, QueueDate
                Dim fileid As Integer = (row1("ID"))
                Dim postdate As Date = (row1("PostDate"))
                Dim queuedate As String = (row1("QueueDate"))

                Dim datefile As String = Format(postdate, "MMddyy").ToString

                RichTextBox4.Text = RichTextBox4.Text & vbNewLine & "Finding Prices for " & postdate & "..."
                RichTextBox4.SelectionStart = RichTextBox4.TextLength
                Pause(0.01)
                'Update Working = True where ID = ID
                SQLstr = "Update mdb_ETFPricing_DateQueue SET Working = True WHERE ID = " & fileid

                Command = New OleDb.OleDbCommand(SQLstr, Mycn)
                Command.ExecuteNonQuery()
                Pause(0.01)
                'Call LoadDateQueue()

                'Insert Into mdb_ETFPrice_TempLoad query
                sqlstr2 = "INSERT INTO mdb_ETFPrice_TempLoad(FileDate, SecType, Symbol, Price, Source)" & _
                "SELECT mdb_ETFPricing_DateQueue.PostDate, (dbo_AdvSecurity.SecTypeBaseCode & dbo_AdvSecurity.PrincipalCurrencyCode) AS Expr1, dbo_AdvSecurity.Symbol, mdb_ETFPricing_PriceImport.Price, 1" & _
                " FROM ((mdb_ETFPricing_PriceImport INNER JOIN mdb_ETFPricing_ApprovedList ON mdb_ETFPricing_PriceImport.Symbol = mdb_ETFPricing_ApprovedList.APXSymbol) INNER JOIN dbo_AdvSecurity ON mdb_ETFPricing_ApprovedList.APXSymbol = dbo_AdvSecurity.Symbol) INNER JOIN mdb_ETFPricing_DateQueue ON mdb_ETFPricing_PriceImport.CloseDate = mdb_ETFPricing_DateQueue.PostDate" & _
                " WHERE(((mdb_ETFPricing_ApprovedList.Current) = True) And ((mdb_ETFPricing_DateQueue.Working) = True) And ((mdb_ETFPricing_DateQueue.Done) = False))" & _
                " GROUP BY mdb_ETFPricing_DateQueue.PostDate, dbo_AdvSecurity.Symbol, mdb_ETFPricing_PriceImport.Price, dbo_AdvSecurity.SecTypeBaseCode, dbo_AdvSecurity.PrincipalCurrencyCode;"

                command2 = New OleDb.OleDbCommand(sqlstr2, Mycn)
                command2.ExecuteNonQuery()
                RichTextBox4.Text = RichTextBox4.Text & vbNewLine & "Prices Loaded..."
                RichTextBox4.SelectionStart = RichTextBox4.TextLength
                Pause(0.5)

                sqlstr7 = "INSERT INTO mdb_ETFPrice_TempLoad(FileDate, SecType, Symbol, Price, Source)" & _
                "SELECT mdb_ETFPricing_DateQueue.PostDate, mdb_SymbolMapping.NewAPXSecType, dbo_AdvSecurity_1.Symbol, mdb_ETFPricing_PriceImport.Price, 1" & _
                " FROM (mdb_SymbolMapping INNER JOIN dbo_AdvSecurity AS dbo_AdvSecurity_1 ON mdb_SymbolMapping.NewAPXSymbol = dbo_AdvSecurity_1.SecurityID) INNER JOIN (((mdb_ETFPricing_PriceImport INNER JOIN mdb_ETFPricing_ApprovedList ON mdb_ETFPricing_PriceImport.Symbol = mdb_ETFPricing_ApprovedList.APXSymbol) INNER JOIN dbo_AdvSecurity ON mdb_ETFPricing_ApprovedList.APXSymbol = dbo_AdvSecurity.Symbol) INNER JOIN mdb_ETFPricing_DateQueue ON mdb_ETFPricing_PriceImport.CloseDate = mdb_ETFPricing_DateQueue.PostDate) ON mdb_SymbolMapping.APXSymbol = dbo_AdvSecurity.SecurityID" & _
                " WHERE(((mdb_ETFPricing_ApprovedList.Current) = True) And ((mdb_ETFPricing_DateQueue.Working) = True) And ((mdb_ETFPricing_DateQueue.Done) = False))" & _
                " GROUP BY mdb_ETFPricing_DateQueue.PostDate, mdb_SymbolMapping.NewAPXSecType, dbo_AdvSecurity_1.Symbol, mdb_ETFPricing_PriceImport.Price;"

                command7 = New OleDb.OleDbCommand(sqlstr7, Mycn)
                command7.ExecuteNonQuery()

                RichTextBox4.Text = RichTextBox4.Text & vbNewLine & "Prices Loaded..."
                RichTextBox4.SelectionStart = RichTextBox4.TextLength
                Pause(0.5)

                RichTextBox4.Text = RichTextBox4.Text & vbNewLine & "Looking for prices not loaded..."
                RichTextBox4.SelectionStart = RichTextBox4.TextLength

                'Look for prices not loaded due to missing data
                sqlstr5 = "INSERT INTO mdb_ETFPrice_Rejects(NewSecType, NewSymbol, OldSymbol, DateNeeded)" & _
                "SELECT mdb_SymbolMapping.NewAPXSecType, dbo_AdvSecurity.Symbol, dbo_AdvSecurity_1.Symbol, #" & postdate & "#" & _
                " FROM (mdb_SymbolMapping INNER JOIN dbo_AdvSecurity ON mdb_SymbolMapping.NewAPXSymbol = dbo_AdvSecurity.SecurityID) INNER JOIN dbo_AdvSecurity AS dbo_AdvSecurity_1 ON mdb_SymbolMapping.APXSymbol = dbo_AdvSecurity_1.SecurityID" & _
                " WHERE (((dbo_AdvSecurity.[Symbol]) Not In (Select Symbol FROM mdb_ETFPrice_TempLoad)));"


                command5 = New OleDb.OleDbCommand(sqlstr5, Mycn)
                command5.ExecuteNonQuery()

                sqlstr8 = "INSERT INTO mdb_ETFPrice_Rejects(NewSecType, NewSymbol, DateNeeded)" & _
                "SELECT mdb_SymbolMapping.APXSecType, dbo_AdvSecurity.Symbol, #" & postdate & "#" & _
                " FROM mdb_SymbolMapping INNER JOIN dbo_AdvSecurity ON mdb_SymbolMapping.APXSymbol = dbo_AdvSecurity.SecurityID" & _
                " WHERE (((dbo_AdvSecurity.Symbol) Not In (Select Symbol FROM mdb_ETFPrice_TempLoad)));"


                command8 = New OleDb.OleDbCommand(sqlstr5, Mycn)
                command8.ExecuteNonQuery()

                Dim sqlstring As String
                sqlstring = "SELECT * FROM mdb_ETFPrice_Rejects WHERE DateNeeded = #" & postdate & "#"

                Dim queryString1 As String = String.Format(sqlstring)
                Dim cmd1 As New OleDb.OleDbCommand(queryString1, Mycn)
                Dim da1 As New OleDb.OleDbDataAdapter(cmd1)
                Dim ds1 As New DataSet

                da1.Fill(ds1, "User")
                Dim dt1 As DataTable = ds1.Tables("User")
                If dt1.Rows.Count > 0 Then
                    RichTextBox4.Text = RichTextBox4.Text & vbNewLine & "Found " & dt1.Rows.Count.ToString & " Rejected Prices..."
                    RichTextBox4.SelectionStart = RichTextBox4.TextLength
                    Call LoadRejectQueue1()
                    Pause(0.1)
                Else
                    RichTextBox4.Text = RichTextBox4.Text & vbNewLine & "No Rejected Prices Found..."
                    RichTextBox4.SelectionStart = RichTextBox4.TextLength
                    Pause(0.1)
                End If

                Dim sqlstring1 As String
                sqlstring1 = "SELECT * FROM mdb_ETFPrice_TempLoad"

                Dim queryString2 As String = String.Format(sqlstring1)
                Dim cmd2 As New OleDb.OleDbCommand(queryString2, Mycn)
                Dim da2 As New OleDb.OleDbDataAdapter(cmd2)
                Dim ds2 As New DataSet

                da2.Fill(ds2, "User")
                Dim dt2 As DataTable = ds2.Tables("User")
                If dt2.Rows.Count = 0 Then
                    RichTextBox4.Text = RichTextBox4.Text & vbNewLine & "No Price Files Returned.  Moving to next date..."
                    RichTextBox4.SelectionStart = RichTextBox4.TextLength
                    Pause(0.1)
                    GoTo Line1
                Else

                End If


                Dim path As String
                path = "C:\_PriceFiles"

                If (System.IO.Directory.Exists(path)) Then
                    System.IO.Directory.Delete(path, True)
                End If

                Dim AccessConn As New System.Data.OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & My.Settings.dbloc & "';Jet OLEDB:Database Password='" & My.Settings.dbpass & "';")

                AccessConn.Open()
                RichTextBox4.Text = RichTextBox4.Text & vbNewLine & "Creating APX File Nammed '" & datefile & ".pri'..."
                RichTextBox4.SelectionStart = RichTextBox4.TextLength
                Pause(0.01)
                Dim AccessCommand As New OleDb.OleDbCommand("SELECT SecType, Symbol, Price, Message, Source INTO [Text;HDR=NO;DATABASE=" & path & "].[" & datefile & ".csv] FROM mdb_ETFPrice_TempLoad", AccessConn)
                System.IO.Directory.CreateDirectory(path)
                AccessCommand.ExecuteNonQuery()
                AccessConn.Close()

                Dim myFiles As String()

                myFiles = IO.Directory.GetFiles(path, "*.csv")
                Pause(0.5)
                Dim newFilePath As String

                For Each filepath As String In myFiles

                    newFilePath = filepath.Replace(".csv", ".pri")

                    System.IO.File.Move(filepath, newFilePath)

                Next

                Dim FileToDelete As String

                FileToDelete = "\\aamapxapps01\apx$\imp\" & datefile & ".pri"
                Pause(0.5)

                If System.IO.File.Exists(FileToDelete) = True Then

                    System.IO.File.Delete(FileToDelete)
                    'MsgBox("File Deleted")
                End If
                RichTextBox4.Text = RichTextBox4.Text & vbNewLine & "File Created..."
                RichTextBox4.SelectionStart = RichTextBox4.TextLength
                RichTextBox4.Text = RichTextBox4.Text & vbNewLine & "Moving File to APX Server..."
                RichTextBox4.SelectionStart = RichTextBox4.TextLength
                Pause(0.01)
                My.Computer.FileSystem.CopyFile("C:\_PriceFiles\" & datefile & ".pri", "\\aamapxapps01\apx$\imp\" & datefile & ".pri")
                'RichTextBox1.Text = RichTextBox1.Text & vbNewLine & "MOVE DENIED - CODE BLOCKED..."
                RichTextBox4.Text = RichTextBox4.Text & vbNewLine & "File moved to APX Server..."
                Pause(0.5)
                '***********************************User Update***********************************
                RichTextBox4.Text = RichTextBox4.Text & vbNewLine & "Creating Import Script..."
                Pause(0.1)
                '***********************************User Update***********************************
                '////////////////////////////////////////////////////////////////////////////////////////////
                '***********************************CREATE 62 Bit BAT File***********************************
                RichTextBox3.Text = "c:" & vbNewLine & "cd C:\Program Files (x86)\Advent\ApxClient\4.0" & vbNewLine & "AdvScriptRunner Imex -i -s\\aamapxapps01\APX$\imp\ -Aua -f\\aamapxapps01\APX$\imp\" & datefile & ".pri -tcsv4 -u -N_268"
                RichTextBox4.Text = RichTextBox4.Text & vbNewLine & "64 Bit Script Created..."
                Pause(0.1)

                RichTextBox4.Text = RichTextBox4.Text & vbNewLine & "Creating BAT File..."
                Pause(0.1)

                RichTextBox3.SaveFile("\\aamapxapps01\apx$\automation\_AutomatedBATFiles\" & datefile & "64.bat", _
                    RichTextBoxStreamType.PlainText)

                RichTextBox4.Text = RichTextBox4.Text & vbNewLine & "64 Bit BAT file created..."
                Pause(0.5)
                '***********************************CREATE 64 Bit BAT File***********************************
                '////////////////////////////////////////////////////////////////////////////////////////////
                '***********************************CREATE 32 Bit BAT File***********************************
                RichTextBox3.Text = "c:" & vbNewLine & "cd C:\Program Files\Advent\ApxClient\4.0" & vbNewLine & "AdvScriptRunner Imex -i -s\\aamapxapps01\APX$\imp\ -Aua -f\\aamapxapps01\APX$\imp\" & datefile & ".pri -tcsv4 -u -N_268"
                RichTextBox4.Text = RichTextBox4.Text & vbNewLine & "32 Bit Script Created..."
                Pause(0.1)

                RichTextBox4.Text = RichTextBox4.Text & vbNewLine & "Creating 64 Bit BAT File..."
                Pause(0.1)

                RichTextBox3.SaveFile("\\aamapxapps01\apx$\automation\_AutomatedBATFiles\" & datefile & ".bat", _
                    RichTextBoxStreamType.PlainText)

                RichTextBox4.Text = RichTextBox4.Text & vbNewLine & "32 Bit BAT file created..."
                Pause(0.5)
                '***********************************CREATE 32 Bit BAT File***********************************
                '////////////////////////////////////////////////////////////////////////////////////////////
                '***********************************Send BAT files to APX***********************************
                RichTextBox4.Text = RichTextBox4.Text & vbNewLine & "Sending BAT files to RepRunner for Processing..."
                Pause(0.1)
                System.Diagnostics.Process.Start("\\aamapxapps01\apx$\Automation\_AutomatedBATFiles\" & datefile & ".BAT")
                System.Diagnostics.Process.Start("\\aamapxapps01\apx$\Automation\_AutomatedBATFiles\" & datefile & "64.BAT")
                RichTextBox4.Text = RichTextBox4.Text & vbNewLine & "Sent files to RepRunner for Processing..."
                Pause(0.5)

Line1:
                'Update Working = False Done = True where ID = ID
                sqlstr3 = "Update mdb_ETFPricing_DateQueue SET Working = False, Done = True WHERE ID = " & fileid

                command3 = New OleDb.OleDbCommand(sqlstr3, Mycn)
                command3.ExecuteNonQuery()
                Pause(0.01)
                'Update Working = False Done = True where ID = ID
                sqlstr4 = "DELETE * FROM mdb_ETFPrice_TempLoad"

                command4 = New OleDb.OleDbCommand(sqlstr4, Mycn)
                command4.ExecuteNonQuery()
                Pause(0.01)
                Call LoadDateQueue1()
                Pause(0.01)
                RichTextBox4.Text = RichTextBox4.Text & vbNewLine & "Waiting on Next Command..."
                RichTextBox4.SelectionStart = RichTextBox4.TextLength
                Pause(0.01)
            Catch ex As Exception
                MsgBox(ex.Message, MsgBoxStyle.Critical, "ERROR")
            End Try






            'Update working = false and done = true

            'call loaddatequeue

        Loop Until DataGridView5.RowCount = 0

        Mycn.Close()

    End Sub


    Public Sub LoadRejectQueue1()
        Try
            Dim conn As New OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & My.Settings.dbloc & "';Jet OLEDB:Database Password='" & My.Settings.dbpass & "';")

            Dim strSQL As String

            strSQL = "SELECT ID, NewSecType As [Sec Type], NewSymbol As [Security Symbol], OldSymbol As [Symbol Needed], DateNeeded As [Date Needed] FROM mdb_ETFPrice_Rejects"

            Dim da As New OleDb.OleDbDataAdapter(strSQL, conn)
            Dim ds As New DataSet
            da.Fill(ds, "Users")

            With DataGridView4
                .DataSource = ds.Tables("Users")
                .Columns(0).Visible = False
            End With

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "ERROR")
        End Try
    End Sub

    Private Sub Button9_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button9.Click
        Dim child As New ETF_PriceImport_APXPortfolios
        child.MdiParent = Home
        child.Show()
    End Sub

    Private Sub Button10_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Dim child As New ETFPriceImport_Model
        child.MdiParent = Home
        child.ID.Text = "NEW"
        child.Show()

    End Sub

    Private Sub Button12_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button12.Click
        Dim child As New ETFPriceImport_Model
        child.MdiParent = Home
        child.ID.Text = "NEW"
        child.Show()
    End Sub

    Private Sub Button11_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button11.Click
        Call LoadModels()
    End Sub

    Public Sub LoadModels()
        Try
            Dim conn As New OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & My.Settings.dbloc & "';Jet OLEDB:Database Password='" & My.Settings.dbpass & "';")

            Dim SqlString As String


            SqlString = "SELECT mdb_ETFPrice_ModelMain.ID, mdb_ETFPrice_ModelMain.ETFPortfolioID, mdb_ETFPrice_ModelMain.TAMPortfolioID, mdb_ETFPrice_ModelMain.ModelName As [Model Name], mdb_ETFPrice_ModelMain.CashBal As [Model Cash], mdb_ETFPrice_ModelMain.InceptionDate As [Inception Date], (SUM(CurrentYield * Weight) / Sum(Weight)) As [Current Yield], (SUM(YTW * Weight) / Sum(Weight)) As [Yield to Worst], (Format((SUM(Duration * Weight) / Sum(Weight)), '#.##')) As [Model Duration], (Format((SUM(AvgLife * Weight) / Sum(Weight)), '#.##')) As [Avg Life]" & _
            " FROM mdb_ETFPrice_ModelMain INNER JOIN mdb_ETFPricing_ModelHoldings ON mdb_ETFPrice_ModelMain.ID = mdb_ETFPricing_ModelHoldings.ModelID" & _
            " WHERE(((mdb_ETFPrice_ModelMain.Active) = True))" & _
            " GROUP BY mdb_ETFPrice_ModelMain.ID, mdb_ETFPrice_ModelMain.ETFPortfolioID, mdb_ETFPrice_ModelMain.TAMPortfolioID, mdb_ETFPrice_ModelMain.ModelName, mdb_ETFPrice_ModelMain.CashBal, mdb_ETFPrice_ModelMain.InceptionDate;"


            Dim da As New OleDb.OleDbDataAdapter(SqlString, conn)
            Dim ds As New DataSet
            da.Fill(ds, "Users")

            With DataGridView6
                .DataSource = ds.Tables("Users")
                .Columns(0).Visible = False
                .Columns(1).Visible = False
                .Columns(2).Visible = False
                .Columns(4).DefaultCellStyle.Format = "c"
                .Columns(6).DefaultCellStyle.Format = "p"
                .Columns(7).DefaultCellStyle.Format = "p"
            End With

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "ERROR")
        End Try
    End Sub

    Private Sub DataGridView6_CellContentClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView6.CellContentClick

    End Sub

    Private Sub DataGridView6_CellContentDoubleClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView6.CellContentDoubleClick
        If DataGridView6.RowCount = 0 Then

        Else
            Call LoadETFModel()
        End If
    End Sub

    Public Sub LoadETFModel()
        Dim child As New ETFPriceImport_Model
        child.MdiParent = Home

        child.ID.Text = DataGridView6.SelectedCells(0).Value

        child.Show()
        child.Cash.Text = DataGridView6.SelectedCells(4).Value
        child.ModelName.Text = DataGridView6.SelectedCells(3).Value
        child.ETFPortfolioID.SelectedValue = DataGridView6.SelectedCells(1).Value
        child.TAMPortfolioID.SelectedValue = DataGridView6.SelectedCells(2).Value

        Call child.LoadAfterInitialUpdate()
    End Sub
End Class