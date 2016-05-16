Imports System.IO
Public Class AMPFExcelExport

    Dim FindTrades As System.Threading.Thread

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        If DataGridView2.RowCount > 0 Then
            
            Dim wapp As Microsoft.Office.Interop.Excel.Application

            Dim wsheet As Microsoft.Office.Interop.Excel.Worksheet

            Dim wbook As Microsoft.Office.Interop.Excel.Workbook

            wapp = New Microsoft.Office.Interop.Excel.Application

            wapp.Visible = False

            wbook = wapp.Workbooks.Add()

            wsheet = wbook.ActiveSheet

            Try

                Dim iX As Integer

                Dim iY As Integer

                Dim iC As Integer

                Dim CC As Integer = DataGridView2.Columns.Count

                For iC = 0 To DataGridView2.Columns.Count - 1

                    wsheet.Cells(1, iC + 1).Value = DataGridView2.Columns(iC).HeaderText

                    wsheet.Cells(1, iC + 1).font.bold = True

                    wsheet.Cells(1, iC + 1).horizontalalignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter

                    wsheet.Rows(1).autofit()

                Next

                For iX = 0 To DataGridView2.Rows.Count - 1

                    For iY = 0 To DataGridView2.Columns.Count - 1

                        wsheet.Cells(iX + 2, iY + 1).value = DataGridView2(iY, iX).Value.ToString

                        wsheet.Cells(iX + 2, iY + 1).horizontalalignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignLeft

                    Next

                Next

                'wsheet.Columns(2).delete()
                wsheet.Columns(1).delete()
                wsheet.Columns(1).delete()
                wsheet.Rows(1).delete()

                'wsheet.Columns.Delete(1)
                'wsheet.Columns.Delete(2)

                'Dim style As Microsoft.Office.Interop.Excel.Style = wsheet.Application.ActiveWorkbook.Styles.Add("NewStyle")
                'style.Font.Bold = True
                'style.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Gray)

                wsheet.Range("A1", "K1").Insert()
                wsheet.Cells.Cells(1, 1) = "CLIENT" & vbNewLine & "ALLOCATION"
                wsheet.Cells.Cells(1, 1).horizontalalignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
                wsheet.Cells.Cells(1, 1).font.bold = True
                wsheet.Cells(1, 1).Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.LightGray)

                wsheet.Cells.Cells(1, 2) = "CONTRA" & vbNewLine & "DEALER"
                wsheet.Cells.Cells(1, 2).horizontalalignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
                wsheet.Cells.Cells(1, 2).font.bold = True
                wsheet.Cells(1, 2).Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.LightGray)

                wsheet.Cells.Cells(1, 3) = "TYPE" & vbNewLine & "(BUY/ SELL)"
                wsheet.Cells.Cells(1, 3).horizontalalignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
                wsheet.Cells.Cells(1, 3).font.bold = True
                wsheet.Cells(1, 3).Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.LightGray)

                wsheet.Cells.Cells(1, 4) = "QUANTITY"
                wsheet.Cells.Cells(1, 4).horizontalalignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
                wsheet.Cells.Cells(1, 4).font.bold = True
                wsheet.Cells(1, 4).Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.LightGray)

                wsheet.Cells.Cells(1, 5) = "SECURITY" & vbNewLine & "/CUSIP"
                wsheet.Cells.Cells(1, 5).horizontalalignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
                wsheet.Cells.Cells(1, 5).font.bold = True
                wsheet.Cells(1, 5).Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.LightGray)

                wsheet.Cells.Cells(1, 6) = "PRICE"
                wsheet.Cells.Cells(1, 6).horizontalalignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
                wsheet.Cells.Cells(1, 6).font.bold = True
                wsheet.Cells(1, 6).Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.LightGray)

                wsheet.Cells.Cells(1, 7) = "BETA" & vbNewLine & "INVENTORY"
                wsheet.Cells.Cells(1, 7).horizontalalignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
                wsheet.Cells.Cells(1, 7).font.bold = True
                wsheet.Cells(1, 7).Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.LightGray)

                wsheet.Cells.Cells(1, 8) = "TRADE" & vbNewLine & "DATE"
                wsheet.Cells.Cells(1, 8).horizontalalignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
                wsheet.Cells.Cells(1, 8).font.bold = True
                wsheet.Cells(1, 8).Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.LightGray)

                wsheet.Cells.Cells(1, 9) = "SETTLE" & vbNewLine & "DATE"
                wsheet.Cells.Cells(1, 9).horizontalalignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
                wsheet.Cells.Cells(1, 9).font.bold = True
                wsheet.Cells(1, 9).Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.LightGray)

                wsheet.Cells.Cells(1, 10) = "PRINCIPAL"
                wsheet.Cells.Cells(1, 10).horizontalalignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
                wsheet.Cells.Cells(1, 10).font.bold = True
                wsheet.Cells(1, 10).Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.LightGray)

                wsheet.Cells.Cells(1, 11) = "INTEREST"
                wsheet.Cells.Cells(1, 11).horizontalalignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
                wsheet.Cells.Cells(1, 11).font.bold = True
                wsheet.Cells(1, 11).Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.LightGray)


                wsheet.Range("A1", "K1").Insert()
                wsheet.Range(wsheet.Cells(1, 1), wsheet.Cells(1, 11)).Merge()
                wsheet.Cells.Cells(1, 1) = "STEP OUT REQUEST"
                wsheet.Cells.Cells(1, 1).horizontalalignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
                wsheet.Cells.Cells(1, 1).font.bold = True
                wsheet.Cells(1, 1).Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.LightGray)

                wsheet.Range("A1", "K1").Insert()
                wsheet.Range(wsheet.Cells(1, 1), wsheet.Cells(1, 11)).Merge()
                wsheet.Cells.Cells(1, 1) = "MONEY MANAGER"
                wsheet.Cells.Cells(1, 1).horizontalalignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
                wsheet.Cells.Cells(1, 1).font.bold = True
                wsheet.Cells(1, 1).Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.LightGray)

                wsheet.Rows(3).rowheight = 30
                wsheet.Columns(1).columnwidth = 11.57
                wsheet.Columns(2).columnwidth = 8.43
                wsheet.Columns(3).columnwidth = 10.43
                wsheet.Columns(4).columnwidth = 13.71
                wsheet.Columns(5).columnwidth = 12.0
                wsheet.Columns(6).columnwidth = 9
                wsheet.Columns(7).columnwidth = 10.71
                wsheet.Columns(8).columnwidth = 10
                wsheet.Columns(9).columnwidth = 10
                wsheet.Columns(10).columnwidth = 10
                wsheet.Columns(11).columnwidth = 9


                With wsheet.Range("A1:K30")
                    .Borders.Weight = 2
                End With

                With wsheet.PageSetup
                    .PrintArea = "$A$1:$K$30"
                    .PaperSize = Microsoft.Office.Interop.Excel.XlPaperSize.xlPaperA4
                    .FitToPagesWide = 1
                    .Zoom = False
                    .Orientation = Microsoft.Office.Interop.Excel.XlPageOrientation.xlLandscape
                End With

                'wapp.Visible = True
                Dim dte As String
                dte = Format(Now(), "MMddyyyy")

                Dim path As String
                path = "\\monfp01\data\AAM Only\_AMPTrades"

                If (System.IO.Directory.Exists(path)) Then
                    'System.IO.Directory.Delete(path, True)
                Else
                    System.IO.Directory.CreateDirectory(path)
                End If

                Dim fname As String
                fname = "AMP STEPOUT TEMPLETE " & dte & " FISA"

                Dim fullnme As String
                fullnme = path & "\" & fname

                wsheet.SaveAs(fullnme)
                'wapp.SaveWorkspace("C:\AMP STEPOUT TEMPLETE " & dte & " FISA")
                wbook.Close()
                wapp.Quit()

                'Catch ex As Exception

                'MsgBox(ex.Message)

                'End Try
                Dim ir As MsgBoxResult
                ir = MsgBox("Trade file sucessfully saved.  Would you like to email these trades now?", MsgBoxStyle.YesNo + MsgBoxStyle.Information, "Email File?")

                If ir = MsgBoxResult.Yes Then
                    Dim app As Microsoft.Office.Interop.Outlook.Application
                    Dim appNameSpace As Microsoft.Office.Interop.Outlook._NameSpace
                    Dim memo As Microsoft.Office.Interop.Outlook.MailItem
                    Dim bdy As String

                    app = New Microsoft.Office.Interop.Outlook.Application
                    appNameSpace = app.GetNamespace("MAPI")
                    appNameSpace.Logon(Nothing, Nothing, False, False)
                    memo = app.CreateItem(Microsoft.Office.Interop.Outlook.OlItemType.olMailItem)

                    memo.To = "SMAAdministration@aamlive.com"
                    'memo.To = "traders@envestnet.com"
                    'memo.CC = "SMAAdministration@aamlive.com"

                    memo.Subject = "Allocation for Ameriprise Street Side Allocation"

                    bdy = "Please see attached allocation for Ameriprise." & vbNewLine & vbNewLine & "Thank you," & vbNewLine & vbNewLine & "SMA Operations Team" & vbNewLine & "Advisors Asset Management, Inc." & vbNewLine & "1-866-259-2427" & vbNewLine & "SMAAdministration@aamlive.com"


                    memo.Body = bdy
                    memo.Attachments.Add(fullnme & ".xlsx")
                    memo.Send()

                Else

                End If

                ir = MsgBox("Trade file sucessfully saved at '" & fullnme & "'.  Would you like to mark these trades as uploaded?", MsgBoxStyle.YesNo + MsgBoxStyle.Information, "File Created!")

                If ir = MsgBoxResult.Yes Then
                    Dim Mycn As OleDb.OleDbConnection
                    Dim Command As OleDb.OleDbCommand
                    Dim Command1 As OleDb.OleDbCommand
                    Dim SQLstr As String
                    Try
                        Mycn = New OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & My.Settings.dbloc & "';Jet OLEDB:Database Password='" & My.Settings.dbpass & "';")
                        Mycn.Open()

                        SQLstr = "INSERT INTO AMP_tradetracker(MoxyOrderID, DateReady, ReadyBy, TradeUploaded, UploadFile)" & _
                        "SELECT MoxyOrderID,#" & Format(Now()) & "#,'" & Environ("USERNAME") & "',-1,'" & fullnme & "' FROM AMP_TradeReady"

                        Command = New OleDb.OleDbCommand(SQLstr, Mycn)
                        Command.ExecuteNonQuery()

                        SQLstr = "DELETE * FROM AMP_TradeReady"

                        Command1 = New OleDb.OleDbCommand(SQLstr, Mycn)
                        Command1.ExecuteNonQuery()

                        Mycn.Close()
                    Catch ex As Exception
                        MsgBox(ex.Message, MsgBoxStyle.Critical, "ERROR")
                    End Try
                Else

                End If


                wapp.UserControl = True

            Catch ex As Exception
                MsgBox(ex.Message, MsgBoxStyle.Critical, "ERROR")
            End Try
        Else

            MsgBox("Cannot process request.  Please check data.", MsgBoxStyle.Information, "Invalid Results.")

        End If


    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        Control.CheckForIllegalCrossThreadCalls = False
        FindTrades = New System.Threading.Thread(AddressOf LoadPendingTrades)
        FindTrades.Start()
    End Sub

    Public Sub LoadPendingTrades()
        Dim Mycn As OleDb.OleDbConnection
        Dim Command As OleDb.OleDbCommand
        Dim Command1 As OleDb.OleDbCommand
        Dim Command2 As OleDb.OleDbCommand
        'Dim Command3 As OleDb.OleDbCommand
        'Dim Command4 As OleDb.OleDbCommand
        'Dim Command5 As OleDb.OleDbCommand
        Dim SQLstr As String
        Try
            Mycn = New OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & My.Settings.dbloc & "';Jet OLEDB:Database Password='" & My.Settings.dbpass & "';")
            Dim ds1 As New DataSet
            Dim eds1 As New DataGridView
            Dim dv1 As New DataView
            Mycn.Open()

            'TextBox5.Text = "Cleaning Table..."
            SQLstr = "DELETE * FROM AMP_TradePending"
            Command = New OleDb.OleDbCommand(SQLstr, Mycn)
            Command.ExecuteNonQuery()

            'TextBox5.Text = "Searching Trades..."
            'SQLstr = "INSERT INTO AMP_TradePending(MoxyOrderID,ClientAllocation, ContraDealer, Type, Quantity, Security, Price, BetaInventory, TradeDate, SettleDate, Principal, Interest)" & _
            '" SELECT MxApp_vAllocation.OrderID, dbo_vQbRowDefPortfolio.PortfolioCode, 'FISA0443' AS Expr0, uma_tradetranslator_1.AMPCode, MxApp_vAllocation.AllocQty, AdvApp_vSecurity.Symbol, MxApp_vAllocation.AllocPrice, '8613' AS Expr2, Format([MxApp_vAllocation].[TradeDate],'mm/dd/yyyy') AS Expr3, Format([MxApp_vAllocation].[SettleDate],'mm/dd/yyyy') AS Expr4, MxApp_vAllocation.PrincipalAmt, IIf([MxApp_vAllocation].[AIAmt]<=0,0,Format([MxApp_vAllocation].[AIAmt],'#.##')) AS Expr5" & _
            '" FROM (((MxApp_vAllocation INNER JOIN MxApp_vSecurity ON MxApp_vAllocation.SecurityID = MxApp_vSecurity.SecurityID) INNER JOIN AdvApp_vSecurity ON MxApp_vSecurity.Symbol = AdvApp_vSecurity.Symbol) INNER JOIN dbo_vQbRowDefPortfolio ON MxApp_vAllocation.PortfolioCode = dbo_vQbRowDefPortfolio.PortfolioCode) INNER JOIN uma_tradetranslator AS uma_tradetranslator_1 ON MxApp_vAllocation.TransactionCode = uma_tradetranslator_1.APXCode" & _
            '" WHERE (((dbo_vQbRowDefPortfolio.PortfolioCode) Not In (SELECT PortfolioCodeReal FROM uma_portfolios_envestnet)) AND ((dbo_vQbRowDefPortfolio.ExternalFirm)='Ameriprise_Financial_Services'))" & _
            '" GROUP BY MxApp_vAllocation.OrderID, dbo_vQbRowDefPortfolio.PortfolioCode, 'FISA0443', uma_tradetranslator_1.AMPCode, MxApp_vAllocation.AllocQty, AdvApp_vSecurity.Symbol, MxApp_vAllocation.AllocPrice, '8613', Format([MxApp_vAllocation].[TradeDate],'mm/dd/yyyy'), Format([MxApp_vAllocation].[SettleDate],'mm/dd/yyyy'), MxApp_vAllocation.PrincipalAmt, IIf([MxApp_vAllocation].[AIAmt]<=0,0,Format([MxApp_vAllocation].[AIAmt],'#.##')), Format([MxApp_vAllocation].[TradeDate],'mm/dd/yyyy');"
            'Command1 = New OleDb.OleDbCommand(SQLstr, Mycn)
            'Command1.ExecuteNonQuery()

            SQLstr = "INSERT INTO AMP_TradePending(MoxyOrderID,ClientAllocation, ContraDealer, Type, Quantity, Security, Price, BetaInventory, TradeDate, SettleDate, Principal, Interest)" & _
            " SELECT MxApp_vAllocation.OrderID, dbo_vQbRowDefPortfolio.PortfolioCode, 'FISA0443' AS Expr0, uma_tradetranslator_1.AMPCode, MxApp_vAllocation.AllocQty, AdvApp_vSecurity.Symbol, MxApp_vAllocation.AllocPrice, '8613' AS Expr2, Format([MxApp_vAllocation].[TradeDate],'mm/dd/yyyy') AS Expr3, Format([MxApp_vAllocation].[SettleDate],'mm/dd/yyyy') AS Expr4, MxApp_vAllocation.PrincipalAmt, IIf([MxApp_vAllocation].[AIAmt]<=0,0,Format([MxApp_vAllocation].[AIAmt],'#.##')) AS Expr5" & _
            " FROM ((((MxApp_vAllocation INNER JOIN MxApp_vSecurity ON MxApp_vAllocation.SecurityID = MxApp_vSecurity.SecurityID) INNER JOIN AdvApp_vSecurity ON MxApp_vSecurity.Symbol = AdvApp_vSecurity.Symbol) INNER JOIN dbo_vQbRowDefPortfolio ON MxApp_vAllocation.PortfolioCode = dbo_vQbRowDefPortfolio.PortfolioCode) INNER JOIN uma_tradetranslator AS uma_tradetranslator_1 ON MxApp_vAllocation.TransactionCode = uma_tradetranslator_1.APXCode) INNER JOIN AMP_SecTypes ON MxApp_vSecurity.SecTypeCode = AMP_SecTypes.APXSecType" & _
            " WHERE (((dbo_vQbRowDefPortfolio.PortfolioCode) Not In (SELECT PortfolioCodeReal FROM uma_portfolios_envestnet)) AND ((dbo_vQbRowDefPortfolio.ExternalFirm)='Ameriprise_Financial_Services'))" & _
            " GROUP BY MxApp_vAllocation.OrderID, dbo_vQbRowDefPortfolio.PortfolioCode, 'FISA0443', uma_tradetranslator_1.AMPCode, MxApp_vAllocation.AllocQty, AdvApp_vSecurity.Symbol, MxApp_vAllocation.AllocPrice, '8613', Format([MxApp_vAllocation].[TradeDate],'mm/dd/yyyy'), Format([MxApp_vAllocation].[SettleDate],'mm/dd/yyyy'), MxApp_vAllocation.PrincipalAmt, IIf([MxApp_vAllocation].[AIAmt]<=0,0,Format([MxApp_vAllocation].[AIAmt],'#.##')), Format([MxApp_vAllocation].[TradeDate],'mm/dd/yyyy');"
            Command1 = New OleDb.OleDbCommand(SQLstr, Mycn)
            Command1.ExecuteNonQuery()

            'TextBox5.Text = "Removing Processed Trades..."
            SQLstr = "DELETE * FROM AMP_TradePending WHERE MoxyOrderID In (SELECT MoxyOrderID FROM AMP_tradetracker)"
            Command2 = New OleDb.OleDbCommand(SQLstr, Mycn)
            Command2.ExecuteNonQuery()

            'TextBox5.Text = "Removing Trades already on upload list..."
            SQLstr = "DELETE * FROM AMP_TradePending WHERE MoxyOrderID In (SELECT MoxyOrderID FROM AMP_TradeReady)"
            Command2 = New OleDb.OleDbCommand(SQLstr, Mycn)
            Command2.ExecuteNonQuery()

            'TextBox5.Text = "Finised loading data!"
            Mycn.Close()

            'MsgBox("done")
            'TextBox5.Text = "Drawing grid..."
            Call LoadPendingTradesGrid()

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "ERROR")
        End Try
    End Sub

    Public Sub LoadPendingTradesGrid()
        Try

            DataGridView1.Enabled = False

            Dim conn As New OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & My.Settings.dbloc & "';Jet OLEDB:Database Password='" & My.Settings.dbpass & "';")

            Dim strSQL As String = "SELECT ClientAllocation, ContraDealer, Type, Quantity, Security, Price, BetaInventory, TradeDate, SettleDate, Principal, Interest FROM AMP_TradePending"
            Dim da As New OleDb.OleDbDataAdapter(strSQL, conn)
            Dim ds As New DataSet
            da.Fill(ds, "Users")

            With DataGridView1
                .DataSource = ds.Tables("Users")
                '.Columns(0).Visible = False
            End With

            'Label10.Text = "Loaded " & DataGridView1.RowCount & " pending trades."

            DataGridView1.Enabled = True
            'DataGridView1.Visible = True
            'TextBox4.Visible = False
            'TextBox5.Visible = False
            'Timer2.Enabled = False

            DataGridView1.ScrollBars = ScrollBars.Both
            DataGridView1.Refresh()

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "ERROR")
        End Try
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Call ApproveAll()
    End Sub

    Public Sub ApproveAll()
        Dim ir As MsgBoxResult
        If DataGridView1.RowCount = "0" Then
            'do nothing
            MsgBox("Please load positions.", MsgBoxStyle.Information, "No Data")
        Else
            ir = MsgBox("You are about to Approve all pending trades.  Are you sure you want to continue?" & vbNewLine & vbNewLine & "NOTE: Trades without a status will not be moved to approved.", MsgBoxStyle.YesNo + MsgBoxStyle.Exclamation, "Confirm")

            If ir = MsgBoxResult.Yes Then
                If DataGridView1.RowCount = "0" Then
                    'do nothing
                Else
                    Dim Mycn As OleDb.OleDbConnection
                    Dim Command As OleDb.OleDbCommand
                    Dim SQLstr As String
                    Try
                        Mycn = New OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & My.Settings.dbloc & "';Jet OLEDB:Database Password='" & My.Settings.dbpass & "';")
                        Dim ds1 As New DataSet
                        Dim eds1 As New DataGridView
                        Dim dv1 As New DataView
                        Mycn.Open()

                        SQLstr = "Update AMP_TradePending SET TradeReady = -1"

                        Command = New OleDb.OleDbCommand(SQLstr, Mycn)
                        Command.ExecuteNonQuery()

                        Mycn.Close()

                        Call MoveReadyTrades()
                        'Call LoadPendingTradesGrid()

                    Catch ex As Exception
                        MsgBox(ex.Message, MsgBoxStyle.Critical, "ERROR")
                    End Try
                End If
            Else
                Call LoadPendingTradesGrid()
            End If
        End If
    End Sub

    Public Sub MoveReadyTrades()
        Dim Mycn As OleDb.OleDbConnection
        Dim Command As OleDb.OleDbCommand
        Dim Command1 As OleDb.OleDbCommand
        Dim SQLstr As String
        Try
            Mycn = New OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & My.Settings.dbloc & "';Jet OLEDB:Database Password='" & My.Settings.dbpass & "';")
            Mycn.Open()

            SQLstr = "INSERT INTO AMP_TradeReady(MoxyOrderID,ClientAllocation, ContraDealer, Type, Quantity, Security, Price, BetaInventory, TradeDate, SettleDate, Principal, Interest)" & _
            " SELECT MoxyOrderID,ClientAllocation, ContraDealer, Type, Quantity, Security, Price, BetaInventory, TradeDate, SettleDate, Principal, Interest" & _
            " FROM AMP_TradePending" & _
            " WHERE TradeReady = -1"

            Command = New OleDb.OleDbCommand(SQLstr, Mycn)
            Command.ExecuteNonQuery()

            SQLstr = "DELETE * FROM AMP_TradePending WHERE TradeReady = -1"

            Command1 = New OleDb.OleDbCommand(SQLstr, Mycn)
            Command1.ExecuteNonQuery()

            Mycn.Close()

            'Call MoveReadyTrades()
            Call LoadPendingTradesGrid()
            Call LoadReadyTrades()

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "ERROR")
        End Try
    End Sub

    Public Sub LoadReadyTrades()
        Try

            DataGridView2.Enabled = False

            Dim conn As New OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & My.Settings.dbloc & "';Jet OLEDB:Database Password='" & My.Settings.dbpass & "';")

            Dim strSQL As String = "SELECT * FROM AMP_TradeReady"
            Dim da As New OleDb.OleDbDataAdapter(strSQL, conn)
            Dim ds As New DataSet
            da.Fill(ds, "Users")

            With DataGridView2
                .DataSource = ds.Tables("Users")
                .Columns(0).Visible = False
            End With

            'Label1.Text = "Loaded " & DataGridView2.RowCount & " ready trades."

            DataGridView2.Enabled = True

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "ERROR")
        End Try
    End Sub

    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        Call LoadReadyTrades()
    End Sub

    Private Sub AddSecTypesToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub

    Private Sub Button5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button5.Click
        Dim child As New AMP_SecTypes
        child.MdiParent = Home
        child.Show()
    End Sub

    Private Sub ApproveTradeToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ApproveTradeToolStripMenuItem.Click
        If DataGridView1.RowCount = "0" Then
            'do nothing
        Else
            Dim Mycn As OleDb.OleDbConnection
            Dim Command As OleDb.OleDbCommand
            Dim SQLstr As String
            Try
                Mycn = New OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & My.Settings.dbloc & "';Jet OLEDB:Database Password='" & My.Settings.dbpass & "';")
                Dim ds1 As New DataSet
                Dim eds1 As New DataGridView
                Dim dv1 As New DataView
                Mycn.Open()

                SQLstr = "Update AMP_TradePending SET TradeReady = -1 WHERE ID = " & DataGridView1.SelectedCells(0).Value

                Command = New OleDb.OleDbCommand(SQLstr, Mycn)
                Command.ExecuteNonQuery()

                Mycn.Close()

                Call MoveReadyTrades()
                'Call LoadPendingTradesGrid()

            Catch ex As Exception
                MsgBox(ex.Message, MsgBoxStyle.Critical, "ERROR")
            End Try
        End If
    End Sub

    Private Sub ApproveAllToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ApproveAllToolStripMenuItem.Click
        Call ApproveAll()
    End Sub

    Private Sub ResetApprovalToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ResetApprovalToolStripMenuItem.Click
        If DataGridView2.RowCount = "0" Then
            'do nothing
        Else
            Dim Mycn As OleDb.OleDbConnection
            Dim Command As OleDb.OleDbCommand
            Dim Command1 As OleDb.OleDbCommand
            Dim SQLstr As String
            Try
                Mycn = New OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & My.Settings.dbloc & "';Jet OLEDB:Database Password='" & My.Settings.dbpass & "';")
                Mycn.Open()

                SQLstr = "INSERT INTO AMP_TradePending(MoxyOrderID,ClientAllocation,ContraDealer, Type, Quantity, Security, Price, BetaInventory, TradeDate, SettleDate, Principal, Interest)" & _
                " SELECT MoxyOrderID,ClientAllocation,ContraDealer, Type, Quantity, Security, Price, BetaInventory, TradeDate, SettleDate, Principal, Interest" & _
                " FROM AMP_TradeReady" & _
                " WHERE ID = " & DataGridView2.SelectedCells(0).Value

                Command = New OleDb.OleDbCommand(SQLstr, Mycn)
                Command.ExecuteNonQuery()

                SQLstr = "DELETE * FROM AMP_TradeReady WHERE ID = " & DataGridView2.SelectedCells(0).Value

                Command1 = New OleDb.OleDbCommand(SQLstr, Mycn)
                Command1.ExecuteNonQuery()

                Mycn.Close()

                'Call MoveReadyTrades()
                Call LoadPendingTradesGrid()
                Call LoadReadyTrades()

            Catch ex As Exception
                MsgBox(ex.Message, MsgBoxStyle.Critical, "ERROR")
            End Try
        End If
    End Sub

    Private Sub AMPFExcelExport_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        If Permissions.AMPSecCodes.Checked Then
            Button5.Enabled = True
        Else
            Button5.Enabled = False
        End If
    End Sub
End Class