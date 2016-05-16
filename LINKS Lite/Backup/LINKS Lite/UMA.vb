Imports System.IO
Imports System.Data.OleDb
Imports System.Web.Mail


Public Class UMA
    Dim TransPositions As System.Threading.Thread
    Dim FindTrades As System.Threading.Thread
    Dim transoverview As System.Threading.Thread
    Dim reconthread As System.Threading.Thread
    Dim recontranslate As System.Threading.Thread

    Private checkPrint As Integer

    Private Sub PrintDocument1_BeginPrint(ByVal sender As Object, ByVal e As System.Drawing.Printing.PrintEventArgs) Handles PrintDocument1.BeginPrint
        checkPrint = 0
    End Sub

    Private Sub PrintDocument1_PrintPage(ByVal sender As Object, ByVal e As System.Drawing.Printing.PrintPageEventArgs) Handles PrintDocument1.PrintPage
        ' Print the content of the RichTextBox. Store the last character printed.
        checkPrint = RichTextBoxPrintCtrl1.Print(checkPrint, RichTextBoxPrintCtrl1.TextLength, e)

        ' Look for more pages
        If checkPrint < RichTextBoxPrintCtrl1.TextLength Then
            e.HasMorePages = True
        Else
            e.HasMorePages = False
        End If
    End Sub

    Private Sub btnPageSetup_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnPageSetup.Click
        RichTextBoxPrintCtrl1.Text = RichTextBox1.Text
        PageSetupDialog1.ShowDialog()
    End Sub

    Private Sub btnPrint_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnPrint.Click
        RichTextBoxPrintCtrl1.Text = RichTextBox1.Text
        If PrintDialog1.ShowDialog() = DialogResult.OK Then
            PrintDocument1.Print()
        End If
    End Sub

    Private Sub btnPrintPreview_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnPrintPreview.Click
        RichTextBoxPrintCtrl1.Text = RichTextBox1.Text
        MsgBox(RichTextBoxPrintCtrl1.Text)
        PrintPreviewDialog1.ShowDialog()
    End Sub
    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        OpenFileDialog1.Title = "Please Select a File"
        OpenFileDialog1.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.Desktop) '"C:"
        OpenFileDialog1.DefaultExt = "CSV"
        OpenFileDialog1.Filter = "CSV files (*.csv)|*.csv|All files (*.*)|*.*"
        Dim fdate As Date
        fdate = DateTimePicker1.Text
        OpenFileDialog1.FileName = "account_positions_" & Format(fdate, "yyyy") & "_" & Format(fdate, "MM") & "_" & Format(fdate, "dd") & ".csv"
        OpenFileDialog1.ShowDialog()

    End Sub

    Private Sub OpenFileDialog1_FileOk(ByVal sender As System.Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles OpenFileDialog1.FileOk

        Dim strm As System.IO.Stream
        strm = OpenFileDialog1.OpenFile()

        TextBox1.Text = OpenFileDialog1.FileName.ToString

        If Not (strm Is Nothing) Then
            'insert code to read the file data
            strm.Close()
        End If

    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        Dim Mycn As OleDb.OleDbConnection
        Dim Command As OleDb.OleDbCommand
        Dim Command1 As OleDb.OleDbCommand
        Dim Command2 As OleDb.OleDbCommand
        Dim Command3 As OleDb.OleDbCommand
        Dim Command4 As OleDb.OleDbCommand
        Dim command5 As OleDb.OleDbCommand
        Dim SQLstr As String

        Try
            Mycn = New OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & My.Settings.dbloc & "';Jet OLEDB:Database Password='" & My.Settings.dbpass & "';")

            Dim ds1 As New DataSet
            Dim eds1 As New DataGridView
            Dim dv1 As New DataView

            Mycn.Open()

            Dim fpath As String
            Dim fname As String

            fpath = Path.GetDirectoryName(TextBox1.Text)
            fname = Path.GetFileName(TextBox1.Text)

            SQLstr = "DELETE * FROM uma_positions_Envestnet_Initial"

            Command = New OleDb.OleDbCommand(SQLstr, Mycn)
            Command.ExecuteNonQuery()

            '==========================================================================================
            'Create a Schema file to force field types of csv file
            '==========================================================================================

            Dim FileToDelete As String

            FileToDelete = fpath & "\Schema.ini"

            If System.IO.File.Exists(FileToDelete) = True Then

                System.IO.File.Delete(FileToDelete)
                'MsgBox("File Deleted")
            End If

            RichTextBox2.Text = "[" & fname & "]" & vbNewLine & "   ColNameHeader=False" & vbNewLine & "   Format=CSVDelimited" & vbNewLine & "   MaxScanRows=0" & vbNewLine & "   CharacterSet=OEM" & vbNewLine & "   Col1=F1 Char Width 255" & vbNewLine & "   Col2=F2 Char Width 255" & vbNewLine & "   Col3=F3 Double Width 25" & vbNewLine & "   Col4=F4 Currency Width 25" & vbNewLine & "   Col5=F5 Char Width 255"

            RichTextBox2.SaveFile(fpath & "\Schema.ini", _
                RichTextBoxStreamType.PlainText)

            '==========================================================================================
            'File Created
            '==========================================================================================

            SQLstr = "Insert INTO uma_positions_Envestnet_Initial([F1], [F2], [F3], [F4], [F5])" & _
            " SELECT [F1], [F2], [F3], [F4], [F5] FROM [Text;FMT=Delimited;HDR=NO;DATABASE=" & fpath & "]." & fname & ";"

            Command1 = New OleDb.OleDbCommand(SQLstr, Mycn)
            Command1.ExecuteNonQuery()

            SQLstr = "Delete * FROM uma_positions_Envestnet"

            Command2 = New OleDb.OleDbCommand(SQLstr, Mycn)
            Command2.ExecuteNonQuery()

            SQLstr = "UPDATE uma_positions_Envestnet_Initial SET F4 = 1 WHERE F4 < 1 OR F4 Is Null"

            command5 = New OleDb.OleDbCommand(SQLstr, Mycn)
            command5.ExecuteNonQuery()

            SQLstr = "Insert INTO uma_positions_Envestnet([F1], [F2], [F3], [F4], [F5])" & _
            " SELECT uma_portfolios_envestnet.PortfolioCode, uma_positions_envestnet_Initial.F2, uma_positions_envestnet_Initial.F3, uma_positions_envestnet_Initial.F4, uma_products_envestnet.ProductID" & _
            " FROM (uma_positions_envestnet_Initial INNER JOIN uma_portfolios_envestnet ON uma_positions_envestnet_Initial.F1 = uma_portfolios_envestnet.PortfolioCodeReal) INNER JOIN uma_products_envestnet ON (uma_positions_envestnet_Initial.F5 = uma_products_envestnet.ProductID) AND (uma_portfolios_envestnet.ProductID = uma_products_envestnet.ID);"

            Command3 = New OleDb.OleDbCommand(SQLstr, Mycn)
            Command3.ExecuteNonQuery()

            SQLstr = "DELETE * FROM uma_positions_envestnet WHERE F3=0"

            Command4 = New OleDb.OleDbCommand(SQLstr, Mycn)
            Command4.ExecuteNonQuery()

            Mycn.Close()

            MsgBox("Import Sucessful!", MsgBoxStyle.Information, "IMPORT")
            Button5.Enabled = True

            'Me.Close()

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "ERROR")

        End Try

    End Sub

    Private Sub UMA_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        If Permissions.UMALaunch.Checked Then
        Else
            MsgBox("You do not have permissions to view this system.", MsgBoxStyle.Critical, "Not Allowed")
            Me.Close()
        End If

        If Permissions.UMAImport.Checked Then
            Button1.Enabled = True
            Button2.Enabled = True
            Button3.Enabled = True
            Button4.Enabled = True

            TextBox1.Text = ""
            TextBox2.Text = ""
            TextBox1.Enabled = True
            TextBox2.Enabled = True

            DateTimePicker1.Enabled = True
            DateTimePicker2.Enabled = True

        Else
            Button1.Enabled = False
            Button2.Enabled = False
            Button3.Enabled = False
            Button4.Enabled = False

            TextBox1.Text = "Import disabled."
            TextBox2.Text = "Import disabled."
            TextBox1.Enabled = False
            TextBox2.Enabled = False

            DateTimePicker1.Enabled = False
            DateTimePicker2.Enabled = False
        End If

        If Permissions.UMATranslate.Checked Then
            Button5.Visible = True
            'Button6.Visible = True
        Else
            Button5.Visible = False
            'Button6.Visible = False
        End If

        If Permissions.UMASkipCheck.Checked Then
            CheckBox1.Enabled = True
        Else
            CheckBox1.Enabled = False
        End If

        If Permissions.UMAPortfolioFunctions.Checked Then
            Button15.Enabled = True
            Button16.Enabled = True
            Button17.Enabled = True
        Else
            Button15.Enabled = False
            Button16.Enabled = False
            Button17.Enabled = False
        End If

        If Permissions.UMATradeFunctions.Checked Then
            Button18.Enabled = True
            Button19.Enabled = True
            Button20.Enabled = True
            Button24.Enabled = True
        Else
            Button18.Enabled = False
            Button19.Enabled = False
            Button20.Enabled = False
            Button24.Enabled = False
        End If

        If Permissions.UMASystemFunctions.Checked Then
            Button22.Enabled = True
            Button23.Enabled = True
        Else
            Button22.Enabled = False
            Button23.Enabled = False
        End If

        If Permissions.UMAAutoRecon.Checked Then
            Button26.Enabled = True
            Button25.Enabled = True
            Button27.Enabled = True
            Button28.Enabled = True
            DataGridView5.Enabled = True
            DataGridView6.Enabled = True
        Else
            Button26.Enabled = False
            Button25.Enabled = False
            Button27.Enabled = False
            Button28.Enabled = False
            DataGridView5.Enabled = False
            DataGridView6.Enabled = False
        End If

        Dim dte1 As Date
        Dim dte2 As String
        Dim dte3 As Date
        dte3 = Format(Now())
        dte1 = DateAdd("d", -1, dte3)
        dte2 = Format(dte1, "dddd")

        Do Until dte2 = "Monday" Or dte2 = "Tuesday" Or dte2 = "Wednesday" Or dte2 = "Thursday" Or dte2 = "Friday"
            dte1 = DateAdd("d", -1, dte1)
            dte2 = Format(dte1, "dddd")
        Loop



        Me.DateTimePicker1.Text = dte1
        DateTimePicker2.Text = dte1
        'Dim fdate As Date
        'fdate = DateTimePicker1.Text

        'TextBox1.Text = "C:\account_positions_" & Format(fdate, "yyyy") & "_" & Format(fdate, "MM") & "_" & Format(fdate, "dd") & ".csv"
        'TextBox2.Text = "C:\account_trades_" & Format(fdate, "yyyy") & "_" & Format(fdate, "MM") & "_" & Format(fdate, "dd") & ".csv"
    End Sub

    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        OpenFileDialog2.Title = "Please Select a File"
        OpenFileDialog2.InitialDirectory = "C:"
        OpenFileDialog2.DefaultExt = "CSV"
        OpenFileDialog2.Filter = "CSV files (*.csv)|*.csv|All files (*.*)|*.*"
        Dim fdate As Date
        fdate = DateTimePicker2.Text
        OpenFileDialog2.FileName = "account_trades_" & Format(fdate, "yyyy") & "_" & Format(fdate, "MM") & "_" & Format(fdate, "dd") & ".csv"
        OpenFileDialog2.ShowDialog()
    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        Dim Mycn As OleDb.OleDbConnection
        Dim Command As OleDb.OleDbCommand
        Dim Command1 As OleDb.OleDbCommand
        Dim SQLstr As String

        Try
            Mycn = New OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & My.Settings.dbloc & "';Jet OLEDB:Database Password='" & My.Settings.dbpass & "';")

            Dim ds1 As New DataSet
            Dim eds1 As New DataGridView
            Dim dv1 As New DataView

            Mycn.Open()

            Dim fpath As String
            Dim fname As String

            fpath = Path.GetDirectoryName(TextBox2.Text)
            fname = Path.GetFileName(TextBox2.Text)

            SQLstr = "DELETE * FROM uma_trades_Envestnet"

            Command = New OleDb.OleDbCommand(SQLstr, Mycn)
            Command.ExecuteNonQuery()

            SQLstr = "Insert INTO uma_trades_Envestnet" & _
            " SELECT * FROM [Text;FMT=Delimited;HDR=NO;DATABASE=" & fpath & "]." & fname & ";"

            Command1 = New OleDb.OleDbCommand(SQLstr, Mycn)
            Command1.ExecuteNonQuery()

            Mycn.Close()

            Button6.Enabled = True
            'Me.Close()
            MsgBox("Import Sucessful!", MsgBoxStyle.Information, "IMPORT")
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "ERROR")

        End Try
    End Sub

    Private Sub Button8_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnPageSetup.Click

    End Sub

    Private Sub Button5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button5.Click
        If Label6.Visible = True Then
            MsgBox("The translator is currently busy. Please try again later.", MsgBoxStyle.Critical, "BUSY.")
        Else
            Label6.Text = "Working"
            Label6.Visible = True
            Timer1.Enabled = True
            Control.CheckForIllegalCrossThreadCalls = False
            TransPositions = New System.Threading.Thread(AddressOf TransPositionsGo)
            TransPositions.Start()
        End If
    End Sub

    Public Sub TransPositionsGo()
        RichTextBox1.Text = "Starting translation of Position File..."


        Dim stop1 As Integer
        stop1 = 0
        'Dim excount As String

        If CheckBox1.Checked Then
            RichTextBox1.Text = RichTextBox1.Text & vbNewLine & "Skipping data verification..."
            'excount = "0"
        Else
            '****BEGAIN DATA VALIDATION****
            RichTextBox1.Text = RichTextBox1.Text & vbNewLine & "Verifying Security and Account Data..."
            RichTextBox1.Text = RichTextBox1.Text & vbNewLine & "Checking Portfolio Codes..."
            Call PortfolioCodeExceptions()
            RichTextBox1.Text = RichTextBox1.Text & vbNewLine & "Checking Security Data..."
            Call SecurityDataExceptions()
            RichTextBox1.Text = RichTextBox1.Text & vbNewLine & "Finished with data verification!"


            Dim Mycn As OleDb.OleDbConnection

            Try
                Mycn = New OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & My.Settings.dbloc & "';Jet OLEDB:Database Password='" & My.Settings.dbpass & "';")
                Mycn.Open()
                Dim sqlstring As String
                sqlstring = "SELECT * FROM uma_exceptions;"
                Dim queryString As String = String.Format(sqlstring)
                Dim cmd As New OleDb.OleDbCommand(queryString, Mycn)
                Dim da As New OleDb.OleDbDataAdapter(cmd)
                Dim ds As New DataSet
                da.Fill(ds, "User")
                Dim dt As DataTable = ds.Tables("User")
                If dt.Rows.Count > 0 Then
                    RichTextBox1.Text = RichTextBox1.Text & vbNewLine & "ERROR: Misssing account or security data.  Check Exception Report."
                    stop1 = 1
                    Button7.ForeColor = Color.Red
                    CheckBox2.Checked = True
                    Timer5.Enabled = True
                    'excount = dt.Rows.Count.ToString
                    Mycn.Close()
                    Mycn.Dispose()
                Else
                    RichTextBox1.Text = RichTextBox1.Text & vbNewLine & "Ready to translate!"
                    'excount = "0"
                    Mycn.Close()
                    Mycn.Dispose()
                End If
            Catch ex As Exception
                MsgBox(ex.Message)
            End Try
            '****END DATA VERIFICIATION****
        End If
        'Check if exception was raised.  If exception was raised, stop process.
        '''''''''No longer applies, an exception will not stop the import.
        'If stop1 = 0 Then
        'Else
        'GoTo Line1
        'End If

        '****START TRANSLATION****
        RichTextBox1.Text = RichTextBox1.Text & vbNewLine & "Preparing tables for translation..."
        Call ClearBlotterTable()
        RichTextBox1.Text = RichTextBox1.Text & vbNewLine & "Tables ready!"
        RichTextBox1.Text = RichTextBox1.Text & vbNewLine & "Translation started..."
        Call PositionTranslator()
        RichTextBox1.Text = RichTextBox1.Text & vbNewLine & "Translation Finished!"
        RichTextBox1.Text = RichTextBox1.Text & vbNewLine & "Saving File..."
        Call SavePositionBlotter1()
        RichTextBox1.Text = RichTextBox1.Text & vbNewLine & "File Saved!"
        RichTextBox1.Text = RichTextBox1.Text & vbNewLine & "Cleaning Blotter File..."
        'Call CleanBlotterFile()
        RichTextBox1.Text = RichTextBox1.Text & vbNewLine & "Blotter File Saved!"
        RichTextBox1.Text = RichTextBox1.Text & vbNewLine & "Importing File to APX"
        'Dim psi As New ProcessStartInfo("\\aamapxapps01\apx$\Automation\UMAPositionBlotter.BAT")
        'psi.RedirectStandardError = True
        'psi.RedirectStandardOutput = True
        'psi.CreateNoWindow = False
        'psi.WindowStyle = ProcessWindowStyle.Hidden
        'psi.UseShellExecute = False

        'Dim process As Process = process.Start(psi)

        System.Diagnostics.Process.Start("\\aamapxapps01\apx$\Automation\UMAPositionBlotter.BAT")
        System.Diagnostics.Process.Start("\\aamapxapps01\apx$\Automation\UMAPositionBlotter64.BAT")
        RichTextBox1.Text = RichTextBox1.Text & vbNewLine & "Import Successful!"

        If stop1 = 0 Then
            RichTextBox1.Text = RichTextBox1.Text & vbNewLine & "Process completed successfully."
        Else
            RichTextBox1.Text = RichTextBox1.Text & vbNewLine & "Process completed with errors.  Check exception report."
        End If

        GoTo Line2
Line1:
        RichTextBox1.Text = RichTextBox1.Text & vbNewLine & "Translation stopped due to errors."
Line2:
        Label6.Visible = False
        Timer1.Enabled = False
        RichTextBoxPrintCtrl1.Text = RichTextBox1.Text
    End Sub
    Public Sub SavePositionBlotter1()
        Dim path As String
        path = "C:\_Blotters"

        If (System.IO.Directory.Exists(path)) Then
            System.IO.Directory.Delete(path, True)
        End If
        
        Dim AccessConn As New System.Data.OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & My.Settings.dbloc & "';Jet OLEDB:Database Password='" & My.Settings.dbpass & "';")

        AccessConn.Open()

        Dim AccessCommand As New OleDb.OleDbCommand("SELECT Field1,Field2,Field3,Field4,Field5,Field6,Field7,Field8,Field9,Field10,Field11,Field12,Field13,Field14,Field15,Field16,Field17,Field18,Field19,Field20,Field21,Field22,Field23,Field24,Field25,Field26,Field27,Field28,Field29,Field30,Field31,Field32,Field33,Field34,Field35,Field36,Field37,Field38,Field39,Field40,Field41,Field42,Field43,Field44,Field45,Field46,Field47,Field48,Field49,Field50,Field51,Field52,Field53,Field54,Field55,Field56,Field57,Field58,Field59 INTO [Text;HDR=NO;DATABASE=" & path & "].[Envestnet_UMA_Position.csv] FROM uma_positionblotter", AccessConn)
        System.IO.Directory.CreateDirectory(path)
        AccessCommand.ExecuteNonQuery()
        AccessConn.Close()

        Call CleanBlotterFile()

    End Sub
    Public Sub SaveTradeBlotter()
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

    Public Sub CleanBlotterFile()
        Dim myFiles As String()
        Dim path As String
        path = "C:\_Blotters"
        myFiles = IO.Directory.GetFiles(path, "*.csv")

        Dim newFilePath As String
        Try
            For Each filepath As String In myFiles

                newFilePath = filepath.Replace(".csv", ".trn")

                System.IO.File.Move(filepath, newFilePath)

            Next

            Dim FileToDelete As String

            FileToDelete = "\\aamapxapps01\apx$\imp\Envestnet_UMA_Position.trn"

            If System.IO.File.Exists(FileToDelete) = True Then

                System.IO.File.Delete(FileToDelete)
                'MsgBox("File Deleted")
            End If

            My.Computer.FileSystem.CopyFile("C:\_Blotters\Envestnet_UMA_Position.trn", "\\aamapxapps01\apx$\imp\Envestnet_UMA_Position.trn")

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "ERROR")

        End Try

    End Sub

    Public Sub SavePositionBlotter()
        Dim Mycn As OleDb.OleDbConnection
        Dim Command As OleDb.OleDbCommand
        Dim SQLstr As String

        Try
            Mycn = New OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & My.Settings.dbloc & "';Jet OLEDB:Database Password='" & My.Settings.dbpass & "';")

            Dim ds1 As New DataSet
            Dim eds1 As New DataGridView
            Dim dv1 As New DataView

            Mycn.Open()

            Dim fpath As String
            Dim fname As String

            fpath = Path.GetDirectoryName(TextBox1.Text)
            fname = Path.GetFileName(TextBox1.Text)

            'Dim AccessCommand As New System.Data.OleDb.OleDbCommand(

            'SQLstr = "SELECT * INTO" & _
            '"[Text;HDR=No;DATABASE=e:\My Documents\TextFiles].[td.txt] FROM uma_positionblotter", AccessConn)

            SQLstr = "SELECT * INTO" & _
            " [Text;FMT=Delimited;HDR=No;DATABASE=C:\_Blotters].[test.csv]" & _
            " FROM uma_positionblotter;"

            'SQLstr = "Insert INTO uma_positions_Envestnet" & _
            '" SELECT * FROM [Text;FMT=Delimited;HDR=NO;DATABASE=" & fpath & "]." & fname & ";"

            Command = New OleDb.OleDbCommand(SQLstr, Mycn)
            Command.ExecuteNonQuery()

            Mycn.Close()

            Button5.Visible = True

            'Me.Close()

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "ERROR")

        End Try
    End Sub

    Public Sub ClearBlotterTable()
        Dim Mycn As OleDb.OleDbConnection
        Dim Command As OleDb.OleDbCommand
        Dim Command1 As OleDb.OleDbCommand
        Dim SQLstr As String
        Try
            Mycn = New OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & My.Settings.dbloc & "';Jet OLEDB:Database Password='" & My.Settings.dbpass & "';")
            'Dim ds1 As New DataSet
            'Dim eds1 As New DataGridView
            'Dim dv1 As New DataView
            Mycn.Open()
            SQLstr = "DELETE * FROM uma_positionblotter;"
            Command = New OleDb.OleDbCommand(SQLstr, Mycn)
            Command.ExecuteNonQuery()

            SQLstr = "DELETE * FROM uma_tradeblotter;"
            Command1 = New OleDb.OleDbCommand(SQLstr, Mycn)
            Command1.ExecuteNonQuery()
            Mycn.Close()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "ERROR")
        End Try
    End Sub

    Public Sub PositionTranslator()
        Dim Mycn As OleDb.OleDbConnection
        Dim Command As OleDb.OleDbCommand
        Dim SQLstr As String
        Dim dte As Date
        dte = DateTimePicker1.Text
        Dim cdte As String
        cdte = Format(dte, "MMddyyyy")
        Try
            Mycn = New OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & My.Settings.dbloc & "';Jet OLEDB:Database Password='" & My.Settings.dbpass & "';")
            Dim ds1 As New DataSet
            Dim eds1 As New DataGridView
            Dim dv1 As New DataView
            Mycn.Open()
            'Check for and insert new exceptions
            SQLstr = "Insert INTO uma_positionblotter([Field1], [Field2], [Field4], [Field5], [Field6], [Field9], [Field18], [Field29], [Field30], [Field42], [Field45], [Field46], [Field59])" & _
            " SELECT uma_positions_envestnet.F1, 'li', AdvApp_vSecurity.SecTypeBaseCode + 'us', uma_positions_envestnet.F2,'" & cdte & "',uma_positions_envestnet.F3,uma_positions_envestnet.F4,'n','65533','1','n','y','y'" & _
            " FROM uma_positions_envestnet INNER JOIN AdvApp_vSecurity ON uma_positions_envestnet.F2 = AdvApp_vSecurity.Symbol;"
            Command = New OleDb.OleDbCommand(SQLstr, Mycn)
            Command.ExecuteNonQuery()
            Mycn.Close()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "ERROR")
        End Try
    End Sub

    Public Sub TradeTranslator()
        Dim Mycn As OleDb.OleDbConnection
        Dim Command As OleDb.OleDbCommand
        Dim Command1 As OleDb.OleDbCommand
        Dim SQLstr As String
        Dim dte As Date
        dte = DateTimePicker1.Text
        Dim cdte As String
        cdte = Format(dte, "MMddyyyy")
        Try
            Mycn = New OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & My.Settings.dbloc & "';Jet OLEDB:Database Password='" & My.Settings.dbpass & "';")
            Dim ds1 As New DataSet
            Dim eds1 As New DataGridView
            Dim dv1 As New DataView
            Mycn.Open()
            'Check for and insert new exceptions
            SQLstr = "Insert INTO uma_tradeblotter([Field1], [Field2], [Field4], [Field5], [Field6], [Field7], [Field9], [Field12], [Field13], [Field18], [Field22], [Field24], [Field26], [Field29], [Field30], [Field42], [Field45], [Field46], [Field78])" & _
            " SELECT uma_trades_envestnet.F1, uma_trades_envestnet.F2, AdvApp_vSecurity.SecTypeBaseCode + 'us', uma_trades_envestnet.F6, uma_trades_envestnet.F3, uma_trades_envestnet.F4,uma_trades_envestnet.F5, 'caus', 'CASH','@' + uma_trades_envestnet.F7, AdvApp_vSecurity.ExchangeID, '0', 'n','n','65533','1','n','y','brok'" & _
            " FROM uma_trades_envestnet INNER JOIN AdvApp_vSecurity ON uma_trades_envestnet.F6 = AdvApp_vSecurity.Symbol" & _
            " WHERE uma_trades_envestnet.F2 = 'by';"
            Command = New OleDb.OleDbCommand(SQLstr, Mycn)
            Command.ExecuteNonQuery()

            SQLstr = "Insert INTO uma_tradeblotter([Field1], [Field2], [Field4], [Field5], [Field6], [Field7], [Field9], [Field12], [Field13], [Field18], [Field22], [Field23], [Field24], [Field26], [Field29], [Field42], [Field45], [Field46], [Field78])" & _
            " SELECT uma_trades_envestnet.F1, uma_trades_envestnet.F2, AdvApp_vSecurity.SecTypeBaseCode + 'us', uma_trades_envestnet.F6, uma_trades_envestnet.F3, uma_trades_envestnet.F4,uma_trades_envestnet.F5, 'caus', 'CASH','@' + uma_trades_envestnet.F7, AdvApp_vSecurity.ExchangeID, 'y','0', 'n','n','1','n','y','brok'" & _
            " FROM uma_trades_envestnet INNER JOIN AdvApp_vSecurity ON uma_trades_envestnet.F6 = AdvApp_vSecurity.Symbol" & _
            " WHERE uma_trades_envestnet.F2 = 'sl';"
            Command1 = New OleDb.OleDbCommand(SQLstr, Mycn)
            Command1.ExecuteNonQuery()

            Mycn.Close()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "ERROR")
        End Try
    End Sub

    Public Sub PortfolioCodeExceptions()
        Dim Mycn As OleDb.OleDbConnection
        Dim Command As OleDb.OleDbCommand
        Dim Command1 As OleDb.OleDbCommand
        Dim command2 As OleDb.OleDbCommand
        Dim SQLstr As String
        Try
            Mycn = New OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & My.Settings.dbloc & "';Jet OLEDB:Database Password='" & My.Settings.dbpass & "';")
            Dim ds1 As New DataSet
            Dim eds1 As New DataGridView
            Dim dv1 As New DataView
            'Clear out old exceptions
            Mycn.Open()
            SQLstr = "DELETE * FROM uma_exceptions"
            Command = New OleDb.OleDbCommand(SQLstr, Mycn)
            Command.ExecuteNonQuery()
            'Check for and insert new exceptions
            SQLstr = "Insert INTO uma_exceptions([PortfolioCode], [Security], [Exception])" & _
            " SELECT uma_positions_envestnet_Initial.[F1], uma_positions_envestnet_Initial.[F2], 'Portfolio Code not known in Advent'" & _
            " FROM uma_positions_envestnet_Initial" & _
            " WHERE uma_positions_envestnet_Initial.[F1] Not In (Select PortfolioCodeReal from uma_portfolios_envestnet)"
            Command1 = New OleDb.OleDbCommand(SQLstr, Mycn)
            Command1.ExecuteNonQuery()
            'Remove rejected portfolios from feed.
            SQLstr = "DELETE * FROM uma_positions_envestnet WHERE [F1] In (Select PortfolioCode FROM uma_exceptions)"
            command2 = New OleDb.OleDbCommand(SQLstr, Mycn)
            command2.ExecuteNonQuery()
            Mycn.Close()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "ERROR")
        End Try
    End Sub
    Public Sub PortfolioCodeExceptionsT()
        Dim Mycn As OleDb.OleDbConnection
        Dim Command As OleDb.OleDbCommand
        Dim Command1 As OleDb.OleDbCommand
        Dim Command2 As OleDb.OleDbCommand
        Dim SQLstr As String
        Try
            Mycn = New OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & My.Settings.dbloc & "';Jet OLEDB:Database Password='" & My.Settings.dbpass & "';")
            Dim ds1 As New DataSet
            Dim eds1 As New DataGridView
            Dim dv1 As New DataView
            'Clear out old exceptions
            Mycn.Open()
            SQLstr = "DELETE * FROM uma_exceptions"
            Command = New OleDb.OleDbCommand(SQLstr, Mycn)
            Command.ExecuteNonQuery()
            'Check for and insert new exceptions
            SQLstr = "Insert INTO uma_exceptions([PortfolioCode], [Security], [Exception])" & _
            " SELECT uma_trades_envestnet.[F1], uma_trades_envestnet.[F6], 'Portfolio Code not known in Advent'" & _
            " FROM uma_trades_envestnet" & _
            " WHERE F1 Not In (Select PortfolioCode from uma_portfolios_envestnet)"
            Command1 = New OleDb.OleDbCommand(SQLstr, Mycn)
            Command1.ExecuteNonQuery()
            'Remove rejected portfolios from feed.
            SQLstr = "DELETE * FROM uma_trades_envestnet WHERE [F1] In (Select PortfolioCode FROM uma_exceptions)"
            command2 = New OleDb.OleDbCommand(SQLstr, Mycn)
            command2.ExecuteNonQuery()
            Mycn.Close()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "ERROR")
        End Try
    End Sub
    Public Sub SecurityDataExceptions()
        Dim Mycn As OleDb.OleDbConnection
        Dim Command1 As OleDb.OleDbCommand
        Dim Command2 As OleDb.OleDbCommand
        Dim Command3 As OleDb.OleDbCommand

        Dim SQLstr As String
        Try
            Mycn = New OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & My.Settings.dbloc & "';Jet OLEDB:Database Password='" & My.Settings.dbpass & "';")
            Dim ds1 As New DataSet
            Dim eds1 As New DataGridView
            Dim dv1 As New DataView
            Mycn.Open()
            SQLstr = "Insert INTO uma_exceptions([PortfolioCode], [Security], [Exception])" & _
            " SELECT uma_positions_envestnet.[F1], uma_positions_envestnet.[F2], 'Symbol not known in Advent'" & _
            " FROM uma_positions_envestnet" & _
            " WHERE F2 Not In (Select Symbol from AdvApp_vSecurity)"
            Command1 = New OleDb.OleDbCommand(SQLstr, Mycn)
            Command1.ExecuteNonQuery()
            SQLstr = "Insert INTO uma_exceptions([PortfolioCode], [Security], [Exception])" & _
            " SELECT uma_positions_envestnet.[F1], uma_positions_envestnet.[F2], 'Missing security value on feed'" & _
            " FROM uma_positions_envestnet" & _
            " WHERE F4 Is Null"
            Command3 = New OleDb.OleDbCommand(SQLstr, Mycn)
            Command3.ExecuteNonQuery()


            'Remove rejected portfolios from feed.
            SQLstr = "DELETE * FROM uma_positions_envestnet WHERE [F1] In (Select PortfolioCode FROM uma_exceptions)"
            command2 = New OleDb.OleDbCommand(SQLstr, Mycn)
            command2.ExecuteNonQuery()
            Mycn.Close()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "ERROR")
        End Try
    End Sub

    Public Sub SecurityDataExceptionsT()
        Dim Mycn As OleDb.OleDbConnection
        Dim Command1 As OleDb.OleDbCommand
        Dim Command2 As OleDb.OleDbCommand
        Dim SQLstr As String
        Try
            Mycn = New OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & My.Settings.dbloc & "';Jet OLEDB:Database Password='" & My.Settings.dbpass & "';")
            Dim ds1 As New DataSet
            Dim eds1 As New DataGridView
            Dim dv1 As New DataView
            Mycn.Open()
            SQLstr = "Insert INTO uma_exceptions([PortfolioCode], [Security], [Exception])" & _
            " SELECT uma_trades_envestnet.[F1], uma_trades_envestnet.[F2], 'Symbol not known in Advent'" & _
            " FROM uma_trades_envestnet" & _
            " WHERE F2 Not In (Select Symbol from AdvApp_vSecurity)"
            Command1 = New OleDb.OleDbCommand(SQLstr, Mycn)
            Command1.ExecuteNonQuery()
            'Remove rejected portfolios from feed.
            SQLstr = "DELETE * FROM uma_trades_envestnet WHERE [F1] In (Select PortfolioCode FROM uma_exceptions)"
            command2 = New OleDb.OleDbCommand(SQLstr, Mycn)
            command2.ExecuteNonQuery()
            Mycn.Close()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "ERROR")
        End Try
    End Sub

    Private Sub Timer1_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer1.Tick

        If Label6.Text = "Working" Then
            Label6.Text = "Working."
        Else
            If Label6.Text = "Working." Then
                Label6.Text = "Working.."
            Else
                If Label6.Text = "Working.." Then
                    Label6.Text = "Working..."
                Else
                    If Label6.Text = "Working..." Then
                        Label6.Text = "Working"
                    End If
                End If
            End If

        End If
    End Sub

    Private Sub OpenFileDialog2_FileOk(ByVal sender As System.Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles OpenFileDialog2.FileOk
        Dim strm As System.IO.Stream
        strm = OpenFileDialog2.OpenFile()

        TextBox2.Text = OpenFileDialog2.FileName.ToString

        If Not (strm Is Nothing) Then
            'insert code to read the file data
            strm.Close()
        End If
    End Sub

    Private Sub Button6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button6.Click
        If Label6.Visible = True Then
            MsgBox("The translator is currently busy. Please try again later.", MsgBoxStyle.Critical, "BUSY.")
        Else
            Label6.Text = "Working"
            Label6.Visible = True
            Timer1.Enabled = True
            Control.CheckForIllegalCrossThreadCalls = False
            TransPositions = New System.Threading.Thread(AddressOf TransTradesGo)
            TransPositions.Start()
        End If
    End Sub

    Public Sub TransTradesGo()
        RichTextBox1.Text = "Starting translation of Trade File..."


        Dim stop1 As Integer
        stop1 = 0

        If CheckBox1.Checked Then
            RichTextBox1.Text = RichTextBox1.Text & vbNewLine & "Skipping data verification..."
        Else
            '****BEGAIN DATA VALIDATION****
            RichTextBox1.Text = RichTextBox1.Text & vbNewLine & "Verifying Security and Account Data..."
            RichTextBox1.Text = RichTextBox1.Text & vbNewLine & "Checking Portfolio Codes..."
            Call PortfolioCodeExceptionsT()
            RichTextBox1.Text = RichTextBox1.Text & vbNewLine & "Checking Security Data..."
            Call SecurityDataExceptionsT()
            RichTextBox1.Text = RichTextBox1.Text & vbNewLine & "Finished with data verification!"


            Dim Mycn As OleDb.OleDbConnection
            Try
                Mycn = New OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & My.Settings.dbloc & "';Jet OLEDB:Database Password='" & My.Settings.dbpass & "';")
                Mycn.Open()
                Dim sqlstring As String
                sqlstring = "SELECT * FROM uma_exceptions;"
                Dim queryString As String = String.Format(sqlstring)
                Dim cmd As New OleDb.OleDbCommand(queryString, Mycn)
                Dim da As New OleDb.OleDbDataAdapter(cmd)
                Dim ds As New DataSet
                da.Fill(ds, "User")
                Dim dt As DataTable = ds.Tables("User")
                If dt.Rows.Count > 0 Then
                    RichTextBox1.Text = RichTextBox1.Text & vbNewLine & "ERROR: Misssing account or security data."
                    stop1 = 1
                    Mycn.Close()
                    Mycn.Dispose()
                Else
                    RichTextBox1.Text = RichTextBox1.Text & vbNewLine & "Ready to translate!"
                    Mycn.Close()
                    Mycn.Dispose()
                End If
            Catch ex As Exception
                MsgBox(ex.Message)
            End Try
            '****END DATA VERIFICIATION****
        End If
        'Check if exception was raised.  If exception was raised, stop process.
        If stop1 = 0 Then
        Else
            GoTo Line1
        End If

        '****START TRANSLATION****
        RichTextBox1.Text = RichTextBox1.Text & vbNewLine & "Preparing tables for translation..."
        Call ClearBlotterTable()
        RichTextBox1.Text = RichTextBox1.Text & vbNewLine & "Tables ready!"
        RichTextBox1.Text = RichTextBox1.Text & vbNewLine & "Converting codes to APX language..."
        Call updatetrades()
        RichTextBox1.Text = RichTextBox1.Text & vbNewLine & "Codes converted!"
        RichTextBox1.Text = RichTextBox1.Text & vbNewLine & "Translation started..."
        Call TradeTranslator()
        RichTextBox1.Text = RichTextBox1.Text & vbNewLine & "Translation Finished!"
        RichTextBox1.Text = RichTextBox1.Text & vbNewLine & "Saving File..."
        Call SaveTradeBlotter()
        RichTextBox1.Text = RichTextBox1.Text & vbNewLine & "File Saved!"
        RichTextBox1.Text = RichTextBox1.Text & vbNewLine & "Cleaning Blotter File..."
        'Call CleanBlotterFile()
        RichTextBox1.Text = RichTextBox1.Text & vbNewLine & "Blotter File Saved!"
        RichTextBox1.Text = RichTextBox1.Text & vbNewLine & "Importing File to APX"
        'Dim psi As New ProcessStartInfo("\\aamapxapps01\apx$\Automation\UMAPositionBlotter.BAT")
        'psi.RedirectStandardError = True
        'psi.RedirectStandardOutput = True
        'psi.CreateNoWindow = False
        'psi.WindowStyle = ProcessWindowStyle.Hidden
        'psi.UseShellExecute = False

        'Dim process As Process = process.Start(psi)

        System.Diagnostics.Process.Start("\\aamapxapps01\apx$\Automation\UMATradeBlotter.BAT")
        System.Diagnostics.Process.Start("\\aamapxapps01\apx$\Automation\UMATradeBlotter64.BAT")
        RichTextBox1.Text = RichTextBox1.Text & vbNewLine & "Import Successful!"
        GoTo Line2
Line1:
        RichTextBox1.Text = RichTextBox1.Text & vbNewLine & "Translation stopped due to errors."
Line2:
        Label6.Visible = False
        Timer1.Enabled = False
    End Sub

    Public Sub updatetrades()

        Dim Mycn As OleDb.OleDbConnection
        Dim Command As OleDb.OleDbCommand
        Dim Command1 As OleDb.OleDbCommand
        Dim Command2 As OleDb.OleDbCommand
        Dim Command3 As OleDb.OleDbCommand
        Dim SQLstr As String
        Try
            Mycn = New OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & My.Settings.dbloc & "';Jet OLEDB:Database Password='" & My.Settings.dbpass & "';")
            Dim ds1 As New DataSet
            Dim eds1 As New DataGridView
            Dim dv1 As New DataView
            Mycn.Open()
            SQLstr = "Update uma_trades_envestnet SET F2 = 'by' WHERE F2 = 'B' AND F2 <> NULL;"
            Command1 = New OleDb.OleDbCommand(SQLstr, Mycn)
            Command1.ExecuteNonQuery()

            SQLstr = "Update uma_trades_envestnet SET F2 = 'sl' WHERE F2 = 'S' AND F2 <> NULL;"
            Command = New OleDb.OleDbCommand(SQLstr, Mycn)
            Command.ExecuteNonQuery()

            SQLstr = "Update uma_trades_envestnet SET F3 = Format(F3, 'MMddyyyy')" & _
            " WHERE F3 <> Format(F3, 'MMddyyyy') AND F3 <> NULL;"
            Command2 = New OleDb.OleDbCommand(SQLstr, Mycn)
            Command2.ExecuteNonQuery()

            SQLstr = "Update uma_trades_envestnet SET F4 = Format(F4, 'MMddyyyy')" & _
            " WHERE F4 <> Format(F4, 'MMddyyyy') AND F4 <> NULL;"
            Command3 = New OleDb.OleDbCommand(SQLstr, Mycn)
            Command3.ExecuteNonQuery()

            Mycn.Close()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "ERROR")
        End Try

    End Sub

    Private Sub Button8_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button8.Click
        Label10.Visible = False
        TextBox4.Visible = True
        TextBox5.Visible = True
        Timer2.Enabled = True

        Control.CheckForIllegalCrossThreadCalls = False
        FindTrades = New System.Threading.Thread(AddressOf LoadPendingTrades)
        FindTrades.Start()
        TextBox5.Text = "Starting..."
        'Call LoadPendingTrades()

    End Sub

    Public Sub LoadPendingTrades()
        Dim Mycn As OleDb.OleDbConnection
        Dim Command As OleDb.OleDbCommand
        Dim Command1 As OleDb.OleDbCommand
        Dim Command2 As OleDb.OleDbCommand
        Dim Command3 As OleDb.OleDbCommand
        Dim Command4 As OleDb.OleDbCommand
        Dim Command5 As OleDb.OleDbCommand
        Dim SQLstr As String
        Try
            Mycn = New OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & My.Settings.dbloc & "';Jet OLEDB:Database Password='" & My.Settings.dbpass & "';")
            Dim ds1 As New DataSet
            Dim eds1 As New DataGridView
            Dim dv1 As New DataView
            Mycn.Open()
            TextBox5.Text = "Cleaning Table..."
            SQLstr = "DELETE * FROM uma_envestnet_TradePending"
            Command = New OleDb.OleDbCommand(SQLstr, Mycn)
            Command.ExecuteNonQuery()
            TextBox5.Text = "Cleaning Factors..."
            SQLstr = "DELETE * FROM uma_tempfactortbl"
            Command5 = New OleDb.OleDbCommand(SQLstr, Mycn)
            Command5.ExecuteNonQuery()
            TextBox5.Text = "Searching Trades..."
            SQLstr = "INSERT INTO uma_envestnet_TradePending(MoxyOrderID,Field1,Field2,Field3,Field4,Field5,Field6,Field7,Field8,Field9,Field11,Field12,Field13,Field15,Field16,Field17,Field18)" & _
            " SELECT MxApp_vAllocation.OrderID, uma_portfolios_envestnet.PortfolioCodeReal, uma_tradetranslator.EnvestnetCode, Format([MxApp_vAllocation].[TradeDate],'mm/dd/yyyy'), Format([MxApp_vAllocation].[SettleDate],'mm/dd/yyyy'), MxApp_vAllocation.AllocQty, MxApp_vSecurity.Symbol, MxApp_vAllocation.AllocPrice, MxApp_vAllocation.PrincipalAmt, IIF([MxApp_vAllocation].[AIAmt]<=0,0, Format([MxApp_vAllocation].[AIAmt], '#.##')), uma_envestnet_sectype.EnvestnetSecType, AdvApp_vSecurity.FullName, uma_products_envestnet.ProductID, 'FISA', '0443', 'Advisors Asset Management', dbo_vQbRowDefPortfolio.Custodian7" & _
            " FROM (((((MxApp_vAllocation INNER JOIN MxApp_vSecurity ON MxApp_vAllocation.SecurityID = MxApp_vSecurity.SecurityID) INNER JOIN uma_envestnet_sectype ON MxApp_vSecurity.SecTypeCode = uma_envestnet_sectype.APXSecType) INNER JOIN AdvApp_vSecurity ON MxApp_vSecurity.Symbol = AdvApp_vSecurity.Symbol) INNER JOIN (uma_portfolios_envestnet INNER JOIN uma_products_envestnet ON uma_portfolios_envestnet.ProductID = uma_products_envestnet.ID) ON MxApp_vAllocation.PortfolioCode = uma_portfolios_envestnet.PortfolioCode) INNER JOIN dbo_vQbRowDefPortfolio ON MxApp_vAllocation.PortfolioCode = dbo_vQbRowDefPortfolio.PortfolioCode) INNER JOIN uma_tradetranslator ON MxApp_vAllocation.TransactionCode = uma_tradetranslator.APXCode;"
            '" SELECT MxOm_vTskMoxyAllocation.OrderID, MxOm_vTskMoxyAllocation.PortID, uma_tradetranslator.EnvestnetCode, Format(MxOm_vTskMoxyAllocation.TradeDate, 'MM/dd/yyyy'), Format(MxOm_vTskMoxyAllocation.SettleDate, 'MM/dd/yyyy'), MxOm_vTskMoxyAllocation.AllocQty, MxOm_vTskMoxyAllocation.Symbol, MxOm_vTskMoxyAllocation.AllocPrice, Format(((MxOm_vTskMoxyAllocation.AllocPrice * MxOm_vTskMoxyAllocation.AllocQty)/100), '#.00') As Principal, Format(((MxOm_vTskMoxyOrders.AIAmount * MxOm_vTskMoxyAllocation.AllocQty)/100), '#.00') As Accrued, uma_envestnet_sectype.EnvestnetSecType, AdvApp_vSecurity.FullName, uma_products_envestnet.ProductID,'1234','0446','Advisors Asset Management', dbo_vQbRowDefPortfolio.Custodian7" & _
            '" FROM ((((((MxOm_vTskMoxyAllocation INNER JOIN MxOm_vTskMoxyOrders ON MxOm_vTskMoxyAllocation.OrderID = MxOm_vTskMoxyOrders.OrderID) INNER JOIN uma_tradetranslator ON MxOm_vTskMoxyAllocation.TranCode = uma_tradetranslator.MoxyCode) INNER JOIN AdvApp_vSecurity ON MxOm_vTskMoxyAllocation.Symbol = AdvApp_vSecurity.Symbol) INNER JOIN dbo_vQbRowDefPortfolio ON MxOm_vTskMoxyAllocation.PortID = dbo_vQbRowDefPortfolio.PortfolioCode) INNER JOIN uma_envestnet_sectype ON MxOm_vTskMoxyAllocation.SecType = uma_envestnet_sectype.APXSecType) INNER JOIN uma_portfolios_envestnet ON MxOm_vTskMoxyAllocation.PortID = uma_portfolios_envestnet.PortfolioCode) INNER JOIN uma_products_envestnet ON uma_portfolios_envestnet.ProductID = uma_products_envestnet.ID;"
            Command1 = New OleDb.OleDbCommand(SQLstr, Mycn)
            Command1.ExecuteNonQuery()
            TextBox5.Text = "Removing Processed Trades..."
            SQLstr = "DELETE * FROM uma_envestnet_TradePending WHERE MoxyOrderID In (SELECT MoxyOrderID FROM uma_tradetracker)"
            Command2 = New OleDb.OleDbCommand(SQLstr, Mycn)
            Command2.ExecuteNonQuery()
            TextBox5.Text = "Removing Trades already on upload list..."
            SQLstr = "DELETE * FROM uma_envestnet_TradePending WHERE MoxyOrderID In (SELECT MoxyOrderID FROM uma_envestnet_TradeReady)"
            Command2 = New OleDb.OleDbCommand(SQLstr, Mycn)
            Command2.ExecuteNonQuery()
            TextBox5.Text = "Searching for factors..."
            SQLstr = "INSERT INTO uma_tempfactortbl(PendingID, Factor)" & _
            " SELECT Top 1 uma_envestnet_TradePending.ID, AdvApp_vSecurityFactor.FactorValue" & _
            " FROM (uma_envestnet_TradePending INNER JOIN AdvApp_vSecurity ON uma_envestnet_TradePending.Field6 = AdvApp_vSecurity.Symbol) INNER JOIN AdvApp_vSecurityFactor ON AdvApp_vSecurity.SecurityID = AdvApp_vSecurityFactor.SecurityID" & _
            " ORDER BY AdvApp_vSecurityFactor.ThruDate DESC;"
            Command3 = New OleDb.OleDbCommand(SQLstr, Mycn)
            Command3.ExecuteNonQuery()
            TextBox5.Text = "Loading factors..."
            SQLstr = "Update uma_envestnet_TradePending" & _
            " INNER JOIN uma_tempfactortbl ON uma_tempfactortbl.PendingID = uma_envestnet_TradePending.ID" & _
            " SET uma_envestnet_TradePending.Field10 = uma_tempfactortbl.Factor"
            Command4 = New OleDb.OleDbCommand(SQLstr, Mycn)
            Command4.ExecuteNonQuery()
            TextBox5.Text = "Finised loading data!"
            Mycn.Close()

            'MsgBox("done")
            TextBox5.Text = "Drawing grid..."
            Call LoadPendingTradesGrid()

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "ERROR")
        End Try
    End Sub

    Public Sub LoadPendingTradesGrid()
        Try

            DataGridView1.Enabled = False

            Dim conn As New OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & My.Settings.dbloc & "';Jet OLEDB:Database Password='" & My.Settings.dbpass & "';")

            Dim strSQL As String = "SELECT * FROM uma_envestnet_TradePending"
            Dim da As New OleDb.OleDbDataAdapter(strSQL, conn)
            Dim ds As New DataSet
            da.Fill(ds, "Users")

            With DataGridView1
                .DataSource = ds.Tables("Users")
                .Columns(0).Visible = False
            End With

            Label10.Text = "Loaded " & DataGridView1.RowCount & " pending trades."

            DataGridView1.Enabled = True
            'DataGridView1.Visible = True
            TextBox4.Visible = False
            TextBox5.Visible = False
            Timer2.Enabled = False

            DataGridView1.ScrollBars = ScrollBars.Both
            DataGridView1.Refresh()

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "ERROR")
        End Try
    End Sub

    Public Sub LoadReadyTrades()
        Try

            DataGridView2.Enabled = False

            Dim conn As New OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & My.Settings.dbloc & "';Jet OLEDB:Database Password='" & My.Settings.dbpass & "';")

            Dim strSQL As String = "SELECT * FROM uma_envestnet_TradeReady"
            Dim da As New OleDb.OleDbDataAdapter(strSQL, conn)
            Dim ds As New DataSet
            da.Fill(ds, "Users")

            With DataGridView2
                .DataSource = ds.Tables("Users")
                .Columns(0).Visible = False
            End With

            Label1.Text = "Loaded " & DataGridView2.RowCount & " ready trades."

            DataGridView2.Enabled = True

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "ERROR")
        End Try
    End Sub

    Private Sub Timer2_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer2.Tick
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

    Private Sub EditStatusToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles EditStatusToolStripMenuItem.Click

    End Sub

    Private Sub Button9_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button9.Click
        If DataGridView1.RowCount = "0" Then
            'do nothing
            MsgBox("Please load positions.", MsgBoxStyle.Information, "No Data")
        Else
            Dim ir As MsgBoxResult

            ir = MsgBox("You are about to change the status on all pending trades.  Are you sure you want to continue?", MsgBoxStyle.YesNo + MsgBoxStyle.Exclamation, "Confirm")

            If ir = MsgBoxResult.Yes Then

                Dim Mycn As OleDb.OleDbConnection
                Dim Command As OleDb.OleDbCommand
                Dim SQLstr As String
                Try
                    Mycn = New OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & My.Settings.dbloc & "';Jet OLEDB:Database Password='" & My.Settings.dbpass & "';")
                    Dim ds1 As New DataSet
                    Dim eds1 As New DataGridView
                    Dim dv1 As New DataView
                    Mycn.Open()

                    SQLstr = "Update uma_envestnet_TradePending SET Field14 = '" & ComboBox1.SelectedItem & "'" ' WHERE Field14 <> '" & ComboBox1.SelectedText & "'"

                    Command = New OleDb.OleDbCommand(SQLstr, Mycn)
                    Command.ExecuteNonQuery()

                    Mycn.Close()

                    Call LoadPendingTradesGrid()

                Catch ex As Exception
                    MsgBox(ex.Message, MsgBoxStyle.Critical, "ERROR")
                End Try
            Else
                Call LoadPendingTradesGrid()
            End If
        End If
    End Sub

    Private Sub NEWToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles NEWToolStripMenuItem.Click
        If DataGridView1.RowCount = "0" Then
            'do nothing
        Else

            Dim ir As MsgBoxResult

            ir = MsgBox("You are about to change the status of order number " & DataGridView1.SelectedCells(1).Value & " to 'NEW'.  Are you sure you want to continue?", MsgBoxStyle.YesNo + MsgBoxStyle.Exclamation, "Confirm")

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

                        SQLstr = "Update uma_envestnet_TradePending SET Field14 = 'NEW' WHERE ID = " & DataGridView1.SelectedCells(0).Value

                        Command = New OleDb.OleDbCommand(SQLstr, Mycn)
                        Command.ExecuteNonQuery()

                        Mycn.Close()

                        Call LoadPendingTradesGrid()

                    Catch ex As Exception
                        MsgBox(ex.Message, MsgBoxStyle.Critical, "ERROR")
                    End Try

                End If
            Else
                Call LoadPendingTradesGrid()
            End If
        End If
    End Sub

    Private Sub BSENTToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BSENTToolStripMenuItem.Click
        If DataGridView1.RowCount = "0" Then
            'do nothing
        Else
            Dim ir As MsgBoxResult

            ir = MsgBox("You are about to change the status of order number " & DataGridView1.SelectedCells(1).Value & " to 'BSENT'.  Are you sure you want to continue?", MsgBoxStyle.YesNo + MsgBoxStyle.Exclamation, "Confirm")

            If ir = MsgBoxResult.Yes Then

                Dim Mycn As OleDb.OleDbConnection
                Dim Command As OleDb.OleDbCommand
                Dim SQLstr As String
                Try
                    Mycn = New OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & My.Settings.dbloc & "';Jet OLEDB:Database Password='" & My.Settings.dbpass & "';")
                    Dim ds1 As New DataSet
                    Dim eds1 As New DataGridView
                    Dim dv1 As New DataView
                    Mycn.Open()

                    SQLstr = "Update uma_envestnet_TradePending SET Field14 = 'BSENT' WHERE ID = " & DataGridView1.SelectedCells(0).Value

                    Command = New OleDb.OleDbCommand(SQLstr, Mycn)
                    Command.ExecuteNonQuery()

                    Mycn.Close()

                    Call LoadPendingTradesGrid()

                Catch ex As Exception
                    MsgBox(ex.Message, MsgBoxStyle.Critical, "ERROR")
                End Try
            Else
                Call LoadPendingTradesGrid()
            End If
        End If
    End Sub

    Private Sub SENTToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles SENTToolStripMenuItem.Click
        If DataGridView1.RowCount = "0" Then
            'do nothing
        Else
            Dim ir As MsgBoxResult

            ir = MsgBox("You are about to change the status of order number " & DataGridView1.SelectedCells(1).Value & " to 'SENT'.  Are you sure you want to continue?", MsgBoxStyle.YesNo + MsgBoxStyle.Exclamation, "Confirm")

            If ir = MsgBoxResult.Yes Then

                Dim Mycn As OleDb.OleDbConnection
                Dim Command As OleDb.OleDbCommand
                Dim SQLstr As String
                Try
                    Mycn = New OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & My.Settings.dbloc & "';Jet OLEDB:Database Password='" & My.Settings.dbpass & "';")
                    Dim ds1 As New DataSet
                    Dim eds1 As New DataGridView
                    Dim dv1 As New DataView
                    Mycn.Open()

                    SQLstr = "Update uma_envestnet_TradePending SET Field14 = 'SENT' WHERE ID = " & DataGridView1.SelectedCells(0).Value

                    Command = New OleDb.OleDbCommand(SQLstr, Mycn)
                    Command.ExecuteNonQuery()

                    Mycn.Close()

                    Call LoadPendingTradesGrid()

                Catch ex As Exception
                    MsgBox(ex.Message, MsgBoxStyle.Critical, "ERROR")
                End Try
            Else
                Call LoadPendingTradesGrid()
            End If
        End If
    End Sub

    Private Sub Button10_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button10.Click
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

                        SQLstr = "Update uma_envestnet_TradePending SET TradeReady = -1 WHERE Field14 <> NULL"

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

            SQLstr = "INSERT INTO uma_envestnet_TradeReady(MoxyOrderID,Field1,Field2,Field3,Field4,Field5,Field6,Field7,Field8,Field9,Field10,Field11,Field12,Field13,Field14,Field15,Field16,Field17,Field18)" & _
            " SELECT MoxyOrderID,Field1,Field2,Field3,Field4,Field5,Field6,Field7,Field8,Field9,Field10,Field11,Field12,Field13,Field14,Field15,Field16,Field17,Field18" & _
            " FROM uma_envestnet_TradePending" & _
            " WHERE TradeReady = -1"

            Command = New OleDb.OleDbCommand(SQLstr, Mycn)
            Command.ExecuteNonQuery()

            SQLstr = "DELETE * FROM uma_envestnet_TradePending WHERE TradeReady = -1"

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

    Private Sub MarkProcessedToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MarkProcessedToolStripMenuItem.Click
        Dim ir As MsgBoxResult
        If DataGridView1.RowCount = "0" Then
            'do nothing
        Else
            ir = MsgBox("You are about to Approve this trade.  Are you sure you want to continue?" & vbNewLine & vbNewLine & "NOTE: Trades without a status will not be moved to approved.", MsgBoxStyle.YesNo + MsgBoxStyle.Exclamation, "Confirm")

            If ir = MsgBoxResult.Yes Then
                Dim Mycn As OleDb.OleDbConnection
                Dim Command As OleDb.OleDbCommand
                Dim SQLstr As String
                Try
                    Mycn = New OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & My.Settings.dbloc & "';Jet OLEDB:Database Password='" & My.Settings.dbpass & "';")
                    Dim ds1 As New DataSet
                    Dim eds1 As New DataGridView
                    Dim dv1 As New DataView
                    Mycn.Open()

                    SQLstr = "Update uma_envestnet_TradePending SET TradeReady = -1 WHERE Field14 <> NULL AND ID = " & DataGridView1.SelectedCells(0).Value

                    Command = New OleDb.OleDbCommand(SQLstr, Mycn)
                    Command.ExecuteNonQuery()

                    Mycn.Close()

                    Call MoveReadyTrades()
                    'Call LoadPendingTradesGrid()
                    'Call LoadReadyTrades()

                Catch ex As Exception
                    MsgBox(ex.Message, MsgBoxStyle.Critical, "ERROR")
                End Try
            Else
            End If
        End If
    End Sub

    Private Sub RemoveTradeToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RemoveTradeToolStripMenuItem.Click
        Dim ir As MsgBoxResult
        If DataGridView1.RowCount = "0" Then
            'do nothing
        Else
            ir = MsgBox("WARNING:" & vbNewLine & "You are about to remove this trade from the processing list.  This process CANNOT be undone!  Are you sure that this trade should NEVER be uploaded?", MsgBoxStyle.YesNo + MsgBoxStyle.Critical, "Confirm")

            If ir = MsgBoxResult.Yes Then

                Dim Mycn As OleDb.OleDbConnection
                Dim Command As OleDb.OleDbCommand
                Dim Command1 As OleDb.OleDbCommand
                Dim SQLstr As String
                Try
                    Mycn = New OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & My.Settings.dbloc & "';Jet OLEDB:Database Password='" & My.Settings.dbpass & "';")
                    Mycn.Open()

                    SQLstr = "INSERT INTO uma_tradetracker(MoxyOrderID, DateReady, ReadyBy, TradeRemoved)" & _
                    "VALUES(" & DataGridView1.SelectedCells(0).Value & ",#" & Format(Now()) & "#,'" & Environ("USERNAME") & "',-1)"

                    Command = New OleDb.OleDbCommand(SQLstr, Mycn)
                    Command.ExecuteNonQuery()

                    SQLstr = "DELETE * FROM uma_envestnet_TradePending WHERE ID = " & DataGridView1.SelectedCells(0).Value

                    Command1 = New OleDb.OleDbCommand(SQLstr, Mycn)
                    Command1.ExecuteNonQuery()

                    Mycn.Close()

                    'Call MoveReadyTrades()
                    Call LoadPendingTradesGrid()

                Catch ex As Exception
                    MsgBox(ex.Message, MsgBoxStyle.Critical, "ERROR")
                End Try
            Else
                Call LoadPendingTradesGrid()
            End If
        End If
    End Sub

    Private Sub Button11_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button11.Click
        If DataGridView2.RowCount = "0" Then
            'do nothing
        Else
            Call SaveTradeReady()
        End If
    End Sub

    Public Sub SaveTradeReady()
        Dim path As String
        path = "\\monfp01\data\AAM Only\_UMATrades"

        If (System.IO.Directory.Exists(path)) Then
            'System.IO.Directory.Delete(path, True)
        Else
            System.IO.Directory.CreateDirectory(path)
        End If

        Dim dte As String
        dte = Format(Now(), "yyyyMMdd")

        Dim tme As String
        tme = Format(Now(), "HHmm")

        Dim fname As String
        fname = "trades_" & dte & "_" & tme & ".csv"

        Dim AccessConn As New System.Data.OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & My.Settings.dbloc & "';Jet OLEDB:Database Password='" & My.Settings.dbpass & "';")

        AccessConn.Open()

        Dim AccessCommand As New OleDb.OleDbCommand("SELECT Field1,Field2,Field3,Field4,Field5,Field6,Field7,Field8,Field9,Field10,Field11,Field12,Field13,Field14,Field15,Field16,Field17,Field18 INTO [Text;HDR=NO;DATABASE=" & path & "].[" & fname & "] FROM uma_envestnet_TradeReady", AccessConn)

        AccessCommand.ExecuteNonQuery()
        AccessConn.Close()

        Dim ir As MsgBoxResult
        Dim fullnme As String
        fullnme = path & "\" & fname

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

            'memo.To = "SMAOperations@aam.us.com"
            memo.To = "traders@envestnet.com"
            memo.CC = "SMAAdministration@aamlive.com"

            memo.Subject = "Uploaded UMA Trades"

            bdy = "Please see the attached file for UMA trades executed today.  These trades will be uploaded to the website shortly." & vbNewLine & vbNewLine & "Thank you," & vbNewLine & vbNewLine & "SMA Operations Team" & vbNewLine & "Advisors Asset Management, Inc." & vbNewLine & "1-866-259-2427" & vbNewLine & "SMAAdministration@aamlive.com"


            memo.Body = bdy
            memo.Attachments.Add(fullnme)
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

                SQLstr = "INSERT INTO uma_tradetracker(MoxyOrderID, DateReady, ReadyBy, TradeUploaded, UploadFile)" & _
                "SELECT MoxyOrderID,#" & Format(Now()) & "#,'" & Environ("USERNAME") & "',-1,'" & fullnme & "' FROM uma_envestnet_TradeReady"

                Command = New OleDb.OleDbCommand(SQLstr, Mycn)
                Command.ExecuteNonQuery()

                SQLstr = "DELETE * FROM uma_envestnet_TradeReady"

                Command1 = New OleDb.OleDbCommand(SQLstr, Mycn)
                Command1.ExecuteNonQuery()

                Mycn.Close()
            Catch ex As Exception
                MsgBox(ex.Message, MsgBoxStyle.Critical, "ERROR")
            End Try
        Else

        End If

    End Sub

    Private Sub Button12_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button12.Click
        Call LoadPendingTradesGrid()
    End Sub

    Private Sub Button13_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button13.Click
        Call LoadReadyTrades()
    End Sub

    Private Sub ChangeToPendingToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ChangeToPendingToolStripMenuItem.Click
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

                SQLstr = "INSERT INTO uma_envestnet_TradePending(MoxyOrderID,Field1,Field2,Field3,Field4,Field5,Field6,Field7,Field8,Field9,Field10,Field11,Field12,Field13,Field14,Field15,Field16,Field17,Field18)" & _
                " SELECT MoxyOrderID,Field1,Field2,Field3,Field4,Field5,Field6,Field7,Field8,Field9,Field10,Field11,Field12,Field13,Field14,Field15,Field16,Field17,Field18" & _
                " FROM uma_envestnet_TradeReady" & _
                " WHERE ID = " & DataGridView2.SelectedCells(0).Value

                Command = New OleDb.OleDbCommand(SQLstr, Mycn)
                Command.ExecuteNonQuery()

                SQLstr = "DELETE * FROM uma_envestnet_TradeReady WHERE ID = " & DataGridView2.SelectedCells(0).Value

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

    Private Sub RemoveTradeToolStripMenuItem1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RemoveTradeToolStripMenuItem1.Click
        Dim ir As MsgBoxResult
        If DataGridView2.RowCount = "0" Then
            'do nothing
        Else
            ir = MsgBox("WARNING:" & vbNewLine & "You are about to remove this trade from the processing list.  This process CANNOT be undone!  Are you sure that this trade should NEVER be uploaded?", MsgBoxStyle.YesNo + MsgBoxStyle.Critical, "Confirm")

            If ir = MsgBoxResult.Yes Then

                Dim Mycn As OleDb.OleDbConnection
                Dim Command As OleDb.OleDbCommand
                Dim Command1 As OleDb.OleDbCommand
                Dim SQLstr As String
                Try
                    Mycn = New OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & My.Settings.dbloc & "';Jet OLEDB:Database Password='" & My.Settings.dbpass & "';")
                    Mycn.Open()

                    SQLstr = "INSERT INTO uma_tradetracker(MoxyOrderID, DateReady, ReadyBy, TradeRemoved)" & _
                    "VALUES(" & DataGridView2.SelectedCells(0).Value & ",#" & Format(Now()) & "#,'" & Environ("USERNAME") & "',-1)"

                    Command = New OleDb.OleDbCommand(SQLstr, Mycn)
                    Command.ExecuteNonQuery()

                    SQLstr = "DELETE * FROM uma_envestnet_TradeReady WHERE ID = " & DataGridView2.SelectedCells(0).Value

                    Command1 = New OleDb.OleDbCommand(SQLstr, Mycn)
                    Command1.ExecuteNonQuery()

                    Mycn.Close()

                    'Call MoveReadyTrades()
                    Call LoadReadyTrades()

                Catch ex As Exception
                    MsgBox(ex.Message, MsgBoxStyle.Critical, "ERROR")
                End Try
            Else
                Call LoadReadyTrades()
            End If

        End If

    End Sub

    Private Sub EditTradeToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles EditTradeToolStripMenuItem.Click
        If DataGridView1.RowCount = "0" Then
            'do nothing
        Else
            Dim child As New UMATradeEdit
            child.MdiParent = Home
            child.Show()

            child.ID.Text = DataGridView1.SelectedCells(0).Value
            child.Label19.Text = "Editing Order Number: " & DataGridView1.SelectedCells(1).Value
            child.MoxyOrderID.Text = DataGridView1.SelectedCells(1).Value

            If IsDBNull(DataGridView1.SelectedCells(2).Value) Then
                child.Field1.Text = ""
            Else
                child.Field1.Text = DataGridView1.SelectedCells(2).Value

            End If

            If IsDBNull(DataGridView1.SelectedCells(3).Value) Then
                child.Field2.Text = ""
            Else
                child.Field2.Text = DataGridView1.SelectedCells(3).Value
            End If

            If IsDBNull(DataGridView1.SelectedCells(4).Value) Then
                child.Field3.Text = ""
            Else
                child.Field3.Text = DataGridView1.SelectedCells(4).Value
            End If

            If IsDBNull(DataGridView1.SelectedCells(5).Value) Then
                child.Field4.Text = ""
            Else
                child.Field4.Text = DataGridView1.SelectedCells(5).Value
            End If

            If IsDBNull(DataGridView1.SelectedCells(6).Value) Then
                child.Field5.Text = ""
            Else
                child.Field5.Text = DataGridView1.SelectedCells(6).Value
            End If

            If IsDBNull(DataGridView1.SelectedCells(7).Value) Then
                child.Field6.Text = ""
            Else
                child.Field6.Text = DataGridView1.SelectedCells(7).Value
            End If

            If IsDBNull(DataGridView1.SelectedCells(8).Value) Then
                child.Field7.Text = ""
            Else
                child.Field7.Text = DataGridView1.SelectedCells(8).Value
            End If

            If IsDBNull(DataGridView1.SelectedCells(9).Value) Then
                child.Field8.Text = ""
            Else
                child.Field8.Text = DataGridView1.SelectedCells(9).Value
            End If

            If IsDBNull(DataGridView1.SelectedCells(10).Value) Then
                child.Field9.Text = ""
            Else
                child.Field9.Text = DataGridView1.SelectedCells(10).Value
            End If

            If IsDBNull(DataGridView1.SelectedCells(11).Value) Then
                child.Field10.Text = ""
            Else
                child.Field10.Text = DataGridView1.SelectedCells(11).Value
            End If

            If IsDBNull(DataGridView1.SelectedCells(12).Value) Then
                child.Field11.Text = ""
            Else
                child.Field11.Text = DataGridView1.SelectedCells(12).Value
            End If

            If IsDBNull(DataGridView1.SelectedCells(13).Value) Then
                child.Field12.Text = ""
            Else
                child.Field12.Text = DataGridView1.SelectedCells(13).Value
            End If

            If IsDBNull(DataGridView1.SelectedCells(14).Value) Then
                child.Field13.Text = ""
            Else
                child.Field13.Text = DataGridView1.SelectedCells(14).Value
            End If

            If IsDBNull(DataGridView1.SelectedCells(15).Value) Then
                child.Field14.Text = ""
            Else
                child.Field14.Text = DataGridView1.SelectedCells(15).Value
            End If

            If IsDBNull(DataGridView1.SelectedCells(16).Value) Then
                child.Field15.Text = ""
            Else
                child.Field15.Text = DataGridView1.SelectedCells(16).Value
            End If

            If IsDBNull(DataGridView1.SelectedCells(17).Value) Then
                child.Field16.Text = ""
            Else
                child.Field16.Text = DataGridView1.SelectedCells(17).Value
            End If

            If IsDBNull(DataGridView1.SelectedCells(18).Value) Then
                child.Field17.Text = ""
            Else
                child.Field17.Text = DataGridView1.SelectedCells(18).Value
            End If

            If IsDBNull(DataGridView1.SelectedCells(19).Value) Then
                child.Field18.Text = ""
            Else
                child.Field18.Text = DataGridView1.SelectedCells(19).Value
            End If


            child.Old1.Text = child.Field1.Text
            child.Old2.Text = child.Field2.Text
            child.Old3.Text = child.Field3.Text
            child.Old4.Text = child.Field4.Text
            child.Old5.Text = child.Field5.Text
            child.Old6.Text = child.Field6.Text
            child.Old7.Text = child.Field7.Text
            child.Old8.Text = child.Field8.Text
            child.Old9.Text = child.Field9.Text
            child.Old10.Text = child.Field10.Text
            child.Old11.Text = child.Field11.Text
            child.Old12.Text = child.Field12.Text
            child.Old13.Text = child.Field13.Text
            child.Old14.Text = child.Field14.Text
            child.Old15.Text = child.Field15.Text
            child.Old16.Text = child.Field16.Text
            child.Old17.Text = child.Field17.Text
            child.Old18.Text = child.Field18.Text
        End If
    End Sub

    Private Sub Button14_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button14.Click
        If DataGridView2.RowCount = "0" Then
            'do nothing
        Else

            Dim ir As MsgBoxResult

            ir = MsgBox("Are you sure you would you like to mark these trades as uploaded?", MsgBoxStyle.YesNo + MsgBoxStyle.Information, "Change Status")

            If ir = MsgBoxResult.Yes Then
                Dim Mycn As OleDb.OleDbConnection
                Dim Command As OleDb.OleDbCommand
                Dim Command1 As OleDb.OleDbCommand
                Dim SQLstr As String
                Try
                    Mycn = New OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & My.Settings.dbloc & "';Jet OLEDB:Database Password='" & My.Settings.dbpass & "';")
                    Mycn.Open()

                    SQLstr = "INSERT INTO uma_tradetracker(MoxyOrderID, DateReady, ReadyBy, TradeUploaded, UploadFile)" & _
                    "SELECT MoxyOrderID,#" & Format(Now()) & "#,'" & Environ("USERNAME") & "',-1,'**UNKNOWN**' FROM uma_envestnet_TradeReady"

                    Command = New OleDb.OleDbCommand(SQLstr, Mycn)
                    Command.ExecuteNonQuery()

                    SQLstr = "DELETE * FROM uma_envestnet_TradeReady"

                    Command1 = New OleDb.OleDbCommand(SQLstr, Mycn)
                    Command1.ExecuteNonQuery()

                    Mycn.Close()

                    Call LoadPendingTradesGrid()
                    Call LoadReadyTrades()

                Catch ex As Exception
                    MsgBox(ex.Message, MsgBoxStyle.Critical, "ERROR")
                End Try
            Else

            End If
        End If
    End Sub

    Private Sub MarkAsUploadedToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MarkAsUploadedToolStripMenuItem.Click
        If DataGridView2.RowCount = "0" Then
            'do nothing
        Else
            Dim ir As MsgBoxResult

            ir = MsgBox("Are you sure you would you like to mark this trade as uploaded?", MsgBoxStyle.YesNo + MsgBoxStyle.Information, "Change Status")

            If ir = MsgBoxResult.Yes Then
                Dim Mycn As OleDb.OleDbConnection
                Dim Command As OleDb.OleDbCommand
                Dim Command1 As OleDb.OleDbCommand
                Dim SQLstr As String
                Try
                    Mycn = New OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & My.Settings.dbloc & "';Jet OLEDB:Database Password='" & My.Settings.dbpass & "';")
                    Mycn.Open()

                    SQLstr = "INSERT INTO uma_tradetracker(MoxyOrderID, DateReady, ReadyBy, TradeUploaded, UploadFile)" & _
                    "SELECT MoxyOrderID,#" & Format(Now()) & "#,'" & Environ("USERNAME") & "',-1,'**UNKNOWN**' FROM uma_envestnet_TradeReady WHERE ID = " & DataGridView2.SelectedCells(0).Value

                    Command = New OleDb.OleDbCommand(SQLstr, Mycn)
                    Command.ExecuteNonQuery()

                    SQLstr = "DELETE * FROM uma_envestnet_TradeReady WHERE ID = " & DataGridView2.SelectedCells(0).Value

                    Command1 = New OleDb.OleDbCommand(SQLstr, Mycn)
                    Command1.ExecuteNonQuery()

                    Mycn.Close()

                    Call LoadPendingTradesGrid()
                    Call LoadReadyTrades()

                Catch ex As Exception
                    MsgBox(ex.Message, MsgBoxStyle.Critical, "ERROR")
                End Try
            Else

            End If
        End If
    End Sub

    Private Sub Button7_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button7.Click
        'Timer5.Enabled = False
        'Button7.BackColor = Color.White
        Button7.ForeColor = Color.Black
        CheckBox2.Checked = False
        Dim child As New UMAExceptons
        child.MdiParent = Home
        child.Show()
        Call child.LoadExceptions()
    End Sub

    Private Sub Button15_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button15.Click
        Dim child As New UMAPortfolio
        child.MdiParent = Home
        child.Show()
        child.ID.Text = "NEW"
    End Sub

    Private Sub Button16_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button16.Click
        Dim child As New UMAPortfolioWorker
        child.MdiParent = Home
        child.Show()
    End Sub

    Private Sub Button17_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button17.Click
        Dim child As New UMAProductWorker
        child.MdiParent = Home
        child.Show()
    End Sub

    Private Sub Button20_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button20.Click
        Dim child As New UMATradeTracker
        child.MdiParent = Home
        child.Show()
    End Sub

    Private Sub Button19_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button19.Click
        Dim child As New UMAChangeLog
        child.MdiParent = Home
        child.Show()
    End Sub

    Private Sub Button18_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button18.Click
        Dim ir As MsgBoxResult

        ir = MsgBox("Clearing the Trade Pending table may have a negative impact on system functionality.  Only use this function if you are sure you need the table cleared before you can proceed.  Are you sure you want to delete everything on this table?", MsgBoxStyle.YesNo + MsgBoxStyle.Critical, "Are you sure?")

        If ir = MsgBoxResult.Yes Then
            Dim Mycn As OleDb.OleDbConnection

            Dim Command1 As OleDb.OleDbCommand
            Dim SQLstr As String
            Try
                Mycn = New OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & My.Settings.dbloc & "';Jet OLEDB:Database Password='" & My.Settings.dbpass & "';")
                Mycn.Open()

                SQLstr = "DELETE * FROM uma_envestnet_TradePending"

                Command1 = New OleDb.OleDbCommand(SQLstr, Mycn)
                Command1.ExecuteNonQuery()

                Mycn.Close()
            Catch ex As Exception
                MsgBox(ex.Message, MsgBoxStyle.Critical, "ERROR")
            End Try
        Else

        End If
    End Sub

    Private Sub Button24_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button24.Click
        Dim ir As MsgBoxResult

        ir = MsgBox("Clearing the Trade Ready table may have a negative impact on system functionality.  Only use this function if you are sure you need the table cleared before you can proceed.  Are you sure you want to delete everything on this table?", MsgBoxStyle.YesNo + MsgBoxStyle.Critical, "Are you sure?")

        If ir = MsgBoxResult.Yes Then
            Dim Mycn As OleDb.OleDbConnection

            Dim Command1 As OleDb.OleDbCommand
            Dim SQLstr As String
            Try
                Mycn = New OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & My.Settings.dbloc & "';Jet OLEDB:Database Password='" & My.Settings.dbpass & "';")
                Mycn.Open()

                SQLstr = "DELETE * FROM uma_envestnet_TradeReady"

                Command1 = New OleDb.OleDbCommand(SQLstr, Mycn)
                Command1.ExecuteNonQuery()

                Mycn.Close()
            Catch ex As Exception
                MsgBox(ex.Message, MsgBoxStyle.Critical, "ERROR")
            End Try
        Else

        End If
    End Sub

    Private Sub Button23_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button23.Click
        Dim child As New UMATransCodeTranslator
        child.MdiParent = Home
        child.Show()
    End Sub

    Private Sub Button22_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button22.Click
        Dim child As New UMASecurityTypeTranslator
        child.MdiParent = Home
        child.Show()
    End Sub

    Private Sub Button21_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button21.Click
        Control.CheckForIllegalCrossThreadCalls = False
        transoverview = New System.Threading.Thread(AddressOf umaoverview)
        transoverview.Start()
    End Sub

    Public Sub umaoverview()
        Call LoadProductsOverview()
        Call LoadProductsOverview()
        Call LoadAccountsOVerview()
        Call FindAUMOverview()
        Call LoadTradeCount()
        Call LoadEFCount()
        Call LoadERCount()
        Call LoadIRCount()
    End Sub

    Public Sub LoadAccountsOVerview()
        Try

            Dim conn As New OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & My.Settings.dbloc & "';Jet OLEDB:Database Password='" & My.Settings.dbpass & "';")

            Dim strSQL As String = "SELECT uma_portfolios_envestnet.ID, uma_portfolios_envestnet.PortfolioCode As [Portfolio Code], dbo_vQbRowDefPortfolio.ReportHeading1 As [Account Name], dbo_vQbRowDefPortfolio.AAMRepName As [AAM Rep], dbo_vQbRowDefPortfolio.ExternalAdvisor As [External Advisor], dbo_vQbRowDefPortfolio.ExternalFirm As [External Firm], dbo_vQbRowDefPortfolio.TotalMarketValue As [Market Value], dbo_vQbRowDefPortfolio.TotalTradableCash As [Cash], uma_products_envestnet.ProductID As [UMA Product ID], uma_products_envestnet.Discipline As [Discipline]" & _
            " FROM (uma_portfolios_envestnet INNER JOIN uma_products_envestnet ON uma_portfolios_envestnet.ProductID = uma_products_envestnet.ID) INNER JOIN dbo_vQbRowDefPortfolio ON uma_portfolios_envestnet.PortfolioCode = dbo_vQbRowDefPortfolio.PortfolioCode" & _
            " WHERE uma_portfolios_envestnet.Active = True"

            Dim da As New OleDb.OleDbDataAdapter(strSQL, conn)
            Dim ds As New DataSet
            da.Fill(ds, "Users")

            With DataGridView4
                .DataSource = ds.Tables("Users")
                .Columns(0).Visible = False
                .Columns(6).DefaultCellStyle.Format = "c"
                .Columns(7).DefaultCellStyle.Format = "c"
                .Columns(6).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                .Columns(7).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
            End With

            TextBox6.Text = DataGridView4.RowCount

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "ERROR")
        End Try
    End Sub

    Public Sub LoadProductsOverview()
        Try

            Dim conn As New OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & My.Settings.dbloc & "';Jet OLEDB:Database Password='" & My.Settings.dbpass & "';")

            Dim strSQL As String = "SELECT * FROM uma_products_envestnet"

            Dim da As New OleDb.OleDbDataAdapter(strSQL, conn)
            Dim ds As New DataSet
            da.Fill(ds, "Users")

            With DataGridView3
                .DataSource = ds.Tables("Users")
                .Columns(0).Visible = False
            End With

            TextBox8.Text = DataGridView3.RowCount

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "ERROR")
        End Try
    End Sub

    Public Sub FindAUMOverview()
        Dim Mycn1 As OleDb.OleDbConnection

        Try

            Mycn1 = New OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & My.Settings.dbloc & "';Jet OLEDB:Database Password='" & My.Settings.dbpass & "';")

            Mycn1.Open()

            Dim sqlstring As String

            sqlstring = "SELECT Sum(dbo_vQbRowDefPortfolio.TotalMarketValue) AS SumOfTotalMarketValue" & _
            " FROM uma_portfolios_envestnet INNER JOIN dbo_vQbRowDefPortfolio ON uma_portfolios_envestnet.PortfolioCode = dbo_vQbRowDefPortfolio.PortfolioCode" & _
            " WHERE uma_portfolios_envestnet.Active = True"

            Dim queryString As String = String.Format(sqlstring)
            Dim cmd As New OleDb.OleDbCommand(queryString, Mycn1)
            Dim da As New OleDb.OleDbDataAdapter(cmd)
            Dim ds As New DataSet

            da.Fill(ds, "User")
            Dim dt As DataTable = ds.Tables("User")
            If dt.Rows.Count > 0 Then

                Dim row As DataRow = dt.Rows(0)

                Dim mv As Double

                mv = (row("SumOfTotalMarketValue"))

                TextBox7.Text = Format(mv, "$#,###.##")

            Else

            End If
            Mycn1.Close()
            Mycn1.Dispose()


        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "ERROR")
        End Try
    End Sub

    Public Sub LoadIRCount()
        Dim Mycn1 As OleDb.OleDbConnection

        Try

            Mycn1 = New OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & My.Settings.dbloc & "';Jet OLEDB:Database Password='" & My.Settings.dbpass & "';")

            Mycn1.Open()

            Dim sqlstring As String

            sqlstring = "SELECT Count(dbo_vQBRowDefPortfolio.AAMRepName) As CountOfIR" & _
            " FROM(dbo_vQBRowDefPortfolio)" & _
            " WHERE (dbo_vQBRowDefPortfolio.PortfolioCode IN (Select PortfolioCode FROM uma_portfolios_envestnet WHERE uma_portfolios_envestnet.Active = True))" & _
            " GROUP BY dbo_vQBRowDefPortfolio.AAMRepName"

            Dim queryString As String = String.Format(sqlstring)
            Dim cmd As New OleDb.OleDbCommand(queryString, Mycn1)
            Dim da As New OleDb.OleDbDataAdapter(cmd)
            Dim ds As New DataSet

            da.Fill(ds, "User")
            Dim dt As DataTable = ds.Tables("User")
            If dt.Rows.Count > 0 Then

                Dim row As DataRow = dt.Rows(0)

                InternalReps.Text = dt.Rows.Count

            Else

            End If
            Mycn1.Close()
            Mycn1.Dispose()


        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "ERROR")
        End Try
    End Sub


    Public Sub LoadERCount()
        Dim Mycn1 As OleDb.OleDbConnection

        Try

            Mycn1 = New OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & My.Settings.dbloc & "';Jet OLEDB:Database Password='" & My.Settings.dbpass & "';")

            Mycn1.Open()

            Dim sqlstring As String

            sqlstring = "SELECT Count(dbo_vQBRowDefPortfolio.ExternalAdvisor) As CountOfER" & _
            " FROM(dbo_vQBRowDefPortfolio)" & _
            " WHERE (dbo_vQBRowDefPortfolio.PortfolioCode IN (Select PortfolioCode FROM uma_portfolios_envestnet WHERE uma_portfolios_envestnet.Active = True))" & _
            " GROUP By dbo_vQBRowDefPortfolio.ExternalAdvisor"

            Dim queryString As String = String.Format(sqlstring)
            Dim cmd As New OleDb.OleDbCommand(queryString, Mycn1)
            Dim da As New OleDb.OleDbDataAdapter(cmd)
            Dim ds As New DataSet

            da.Fill(ds, "User")
            Dim dt As DataTable = ds.Tables("User")
            If dt.Rows.Count > 0 Then

                Dim row As DataRow = dt.Rows(0)

                ExternalReps.Text = dt.Rows.Count

            Else

            End If
            Mycn1.Close()
            Mycn1.Dispose()


        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "ERROR")
        End Try
    End Sub

    Public Sub LoadEFCount()
        Dim Mycn1 As OleDb.OleDbConnection

        Try

            Mycn1 = New OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & My.Settings.dbloc & "';Jet OLEDB:Database Password='" & My.Settings.dbpass & "';")

            Mycn1.Open()

            Dim sqlstring As String

            sqlstring = "SELECT Count(dbo_vQBRowDefPortfolio.ExternalFirm) As CountOfEF" & _
            " FROM(dbo_vQBRowDefPortfolio)" & _
            " WHERE (dbo_vQBRowDefPortfolio.PortfolioCode IN (Select PortfolioCode FROM uma_portfolios_envestnet WHERE uma_portfolios_envestnet.Active = True))" & _
            " GROUP BY dbo_vQBRowDefPortfolio.ExternalFirm"

            Dim queryString As String = String.Format(sqlstring)
            Dim cmd As New OleDb.OleDbCommand(queryString, Mycn1)
            Dim da As New OleDb.OleDbDataAdapter(cmd)
            Dim ds As New DataSet

            da.Fill(ds, "User")
            Dim dt As DataTable = ds.Tables("User")
            If dt.Rows.Count > 0 Then

                Dim row As DataRow = dt.Rows(0)

                ExternalFirms.Text = dt.Rows.Count

            Else

            End If
            Mycn1.Close()
            Mycn1.Dispose()


        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "ERROR")
        End Try
    End Sub

    Public Sub LoadTradeCount()
        Dim Mycn1 As OleDb.OleDbConnection

        Try

            Mycn1 = New OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & My.Settings.dbloc & "';Jet OLEDB:Database Password='" & My.Settings.dbpass & "';")

            Mycn1.Open()

            Dim dte1 As Date
            Dim dte2 As String
            Dim dte3 As Date
            dte3 = Format(Now())
            dte1 = DateAdd("d", -1, dte3)
            dte2 = Format(dte1, "dddd")

            Do Until dte2 = "Monday" Or dte2 = "Tuesday" Or dte2 = "Wednesday" Or dte2 = "Thursday" Or dte2 = "Friday"
                dte1 = DateAdd("d", -1, dte1)
                dte2 = Format(dte1, "dddd")
            Loop

            Dim sqlstring As String

            sqlstring = "SELECT Count(uma_tradetracker.ID) As CountOfID" & _
            " FROM(uma_tradetracker)" & _
            " WHERE (((uma_tradetracker.DateReady)>#" & dte1 & "#));"

            Dim queryString As String = String.Format(sqlstring)
            Dim cmd As New OleDb.OleDbCommand(queryString, Mycn1)
            Dim da As New OleDb.OleDbDataAdapter(cmd)
            Dim ds As New DataSet

            da.Fill(ds, "User")
            Dim dt As DataTable = ds.Tables("User")
            If dt.Rows.Count > 0 Then

                Dim row As DataRow = dt.Rows(0)

                TextBox9.Text = dt.Rows.Count

            Else

            End If
            Mycn1.Close()
            Mycn1.Dispose()


        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "ERROR")
        End Try
    End Sub

    Private Sub Button25_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub

    Private Sub Button26_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button26.Click
        'TextBox10.Text = "Loading Interface Positions."
        'Call LoadReconPositions()

        DataGridView5.Visible = False
        TextBox11.Visible = True
        TextBox10.Visible = True
        Timer3.Enabled = True

        'Control.CheckForIllegalCrossThreadCalls = False
        'reconthread = New System.Threading.Thread(AddressOf LoadReconPositions)
        'reconthread.Start()

        Call LoadReconPositions()
        TextBox10.Text = "Starting..."
        'Call LoadPendingTrades()
        'Call LoadReconPositions()

    End Sub

    Public Sub LoadReconPositions()
        'Load ReconPositions
        Dim Mycn As OleDb.OleDbConnection

        Try
            Mycn = New OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & My.Settings.dbloc & "';Jet OLEDB:Database Password='" & My.Settings.dbpass & "';")
            Mycn.Open()

            TextBox10.Text = "Priming Interface Tables."
            Dim Command1 As OleDb.OleDbCommand
            Dim SQLstr As String
            SQLstr = "DELETE * FROM uma_PositionRecon"
            Command1 = New OleDb.OleDbCommand(SQLstr, Mycn)
            Command1.ExecuteNonQuery()
            TextBox10.Text = "Interface Tables Primed."

            TextBox10.Text = "Loading Interface Positions."
            Dim Command2 As OleDb.OleDbCommand
            Dim SQLstr2 As String
            SQLstr2 = "INSERT INTO uma_PositionRecon ( PortfolioCode, SecType, Symbol, Quantity, MarketValue, PositionDate )" & _
            " SELECT uma_portfolios_envestnet.PortfolioCode, AdvApp_vSecurity.SecTypeBaseCode & AdvApp_vSecurity.PrincipalCurrencyCode, AdvApp_vSecurity.Symbol, AdvApp_vPositionRecon.Quantity, AdvApp_vPositionRecon.MarketValue, AdvApp_vPositionRecon.PositionDate" & _
            " FROM (AdvApp_vPositionRecon INNER JOIN uma_portfolios_envestnet ON AdvApp_vPositionRecon.PortfolioCode = uma_portfolios_envestnet.PortfolioCode) INNER JOIN AdvApp_vSecurity ON AdvApp_vPositionRecon.SecurityID = AdvApp_vSecurity.SecurityID;"
            Command2 = New OleDb.OleDbCommand(SQLstr2, Mycn)
            Command2.ExecuteNonQuery()
            TextBox10.Text = "Interface Positions Loaded."

            TextBox10.Text = "Priming APX Tables."
            Dim Command3 As OleDb.OleDbCommand
            Dim SQLstr3 As String
            SQLstr3 = "DELETE * FROM uma_Position"
            Command3 = New OleDb.OleDbCommand(SQLstr3, Mycn)
            Command3.ExecuteNonQuery()
            TextBox10.Text = "APX Tables Primed."

            TextBox10.Text = "Loading APX Positions."
            Dim Command4 As OleDb.OleDbCommand
            Dim SQLstr4 As String
            SQLstr4 = "INSERT INTO uma_Position ( PortfolioCode, SecType, Symbol, Quantity, MarketValue )" & _
            " SELECT uma_portfolios_envestnet.PortfolioCode, dbo_vQbRowDefPosition.SecTypeCode, dbo_vQbRowDefPosition.Symbol, dbo_vQbRowDefPosition.Quantity,dbo_vQbRowDefPosition.MarketValue" & _
            " FROM dbo_vQbRowDefPosition INNER JOIN uma_portfolios_envestnet ON dbo_vQbRowDefPosition.PortfolioCode = uma_portfolios_envestnet.PortfolioCode;"
            Command4 = New OleDb.OleDbCommand(SQLstr4, Mycn)
            Command4.ExecuteNonQuery()
            TextBox10.Text = "APX Positions Loaded."

            TextBox10.Text = "Priming Position Load Tables."
            Dim Command11 As OleDb.OleDbCommand
            Dim SQLstr11 As String
            SQLstr11 = "DELETE * FROM uma_PositionLoad WHERE Approved = False"
            Command11 = New OleDb.OleDbCommand(SQLstr11, Mycn)
            Command11.ExecuteNonQuery()
            TextBox10.Text = "Position Load Tables Primed."

            TextBox10.Text = "Loading Cash Breaks - lo's."
            Dim Command5 As OleDb.OleDbCommand
            Dim SQLstr5 As String
            SQLstr5 = "INSERT INTO uma_PositionLoad ( PortfolioCode, Symbol, InterfaceQnty, InterfaceMarketVal, APXQnty, APXMarketVal, InterfaceDate, QntyDiff, MarketValDiff, [Action], SecType )" & _
            " SELECT uma_PositionRecon.PortfolioCode, uma_PositionRecon.Symbol, uma_PositionRecon.Quantity, uma_PositionRecon.MarketValue, uma_Position.Quantity, uma_Position.MarketValue, uma_PositionRecon.PositionDate, 0 AS QntyDiff, (uma_Position.MarketValue-uma_PositionRecon.MarketValue) AS MarketValDiff, 'lo' AS Expr1, uma_PositionRecon.SecType" & _
            " FROM uma_PositionRecon INNER JOIN uma_Position ON (uma_PositionRecon.Symbol = uma_Position.Symbol) AND (uma_PositionRecon.PortfolioCode = uma_Position.PortfolioCode)" & _
            " WHERE(((uma_PositionRecon.SecType) = 'ca') And ((uma_PositionRecon.MarketValue) < [uma_Position].[MarketValue]))" & _
            " GROUP BY uma_PositionRecon.PortfolioCode, uma_PositionRecon.Symbol, uma_PositionRecon.Quantity, uma_PositionRecon.MarketValue, uma_Position.Quantity, uma_Position.MarketValue, uma_PositionRecon.PositionDate, uma_PositionRecon.SecType;"
            Command5 = New OleDb.OleDbCommand(SQLstr5, Mycn)
            Command5.ExecuteNonQuery()
            TextBox10.Text = "Loading Cash Breaks - li's."
            Dim Command6 As OleDb.OleDbCommand
            Dim SQLstr6 As String
            SQLstr6 = "INSERT INTO uma_PositionLoad ( PortfolioCode, Symbol, InterfaceQnty, InterfaceMarketVal, APXQnty, APXMarketVal, InterfaceDate, QntyDiff, MarketValDiff, [Action], SecType )" & _
            " SELECT uma_PositionRecon.PortfolioCode, uma_PositionRecon.Symbol, uma_PositionRecon.Quantity, uma_PositionRecon.MarketValue, uma_Position.Quantity, uma_Position.MarketValue, uma_PositionRecon.PositionDate, 0 AS QntyDiff, (uma_PositionRecon.MarketValue-uma_Position.MarketValue) AS MarketValDiff, 'li' AS Expr1, uma_PositionRecon.SecType" & _
            " FROM uma_PositionRecon INNER JOIN uma_Position ON (uma_PositionRecon.Symbol = uma_Position.Symbol) AND (uma_PositionRecon.PortfolioCode = uma_Position.PortfolioCode)" & _
            " WHERE(((uma_PositionRecon.SecType) = 'ca') And ((uma_PositionRecon.MarketValue) > [uma_Position].[MarketValue]))" & _
            " GROUP BY uma_PositionRecon.PortfolioCode, uma_PositionRecon.Symbol, uma_PositionRecon.Quantity, uma_PositionRecon.MarketValue, uma_Position.Quantity, uma_Position.MarketValue, uma_PositionRecon.PositionDate, uma_PositionRecon.SecType;"
            Command6 = New OleDb.OleDbCommand(SQLstr6, Mycn)
            Command6.ExecuteNonQuery()
            TextBox10.Text = "Cash Breaks Loaded."

            TextBox10.Text = "Loading Position Breaks - lo's."
            Dim Command7 As OleDb.OleDbCommand
            Dim SQLstr7 As String
            SQLstr7 = "INSERT INTO uma_PositionLoad ( PortfolioCode, Symbol, InterfaceQnty, InterfaceMarketVal, APXQnty, APXMarketVal, InterfaceDate, QntyDiff, MarketValDiff, [Action], SecType )" & _
            " SELECT uma_PositionRecon.PortfolioCode, uma_PositionRecon.Symbol, uma_PositionRecon.Quantity, uma_PositionRecon.MarketValue, uma_Position.Quantity, uma_Position.MarketValue, uma_PositionRecon.PositionDate, (uma_Position.Quantity-uma_PositionRecon.Quantity) AS QntyDiff, (uma_Position.MarketValue-uma_PositionRecon.MarketValue) AS MarketValDiff, 'lo' AS Expr1, uma_PositionRecon.SecType" & _
            " FROM uma_PositionRecon INNER JOIN uma_Position ON (uma_PositionRecon.Symbol = uma_Position.Symbol) AND (uma_PositionRecon.PortfolioCode = uma_Position.PortfolioCode)" & _
            " WHERE(((uma_PositionRecon.SecType) <> 'ca') And ((uma_PositionRecon.Quantity) < [uma_Position].[Quantity]))" & _
            " GROUP BY uma_PositionRecon.PortfolioCode, uma_PositionRecon.Symbol, uma_PositionRecon.Quantity, uma_PositionRecon.MarketValue, uma_Position.Quantity, uma_Position.MarketValue, uma_PositionRecon.PositionDate, uma_PositionRecon.SecType;"
            Command7 = New OleDb.OleDbCommand(SQLstr7, Mycn)
            Command7.ExecuteNonQuery()
            TextBox10.Text = "Loading Position Breaks - li's."
            Dim Command8 As OleDb.OleDbCommand
            Dim SQLstr8 As String
            SQLstr8 = "INSERT INTO uma_PositionLoad ( PortfolioCode, Symbol, InterfaceQnty, InterfaceMarketVal, APXQnty, APXMarketVal, InterfaceDate, QntyDiff, MarketValDiff, [Action], SecType )" & _
            " SELECT uma_PositionRecon.PortfolioCode, uma_PositionRecon.Symbol, uma_PositionRecon.Quantity, uma_PositionRecon.MarketValue, uma_Position.Quantity, uma_Position.MarketValue, uma_PositionRecon.PositionDate, (uma_PositionRecon.Quantity-uma_Position.Quantity) AS QntyDiff, (uma_PositionRecon.MarketValue-uma_Position.MarketValue) AS MarketValDiff, 'li' AS Expr1, uma_PositionRecon.SecType" & _
            " FROM uma_PositionRecon INNER JOIN uma_Position ON (uma_PositionRecon.Symbol = uma_Position.Symbol) AND (uma_PositionRecon.PortfolioCode = uma_Position.PortfolioCode)" & _
            " WHERE(((uma_PositionRecon.SecType) <> 'ca') And ((uma_PositionRecon.Quantity) > [uma_Position].[Quantity]))" & _
            " GROUP BY uma_PositionRecon.PortfolioCode, uma_PositionRecon.Symbol, uma_PositionRecon.Quantity, uma_PositionRecon.MarketValue, uma_Position.Quantity, uma_Position.MarketValue, uma_PositionRecon.PositionDate, uma_PositionRecon.SecType;"
            Command8 = New OleDb.OleDbCommand(SQLstr8, Mycn)
            Command8.ExecuteNonQuery()
            TextBox10.Text = "Position Breaks Loaded."

            TextBox10.Text = "Loading Position only known to Interface."
            Dim Command9 As OleDb.OleDbCommand
            Dim SQLstr9 As String
            SQLstr9 = "INSERT INTO uma_PositionLoad ( PortfolioCode, Symbol, InterfaceQnty, InterfaceMarketVal, APXQnty, APXMarketVal, InterfaceDate, QntyDiff, MarketValDiff, [Action], SecType )" & _
            " SELECT uma_PositionRecon.PortfolioCode, uma_PositionRecon.Symbol, uma_PositionRecon.Quantity, uma_PositionRecon.MarketValue, 0, 0, uma_PositionRecon.PositionDate, uma_PositionRecon.Quantity, uma_PositionRecon.MarketValue, 'li', uma_PositionRecon.SecType" & _
            " FROM(uma_PositionRecon)" & _
            " WHERE uma_PositionRecon.Symbol Not In (Select Symbol FROM uma_Position WHERE PortfolioCode = uma_PositionRecon.PortfolioCode)"
            Command9 = New OleDb.OleDbCommand(SQLstr9, Mycn)
            Command9.ExecuteNonQuery()
            TextBox10.Text = "Loading Position Breaks - li's."
            Dim Command10 As OleDb.OleDbCommand
            Dim SQLstr10 As String
            SQLstr10 = "INSERT INTO uma_PositionLoad ( PortfolioCode, Symbol, InterfaceQnty, InterfaceMarketVal, APXQnty, APXMarketVal, QntyDiff, MarketValDiff, [Action], SecType )" & _
            " SELECT uma_Position.PortfolioCode, uma_Position.Symbol, 0, 0, uma_Position.Quantity, uma_Position.MarketValue, uma_Position.Quantity, uma_Position.MarketValue, 'lo', uma_Position.SecType" & _
            " FROM(uma_Position)" & _
            " WHERE uma_Position.Symbol Not In (Select Symbol FROM uma_PositionRecon WHERE PortfolioCode = uma_Position.PortfolioCode)"
            Command10 = New OleDb.OleDbCommand(SQLstr10, Mycn)
            Command10.ExecuteNonQuery()
            TextBox10.Text = "Position Breaks Loaded."

            TextBox10.Text = "Cleaning table data."
            Dim Command12 As OleDb.OleDbCommand
            Dim SQLstr12 As String
            SQLstr12 = "DELETE * FROM uma_PositionLoad WHERE ((QntyDiff < 1) AND (MarketValDiff < 1))"
            Command12 = New OleDb.OleDbCommand(SQLstr12, Mycn)
            Command12.ExecuteNonQuery()
            TextBox10.Text = "Table data cleaned."

            TextBox10.Text = "Removing records already approved."
            Dim Command13 As OleDb.OleDbCommand
            Dim SQLstr13 As String
            SQLstr13 = "DELETE * FROM uma_PositionLoad WHERE ID In (SELECT uma_PositionLoad.ID" & _
            " FROM uma_PositionLoad INNER JOIN uma_PositionLoad AS uma_PositionLoad_1 ON (uma_PositionLoad.Symbol = uma_PositionLoad_1.Symbol) AND (uma_PositionLoad.SecType = uma_PositionLoad_1.SecType) AND (uma_PositionLoad.PortfolioCode = uma_PositionLoad_1.PortfolioCode)" & _
            " WHERE(((uma_PositionLoad.Approved) = False) And ((uma_PositionLoad_1.Approved) = True))" & _
            " GROUP BY uma_PositionLoad.ID)"

            Command13 = New OleDb.OleDbCommand(SQLstr13, Mycn)
            Command13.ExecuteNonQuery()
            TextBox10.Text = "Records Cleared."

            TextBox10.Text = "Drawing Grid."

            Call LoadReconBreaks()

            Mycn.Close()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "ERROR")
        End Try

    End Sub

    Public Sub LoadReconBreaks()
        Try

            Dim conn As New OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & My.Settings.dbloc & "';Jet OLEDB:Database Password='" & My.Settings.dbpass & "';")

            Dim strSQL As String = "SELECT *" & _
            " FROM uma_PositionLoad;"

            Dim da As New OleDb.OleDbDataAdapter(strSQL, conn)
            Dim ds As New DataSet
            da.Fill(ds, "Users")

            With DataGridView5
                .DataSource = ds.Tables("Users")
                .Columns(0).Visible = False
            End With

            DataGridView5.Visible = True
            TextBox11.Visible = False
            TextBox10.Visible = False
            Timer3.Enabled = False

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "ERROR")
        End Try
    End Sub

    Private Sub Timer3_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer3.Tick
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

    Private Sub Button25_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button25.Click
        If DataGridView5.RowCount = "0" Then
            MsgBox("No breaks found.  Please check for breaks before running Auto Recon.", MsgBoxStyle.Information, "No Data")
        Else
            DataGridView6.Visible = False
            TextBox13.Visible = True
            TextBox12.Visible = True
            Timer4.Enabled = True

            'Control.CheckForIllegalCrossThreadCalls = False
            'recontranslate = New System.Threading.Thread(AddressOf TradeBreakTranslate)
            'recontranslate.Start()
            Call TradeBreakTranslate()
            TextBox12.Text = "Starting..."
        End If
    End Sub

    Public Sub TradeBreakTranslate()
        Dim Mycn As OleDb.OleDbConnection
        Dim Command As OleDb.OleDbCommand
        Dim Command1 As OleDb.OleDbCommand
        Dim Command2 As OleDb.OleDbCommand
        Dim Command3 As OleDb.OleDbCommand
        Dim SQLstr As String
        'Dim dte As Date
        'dte = Format(Now())
        'Dim ystrd As Date = DateAdd(DateInterval.Day, -1, dte)
        Dim cdte As String
        'cdte = Format(ystrd, "MMddyyyy")

        Dim dte1 As Date
        Dim dte2 As String
        Dim dte3 As Date
        dte3 = Format(Now())
        dte1 = DateAdd("d", -1, dte3)
        dte2 = Format(dte1, "dddd")

        Do Until dte2 = "Monday" Or dte2 = "Tuesday" Or dte2 = "Wednesday" Or dte2 = "Thursday" Or dte2 = "Friday"
            dte1 = DateAdd("d", -1, dte1)
            dte2 = Format(dte1, "dddd")
        Loop

        Dim mnth As Integer
        Dim day1 As Integer
        Dim yer As Integer

        mnth = DatePart(DateInterval.Month, dte1)
        day1 = DatePart(DateInterval.Day, dte1)
        yer = DatePart(DateInterval.Year, dte1)

        If mnth <= 9 Then
            cdte = "0" & mnth & day1 & yer
        Else
            cdte = mnth & day1 & yer
        End If

        Try
            Mycn = New OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & My.Settings.dbloc & "';Jet OLEDB:Database Password='" & My.Settings.dbpass & "';")
            Dim ds1 As New DataSet
            Dim eds1 As New DataGridView
            Dim dv1 As New DataView
            Mycn.Open()

            'TextBox12.Text = "Priming Translation Table..."
            'SQLstr = "DELETE * FROM uma_tradeblotter;"
            'Command3 = New OleDb.OleDbCommand(SQLstr, Mycn)
            'Command3.ExecuteNonQuery()
            'TextBox12.Text = "Translation Table Primed..."

            TextBox12.Text = "Translating security li's..."
            SQLstr = "Insert INTO uma_tradeblotter([Field1], [Field2], [Field4], [Field5], [Field6], [Field9], [Field17], [Field18], [Field29], [Field30], [Field42], [Field45], [Field46], [Field59])" & _
            " SELECT uma_PositionLoad.PortfolioCode, uma_PositionLoad.Action, uma_PositionLoad.SecType, uma_PositionLoad.Symbol, '" & cdte & "', uma_PositionLoad.QntyDiff,'y',uma_PositionLoad.MarketValDiff,'n','65533','1','n','y','y'" & _
            " FROM uma_PositionLoad" & _
            " WHERE ((uma_PositionLoad.Action = 'li') AND (uma_PositionLoad.Approved = True) AND (uma_PositionLoad.SecType <> 'caus'))" & _
            " GROUP BY uma_PositionLoad.PortfolioCode, uma_PositionLoad.Action, uma_PositionLoad.SecType, uma_PositionLoad.Symbol, uma_PositionLoad.QntyDiff,uma_PositionLoad.MarketValDiff;"
            Command = New OleDb.OleDbCommand(SQLstr, Mycn)
            Command.ExecuteNonQuery()
            TextBox12.Text = "Finished Translating security li's..."

            TextBox12.Text = "Translating security lo's..."
            SQLstr = "Insert INTO uma_tradeblotter([Field1], [Field2], [Field4], [Field5], [Field6], [Field9], [Field17], [Field18], [Field29], [Field42], [Field45], [Field46], [Field59])" & _
            " SELECT uma_PositionLoad.PortfolioCode, uma_PositionLoad.Action, uma_PositionLoad.SecType, uma_PositionLoad.Symbol, '" & cdte & "', uma_PositionLoad.QntyDiff,'y',uma_PositionLoad.MarketValDiff,'n','1','n','y','y'" & _
            " FROM uma_PositionLoad" & _
            " WHERE ((uma_PositionLoad.Action = 'lo') AND (uma_PositionLoad.Approved = True) AND (uma_PositionLoad.SecType <> 'caus'))" & _
            " GROUP BY uma_PositionLoad.PortfolioCode, uma_PositionLoad.Action, uma_PositionLoad.SecType, uma_PositionLoad.Symbol, uma_PositionLoad.QntyDiff,uma_PositionLoad.MarketValDiff;"
            Command2 = New OleDb.OleDbCommand(SQLstr, Mycn)
            Command2.ExecuteNonQuery()
            TextBox12.Text = "Finished Translating security lo's..."

            TextBox12.Text = "Translating cash transactions..."
            SQLstr = "Insert INTO uma_tradeblotter([Field1], [Field2], [Field4], [Field5], [Field6], [Field17], [Field18], [Field42], [Field45], [Field46])" & _
            " SELECT uma_PositionLoad.PortfolioCode, uma_PositionLoad.Action, uma_PositionLoad.SecType, uma_PositionLoad.Symbol, '" & cdte & "','y',uma_PositionLoad.MarketValDiff,'1','n','y'" & _
            " FROM uma_PositionLoad" & _
            " WHERE uma_PositionLoad.SecType = 'caus' AND (uma_PositionLoad.Approved = True)" & _
            " GROUP BY uma_PositionLoad.PortfolioCode, uma_PositionLoad.Action, uma_PositionLoad.SecType, uma_PositionLoad.Symbol, uma_PositionLoad.MarketValDiff;"
            Command1 = New OleDb.OleDbCommand(SQLstr, Mycn)
            Command1.ExecuteNonQuery()
            TextBox12.Text = "Finished Translating cash transactions..."


            TextBox12.Text = "Cleaning Recon Load Table..."
            SQLstr = "DELETE * FROM uma_PositionLoad WHERE Approved = True"
            Command3 = New OleDb.OleDbCommand(SQLstr, Mycn)
            Command3.ExecuteNonQuery()
            TextBox12.Text = "Table Cleaned..."


            TextBox12.Text = "Drawing Grid..."
            Call DrawReconTransGrid()
            Call LoadReconBreaks()

            Mycn.Close()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "ERROR")
        End Try
    End Sub

    Public Sub DrawReconTransGrid()
        Try

            Dim conn As New OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & My.Settings.dbloc & "';Jet OLEDB:Database Password='" & My.Settings.dbpass & "';")

            Dim strSQL As String = "SELECT *" & _
            " FROM uma_tradeblotter;"

            Dim da As New OleDb.OleDbDataAdapter(strSQL, conn)
            Dim ds As New DataSet
            da.Fill(ds, "Users")

            With DataGridView6
                .DataSource = ds.Tables("Users")
                .Columns(0).Visible = False
            End With

            DataGridView6.Visible = True
            TextBox13.Visible = False
            TextBox12.Visible = False
            Timer4.Enabled = False

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "ERROR")
        End Try
    End Sub

    Private Sub Timer4_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer4.Tick
        If TextBox13.Text = "Working" Then
            TextBox13.Text = "Working."
        Else
            If TextBox13.Text = "Working." Then
                TextBox13.Text = "Working.."
            Else
                If TextBox13.Text = "Working.." Then
                    TextBox13.Text = "Working..."
                Else
                    If TextBox13.Text = "Working..." Then
                        TextBox13.Text = "Working"
                    End If
                End If
            End If

        End If
    End Sub

    Private Sub Button27_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button27.Click
        Call DrawReconTransGrid()
    End Sub

    Public Sub ExportReconTrade()
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

        Call ClearReconExport()
        Call DrawReconTransGrid()

    End Sub

    Private Sub Button28_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button28.Click
        If DataGridView6.RowCount = "0" Then
            MsgBox("Nothing to send to APX.  Please run Auto Recon before attempting to export.", MsgBoxStyle.Information, "No Data")
        Else
            Call ExportReconTrade()
        End If
    End Sub

    Public Sub ClearReconExport()
        Dim Mycn As OleDb.OleDbConnection
        Dim Command As OleDb.OleDbCommand
        Dim SQLstr As String

        Try
            Mycn = New OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & My.Settings.dbloc & "';Jet OLEDB:Database Password='" & My.Settings.dbpass & "';")
            Dim ds1 As New DataSet
            Dim eds1 As New DataGridView
            Dim dv1 As New DataView
            Mycn.Open()

            SQLstr = "DELETE * FROM uma_tradeblotter;"
            Command = New OleDb.OleDbCommand(SQLstr, Mycn)
            Command.ExecuteNonQuery()

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "ERROR")
        End Try
    End Sub

    Private Sub TabPage1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TabPage1.Click

    End Sub

    Private Sub Button29_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button29.Click
        Call LoadReconBreaks()
    End Sub

    Private Sub ChangeApprovalStatusToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ChangeApprovalStatusToolStripMenuItem.Click
        If DataGridView5.RowCount = 0 Then

        Else
            If DataGridView5.SelectedCells(12).Value = True Then
                Call ChangeUMATradeUnApproved()
            Else
                Call ChangeUMATradeApproved()
            End If
        End If
    End Sub

    Public Sub ChangeUMATradeApproved()
        Dim Mycn As OleDb.OleDbConnection
        Dim Command As OleDb.OleDbCommand
        Dim SQLstr As String

        Try
            Mycn = New OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & My.Settings.dbloc & "';Jet OLEDB:Database Password='" & My.Settings.dbpass & "';")
            Dim ds1 As New DataSet
            Dim eds1 As New DataGridView
            Dim dv1 As New DataView
            Mycn.Open()

            SQLstr = "UPDATE uma_PositionLoad SET Approved = True WHERE ID = " & DataGridView5.SelectedCells(0).Value
            Command = New OleDb.OleDbCommand(SQLstr, Mycn)
            Command.ExecuteNonQuery()

            Call LoadReconBreaks()

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "ERROR")
        End Try
    End Sub

    Public Sub ChangeUMATradeUnApproved()
        Dim Mycn As OleDb.OleDbConnection
        Dim Command As OleDb.OleDbCommand
        Dim SQLstr As String

        Try
            Mycn = New OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & My.Settings.dbloc & "';Jet OLEDB:Database Password='" & My.Settings.dbpass & "';")
            Dim ds1 As New DataSet
            Dim eds1 As New DataGridView
            Dim dv1 As New DataView
            Mycn.Open()

            SQLstr = "UPDATE uma_PositionLoad SET Approved = False WHERE ID = " & DataGridView5.SelectedCells(0).Value
            Command = New OleDb.OleDbCommand(SQLstr, Mycn)
            Command.ExecuteNonQuery()

            Call LoadReconBreaks()

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "ERROR")
        End Try
    End Sub

    Private Sub ApproveAllToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ApproveAllToolStripMenuItem.Click
        Call ApproveAllBreaks()
    End Sub

    Public Sub ApproveAllBreaks()
        Dim ir As MsgBoxResult
        ir = MsgBox("Are you sure you want to approve all trade breaks?", MsgBoxStyle.YesNo + MsgBoxStyle.Exclamation, "Confirm Selection")
        If ir = MsgBoxResult.Yes Then

            Dim Mycn As OleDb.OleDbConnection
            Dim Command As OleDb.OleDbCommand
            Dim SQLstr As String

            Try
                Mycn = New OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & My.Settings.dbloc & "';Jet OLEDB:Database Password='" & My.Settings.dbpass & "';")
                Dim ds1 As New DataSet
                Dim eds1 As New DataGridView
                Dim dv1 As New DataView
                Mycn.Open()

                SQLstr = "UPDATE uma_PositionLoad SET Approved = True"
                Command = New OleDb.OleDbCommand(SQLstr, Mycn)
                Command.ExecuteNonQuery()

                Call LoadReconBreaks()

            Catch ex As Exception
                MsgBox(ex.Message, MsgBoxStyle.Critical, "ERROR")
            End Try

        Else

        End If

    End Sub

    Private Sub DeleteToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles DeleteToolStripMenuItem.Click
        If DataGridView6.RowCount = 0 Then
            'do nothing
        Else
            Dim ir As MsgBoxResult
            ir = MsgBox("Are you sure you want to delete this record?", MsgBoxStyle.YesNo + MsgBoxStyle.Information, "Confirm Delete Record")
            If ir = MsgBoxResult.Yes Then
                Dim Mycn As OleDb.OleDbConnection
                Dim Command As OleDb.OleDbCommand
                Dim SQLstr As String

                Try
                    Mycn = New OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & My.Settings.dbloc & "';Jet OLEDB:Database Password='" & My.Settings.dbpass & "';")
                    Dim ds1 As New DataSet
                    Dim eds1 As New DataGridView
                    Dim dv1 As New DataView
                    Mycn.Open()

                    SQLstr = "DELETE * FROM uma_tradeblotter WHERE ID = " & DataGridView6.SelectedCells(0).Value
                    Command = New OleDb.OleDbCommand(SQLstr, Mycn)
                    Command.ExecuteNonQuery()

                    Call DrawReconTransGrid()

                Catch ex As Exception
                    MsgBox(ex.Message, MsgBoxStyle.Critical, "ERROR")
                End Try
            Else
                'do nothing
            End If
        End If
    End Sub

    Private Sub Timer5_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer5.Tick
        If Button7.ForeColor = Color.Black And CheckBox2.Checked = True Then
            'Button7.BackColor = Color.Red
            Button7.ForeColor = Color.Red
            CheckBox2.Checked = False
        Else
            'If CheckBox2.Checked = False Then
            'Button7.BackColor = Color.White
            Button7.ForeColor = Color.Black
            'CheckBox2.Checked = True
            'Else
            'End If
        End If
    End Sub

    Private Sub Button31_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button31.Click
        OpenFileDialog3.Title = "Please Select a File"
        OpenFileDialog3.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.Desktop) '"C:"
        OpenFileDialog3.DefaultExt = "CSV"
        OpenFileDialog3.Filter = "CSV files (*.csv)|*.csv|All files (*.*)|*.*"
        Dim fdate As Date
        fdate = DateTimePicker1.Text
        OpenFileDialog3.FileName = "account_raise_cash_" & Format(fdate, "yyyy") & "_" & Format(fdate, "MM") & "_" & Format(fdate, "dd") & ".csv"
        OpenFileDialog3.ShowDialog()
    End Sub

    Private Sub OpenFileDialog3_FileOk(ByVal sender As System.Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles OpenFileDialog3.FileOk
        Dim strm As System.IO.Stream
        strm = OpenFileDialog3.OpenFile()

        TextBox4.Text = OpenFileDialog3.FileName.ToString

        If Not (strm Is Nothing) Then
            'insert code to read the file data
            strm.Close()
        End If
    End Sub
End Class