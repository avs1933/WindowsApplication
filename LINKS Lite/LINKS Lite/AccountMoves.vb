Imports System.IO
Imports System.Data.OleDb

Public Class AccountMoves
    Dim ProssVal As System.Threading.Thread
    Dim ProssSales As System.Threading.Thread

    Private Sub AccountMoves_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        If Permissions.EditAccountMovesSettings.Checked Then
            Call LoadExComposite()
            Call LoadExGrid()
            Call LoadMAPComposite()
            Call LoadMAPDisciplines()
            Call LoadMapGrid()
            Call LoadTerritoryCB()
            Call LoadDepartmentGrid()
        Else
            TabControl1.TabPages.Remove(TabPage2)
            TabControl1.TabPages.Remove(TabPage3)
            TabControl1.TabPages.Remove(TabPage5)
        End If


    End Sub

    Public Sub LoadExComposite()
        Try

            Dim conn As New OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & My.Settings.dbloc & "';Jet OLEDB:Database Password='" & My.Settings.dbpass & "';")

            Dim strSQL As String = "SELECT PortfolioCompositeID, PortfolioCompositeCode FROM dbo_vQbRowDefPortfolioComposite ORDER BY PortfolioCompositeCode"
            Dim da As New OleDb.OleDbDataAdapter(strSQL, conn)
            Dim ds As New DataSet
            da.Fill(ds, "Users")

            With cboExComposite
                .DataSource = ds.Tables("Users")
                .DisplayMember = "PortfolioCompositeCode"
                .ValueMember = "PortfolioCompositeID"
                .SelectedIndex = 0
            End With

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "ERROR")
        End Try
    End Sub

    Public Sub LoadMAPDisciplines()
        Try

            Dim conn As New OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & My.Settings.dbloc & "';Jet OLEDB:Database Password='" & My.Settings.dbpass & "';")

            Dim strSQL As String = "SELECT * FROM AAM_DisciplineQuery ORDER BY DisplayValue"
            Dim da As New OleDb.OleDbDataAdapter(strSQL, conn)
            Dim ds As New DataSet
            da.Fill(ds, "Users")

            With cboMapDiscipline
                .DataSource = ds.Tables("Users")
                .DisplayMember = "DisplayValue"
                .ValueMember = "DisplayValue"
                .SelectedIndex = 0
            End With

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "ERROR")
        End Try
    End Sub

    Public Sub LoadMAPComposite()
        Try
            Dim conn As New OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & My.Settings.dbloc & "';Jet OLEDB:Database Password='" & My.Settings.dbpass & "';")
            Dim strSQL As String = "SELECT PortfolioCompositeID, PortfolioCompositeCode FROM dbo_vQbRowDefPortfolioComposite WHERE PortfolioCompositeCode Not In(Select CompositeCode FROM mdb_AccountMovesDisciplineMappings) AND PortfolioCompositeID Not In (Select CompositeID FROM mdb_AccountMovesExcludedComposites) ORDER BY PortfolioCompositeCode"
            Dim da As New OleDb.OleDbDataAdapter(strSQL, conn)
            Dim ds As New DataSet
            da.Fill(ds, "Users")
            Dim dt As DataTable = ds.Tables("Users")

            If dt.Rows.Count > 0 Then
                With cboMapComposite
                    .DataSource = ds.Tables("Users")
                    .DisplayMember = "PortfolioCompositeCode"
                    .ValueMember = "PortfolioCompositeID"
                    .SelectedIndex = 0
                End With
                btnMapSave.Enabled = True
            Else
                btnMapSave.Enabled = False
            End If

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "ERROR")
        End Try
    End Sub

    Public Sub LoadExGrid()
        Try

            Dim conn As New OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & My.Settings.dbloc & "';Jet OLEDB:Database Password='" & My.Settings.dbpass & "';")

            Dim strSQL As String = "SELECT mdb_AccountMovesExcludedComposites.ID, mdb_AccountMovesExcludedComposites.CompositeID, dbo_vQbRowDefPortfolioComposite.PortfolioCompositeCode As [Composite Code], dbo_vQbRowDefPortfolioComposite.ReportHeading1 As [Report Heading 1], dbo_vQbRowDefPortfolioComposite.ReportHeading2 As [Report Heading 2]" & _
            " FROM mdb_AccountMovesExcludedComposites INNER JOIN dbo_vQbRowDefPortfolioComposite ON mdb_AccountMovesExcludedComposites.CompositeID = dbo_vQbRowDefPortfolioComposite.PortfolioCompositeID;"
            Dim da As New OleDb.OleDbDataAdapter(strSQL, conn)
            Dim ds As New DataSet
            da.Fill(ds, "Users")

            With DataGridView1
                .DataSource = ds.Tables("Users")
                .Columns(0).Visible = False
                .Columns(1).Visible = False
            End With

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "ERROR")
        End Try
    End Sub

    Public Sub LoadMapGrid()
        Try

            Dim conn As New OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & My.Settings.dbloc & "';Jet OLEDB:Database Password='" & My.Settings.dbpass & "';")

            Dim strSQL As String = "SELECT ID, CompositeCode As [Composite Code], DisciplineName As [Discipline Name] FROM mdb_AccountMovesDisciplineMappings"
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

    Private Sub EditToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles EditToolStripMenuItem.Click
        If DataGridView1.RowCount < 1 Then

        Else
            ExID.Text = DataGridView1.SelectedCells(0).Value
            cboExComposite.SelectedValue = DataGridView1.SelectedCells(1).Value
        End If
    End Sub

    Private Sub btnExCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnExCancel.Click
        ExID.Text = "NEW"
    End Sub

    Private Sub btnExDelete_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnExDelete.Click
        If ExID.Text = "NEW" Then
            MsgBox("Record has not been saved.", MsgBoxStyle.Information, "Nothing to delete.")
        Else
            Dim ir As MsgBoxResult
            ir = MsgBox("Are you sure you want to delete this record?", MsgBoxStyle.YesNo + MsgBoxStyle.Exclamation, "Confirm Delete.")
            If ir = MsgBoxResult.Yes Then
                Dim Mycn As OleDb.OleDbConnection
                Dim Command As OleDb.OleDbCommand
                Dim SQLstr As String

                Try
                    Mycn = New OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & My.Settings.roledbloc & "';Jet OLEDB:Database Password='" & My.Settings.dbpass & "';")
                    Mycn.Open()

                    SQLstr = "DELETE * FROM mdb_AccountMovesExcludedComposites WHERE [ID] = " & ExID.Text

                    Command = New OleDb.OleDbCommand(SQLstr, Mycn)
                    Command.ExecuteNonQuery()

                    Mycn.Close()

                    Call LoadExGrid()

                Catch ex As Exception
                    MsgBox(ex.Message, MsgBoxStyle.Critical, "ERROR")
                End Try
            Else

            End If
        End If
    End Sub

    Private Sub btnExSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnExSave.Click
        If ExID.Text = "NEW" Then
            Call SaveExNew()
        Else
            Call SaveExOld()
        End If
    End Sub

    Public Sub SaveExNew()
        Dim Mycn As OleDb.OleDbConnection
        Dim Command As OleDb.OleDbCommand
        Dim SQLstr As String

        Try
            Mycn = New OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & My.Settings.roledbloc & "';Jet OLEDB:Database Password='" & My.Settings.dbpass & "';")
            Mycn.Open()

            SQLstr = "INSERT INTO mdb_AccountMovesExcludedComposites([CompositeID])" & _
            "VALUES(" & cboExComposite.SelectedValue & ")"

            Command = New OleDb.OleDbCommand(SQLstr, Mycn)
            Command.ExecuteNonQuery()

            Mycn.Close()

            Call LoadExGrid()

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "ERROR")
        End Try
    End Sub

    Public Sub SaveExOld()
        Dim Mycn As OleDb.OleDbConnection
        Dim Command As OleDb.OleDbCommand
        Dim SQLstr As String

        Try
            Mycn = New OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & My.Settings.roledbloc & "';Jet OLEDB:Database Password='" & My.Settings.dbpass & "';")
            Mycn.Open()

            SQLstr = "Update mdb_AccountMovesExcludedComposites SET [CompositeID] = " & cboExComposite.SelectedValue & " WHERE ID = " & ExID.Text

            Command = New OleDb.OleDbCommand(SQLstr, Mycn)
            Command.ExecuteNonQuery()

            Mycn.Close()

            Call LoadExGrid()

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "ERROR")
        End Try
    End Sub

    Private Sub DataGridView1_CellContentClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick

    End Sub

    Private Sub DataGridView1_CellContentDoubleClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView1.CellContentDoubleClick
        If DataGridView1.RowCount < 1 Then

        Else
            ExID.Text = DataGridView1.SelectedCells(0).Value
            cboExComposite.SelectedValue = DataGridView1.SelectedCells(1).Value
        End If
    End Sub

    Private Sub ComboBox2_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cboMapDiscipline.SelectedIndexChanged

    End Sub

    Private Sub btnMapSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnMapSave.Click
        If txtMapID.Text = "NEW" Then
            Call SaveMapNew()
        Else

        End If
    End Sub

    Public Sub SaveMapNew()
        Dim Mycn As OleDb.OleDbConnection
        Dim Command As OleDb.OleDbCommand
        Dim SQLstr As String

        Try
            Mycn = New OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & My.Settings.roledbloc & "';Jet OLEDB:Database Password='" & My.Settings.dbpass & "';")
            Mycn.Open()

            SQLstr = "INSERT INTO mdb_AccountMovesDisciplineMappings([CompositeCode], DisciplineName)" & _
            "VALUES('" & cboMapComposite.Text & "','" & cboMapDiscipline.Text & "')"

            Command = New OleDb.OleDbCommand(SQLstr, Mycn)
            Command.ExecuteNonQuery()

            Mycn.Close()

            Call LoadMapGrid()
            Call LoadMAPComposite()

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "ERROR")
        End Try
    End Sub

    Private Sub DeleteToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles DeleteToolStripMenuItem.Click
        If DataGridView2.RowCount > 1 Then
            Dim ir As MsgBoxResult
            ir = MsgBox("Are you sure you want to delete this record?", MsgBoxStyle.YesNo + MsgBoxStyle.Exclamation, "Confirm Delete.")
            If ir = MsgBoxResult.Yes Then
                Dim Mycn As OleDb.OleDbConnection
                Dim Command As OleDb.OleDbCommand
                Dim SQLstr As String

                Try
                    Mycn = New OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & My.Settings.roledbloc & "';Jet OLEDB:Database Password='" & My.Settings.dbpass & "';")
                    Mycn.Open()

                    SQLstr = "DELETE * FROM mdb_AccountMovesDisciplineMappings WHERE [ID] = " & DataGridView2.SelectedCells(0).Value

                    Command = New OleDb.OleDbCommand(SQLstr, Mycn)
                    Command.ExecuteNonQuery()

                    Mycn.Close()

                    Call LoadMapGrid()

                Catch ex As Exception
                    MsgBox(ex.Message, MsgBoxStyle.Critical, "ERROR")
                End Try
            Else

            End If
        Else
        End If
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Dim dtevalid As Date
        Dim dtetoday As Date
        dtevalid = DateTimePicker1.Text
        dtetoday = Format(Now())
        PictureBox1.Visible = True
        PictureBox1.Image = Image.FromFile(Application.StartupPath & "\resources\159.gif")
        RichTextBox1.Text = ""
        RichTextBox2.Text = ""
        ckbFileFound.Checked = False
        DataGridView3.DataSource = ""
        RichTextBox1.Text = "Checking Dates..."
        If dtevalid >= DateAdd(DateInterval.Day, -1, dtetoday) Then
            RichTextBox1.Text = RichTextBox1.Text & vbNewLine & "Date not valid!"
            PictureBox1.Visible = False
        Else
            'PictureBox1.Visible = True
            'PictureBox1.Image = Image.FromFile(Application.StartupPath & "\resources\check-mark-8-small.png")
            RichTextBox1.Text = RichTextBox1.Text & vbNewLine & "Date Valid..."
            RichTextBox1.Text = RichTextBox1.Text & vbNewLine & "Setting Date for sales Report..."
            Dim dte1 As Date
            dte1 = DateAdd(DateInterval.Month, -12, dtevalid)
            CheckBox2.Checked = True
            DateTimePicker3.Text = dte1
            DateTimePicker2.Text = dtevalid
            RichTextBox1.Text = RichTextBox1.Text & vbNewLine & "Sales Report Date Set..."
            RichTextBox1.Text = RichTextBox1.Text & vbNewLine & "Finding Values"
            Control.CheckForIllegalCrossThreadCalls = False
            ProssVal = New System.Threading.Thread(AddressOf FindAssets)
            ProssVal.Start()
        End If
    End Sub

    Public Sub FindAssets()
        Dim dateneeded1 As Date = DateTimePicker1.Text
        Dim dateneeded2 As String = Format(dateneeded1, "MMddyyyy")
        'Dim portcount As Integer = (row1("PortfolioCount"))

        Dim datefile As String = Format(dateneeded1, "MMddyy").ToString

        Dim Mycn As OleDb.OleDbConnection
        Dim Command As OleDb.OleDbCommand
        Dim SQLstr As String
        'Dim ir As MsgBoxResult
        'ir = MsgBox(dateneeded2, MsgBoxStyle.YesNo, "Check")
        'If ir = MsgBoxResult.Yes Then
        RichTextBox1.Text = RichTextBox1.Text & vbNewLine & "Creating File for APX..."
        Mycn = New OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & My.Settings.dbloc & "';Jet OLEDB:Database Password='" & My.Settings.dbpass & "';")
        Mycn.Open()
        '***********************************CREATE 62 Bit BAT File***********************************
        If File.Exists("\\aamapxapps01\apx$\_BillingVals\" & dateneeded2 & ".xls") Then
            File.Delete("\\aamapxapps01\apx$\_BillingVals\" & dateneeded2 & ".xls")
        Else

        End If

        'If ((RadioButton2.Checked) And (dateneeded1 = DateTimePicker2.Text)) Then
        RichTextBox2.Text = "c:" & vbNewLine & "cd C:\Program Files (x86)\Advent\ApxClient\4.0" & vbNewLine & "AdvScriptRunner REPRUN -mBillingValsFinalWAI -p@AAMAssets -x -b" & dateneeded2 & " -vs -t\\aamapxapps01\apx$\_BillingVals\" & dateneeded2 & ".xls"
        'Else
        'RichTextBox2.Text = "c:" & vbNewLine & "cd C:\Program Files (x86)\Advent\ApxClient\4.0" & vbNewLine & "AdvScriptRunner REPRUN -mBillingValsFinal " & Chr(34) & "-p" & RichTextBox4.Text & Chr(34) & " -x -b" & dateneeded2 & " -vs -t\\aamapxapps01\apx$\_BillingVals\" & dateneeded2 & ".xls"
        'End If

        'RichTextBox1.Text = RichTextBox1.Text & vbNewLine & "64 Bit Script Created..."

        'RichTextBox1.Text = RichTextBox1.Text & vbNewLine & "Creating BAT File..."
        'Label14.Text = "Creating BAT File"
        RichTextBox2.SaveFile("\\aamapxapps01\apx$\automation\_AutomatedBillingBATS\" & datefile & "64.bat", _
            RichTextBoxStreamType.PlainText)

        RichTextBox1.Text = RichTextBox1.Text & vbNewLine & "64 Bit BAT file created..."

        System.Diagnostics.Process.Start("\\aamapxapps01\apx$\Automation\_AutomatedBillingBATS\" & datefile & "64.BAT")
        'PictureBox2.Visible = True
        RichTextBox1.Text = RichTextBox1.Text & vbNewLine & "Sent files to RepRunner for Processing..."

        'SQLstr = "Update mdb_BillingValQueue SET Working = False, Done = True WHERE ID = " & fileid

        'Command = New OleDb.OleDbCommand(SQLstr, Mycn)
        'Command.ExecuteNonQuery()

        SQLstr = "DELETE * FROM mdb_AccountMovesValueImport"

        Command = New OleDb.OleDbCommand(SQLstr, Mycn)
        Command.ExecuteNonQuery()

        Dim fileloc As String
        fileloc = "\\aamapxapps01\apx$\_BillingVals\" & dateneeded2 & ".xls"
        Do
            If (System.IO.File.Exists(fileloc)) Then
                RichTextBox1.Text = RichTextBox1.Text & vbNewLine & "File found.  Starting Import."
                'Label12.Text = "Found File"
                ckbFileFound.Checked = True
                'PictureBox3.Visible = True

                Pause(0.5)

                '==========================================================================================
                'Create a Schema file to force field types of csv file
                '==========================================================================================

                'Dim FileToDelete As String

                'FileToDelete = "\\aamapxapps01\apx$\_BillingVals\Schema.ini"

                'If System.IO.File.Exists(FileToDelete) = True Then

                'System.IO.File.Delete(FileToDelete)
                'MsgBox("File Deleted")
                'End If

                'RichTextBox2.Text = "[" & dateneeded2 & ".xls" & "]" & vbNewLine & "   ColNameHeader=True" & vbNewLine & "   Format=XLS" & vbNewLine & "   MaxScanRows=0" & vbNewLine & "   CharacterSet=OEM" & vbNewLine & "   Col1=PortfolioID Double Width 25" & vbNewLine & "   Col2=Date1 Char Width 255" & vbNewLine & "   Col3=Symbol Char Width 255" & vbNewLine & "   Col4=Desc Char Width 25" & vbNewLine & "   Col5=Qnty Double Width 25" & vbNewLine & "   Col6=Price Double Width 25" & vbNewLine & "   Col7=AI Double Width 25" & vbNewLine & "   Col8=TMV Double Width 25"

                'RichTextBox2.SaveFile("\\aamapxapps01\apx$\_BillingVals\Schema.ini", _
                'RichTextBoxStreamType.PlainText)

                '==========================================================================================
                'File Created
                '==========================================================================================
                RichTextBox1.Text = RichTextBox1.Text & vbNewLine & "Cleaning File."
                Dim wapp As Microsoft.Office.Interop.Excel.Application

                Dim wsheet As Microsoft.Office.Interop.Excel.Worksheet

                Dim wbook As Microsoft.Office.Interop.Excel.Workbook

                wapp = New Microsoft.Office.Interop.Excel.Application

                wapp.Visible = False

                'wbook = wapp.Workbooks.Add()

                wbook = wapp.Workbooks.Open(fileloc)

                wsheet = wbook.Sheets("Sheet1")
                'wsheet.Rows.Delete(0)
                'wsheet.Rows.Delete(1)
                'wsheet.Rows.Delete(2)
                'wsheet.Rows(1).delete()
                wsheet.Range("H5", "H30000").Replace("", "0")
                'wsheet.Range("H5", "H30000").Replace(" ", "0")
                wsheet.Range("H5", "H30000").Replace("NA", "0")

                wsheet.SaveAs("\\aamapxapps01\apx$\_BillingVals\" & dateneeded2)
                'wapp.SaveWorkspace("C:\AMP STEPOUT TEMPLETE " & dte & " FISA")
                wbook.Close()
                wapp.Quit()

                RichTextBox1.Text = RichTextBox1.Text & vbNewLine & "Placing in Queue."
                '****************** IMPORT FILE TO DATAGRID ***************************
                Dim _filename As String = fileloc
                Dim _conn As String
                Dim ds1 As New DataSet
                Dim ds2 As New DataSet


                _conn = "Provider=Microsoft.Jet.OLEDB.4.0;" & "Data Source=" & _filename & ";" & "Extended Properties=Excel 8.0;"
                Dim _connection As OleDbConnection = New OleDbConnection(_conn)
                Dim da5 As OleDbDataAdapter = New OleDbDataAdapter()
                Dim _command As OleDbCommand = New OleDbCommand()

                _command.Connection = _connection
                _command.CommandText = "SELECT * FROM [Sheet1$] WHERE [PortfolioID] Is Not Null"
                da5.SelectCommand = _command
                da5.Fill(ds1, "sheet1")

                DataGridView3.DataSource = ds1
                DataGridView3.DataMember = "sheet1"
                'Label12.Text = "Found " & DataGridView3.RowCount.ToString & " rows of data."
                '****************** END IMPORT ***************************
                'RichTextBox1.Text = RichTextBox1.Text & vbNewLine & "File '" & fname & "' Imported.  Sending data to database."
                'Label14.Text = "File '" & fname & "' Imported.  Sending data to database"
                '****************** SEND DATA TO DATABASE ***************************
                Dim PID1 As Integer
                Dim Dte1 As Date
                Dim IsCash As String
                Dim Symbol As String
                Dim Desc As String
                Dim Qnty As Double
                Dim Price As Double
                Dim TMV As Double
                Dim AI As Double
                'ProgressBar2.Value = 0
                'ProgressBar2.Maximum = DataGridView3.RowCount

                RichTextBox1.Text = RichTextBox1.Text & vbNewLine & "Importing"
                For i As Integer = 0 To (Me.DataGridView3.Rows.Count - 1)
                    If DataGridView3.Rows(i).Cells(0).ValueType Is GetType(String) Then

                    Else


                        PID1 = DataGridView3.Rows(i).Cells(0).Value
                        Dte1 = DataGridView3.Rows(i).Cells(1).Value
                        IsCash = DataGridView3.Rows(i).Cells(2).Value
                        If IsDBNull(DataGridView3.Rows(i).Cells(3).Value) Then
                            Symbol = ""
                        Else
                            Symbol = DataGridView3.Rows(i).Cells(3).Value
                        End If

                        Desc = DataGridView3.Rows(i).Cells(4).Value
                        If IsDBNull(DataGridView3.Rows(i).Cells(5).Value) Then
                            Qnty = DataGridView3.Rows(i).Cells(8).Value
                        Else
                            Qnty = DataGridView3.Rows(i).Cells(5).Value
                        End If

                        If IsDBNull(DataGridView3.Rows(i).Cells(6).Value) Then
                            Price = 1
                        Else
                            Price = DataGridView3.Rows(i).Cells(6).Value
                        End If

                        If IsDBNull(DataGridView3.Rows(i).Cells(7).Value) Then
                            'Or DataGridView3.Rows(i).Cells(7).ValueType Is GetType(String) 
                            AI = 0
                        Else
                            'If DataGridView3.Rows(i).Cells(7).Value = "NA" Then
                            'AI = 0
                            'Else
                            AI = DataGridView3.Rows(i).Cells(7).Value
                            'End If
                        End If
                        If IsDBNull(DataGridView3.Rows(i).Cells(8).Value) Then
                            'Or DataGridView3.Rows(i).Cells(7).ValueType Is GetType(String) 
                            TMV = 0
                        Else
                            TMV = DataGridView3.Rows(i).Cells(8).Value
                        End If
                        SQLstr = "Insert Into mdb_AccountMovesValueImport ([PortfolioID], [Date1], [IsCash], [Symbol], [Desc], [Qnty], [Price], [AI], [TMV])" & _
                        " VALUES(" & PID1 & ",#" & Dte1 & "#,'" & IsCash & "','" & Symbol & "','" & Desc & "'," & Qnty & "," & Price & "," & AI & "," & TMV & ")"

                        Command = New OleDb.OleDbCommand(SQLstr, Mycn)
                        Command.ExecuteNonQuery()
                        'PictureBox4.Visible = True
                    End If
                    'ProgressBar2.Value = ProgressBar2.Value + 1

                    'Dim nleft As Integer
                    'nleft = ProgressBar2.Maximum - ProgressBar2.Value

                    'Label12.Text = nleft & " records left to import."

                Next
                'RichTextBox1.Text = RichTextBox1.Text & vbNewLine & "File '" & fname & "' loaded and imported to database."
                'Label14.Text = "File '" & fname & "' loaded and imported to database"
                '****************** END DATASEND ***************************

                'SQLstr = "Update mdb_BillingFileQueue SET Done = True, Hits = " & hit & " WHERE ID = " & fileid
                'Command = New OleDb.OleDbCommand(SQLstr, Mycn)
                'Command.ExecuteNonQuery()

            Else

                'hit = hit + 1
                'SQLstr = "Update mdb_BillingFileQueue SET Hits = " & hit & " WHERE ID = " & fileid

                'Command = New OleDb.OleDbCommand(SQLstr, Mycn)
                'Command.ExecuteNonQuery()

                'Label12.Text = "File not found.  # of hits: " & hit + 1

                Pause(0.1)

                'If hit > 750 Then
                'Label12.Text = "ERROR: Maximum number of hits reached."
                'RichTextBox1.Text = RichTextBox1.Text & vbNewLine & "File '" & fname & "' cannot be found."
                'RichTextBox1.Text = RichTextBox1.Text & vbNewLine & "Maximum number of hits reached.  Process Ended in Error."
                'Label14.Text = "Process ended in error."
                'GoTo line2
                'Else

                'End If

            End If

            'Call LoadFileQueue()

            'End If
        Loop Until ckbFileFound.Checked = True

        SQLstr = "DELETE * FROM mdb_AccountMoves_Assets"
        RichTextBox1.Text = RichTextBox1.Text & vbNewLine & "Loading QID 1..."
        Command = New OleDb.OleDbCommand(SQLstr, Mycn)
        Command.ExecuteNonQuery()

        SQLstr = "INSERT INTO mdb_AccountMoves_Assets (PortfolioID, ReportHeading1, SumOfAI, SumOfTMV, TotalAssets, DisciplineName, QID)" & _
        "SELECT dbo_vQbRowDefPortfolio.PortfolioID, dbo_vQbRowDefPortfolio.ReportHeading1, Sum(mdb_AccountMovesValueImport.AI) AS SumOfAI, Sum(mdb_AccountMovesValueImport.TMV) AS SumOfTMV, (SumOfAI + SumOfTMV) As [TotalAssets], mdb_AccountMovesDisciplineMappings.DisciplineName, 1" & _
        " FROM (mdb_AccountMovesValueImport INNER JOIN (dbo_vQbRowDefPortfolioCompositeMember INNER JOIN mdb_AccountMovesDisciplineMappings ON dbo_vQbRowDefPortfolioCompositeMember.PortfolioCompositeCode = mdb_AccountMovesDisciplineMappings.CompositeCode) ON mdb_AccountMovesValueImport.PortfolioID = dbo_vQbRowDefPortfolioCompositeMember.MemberID) INNER JOIN dbo_vQbRowDefPortfolio ON mdb_AccountMovesValueImport.PortfolioID = dbo_vQbRowDefPortfolio.PortfolioID" & _
        " WHERE(((CDate([EntryDate])) <= #" & dateneeded1 & "#) And ((dbo_vQbRowDefPortfolioCompositeMember.ExitDate) Is Not Null) And ((CDate([ExitDate])) >= #" & dateneeded1 & "#))" & _
        " GROUP BY dbo_vQbRowDefPortfolio.PortfolioID, dbo_vQbRowDefPortfolio.ReportHeading1, mdb_AccountMovesDisciplineMappings.DisciplineName;"

        Command = New OleDb.OleDbCommand(SQLstr, Mycn)
        Command.ExecuteNonQuery()

        'PictureBox5.Visible = True
        RichTextBox1.Text = RichTextBox1.Text & vbNewLine & "Loading QID 2..."

        SQLstr = "INSERT INTO mdb_AccountMoves_Assets (PortfolioID, ReportHeading1, SumOfAI, SumOfTMV, TotalAssets, DisciplineName, QID)" & _
        "SELECT dbo_vQbRowDefPortfolio.PortfolioID, dbo_vQbRowDefPortfolio.ReportHeading1, Sum(mdb_AccountMovesValueImport.AI) AS SumOfAI, Sum(mdb_AccountMovesValueImport.TMV) AS SumOfTMV, (SumOfAI + SumOfTMV) As [TotalAssets], mdb_AccountMovesDisciplineMappings.DisciplineName, 2" & _
        " FROM (mdb_AccountMovesValueImport INNER JOIN (dbo_vQbRowDefPortfolioCompositeMember INNER JOIN mdb_AccountMovesDisciplineMappings ON dbo_vQbRowDefPortfolioCompositeMember.PortfolioCompositeCode = mdb_AccountMovesDisciplineMappings.CompositeCode) ON mdb_AccountMovesValueImport.PortfolioID = dbo_vQbRowDefPortfolioCompositeMember.MemberID) INNER JOIN dbo_vQbRowDefPortfolio ON mdb_AccountMovesValueImport.PortfolioID = dbo_vQbRowDefPortfolio.PortfolioID" & _
        " WHERE(((CDate([EntryDate])) <= #" & dateneeded1 & "#) And ((dbo_vQbRowDefPortfolioCompositeMember.ExitDate) Is Null))" & _
        " GROUP BY dbo_vQbRowDefPortfolio.PortfolioID, dbo_vQbRowDefPortfolio.ReportHeading1, mdb_AccountMovesDisciplineMappings.DisciplineName;"

        Command = New OleDb.OleDbCommand(SQLstr, Mycn)
        Command.ExecuteNonQuery()

        'PictureBox6.Visible = True
        RichTextBox1.Text = RichTextBox1.Text & vbNewLine & "Loading QID 3..."

        SQLstr = "INSERT INTO mdb_AccountMoves_Assets (PortfolioID, ReportHeading1, SumOfAI, SumOfTMV, TotalAssets, DisciplineName, QID)" & _
        "SELECT dbo_vQbRowDefPortfolio.PortfolioID, dbo_vQbRowDefPortfolio.ReportHeading1, Sum(mdb_AccountMovesValueImport.AI) AS SumOfAI, Sum(mdb_AccountMovesValueImport.TMV) AS SumOfTMV, (SumOfAI+SumOfTMV) AS TotalAssets, dbo_vQbRowDefPortfolio.Discipline6, 3" & _
        " FROM mdb_AccountMovesValueImport INNER JOIN dbo_vQbRowDefPortfolio ON mdb_AccountMovesValueImport.PortfolioID = dbo_vQbRowDefPortfolio.PortfolioID" & _
        " WHERE (((mdb_AccountMovesValueImport.PortfolioID) Not In (Select PortfolioID FROM mdb_AccountMoves_Assets)))" & _
        " GROUP BY dbo_vQbRowDefPortfolio.PortfolioID, dbo_vQbRowDefPortfolio.ReportHeading1, dbo_vQbRowDefPortfolio.Discipline6"

        '"SELECT dbo_vQbRowDefPortfolio.PortfolioID, dbo_vQbRowDefPortfolio.ReportHeading1, Sum(mdb_AccountMovesValueImport.AI) AS SumOfAI, Sum(mdb_AccountMovesValueImport.TMV) AS SumOfTMV, (SumOfAI+SumOfTMV) AS TotalAssets, dbo_vQbRowDefPortfolio.Discipline6, 3" & _
        '" FROM mdb_AccountMovesValueImport INNER JOIN dbo_vQbRowDefPortfolio ON mdb_AccountMovesValueImport.PortfolioID = dbo_vQbRowDefPortfolio.PortfolioID" & _
        '" WHERE (((mdb_AccountMovesValueImport.PortfolioID) Not In (Select MemberID FROM [dbo_vQbRowDefPortfolioCompositeMember] WHERE CDate([EntryDate]) <#" & dateneeded1 & "# AND ExitDate is Null)))" & _
        '" GROUP BY dbo_vQbRowDefPortfolio.PortfolioID, dbo_vQbRowDefPortfolio.ReportHeading1, dbo_vQbRowDefPortfolio.Discipline6"


        Command = New OleDb.OleDbCommand(SQLstr, Mycn)
        Command.ExecuteNonQuery()

        'PictureBox6.Visible = True

        RichTextBox1.Text = RichTextBox1.Text & vbNewLine & "DONE!"
        Button3.Enabled = True
        Button3.BackColor = Color.Red
        Button3.ForeColor = Color.White
        PictureBox1.Visible = False
        CheckBox1.Checked = False
        'Else
        'End If

    End Sub

    Public Sub LoadAssetsByStrategy()
        Try

            DataGridView1.Enabled = False
            Dim strSQL As String
            Dim conn As New OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & My.Settings.dbloc & "';Jet OLEDB:Database Password='" & My.Settings.dbpass & "';")
            If CheckBox4.Checked = False Then
                strSQL = "SELECT mdb_AccountMoves_Assets.DisciplineName As [Discipline], Sum(mdb_AccountMoves_Assets.SumOfAI) AS [Accrued Interest], Sum(mdb_AccountMoves_Assets.SumOfTMV) AS [Market Value], Sum(mdb_AccountMoves_Assets.TotalAssets) AS [Total Assets]" & _
                " FROM(mdb_AccountMoves_Assets)" & _
                " GROUP BY mdb_AccountMoves_Assets.DisciplineName"
            Else
                strSQL = "SELECT mdb_AccountMoves_Assets.DisciplineName AS Discipline, mdb_AccountMoves_Departments.DepartmentName As [Department], Sum(mdb_AccountMoves_Assets.SumOfAI) AS [Accrued Interest], Sum(mdb_AccountMoves_Assets.SumOfTMV) AS [Market Value], Sum(mdb_AccountMoves_Assets.TotalAssets) AS [Total Assets]" & _
                " FROM (mdb_AccountMoves_Assets INNER JOIN dbo_vQbRowDefPortfolio ON mdb_AccountMoves_Assets.PortfolioID = dbo_vQbRowDefPortfolio.PortfolioID) INNER JOIN mdb_AccountMoves_Departments ON dbo_vQbRowDefPortfolio.AAMTeamName = mdb_AccountMoves_Departments.TerritoryName" & _
                " GROUP BY mdb_AccountMoves_Assets.DisciplineName, mdb_AccountMoves_Departments.DepartmentName;"
            End If

            Dim da As New OleDb.OleDbDataAdapter(strSQL, conn)
            Dim ds As New DataSet
            da.Fill(ds, "Users")

            With DataGridView4
                .DataSource = ds.Tables("Users")
                '.Columns(0).Visible = False
            End With

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

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        Button3.BackColor = Color.White
        Button3.ForeColor = Color.Black
        Call LoadAssetsByStrategy()
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        Dim path As String
        path = "C:\_AccountMoves"

        If (System.IO.Directory.Exists(path)) Then
            'System.IO.Directory.Delete(path, True)
        Else
            System.IO.Directory.CreateDirectory(path)
        End If

        Dim date1 As Date
        date1 = DateTimePicker1.Text

        Dim dte As String
        dte = Format(date1, "yyyyMMdd")

        Dim tme As String
        tme = Format(Now(), "HHmm")

        Dim fname As String
        fname = "AccountMoves_" & dte & "_" & tme & ".csv"

        Dim AccessConn As New System.Data.OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & My.Settings.dbloc & "';Jet OLEDB:Database Password='" & My.Settings.dbpass & "';")

        AccessConn.Open()

        If CheckBox4.Checked = False Then
            'strSQL = "SELECT mdb_AccountMoves_Assets.DisciplineName As [Discipline], Sum(mdb_AccountMoves_Assets.SumOfAI) AS [Accrued Interest], Sum(mdb_AccountMoves_Assets.SumOfTMV) AS [Market Value], Sum(mdb_AccountMoves_Assets.TotalAssets) AS [Total Assets]" & _
            '" FROM(mdb_AccountMoves_Assets)" & _
            '" GROUP BY mdb_AccountMoves_Assets.DisciplineName"
            Dim AccessCommand As New OleDb.OleDbCommand("SELECT mdb_AccountMoves_Assets.DisciplineName As [Discipline Name], Sum(mdb_AccountMoves_Assets.SumOfAI) AS [Accrued Interest], Sum(mdb_AccountMoves_Assets.SumOfTMV) AS [Market Value], Sum(mdb_AccountMoves_Assets.TotalAssets) AS [Total Assets] INTO [Text;HDR=YES;DATABASE=" & path & "].[" & fname & "] FROM [mdb_AccountMoves_Assets] GROUP BY mdb_AccountMoves_Assets.DisciplineName", AccessConn)
            AccessCommand.ExecuteNonQuery()
        Else
            'strSQL = "SELECT mdb_AccountMoves_Assets.DisciplineName AS Discipline, mdb_AccountMoves_Departments.DepartmentName As [Department], Sum(mdb_AccountMoves_Assets.SumOfAI) AS [Accrued Interest], Sum(mdb_AccountMoves_Assets.SumOfTMV) AS [Market Value], Sum(mdb_AccountMoves_Assets.TotalAssets) AS [Total Assets]" & _
            '" FROM (mdb_AccountMoves_Assets INNER JOIN dbo_vQbRowDefPortfolio ON mdb_AccountMoves_Assets.PortfolioID = dbo_vQbRowDefPortfolio.PortfolioID) INNER JOIN mdb_AccountMoves_Departments ON dbo_vQbRowDefPortfolio.AAMTeamName = mdb_AccountMoves_Departments.TerritoryName" & _
            '" GROUP BY mdb_AccountMoves_Assets.DisciplineName, mdb_AccountMoves_Departments.DepartmentName;"

            Dim AccessCommand As New OleDb.OleDbCommand("SELECT mdb_AccountMoves_Assets.DisciplineName As [Discipline Name], mdb_AccountMoves_Departments.DepartmentName As [Department], Sum(mdb_AccountMoves_Assets.SumOfAI) AS [Accrued Interest], Sum(mdb_AccountMoves_Assets.SumOfTMV) AS [Market Value], Sum(mdb_AccountMoves_Assets.TotalAssets) AS [Total Assets] INTO [Text;HDR=YES;DATABASE=" & path & "].[" & fname & "] FROM (mdb_AccountMoves_Assets INNER JOIN dbo_vQbRowDefPortfolio ON mdb_AccountMoves_Assets.PortfolioID = dbo_vQbRowDefPortfolio.PortfolioID) INNER JOIN mdb_AccountMoves_Departments ON dbo_vQbRowDefPortfolio.AAMTeamName = mdb_AccountMoves_Departments.TerritoryName GROUP BY mdb_AccountMoves_Assets.DisciplineName, mdb_AccountMoves_Departments.DepartmentName", AccessConn)
            AccessCommand.ExecuteNonQuery()

        End If

        AccessConn.Close()
        'Dim AccessCommand As New OleDb.OleDbCommand("SELECT mdb_AccountMoves_Assets.DisciplineName As [Discipline Name], Sum(mdb_AccountMoves_Assets.SumOfAI) AS [Accrued Interest], Sum(mdb_AccountMoves_Assets.SumOfTMV) AS [Market Value], Sum(mdb_AccountMoves_Assets.TotalAssets) AS [Total Assets] INTO [Text;HDR=YES;DATABASE=" & path & "].[" & fname & "] FROM [mdb_AccountMoves_Assets] GROUP BY mdb_AccountMoves_Assets.DisciplineName", AccessConn)

        

        Dim ir As MsgBoxResult
        Dim fullnme As String
        fullnme = path & "\" & fname

        ir = MsgBox("Export finished.  File saved as '" & fullnme & "." & vbNewLine & vbNewLine & "Would you like to open the file now?", MsgBoxStyle.YesNo + MsgBoxStyle.Information, "Done")

        If ir = MsgBoxResult.Yes Then

            Dim xlsApp As Microsoft.Office.Interop.Excel.Application
            Dim xlsWorkBook As Microsoft.Office.Interop.Excel.Workbook

            xlsApp = New Microsoft.Office.Interop.Excel.Application

            xlsWorkBook = xlsApp.Workbooks.Open(fullnme)
            xlsApp.Visible = True

        Else
        End If
    End Sub

    Private Sub DateTimePicker2_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles DateTimePicker2.LostFocus
        If CheckBox2.Checked Then

        Else
            Dim dte1 As Date
            dte1 = DateTimePicker2.Text
            Dim dtestart As Date
            dtestart = DateAdd(DateInterval.Month, -12, dte1)

            DateTimePicker3.Text = dtestart
            CheckBox2.Checked = True
        End If
    End Sub

    Private Sub DateTimePicker2_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles DateTimePicker2.ValueChanged



    End Sub

    Private Sub Button6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button6.Click
        RichTextBox4.Text = "Checking Dates..."
        PictureBox2.Visible = True
        PictureBox2.Image = Image.FromFile(Application.StartupPath & "\resources\159.gif")

        Dim dte1 As Date
        Dim dte2 As Date
        Dim dtenow As Date

        dte1 = DateTimePicker3.Text
        dte2 = DateTimePicker2.Text
        dtenow = Format(Now())

        If dte2 < dte1 Then
            RichTextBox4.Text = RichTextBox4.Text & vbNewLine & "Ending Date cannot be before start date!"
            PictureBox2.Visible = False
        Else
            If dte2 > DateAdd(DateInterval.Day, -1, dtenow) Then
                RichTextBox4.Text = RichTextBox4.Text & vbNewLine & "Ending Date cannot be in the future!"
                PictureBox2.Visible = False
            Else
                Control.CheckForIllegalCrossThreadCalls = False
                ProssSales = New System.Threading.Thread(AddressOf LoadSales)
                ProssSales.Start()
            End If
        End If
    End Sub

    Public Sub LoadSales()
        Dim dte1 As Date
        Dim dte2 As Date

        dte1 = DateTimePicker3.Text
        dte2 = DateTimePicker2.Text

        Dim Mycn As OleDb.OleDbConnection
        Dim Command As OleDb.OleDbCommand
        Dim SQLstr As String

        Mycn = New OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & My.Settings.dbloc & "';Jet OLEDB:Database Password='" & My.Settings.dbpass & "';")
        Mycn.Open()

        SQLstr = "DELETE * FROM mdb_AccountMoves_Sales"
        RichTextBox4.Text = RichTextBox4.Text & vbNewLine & "Loading QID 1001..."
        Command = New OleDb.OleDbCommand(SQLstr, Mycn)
        Command.ExecuteNonQuery()

        SQLstr = "INSERT INTO mdb_AccountMoves_Sales (PortfolioID, ReportHeading1, SumOfTrades, DisciplineName, QID,Rate1,BSC)" & _
        "SELECT dbo_vQbRowDefPortfolio.PortfolioID, dbo_vQbRowDefPortfolio.ReportHeading1, Sum(dbo_vQbRowDefPortfolioTransaction.TradeAmount) AS SumOfTradeAmount, mdb_AccountMovesDisciplineMappings.DisciplineName, 1001, TieredBillingRate1, ((SumOfTradeAmount*TieredBillingRate1)/100)" & _
        " FROM (dbo_vQbRowDefPortfolioTransaction INNER JOIN (dbo_vQbRowDefPortfolioCompositeMember INNER JOIN mdb_AccountMovesDisciplineMappings ON dbo_vQbRowDefPortfolioCompositeMember.PortfolioCompositeCode = mdb_AccountMovesDisciplineMappings.CompositeCode) ON dbo_vQbRowDefPortfolioTransaction.PortfolioID = dbo_vQbRowDefPortfolioCompositeMember.MemberID) INNER JOIN dbo_vQbRowDefPortfolio ON dbo_vQbRowDefPortfolioTransaction.PortfolioID = dbo_vQbRowDefPortfolio.PortfolioID" & _
        " WHERE (((dbo_vQbRowDefPortfolioTransaction.TransactionCode)='li' Or (dbo_vQbRowDefPortfolioTransaction.TransactionCode)='ti') AND ((CDate([EntryDate]))<=#" & dte2 & "#) AND ((dbo_vQbRowDefPortfolioCompositeMember.ExitDate) Is Not Null) AND ((CDate([ExitDate]))>=#" & dte2 & "#) AND ((CDate([TradeDate])) Between #" & dte1 & "# And #" & dte2 & "#))" & _
        " GROUP BY dbo_vQbRowDefPortfolio.PortfolioID, dbo_vQbRowDefPortfolio.ReportHeading1, mdb_AccountMovesDisciplineMappings.DisciplineName, TieredBillingRate1;"

        Command = New OleDb.OleDbCommand(SQLstr, Mycn)
        Command.ExecuteNonQuery()

        RichTextBox4.Text = RichTextBox4.Text & vbNewLine & "Loading QID 1002..."

        SQLstr = "INSERT INTO mdb_AccountMoves_Sales (PortfolioID, ReportHeading1, SumOfTrades, DisciplineName, QID,Rate1,BSC)" & _
        "SELECT dbo_vQbRowDefPortfolio.PortfolioID, dbo_vQbRowDefPortfolio.ReportHeading1, Sum(dbo_vQbRowDefPortfolioTransaction.TradeAmount) AS SumOfTradeAmount, mdb_AccountMovesDisciplineMappings.DisciplineName, 1002, TieredBillingRate1,((SumOfTradeAmount*TieredBillingRate1)/100)" & _
        " FROM (dbo_vQbRowDefPortfolioTransaction INNER JOIN (dbo_vQbRowDefPortfolioCompositeMember INNER JOIN mdb_AccountMovesDisciplineMappings ON dbo_vQbRowDefPortfolioCompositeMember.PortfolioCompositeCode = mdb_AccountMovesDisciplineMappings.CompositeCode) ON dbo_vQbRowDefPortfolioTransaction.PortfolioID = dbo_vQbRowDefPortfolioCompositeMember.MemberID) INNER JOIN dbo_vQbRowDefPortfolio ON dbo_vQbRowDefPortfolioTransaction.PortfolioID = dbo_vQbRowDefPortfolio.PortfolioID" & _
        " WHERE (((dbo_vQbRowDefPortfolioTransaction.TransactionCode)='li' Or (dbo_vQbRowDefPortfolioTransaction.TransactionCode)='ti') AND ((CDate([EntryDate]))<=#" & dte2 & "#) AND ((dbo_vQbRowDefPortfolioCompositeMember.ExitDate) Is Null) AND ((CDate([TradeDate])) Between #" & dte1 & "# And #" & dte2 & "#))" & _
        " GROUP BY dbo_vQbRowDefPortfolio.PortfolioID, dbo_vQbRowDefPortfolio.ReportHeading1, mdb_AccountMovesDisciplineMappings.DisciplineName, TieredBillingRate1;"

        Command = New OleDb.OleDbCommand(SQLstr, Mycn)
        Command.ExecuteNonQuery()

        RichTextBox4.Text = RichTextBox4.Text & vbNewLine & "Loading QID 1003..."

        SQLstr = "INSERT INTO mdb_AccountMoves_Sales (PortfolioID, ReportHeading1, SumOfTrades, DisciplineName, QID,Rate1,BSC)" & _
        "SELECT dbo_vQbRowDefPortfolio.PortfolioID, dbo_vQbRowDefPortfolio.ReportHeading1, Sum(dbo_vQbRowDefPortfolioTransaction.TradeAmount) AS SumOfTradeAmount, dbo_vQbRowDefPortfolio.Discipline6, 1003 AS Expr1, TieredBillingRate1, ((SumOfTradeAmount*TieredBillingRate1)/100)" & _
        " FROM dbo_vQbRowDefPortfolio INNER JOIN dbo_vQbRowDefPortfolioTransaction ON dbo_vQbRowDefPortfolio.PortfolioID = dbo_vQbRowDefPortfolioTransaction.PortfolioID" & _
        " WHERE (((dbo_vQbRowDefPortfolioTransaction.TransactionCode)='li' Or (dbo_vQbRowDefPortfolioTransaction.TransactionCode)='ti') AND ((CDate([TradeDate])) Between #" & dte1 & "# And #" & dte2 & "#) AND ((dbo_vQbRowDefPortfolio.PortfolioStatus)='open' Or (dbo_vQbRowDefPortfolio.PortfolioStatus)='closed') AND ((dbo_vQbRowDefPortfolio.ManagerCode)='aam' Or (dbo_vQbRowDefPortfolio.ManagerCode)='u aam') AND ((dbo_vQbRowDefPortfolio.PortfolioID) Not In (Select PortfolioID FROM mdb_AccountMoves_Sales)))" & _
        " GROUP BY dbo_vQbRowDefPortfolio.PortfolioID, dbo_vQbRowDefPortfolio.ReportHeading1, dbo_vQbRowDefPortfolio.Discipline6, TieredBillingRate1"


        '"SELECT dbo_vQbRowDefPortfolio.PortfolioID, dbo_vQbRowDefPortfolio.ReportHeading1, Sum(dbo_vQbRowDefPortfolioTransaction.TradeAmount) AS SumOfTradeAmount, dbo_vQbRowDefPortfolio.Discipline6, 1003 AS Expr1" & _
        '" FROM dbo_vQbRowDefPortfolio INNER JOIN dbo_vQbRowDefPortfolioTransaction ON dbo_vQbRowDefPortfolio.PortfolioID = dbo_vQbRowDefPortfolioTransaction.PortfolioID" & _
        '" WHERE (((dbo_vQbRowDefPortfolioTransaction.TransactionCode)='li' Or (dbo_vQbRowDefPortfolioTransaction.TransactionCode)='ti') AND ((CDate([TradeDate])) Between #" & dte1 & "# And #" & dte2 & "#) AND ((dbo_vQbRowDefPortfolio.PortfolioStatus)='open' Or (dbo_vQbRowDefPortfolio.PortfolioStatus)='closed') AND ((dbo_vQbRowDefPortfolio.ManagerCode)='aam' Or (dbo_vQbRowDefPortfolio.ManagerCode)='u aam') AND ((dbo_vQbRowDefPortfolio.PortfolioID) Not In (Select MemberID FROM [dbo_vQbRowDefPortfolioCompositeMember] WHERE CDate([EntryDate]) <#" & dte2 & "# AND ExitDate is Null)))" & _
        '" GROUP BY dbo_vQbRowDefPortfolio.PortfolioID, dbo_vQbRowDefPortfolio.ReportHeading1, dbo_vQbRowDefPortfolio.Discipline6;"

        Command = New OleDb.OleDbCommand(SQLstr, Mycn)
        Command.ExecuteNonQuery()

        RichTextBox4.Text = RichTextBox4.Text & vbNewLine & "Sales Loaded!"
        Button4.Enabled = True
        Button4.BackColor = Color.Red
        Button4.ForeColor = Color.White
        PictureBox2.Visible = False

    End Sub

    Public Sub LoadSalesGrid()
        Try
            Dim strSQL As String
            Dim conn As New OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & My.Settings.dbloc & "';Jet OLEDB:Database Password='" & My.Settings.dbpass & "';")
            If CheckBox3.Checked = False Then
                strSQL = "SELECT mdb_AccountMoves_Sales.DisciplineName As [Discipline Name], Sum(mdb_AccountMoves_Sales.SumOfTrades) AS [Yearly Sales], Sum(mdb_AccountMoves_Sales.BSC) As [BSC]" & _
            " FROM(mdb_AccountMoves_Sales)" & _
            " GROUP BY mdb_AccountMoves_Sales.DisciplineName" & _
            " ORDER BY mdb_AccountMoves_Sales.DisciplineName;"
            Else
                strSQL = "SELECT mdb_AccountMoves_Sales.DisciplineName AS [Discipline Name], mdb_AccountMoves_Departments.DepartmentName As [Department], Sum(mdb_AccountMoves_Sales.SumOfTrades) AS [Yearly Sales], Sum(mdb_AccountMoves_Sales.BSC) As [BSC]" & _
                " FROM (mdb_AccountMoves_Sales INNER JOIN dbo_vQbRowDefPortfolio ON mdb_AccountMoves_Sales.PortfolioID = dbo_vQbRowDefPortfolio.PortfolioID) INNER JOIN mdb_AccountMoves_Departments ON dbo_vQbRowDefPortfolio.AAMTeamName = mdb_AccountMoves_Departments.TerritoryName" & _
                " GROUP BY mdb_AccountMoves_Sales.DisciplineName, mdb_AccountMoves_Departments.DepartmentName" & _
                " ORDER BY mdb_AccountMoves_Sales.DisciplineName;"

            End If

            Dim da As New OleDb.OleDbDataAdapter(strSQL, conn)
            Dim ds As New DataSet
            da.Fill(ds, "Users")

            With DataGridView5
                .DataSource = ds.Tables("Users")
                '.Columns(0).Visible = False
            End With

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "ERROR")
        End Try
    End Sub

    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        Call LoadSalesGrid()
        Button4.BackColor = Color.White
        Button4.ForeColor = Color.Black
    End Sub

    Private Sub Button5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button5.Click
        Dim path As String
        path = "C:\_AccountMoves"

        If (System.IO.Directory.Exists(path)) Then
            'System.IO.Directory.Delete(path, True)
        Else
            System.IO.Directory.CreateDirectory(path)
        End If

        Dim date1 As Date
        date1 = DateTimePicker1.Text

        Dim dte As String
        dte = Format(date1, "yyyyMMdd")

        Dim tme As String
        tme = Format(Now(), "HHmmss")

        Dim fname As String
        fname = "AccountSales_" & dte & "_" & tme & ".csv"

        Dim AccessConn As New System.Data.OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & My.Settings.dbloc & "';Jet OLEDB:Database Password='" & My.Settings.dbpass & "';")

        AccessConn.Open()

        If CheckBox3.Checked = False Then
            'strSQL = "SELECT mdb_AccountMoves_Sales.DisciplineName As [Discipline Name], Sum(mdb_AccountMoves_Sales.SumOfTrades) AS [Yearly Sales]" & _
            '" FROM(mdb_AccountMoves_Sales)" & _
            '" GROUP BY mdb_AccountMoves_Sales.DisciplineName" & _
            '" ORDER BY mdb_AccountMoves_Sales.DisciplineName;"
            Dim AccessCommand As New OleDb.OleDbCommand("SELECT mdb_AccountMoves_Sales.DisciplineName As [Discipline Name], Sum(mdb_AccountMoves_Sales.SumOfTrades) AS [Yearly Sales], Sum(mdb_AccountMoves_Sales.BSC) As [BSC] INTO [Text;HDR=YES;DATABASE=" & path & "].[" & fname & "] FROM(mdb_AccountMoves_Sales) GROUP BY mdb_AccountMoves_Sales.DisciplineName ORDER BY mdb_AccountMoves_Sales.DisciplineName", AccessConn)
            AccessCommand.ExecuteNonQuery()
        Else
            'strSQL = "SELECT mdb_AccountMoves_Sales.DisciplineName AS [Discipline Name], mdb_AccountMoves_Departments.DepartmentName As [Department], Sum(mdb_AccountMoves_Sales.SumOfTrades) AS [Yearly Sales]" & _
            '" FROM (mdb_AccountMoves_Sales INNER JOIN dbo_vQbRowDefPortfolio ON mdb_AccountMoves_Sales.PortfolioID = dbo_vQbRowDefPortfolio.PortfolioID) INNER JOIN mdb_AccountMoves_Departments ON dbo_vQbRowDefPortfolio.AAMTeamName = mdb_AccountMoves_Departments.TerritoryName" & _
            '" GROUP BY mdb_AccountMoves_Sales.DisciplineName, mdb_AccountMoves_Departments.DepartmentName" & _
            '" ORDER BY mdb_AccountMoves_Sales.DisciplineName;"
            Dim AccessCommand As New OleDb.OleDbCommand("SELECT mdb_AccountMoves_Sales.DisciplineName As [Discipline Name], mdb_AccountMoves_Departments.DepartmentName As [Department], Sum(mdb_AccountMoves_Sales.SumOfTrades) AS [Yearly Sales], Sum(mdb_AccountMoves_Sales.BSC) As [BSC] INTO [Text;HDR=YES;DATABASE=" & path & "].[" & fname & "] FROM (mdb_AccountMoves_Sales INNER JOIN dbo_vQbRowDefPortfolio ON mdb_AccountMoves_Sales.PortfolioID = dbo_vQbRowDefPortfolio.PortfolioID) INNER JOIN mdb_AccountMoves_Departments ON dbo_vQbRowDefPortfolio.AAMTeamName = mdb_AccountMoves_Departments.TerritoryName GROUP BY mdb_AccountMoves_Sales.DisciplineName, mdb_AccountMoves_Departments.DepartmentName ORDER BY mdb_AccountMoves_Sales.DisciplineName", AccessConn)
            AccessCommand.ExecuteNonQuery()
        End If

        'Dim AccessCommand As New OleDb.OleDbCommand("SELECT mdb_AccountMoves_Sales.DisciplineName As [Discipline Name], Sum(mdb_AccountMoves_Sales.SumOfTrades) AS [Yearly Sales] INTO [Text;HDR=YES;DATABASE=" & path & "].[" & fname & "] FROM(mdb_AccountMoves_Sales) GROUP BY mdb_AccountMoves_Sales.DisciplineName ORDER BY mdb_AccountMoves_Sales.DisciplineName", AccessConn)

        'AccessCommand.ExecuteNonQuery()
        AccessConn.Close()

        Dim ir As MsgBoxResult
        Dim fullnme As String
        fullnme = path & "\" & fname

        ir = MsgBox("Export finished.  File saved as '" & fullnme & "." & vbNewLine & vbNewLine & "Would you like to open the file now?", MsgBoxStyle.YesNo + MsgBoxStyle.Information, "Done")

        If ir = MsgBoxResult.Yes Then

            Dim xlsApp As Microsoft.Office.Interop.Excel.Application
            Dim xlsWorkBook As Microsoft.Office.Interop.Excel.Workbook

            xlsApp = New Microsoft.Office.Interop.Excel.Application

            xlsWorkBook = xlsApp.Workbooks.Open(fullnme)
            xlsApp.Visible = True

        Else
        End If
    End Sub

    Private Sub Button7_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button7.Click
        Dim Mycn As OleDb.OleDbConnection
        Dim Command As OleDb.OleDbCommand
        Dim SQLstr As String

        Try
            Mycn = New OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & My.Settings.roledbloc & "';Jet OLEDB:Database Password='" & My.Settings.dbpass & "';")
            Mycn.Open()

            SQLstr = "INSERT INTO mdb_AccountMoves_Departments([TerritoryName], DepartmentName)" & _
            "VALUES('" & cboTerritory.Text & "','" & cboDepartment.Text & "')"

            Command = New OleDb.OleDbCommand(SQLstr, Mycn)
            Command.ExecuteNonQuery()

            Mycn.Close()

            Call LoadTerritoryCB()
            Call LoadDepartmentGrid()

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "ERROR")
        End Try
    End Sub

    Public Sub LoadTerritoryCB()
        Try

            Dim conn As New OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & My.Settings.dbloc & "';Jet OLEDB:Database Password='" & My.Settings.dbpass & "';")

            Dim strSQL As String = "SELECT AAM_TeamQuery.DataValue" & _
            " FROM(AAM_TeamQuery)" & _
            " WHERE (((AAM_TeamQuery.DataValue) Not In (SELECT TerritoryName FROM [mdb_AccountMoves_Departments])))" & _
            " GROUP BY AAM_TeamQuery.DataValue" & _
            " ORDER BY AAM_TeamQuery.DataValue"
            Dim da As New OleDb.OleDbDataAdapter(strSQL, conn)
            Dim ds As New DataSet
            da.Fill(ds, "Users")
            Dim dt As DataTable = ds.Tables("Users")

            If dt.Rows.Count > 0 Then

                With cboTerritory
                    .DataSource = ds.Tables("Users")
                    .DisplayMember = "DataValue"
                    .ValueMember = "DataValue"
                    .SelectedIndex = 0
                End With
                Button7.Enabled = True
            Else
                Button7.Enabled = False
            End If

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "ERROR")
        End Try
    End Sub

    Public Sub LoadDepartmentGrid()
        Try

            Dim conn As New OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & My.Settings.dbloc & "';Jet OLEDB:Database Password='" & My.Settings.dbpass & "';")

            Dim strSQL As String = "SELECT mdb_AccountMoves_Departments.ID, mdb_AccountMoves_Departments.TerritoryName, mdb_AccountMoves_Departments.DepartmentName" & _
            " FROM mdb_AccountMoves_Departments;"
            Dim da As New OleDb.OleDbDataAdapter(strSQL, conn)
            Dim ds As New DataSet
            da.Fill(ds, "Users")

            With dgvDepartmentMappings
                .DataSource = ds.Tables("Users")
                .Columns(0).Visible = False
            End With

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "ERROR")
        End Try
    End Sub

    Private Sub DataGridView2_CellContentClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView2.CellContentClick

    End Sub

    Private Sub RemoveToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RemoveToolStripMenuItem.Click
        If dgvDepartmentMappings.RowCount > 1 Then
            Dim ir As MsgBoxResult
            ir = MsgBox("Are you sure you want to delete this record?", MsgBoxStyle.YesNo + MsgBoxStyle.Exclamation, "Confirm Delete.")
            If ir = MsgBoxResult.Yes Then
                Dim Mycn As OleDb.OleDbConnection
                Dim Command As OleDb.OleDbCommand
                Dim SQLstr As String

                Try
                    Mycn = New OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & My.Settings.roledbloc & "';Jet OLEDB:Database Password='" & My.Settings.dbpass & "';")
                    Mycn.Open()

                    SQLstr = "DELETE * FROM mdb_AccountMoves_Departments WHERE [ID] = " & dgvDepartmentMappings.SelectedCells(0).Value

                    Command = New OleDb.OleDbCommand(SQLstr, Mycn)
                    Command.ExecuteNonQuery()

                    Mycn.Close()

                    Call LoadDepartmentGrid()
                    Call LoadTerritoryCB()

                Catch ex As Exception
                    MsgBox(ex.Message, MsgBoxStyle.Critical, "ERROR")
                End Try
            Else

            End If
        Else
        End If
    End Sub
End Class