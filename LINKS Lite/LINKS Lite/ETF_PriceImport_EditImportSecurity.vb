Public Class ETF_PriceImport_EditImportSecurity

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        Dim child As New ETF_PriceImport_SecData
        child.MdiParent = Home
        child.Show()
        child.ID.Text = "NEW"
    End Sub

    Public Sub LoadNotApproved()
        Try

            Dim conn As New OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & My.Settings.dbloc & "';Jet OLEDB:Database Password='" & My.Settings.dbpass & "';")

            Dim strSQL As String = "SELECT mdb_ETFPrice_SecurityData.ID, dbo_AdvSecurity.Symbol, mdb_ETFPrice_SecurityData.NewSecType, mdb_ETFPrice_SecurityData.NewSymbol, mdb_ETFPrice_SecurityData.MaturityDate, mdb_ETFPrice_SecurityData.InterestRate, mdb_ETFPrice_SecurityData.PaymentFreq, mdb_ETFPrice_SecurityData.Rating, mdb_ETFPrice_SecurityData.Rating2, mdb_ETFPrice_SecurityData.AvgLife, mdb_ETFPrice_SecurityData.YTW, mdb_ETFPrice_SecurityData.Duration, mdb_ETFPrice_SecurityData.CMOResetRule, mdb_ETFPrice_SecurityData.IssueCountry, mdb_ETFPrice_SecurityData.PrimarySymbolType, mdb_ETFPrice_SecurityData.Desc, AdvApp_vAssetClass.AssetClassName, mdb_ETFPrice_SecurityData.Import" & _
            " FROM (dbo_AdvSecurity INNER JOIN mdb_ETFPrice_SecurityData ON dbo_AdvSecurity.SecurityID = mdb_ETFPrice_SecurityData.APXSecurityID) INNER JOIN AdvApp_vAssetClass ON mdb_ETFPrice_SecurityData.AssetClassLongCode = AdvApp_vAssetClass.AssetClassCode" & _
            " WHERE Import = False AND Processed = False"

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

    Public Sub LoadApproved()
        Try

            Dim conn As New OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & My.Settings.dbloc & "';Jet OLEDB:Database Password='" & My.Settings.dbpass & "';")

            Dim strSQL As String = "SELECT mdb_ETFPrice_SecurityData.ID, dbo_AdvSecurity.Symbol, mdb_ETFPrice_SecurityData.NewSecType, mdb_ETFPrice_SecurityData.NewSymbol, mdb_ETFPrice_SecurityData.MaturityDate, mdb_ETFPrice_SecurityData.InterestRate, mdb_ETFPrice_SecurityData.PaymentFreq, mdb_ETFPrice_SecurityData.Rating, mdb_ETFPrice_SecurityData.Rating2, mdb_ETFPrice_SecurityData.AvgLife, mdb_ETFPrice_SecurityData.YTW, mdb_ETFPrice_SecurityData.Duration, mdb_ETFPrice_SecurityData.CMOResetRule, mdb_ETFPrice_SecurityData.IssueCountry, mdb_ETFPrice_SecurityData.PrimarySymbolType, mdb_ETFPrice_SecurityData.Desc, AdvApp_vAssetClass.AssetClassName, mdb_ETFPrice_SecurityData.Import" & _
            " FROM (dbo_AdvSecurity INNER JOIN mdb_ETFPrice_SecurityData ON dbo_AdvSecurity.SecurityID = mdb_ETFPrice_SecurityData.APXSecurityID) INNER JOIN AdvApp_vAssetClass ON mdb_ETFPrice_SecurityData.AssetClassLongCode = AdvApp_vAssetClass.AssetClassCode" & _
            " WHERE Import = True AND Processed = False"

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

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Call LoadApproved()
        Call LoadNotApproved()
    End Sub

    Private Sub ETF_PriceImport_EditImportSecurity_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Call LoadApproved()
        Call LoadNotApproved()

    End Sub

    Private Sub EditToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles EditToolStripMenuItem.Click
        If DataGridView1.RowCount = 0 Then
            'do nothing
        Else


            Try
                Dim conn As New OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & My.Settings.dbloc & "';Jet OLEDB:Database Password='" & My.Settings.dbpass & "';")

                Dim strSQL As String = "SELECT * FROM mdb_ETFPrice_SecurityData WHERE ID = " & DataGridView1.SelectedCells(0).Value
                Dim queryString As String = String.Format(strSQL)
                Dim cmd As New OleDb.OleDbCommand(queryString, conn)
                Dim da As New OleDb.OleDbDataAdapter(cmd)
                Dim ds As New DataSet

                da.Fill(ds, "User")
                Dim dt As DataTable = ds.Tables("User")

                Dim row1 As DataRow = dt.Rows(0)



                If (row1("IMPORT")) = True Then
                    MsgBox("You cannot edit an approved record.", MsgBoxStyle.Exclamation, "Cannot Edit")
                Else
                    Dim child As New ETF_PriceImport_SecData
                    child.MdiParent = Home
                    child.Show()

                    child.ID.Text = (row1("ID"))
                    child.ComboBox1.SelectedValue = (row1("APXSecurityID"))
                    child.MaturityDate.Text = (row1("TrueDate"))
                    child.InterestRate.Text = (row1("InterestRate"))
                    child.PaymentFreq.Text = (row1("PaymentFreq"))
                    child.Rating.Text = (row1("Rating"))
                    child.AvgLife.Text = (row1("AvgLife"))
                    child.YTW.Text = (row1("YTW"))
                    child.Duration.Text = (row1("Duration"))
                    child.CMOResetRule.Text = (row1("CMOResetRule"))
                    child.PrimarySymbolType.Text = (row1("PrimarySymbolType"))
                    child.AssetClass.SelectedValue = (row1("AssetClassLongCode"))
                    child.NewSymbol.Text = (row1("NewSymbol"))
                    child.SecType.SelectedValue = (row1("NewSecType"))
                    child.lblSecName.Text = (row1("Desc"))
                    child.Rating2.Text = (row1("Rating2"))
                End If



            Catch ex As Exception
                MsgBox(ex.Message, MsgBoxStyle.Critical, "ERROR")
            End Try
        End If
    End Sub

    Private Sub DeleteToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles DeleteToolStripMenuItem.Click
        If DataGridView1.RowCount = 0 Then
            'do nothing
        Else
            Dim ir As MsgBoxResult
            ir = MsgBox("Are you sure you want to delete this record?  This cannot be undone!", MsgBoxStyle.YesNo + MsgBoxStyle.Critical, "Confirm Delete")
            If ir = MsgBoxResult.Yes Then
                Dim Mycn As OleDb.OleDbConnection
                Dim Command As OleDb.OleDbCommand
                Dim SQLstr As String
                'Try
                Mycn = New OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & My.Settings.dbloc & "';Jet OLEDB:Database Password='" & My.Settings.dbpass & "';")
                Mycn.Open()

                If DataGridView1.SelectedCells(17).Value = True Then
                    MsgBox("You cannot delete an approved record.", MsgBoxStyle.Information, "Cant Delete")
                    GoTo Line1
                Else

                End If
                SQLstr = "DELETE * FROM mdb_ETFPrice_SecurityData WHERE ID = " & DataGridView1.SelectedCells(0).Value

                Command = New OleDb.OleDbCommand(SQLstr, Mycn)
                Command.ExecuteNonQuery()

Line1:


                Mycn.Close()

                Call LoadApproved()
                Call LoadNotApproved()

                'Catch ex As Exception
                'MsgBox(ex.Message, MsgBoxStyle.Critical, "ERROR")
                'End Try
            Else
                'do nothing
            End If
        End If
    End Sub

    Private Sub ApproveToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ApproveToolStripMenuItem.Click
        If DataGridView1.RowCount = 0 Then
            'do nothing
        Else
            Dim Mycn As OleDb.OleDbConnection
            Dim Command As OleDb.OleDbCommand
            Dim SQLstr As String
            Try
                Mycn = New OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & My.Settings.dbloc & "';Jet OLEDB:Database Password='" & My.Settings.dbpass & "';")
                Mycn.Open()

                If DataGridView1.SelectedCells(16).Value = False Then
                    SQLstr = "UPDATE mdb_ETFPrice_SecurityData SET Import = True WHERE ID = " & DataGridView1.SelectedCells(0).Value
                Else
                    SQLstr = "UPDATE mdb_ETFPrice_SecurityData SET Import = False WHERE ID = " & DataGridView1.SelectedCells(0).Value
                End If

                Command = New OleDb.OleDbCommand(SQLstr, Mycn)
                Command.ExecuteNonQuery()

                Mycn.Close()

                Call LoadApproved()
                Call LoadNotApproved()

            Catch ex As Exception
                MsgBox(ex.Message, MsgBoxStyle.Critical, "ERROR")
            End Try

        End If
    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        Dim ir As MsgBoxResult
        ir = MsgBox("Are you sure you want to approve all securities for import?", MsgBoxStyle.YesNo + MsgBoxStyle.Information, "Confirm Approve All")
        If ir = MsgBoxResult.Yes Then
            Dim Mycn As OleDb.OleDbConnection
            Dim Command As OleDb.OleDbCommand
            Dim SQLstr As String
            Try
                Mycn = New OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & My.Settings.dbloc & "';Jet OLEDB:Database Password='" & My.Settings.dbpass & "';")
                Mycn.Open()


                SQLstr = "UPDATE mdb_ETFPrice_SecurityData SET Import = True"

                Command = New OleDb.OleDbCommand(SQLstr, Mycn)
                Command.ExecuteNonQuery()

                Mycn.Close()

                Call LoadApproved()
                Call LoadNotApproved()

            Catch ex As Exception
                MsgBox(ex.Message, MsgBoxStyle.Critical, "ERROR")
            End Try
        Else
        End If
    End Sub

    Private Sub RemoveApprovalToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RemoveApprovalToolStripMenuItem.Click
        If DataGridView2.RowCount = 0 Then
            'do nothing
        Else
            Dim Mycn As OleDb.OleDbConnection
            Dim Command As OleDb.OleDbCommand
            Dim SQLstr As String
            Try
                Mycn = New OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & My.Settings.dbloc & "';Jet OLEDB:Database Password='" & My.Settings.dbpass & "';")
                Mycn.Open()

                If DataGridView2.SelectedCells(16).Value = False Then
                    SQLstr = "UPDATE mdb_ETFPrice_SecurityData SET Import = True WHERE ID = " & DataGridView2.SelectedCells(0).Value
                Else
                    SQLstr = "UPDATE mdb_ETFPrice_SecurityData SET Import = False WHERE ID = " & DataGridView2.SelectedCells(0).Value
                End If

                Command = New OleDb.OleDbCommand(SQLstr, Mycn)
                Command.ExecuteNonQuery()

                Mycn.Close()

                Call LoadApproved()
                Call LoadNotApproved()

            Catch ex As Exception
                MsgBox(ex.Message, MsgBoxStyle.Critical, "ERROR")
            End Try

        End If
    End Sub

    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        If DataGridView2.RowCount = 0 Then
            'do nothing
        Else
            Dim ir As MsgBoxResult
            ir = MsgBox("Are you sure you are ready to send these securities to APX?", MsgBoxStyle.YesNo + MsgBoxStyle.Information, "Send to APX?")
            If ir = MsgBoxResult.Yes Then
                Call MapSecurities()
            Else
                'do nothing
            End If
        End If
    End Sub

    Public Sub MapSecurities()
        Dim Mycn As OleDb.OleDbConnection
        Dim Command As OleDb.OleDbCommand
        Dim SQLstr As String
        
        Try
            Mycn = New OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & My.Settings.dbloc & "';Jet OLEDB:Database Password='" & My.Settings.dbpass & "';")
            Mycn.Open()

            SQLstr = "INSERT INTO mdb_SecImport(Field1, Field2, Field5, Field6, Field11, Field21, Field22, Field24, Field26, Field29, Field30, Field31, Field32, Field34, Field45, Field82, Field88)" & _
            " SELECT NewSecType, NewSymbol, [Desc], AssetClassLongCode, IssueCountry, MaturityDate, MaturityDate, InterestRate, PaymentFreq, Rating, Rating2, AvgLife, YTW, Duration, CMOResetRule, NewSymbol, PrimarySymbolType FROM mdb_ETFPrice_SecurityData WHERE Import = True AND Processed = False"

            Command = New OleDb.OleDbCommand(SQLstr, Mycn)
            Command.ExecuteNonQuery()

            Mycn.Close()

            Call SaveSecFile()

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "ERROR")
        End Try
    End Sub

    Public Sub SaveSecFile()
        Dim path As String
        path = "C:\_SECFiles"

        If (System.IO.Directory.Exists(path)) Then
            System.IO.Directory.Delete(path, True)
        End If

        Dim AccessConn As New System.Data.OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & My.Settings.dbloc & "';Jet OLEDB:Database Password='" & My.Settings.dbpass & "';")

        AccessConn.Open()

        Dim AccessCommand As New OleDb.OleDbCommand("SELECT Field1,Field2,Field3,Field4,Field5,Field6,Field7,Field8,Field9,Field10,Field11,Field12,Field13,Field14,Field15,Field16,Field17,Field18,Field19,Field20,Field21,Field22,Field23,Field24,Field25,Field26,Field27,Field28,Field29,Field30,Field31,Field32,Field33,Field34,Field35,Field36,Field37,Field38,Field39,Field40,Field41,Field42,Field43,Field44,Field45,Field46,Field47,Field48,Field49,Field50,Field51,Field52,Field53,Field54,Field55,Field56,Field57,Field58,Field59,Field60,Field61,Field62,Field63,Field64,Field65,Field66,Field67,Field68,Field69,Field70,Field71,Field72,Field73,Field74,Field75,Field76,Field77,Field78,Field79,Field80,Field81,Field82,Field83,Field84,Field85,Field86,Field87,Field88,Field89,Field90,Field91,Field92,Field93,Field94,Field95,Field96,Field97,Field98,Field99,Field100,Field101,Field102,Field103,Field104,Field105,Field106,Field107 INTO [Text;HDR=NO;DATABASE=" & path & "].[sec.csv] FROM mdb_SECImport", AccessConn)
        System.IO.Directory.CreateDirectory(path)
        AccessCommand.ExecuteNonQuery()
        AccessConn.Close()

        Call CleanSecFile()

    End Sub

    Public Sub CleanSecFile()
        Dim myFiles As String()
        Dim path As String
        path = "C:\_SECFiles"
        myFiles = IO.Directory.GetFiles(path, "*.csv")

        Dim newFilePath As String

        For Each filepath As String In myFiles

            newFilePath = filepath.Replace(".csv", ".inf")

            System.IO.File.Move(filepath, newFilePath)

        Next

        Dim FileToDelete As String

        FileToDelete = "\\aamapxapps01\apx$\imp\sec.inf"

        If System.IO.File.Exists(FileToDelete) = True Then

            System.IO.File.Delete(FileToDelete)
            'MsgBox("File Deleted")
        End If

        My.Computer.FileSystem.CopyFile("C:\_SECFiles\sec.inf", "\\aamapxapps01\apx$\imp\sec.inf")

        Call ImportSECtoAPX()
    End Sub

    Public Sub ImportSECtoAPX()
        System.Diagnostics.Process.Start("\\aamapxapps01\apx$\Automation\SECImport.BAT")
        System.Diagnostics.Process.Start("\\aamapxapps01\apx$\Automation\SECImport64.BAT")

        Call MarkSECasDone()
    End Sub

    Public Sub MarkSECasDone()
        Dim Mycn As OleDb.OleDbConnection
        Dim Command As OleDb.OleDbCommand
        Dim SQLstr As String
        Try
            Mycn = New OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & My.Settings.dbloc & "';Jet OLEDB:Database Password='" & My.Settings.dbpass & "';")
            Mycn.Open()


            SQLstr = "UPDATE mdb_ETFPrice_SecurityData SET Processed = True, ProcessedDateStamp = Format(Now()), ProcessedBy = '" & Environ("USERNAME") & "' WHERE Import = True"

            Command = New OleDb.OleDbCommand(SQLstr, Mycn)
            Command.ExecuteNonQuery()

            Mycn.Close()

            Call LoadApproved()
            Call LoadNotApproved()

            Call CleanSECImportTable()

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "ERROR")
        End Try
    End Sub

    Public Sub CleanSECImportTable()
        Dim Mycn As OleDb.OleDbConnection
        Dim Command As OleDb.OleDbCommand
        Dim SQLstr As String
        Try
            Mycn = New OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & My.Settings.dbloc & "';Jet OLEDB:Database Password='" & My.Settings.dbpass & "';")
            Mycn.Open()


            SQLstr = "DELETE * FROM mdb_SECImport"

            Command = New OleDb.OleDbCommand(SQLstr, Mycn)
            Command.ExecuteNonQuery()

            Mycn.Close()

            Call LoadApproved()
            Call LoadNotApproved()

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "ERROR")
        End Try
    End Sub

    Private Sub Button5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button5.Click
        Dim ir As MsgBoxResult
        ir = MsgBox("This will reset all securities to be re-imported.  Are you sure you want to continue?", MsgBoxStyle.YesNo + MsgBoxStyle.Critical, "Confirm Reset")
        If ir = MsgBoxResult.Yes Then
            Dim Mycn As OleDb.OleDbConnection
            Dim Command As OleDb.OleDbCommand
            Dim SQLstr As String
            Try
                Mycn = New OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & My.Settings.dbloc & "';Jet OLEDB:Database Password='" & My.Settings.dbpass & "';")
                Mycn.Open()


                SQLstr = "Update mdb_ETFPrice_SecurityData SET Import = False, Processed = False, ProcessedDateStamp = Null, ProcessedBy = Null"

                Command = New OleDb.OleDbCommand(SQLstr, Mycn)
                Command.ExecuteNonQuery()

                Mycn.Close()

            Catch ex As Exception
                MsgBox(ex.Message, MsgBoxStyle.Critical, "ERROR")
            End Try

            Call LoadApproved()
            Call LoadNotApproved()

        Else

        End If
    End Sub
End Class