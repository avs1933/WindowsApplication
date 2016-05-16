Imports System.IO
Imports System.Data.OleDb

Public Class Leads

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        OpenFileDialog1.Title = "Please Select a File"
        OpenFileDialog1.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.Desktop) '"C:"
        OpenFileDialog1.DefaultExt = "CSV"
        OpenFileDialog1.Filter = "CSV files (*.csv)|*.csv|All files (*.*)|*.*"
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
        'Dim Command2 As OleDb.OleDbCommand
        Dim Command3 As OleDb.OleDbCommand
        Dim Command4 As OleDb.OleDbCommand
        Dim command5 As OleDb.OleDbCommand
        Dim command6 As OleDb.OleDbCommand
        Dim command7 As OleDb.OleDbCommand
        Dim command8 As OleDb.OleDbCommand
        Dim SQLstr As String

        'Try
        Mycn = New OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & My.Settings.dbloc & "';Jet OLEDB:Database Password='" & My.Settings.dbpass & "';")

        Dim ds1 As New DataSet
        Dim eds1 As New DataGridView
        Dim dv1 As New DataView

        Mycn.Open()

        Dim fpath As String
        Dim fname As String

        fpath = Path.GetDirectoryName(TextBox1.Text)
        fname = Path.GetFileName(TextBox1.Text)

        SQLstr = "DELETE * FROM leads_reps_working"

        Command = New OleDb.OleDbCommand(SQLstr, Mycn)
        Command.ExecuteNonQuery()

        '==========================================================================================
        'Create a Schema file to force field types of csv file
        '~~~~Commented this out because I dont think its needed.  Add back in if import fails.~~~
        '~~~~Code from UMA System~~~~
        '==========================================================================================

        'Dim FileToDelete As String

        'FileToDelete = fpath & "\Schema.ini"

        'If System.IO.File.Exists(FileToDelete) = True Then

        'System.IO.File.Delete(FileToDelete)
        'MsgBox("File Deleted")
        'End If

        'RichTextBox2.Text = "[" & fname & "]" & vbNewLine & "   ColNameHeader=False" & vbNewLine & "   Format=CSVDelimited" & vbNewLine & "   MaxScanRows=0" & vbNewLine & "   CharacterSet=OEM" & vbNewLine & "   Col1=F1 Char Width 255" & vbNewLine & "   Col2=F2 Char Width 255" & vbNewLine & "   Col3=F3 Double Width 25" & vbNewLine & "   Col4=F4 Currency Width 25" & vbNewLine & "   Col5=F5 Char Width 255"

        'RichTextBox2.SaveFile(fpath & "\Schema.ini", _
        '    RichTextBoxStreamType.PlainText)

        '==========================================================================================
        'File Created
        '==========================================================================================

        SQLstr = "Insert INTO leads_Reps_Working(Advisor, Enterprise, Firm, Phone, Address, City, State, Zip, Proposal, Investment, Target, Status, DaysOld, TMV, DateCreated)" & _
        " SELECT [Advisor], [Enterprise], [Firm], [Phone], [Address], [City], [State], [Zip], [Proposal], [Investment], [Target], [Status], [Days Old] As [DaysOld], [Value], DateAdd('d', -(daysold), Now()) FROM [Text;FMT=Delimited;HDR=YES;DATABASE=" & fpath & "]." & fname & ";"

        Command1 = New OleDb.OleDbCommand(SQLstr, Mycn)
        Command1.ExecuteNonQuery()

        'SQLstr = "Delete * FROM uma_positions_Envestnet"

        'Command2 = New OleDb.OleDbCommand(SQLstr, Mycn)
        'Command2.ExecuteNonQuery()

        SQLstr = "Insert INTO leads_Reps(Advisor, Enterprise, Firm, Phone, Address, City, State, Zip, Proposal, Investment, Target, Status, DaysOld, TMV, DateCreated)" & _
        " SELECT Advisor, Enterprise, Firm, Phone, Address, City, State, Zip, Proposal, Investment, Target, Status, DaysOld, TMV, DateCreated" & _
        " FROM leads_Reps_working" & _
        " Where Advisor Not In (Select Advisor from leads_Reps where Advisor = leads_Reps_working.Advisor AND Enterprise = leads_Reps_working.Enterprise AND day(datecreated) = day(leads_reps_working.datecreated) AND month(datecreated) = month(leads_reps_working.datecreated) AND year(datecreated) = year(leads_reps_working.datecreated) AND Firm = leads_Reps_working.Firm AND Investment = leads_Reps_working.Investment AND TMV = leads_Reps_working.TMV)"

        Command3 = New OleDb.OleDbCommand(SQLstr, Mycn)
        Command3.ExecuteNonQuery()

        SQLstr = "INSERT INTO leads_Changes (LeadID, ChangeType, OldValue, NewValue)" & _
        " SELECT leads_Reps.ID, 'Status' AS Expr1, leads_Reps.Status, leads_Reps_working.Status" & _
        " FROM leads_Reps INNER JOIN leads_Reps_working ON (leads_Reps.TMV = leads_Reps_working.TMV) AND (leads_Reps.Investment = leads_Reps_working.Investment) AND (leads_Reps.Firm = leads_Reps_working.Firm) AND (leads_Reps.Enterprise = leads_Reps_working.Enterprise) AND (leads_Reps.Advisor = leads_Reps_working.Advisor)" & _
        " WHERE (((leads_Reps.Status)<>[leads_Reps_working].[Status]) AND day(leads_Reps.datecreated) = day(leads_reps_working.datecreated) AND month(leads_Reps.datecreated) = month(leads_reps_working.datecreated) AND year(leads_Reps.datecreated) = year(leads_reps_working.datecreated));"


        Command4 = New OleDb.OleDbCommand(SQLstr, Mycn)
        Command4.ExecuteNonQuery()

        SQLstr = "INSERT INTO leads_Changes (LeadID, ChangeType, OldValue, NewValue)" & _
        " SELECT leads_Reps.ID, 'Target' AS Expr1, leads_Reps.Target, leads_Reps_working.Target" & _
        " FROM leads_Reps INNER JOIN leads_Reps_working ON (leads_Reps.TMV = leads_Reps_working.TMV) AND (leads_Reps.Investment = leads_Reps_working.Investment) AND (leads_Reps.Firm = leads_Reps_working.Firm) AND (leads_Reps.Enterprise = leads_Reps_working.Enterprise) AND (leads_Reps.Advisor = leads_Reps_working.Advisor)" & _
        " WHERE (((leads_Reps.Target)<>[leads_Reps_working].[Target]) AND day(leads_Reps.datecreated) = day(leads_reps_working.datecreated) AND month(leads_Reps.datecreated) = month(leads_reps_working.datecreated) AND year(leads_Reps.datecreated) = year(leads_reps_working.datecreated));"


        Command7 = New OleDb.OleDbCommand(SQLstr, Mycn)
        Command7.ExecuteNonQuery()

        SQLstr = "UPDATE leads_Reps_working INNER JOIN leads_Reps ON (leads_Reps_working.Advisor = leads_Reps.Advisor) AND (leads_Reps_working.Enterprise = leads_Reps.Enterprise) AND (leads_Reps_working.Firm = leads_Reps.Firm) AND (leads_Reps_working.Investment = leads_Reps.Investment) AND (leads_Reps_working.TMV = leads_Reps.TMV) SET leads_Reps.Target = Leads_Reps_working.Target, leads_Reps.Status = Leads_Reps_working.status, leads_Reps.ValuesChanged = -1, leads_Reps.DaysOld = [leads_Reps_working].[DaysOld]" & _
        " WHERE (((Day([leads_Reps].[datecreated]))=Day([leads_reps_working].[datecreated])) AND ((Month([leads_Reps].[datecreated]))=Month([leads_reps_working].[datecreated])) AND ((Year([leads_Reps].[datecreated]))=Year([leads_reps_working].[datecreated])));"

        command5 = New OleDb.OleDbCommand(SQLstr, Mycn)
        command5.ExecuteNonQuery()

        SQLstr = "UPDATE leads_Reps SET StatusID = 1 where territoryID Is Null"
        command6 = New OleDb.OleDbCommand(SQLstr, Mycn)
        command6.ExecuteNonQuery()

        SQLstr = "UPDATE leads_Reps SET StatusID = 7 where Status <> 'Invested' AND Advisor Not In (Select Advisor from leads_Reps_working where Advisor = leads_Reps.Advisor AND Enterprise = leads_Reps.Enterprise AND day(datecreated) = day(leads_reps.datecreated) AND month(datecreated) = month(leads_reps.datecreated) AND year(datecreated) = year(leads_reps.datecreated) AND Firm = leads_Reps.Firm AND Investment = leads_Reps.Investment AND TMV = leads_Reps.TMV)"
        command8 = New OleDb.OleDbCommand(SQLstr, Mycn)
        command8.ExecuteNonQuery()

        Mycn.Close()

        MsgBox("Import Sucessful!", MsgBoxStyle.Information, "IMPORT")
        'Button5.Enabled = True

        Call LoadUnassignedLeads()
        Call loadchanges()

        'Me.Close()

        'Catch ex As Exception
        'MsgBox(ex.Message, MsgBoxStyle.Critical, "ERROR")

        'End Try
    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        Call LoadUnassignedLeads()
    End Sub

    Public Sub LoadUnassignedLeads()
        Try

            Dim conn As New OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & My.Settings.dbloc & "';Jet OLEDB:Database Password='" & My.Settings.dbpass & "';")

            Dim strSQL As String = "SELECT * FROM leads_reps WHERE StatusID = 1"

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

    Public Sub loadchanges()
        Try

            Dim conn As New OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & My.Settings.dbloc & "';Jet OLEDB:Database Password='" & My.Settings.dbpass & "';")

            Dim strSQL As String = "SELECT leads_Changes.ID, leads_Changes.LeadID, leads_Changes.ChangeType AS [What Changed], leads_Changes.OldValue AS [Old Value], leads_Changes.NewValue AS [New Value], leads_Reps.Advisor, leads_Reps.Enterprise, leads_Reps.Firm, leads_Reps.Investment, leads_Reps.Status, leads_Reps.DaysOld AS [Days Old], leads_Reps.TMV, leads_Status.StatusName AS [Internal Status], leads_Reps.DateCreated AS [Date Created]" & _
            " FROM (leads_Changes INNER JOIN leads_Reps ON leads_Changes.LeadID = leads_Reps.ID) INNER JOIN leads_Status ON leads_Reps.StatusID = leads_Status.ID" & _
            " WHERE(((leads_Changes.Reviewed) = False))" & _
            " GROUP BY leads_Changes.ID, leads_Changes.LeadID, leads_Changes.ChangeType, leads_Changes.OldValue, leads_Changes.NewValue, leads_Reps.Advisor, leads_Reps.Enterprise, leads_Reps.Firm, leads_Reps.Investment, leads_Reps.Status, leads_Reps.DaysOld, leads_Reps.TMV, leads_Status.StatusName, leads_Reps.DateCreated;"

            Dim da As New OleDb.OleDbDataAdapter(strSQL, conn)
            Dim ds As New DataSet
            da.Fill(ds, "Users")

            With DataGridView2
                .DataSource = ds.Tables("Users")
                .Columns(0).Visible = False
                .Columns(1).Visible = False
            End With

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "ERROR")
        End Try
    End Sub

    Private Sub DataGridView2_CellContentClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView2.CellContentClick

    End Sub

    Private Sub DataGridView2_CellContentDoubleClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView2.CellContentDoubleClick
        If DataGridView2.RowCount = 0 Then

        Else
            Dim child As New Leads_Rep
            child.MdiParent = Home
            child.Show()
            child.ID.Text = DataGridView2.SelectedCells(1).Value
            Call child.LoadAll()
        End If

    End Sub

    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        Call loadchanges()
    End Sub

    Private Sub DataGridView1_CellContentClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick

    End Sub

    Private Sub DataGridView1_CellContentDoubleClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView1.CellContentDoubleClick
        If DataGridView1.RowCount = 0 Then

        Else
            Dim child As New Leads_Rep
            child.MdiParent = Home
            child.Show()
            child.ID.Text = DataGridView1.SelectedCells(0).Value
            Call child.LoadAll()
        End If
    End Sub

    Private Sub MarkReviedToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MarkReviedToolStripMenuItem.Click
        Dim Mycn As OleDb.OleDbConnection
        Dim Command As OleDb.OleDbCommand
        Dim SQLstr As String
        Try
            Mycn = New OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & My.Settings.dbloc & "';Jet OLEDB:Database Password='" & My.Settings.dbpass & "';")
            Mycn.Open()

            'Dim advisor1 As String
            'advisor1 = Advisor.Text
            'advisor1 = Replace(advisor1, "'", "''")

            SQLstr = "UPDATE leads_Changes SET Reviewed = True WHERE ID = " & DataGridView2.SelectedCells(0).Value

            Command = New OleDb.OleDbCommand(SQLstr, Mycn)
            Command.ExecuteNonQuery()

            Mycn.Close()

            Call loadchanges()

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "ERROR")
        End Try
    End Sub

    Private Sub Leads_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        DateTimePicker1.Text = DateAdd(DateInterval.Day, -90, Now())
        Call LoadSearchCriteria()
    End Sub

    Public Sub LoadSearchCriteria()
        Try

            Dim conn As New OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & My.Settings.roledbloc & "';Jet OLEDB:Database Password='" & My.Settings.dbpass & "';")

            Dim strSQL As String = "SELECT Distinct Status FROM leads_reps ORDER BY Status"
            Dim da As New OleDb.OleDbDataAdapter(strSQL, conn)
            Dim ds As New DataSet
            da.Fill(ds, "Users")

            With Status
                .DataSource = ds.Tables("Users")
                .DisplayMember = "Status"
                '.ValueMember = "ID"
                .SelectedIndex = 0
            End With

            Dim strSQL1 As String = "SELECT sys_Territory.ID, sys_Territory.TerritoryName" & _
            " FROM(sys_Territory)" & _
            " WHERE(((sys_Territory.Active) = True))" & _
            " GROUP BY sys_Territory.ID, sys_Territory.TerritoryName" & _
            " ORDER BY sys_Territory.TerritoryName"
            Dim da1 As New OleDb.OleDbDataAdapter(strSQL1, conn)
            Dim ds1 As New DataSet
            da1.Fill(ds1, "Users")

            With TerrID
                .DataSource = ds1.Tables("Users")
                .DisplayMember = "TerritoryName"
                .ValueMember = "ID"
                .SelectedIndex = 0
            End With

            Dim strSQL2 As String = "SELECT sys_Territory.ID, sys_Users.FullName" & _
            " FROM sys_Territory INNER JOIN sys_Users ON sys_Territory.AssignedMember = sys_Users.ID" & _
            " WHERE(((sys_Territory.Active) = True))" & _
            " GROUP BY sys_Territory.ID, sys_Users.FullName" & _
            " ORDER BY sys_Users.FullName;"
            Dim da2 As New OleDb.OleDbDataAdapter(strSQL2, conn)
            Dim ds2 As New DataSet
            da2.Fill(ds2, "Users")

            With TerrRep
                .DataSource = ds2.Tables("Users")
                .DisplayMember = "FullName"
                .ValueMember = "ID"
                .SelectedIndex = 0
            End With

            Dim strSQL3 As String = "SELECT * FROM leads_Status ORDER BY StatusName"
            Dim da3 As New OleDb.OleDbDataAdapter(strSQL3, conn)
            Dim ds3 As New DataSet
            da3.Fill(ds3, "Users")

            With StatusID
                .DataSource = ds3.Tables("Users")
                .DisplayMember = "StatusName"
                .ValueMember = "ID"
                .SelectedIndex = 0
            End With

            Dim strSQL4 As String = "SELECT Distinct State FROM leads_reps ORDER BY State"
            Dim da4 As New OleDb.OleDbDataAdapter(strSQL4, conn)
            Dim ds4 As New DataSet
            da4.Fill(ds4, "Users")

            With State
                .DataSource = ds4.Tables("Users")
                .DisplayMember = "State"
                '.ValueMember = "ID"
                .SelectedIndex = 0
            End With

            Dim strSQL5 As String = "SELECT Distinct Firm FROM leads_reps ORDER BY Firm"
            Dim da5 As New OleDb.OleDbDataAdapter(strSQL5, conn)
            Dim ds5 As New DataSet
            da5.Fill(ds5, "Users")

            With Firm
                .DataSource = ds5.Tables("Users")
                .DisplayMember = "Firm"
                '.ValueMember = "ID"
                .SelectedIndex = 0
            End With

            Dim strSQL6 As String = "SELECT Distinct Investment FROM leads_reps ORDER BY Investment"
            Dim da6 As New OleDb.OleDbDataAdapter(strSQL6, conn)
            Dim ds6 As New DataSet
            da6.Fill(ds6, "Users")

            With Investment
                .DataSource = ds6.Tables("Users")
                .DisplayMember = "Investment"
                '.ValueMember = "ID"
                .SelectedIndex = 0
            End With

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "ERROR")
            'Application.Exit()
        End Try
    End Sub

    Private Sub CheckBox1_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckBox1.CheckedChanged
        If CheckBox1.Checked Then
            Advisor.Enabled = True
        Else
            Advisor.Enabled = False
        End If
    End Sub

    Private Sub CheckBox2_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckBox2.CheckedChanged
        If CheckBox2.Checked Then
            Status.Enabled = True
        Else
            Status.Enabled = False
        End If
    End Sub

    Private Sub CheckBox3_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckBox3.CheckedChanged
        If CheckBox3.Checked Then
            StatusID.Enabled = True
        Else
            StatusID.Enabled = False
        End If
    End Sub

    Private Sub CheckBox4_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckBox4.CheckedChanged
        If CheckBox4.Checked Then
            Investment.Enabled = True
        Else
            Investment.Enabled = False
        End If
    End Sub

    Private Sub CheckBox5_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckBox5.CheckedChanged
        If CheckBox5.Checked Then
            TerrRep.Enabled = True
        Else
            TerrRep.Enabled = False
        End If
    End Sub

    Private Sub CheckBox6_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckBox6.CheckedChanged
        If CheckBox6.Checked Then
            TerrID.Enabled = True
        Else
            TerrID.Enabled = False
        End If
    End Sub

    Private Sub CheckBox7_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckBox7.CheckedChanged
        If CheckBox7.Checked Then
            State.Enabled = True
        Else
            State.Enabled = False
        End If
    End Sub

    Private Sub CheckBox8_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckBox8.CheckedChanged
        If CheckBox8.Checked Then
            Firm.Enabled = True
        Else
            Firm.Enabled = False
        End If
    End Sub

    Private Sub Button5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button5.Click
        Try

            Dim conn As New OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & My.Settings.dbloc & "';Jet OLEDB:Database Password='" & My.Settings.dbpass & "';")

            Dim strSQL As String = "SELECT * FROM leads_reps WHERE DateCreated >= #" & DateTimePicker1.Text & "# AND DateCreated <= #" & DateTimePicker2.Text & "#"

            If CheckBox1.Checked = True Then
                strSQL = strSQL & " AND Advisor Like '%" & Advisor.Text & "%'"
            Else
            End If

            If CheckBox2.Checked = True Then
                strSQL = strSQL & " AND Status = '" & Status.Text & "'"
            Else
            End If

            If CheckBox3.Checked = True Then
                strSQL = strSQL & " AND StatusID = " & StatusID.SelectedValue
            Else
            End If

            If CheckBox4.Checked = True Then
                strSQL = strSQL & " AND Investment = '" & Investment.Text & "'"
            Else
            End If

            If CheckBox5.Checked = True Then
                strSQL = strSQL & " AND TerritoryID = " & TerrRep.SelectedValue
            Else
            End If

            If CheckBox6.Checked = True Then
                strSQL = strSQL & " AND TerritoryID = " & TerrID.SelectedValue
            Else
            End If

            If CheckBox7.Checked = True Then
                strSQL = strSQL & " AND State = '" & State.Text & "'"
            Else
            End If

            If CheckBox8.Checked = True Then
                strSQL = strSQL & " AND Firm = '" & Firm.Text & "'"
            Else
            End If

            Dim da As New OleDb.OleDbDataAdapter(strSQL, conn)
            Dim ds As New DataSet
            da.Fill(ds, "Users")

            With DataGridView3
                .DataSource = ds.Tables("Users")
                .Columns(0).Visible = False
            End With

            Label13.Text = "Found " & DataGridView3.RowCount.ToString & " record(s)."

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "ERROR")
        End Try
    End Sub

    Private Sub DataGridView3_CellContentDoubleClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView3.CellContentDoubleClick
        If DataGridView3.RowCount = 0 Then

        Else
            Dim child As New Leads_Rep
            child.MdiParent = Home
            child.Show()
            child.ID.Text = DataGridView3.SelectedCells(0).Value
            Call child.LoadAll()
        End If
    End Sub
End Class