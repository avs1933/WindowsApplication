Public Class Leads_Rep

    Private Sub Leads_Rep_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Call LoadComboBoxes()
    End Sub

    Public Sub LoadComboBoxes()
        Try

            Dim conn As New OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & My.Settings.roledbloc & "';Jet OLEDB:Database Password='" & My.Settings.dbpass & "';")

            Dim strSQL As String = "SELECT * FROM leads_Status ORDER BY StatusName"
            Dim da As New OleDb.OleDbDataAdapter(strSQL, conn)
            Dim ds As New DataSet
            da.Fill(ds, "Users")

            With cboIStatus
                .DataSource = ds.Tables("Users")
                .DisplayMember = "StatusName"
                .ValueMember = "ID"
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

            With cboTerritory
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

            With cboMember
                .DataSource = ds2.Tables("Users")
                .DisplayMember = "FullName"
                .ValueMember = "ID"
                .SelectedIndex = 0
            End With

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "ERROR")
            'Application.Exit()
        End Try
    End Sub

    Private Sub CheckBox1_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckBox1.CheckedChanged
        If CheckBox1.Checked Then
            cboTerritory.Enabled = True
            cboMember.Enabled = False
            CheckBox2.Checked = False
        Else
            cboTerritory.Enabled = False
            cboMember.Enabled = True
            CheckBox2.Checked = True
            CheckBox1.Checked = False
        End If
    End Sub

    Private Sub CheckBox2_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckBox2.CheckedChanged
        If CheckBox2.Checked Then
            CheckBox1.Checked = False
            cboMember.Enabled = True
            cboTerritory.Enabled = False
        Else
            CheckBox1.Checked = True
        End If
    End Sub

    Public Sub LoadAll()
        Try
            Dim conn As New OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & My.Settings.dbloc & "';Jet OLEDB:Database Password='" & My.Settings.dbpass & "';")

            Dim strSQL As String = "SELECT * from leads_reps where ID = " & ID.Text
            Dim queryString As String = String.Format(strSQL)
            Dim cmd As New OleDb.OleDbCommand(queryString, conn)
            Dim da As New OleDb.OleDbDataAdapter(cmd)
            Dim ds As New DataSet

            da.Fill(ds, "User")
            Dim dt As DataTable = ds.Tables("User")

            Dim row1 As DataRow = dt.Rows(0)

            If IsDBNull(row1("Advisor")) Then
                Advisor.Text = "Unknown"
            Else
                Advisor.Text = (row1("Advisor"))
            End If

            If IsDBNull(row1("Enterprise")) Then
                Enterprise.Text = "Unknown"
            Else
                Enterprise.Text = (row1("Enterprise"))
            End If

            If IsDBNull(row1("Firm")) Then
                Firm.Text = "Unknown"
            Else
                Firm.Text = (row1("Firm"))
            End If

            If IsDBNull(row1("Phone")) Then
                Phone.Text = "Unknown"
            Else
                Phone.Text = (row1("Phone"))
            End If

            If IsDBNull(row1("Address")) Then
                Address.Text = "Unknown"
            Else
                Address.Text = (row1("Address"))
            End If

            If IsDBNull(row1("City")) Then
                City.Text = "Unknown"
            Else
                City.Text = (row1("City"))
            End If

            If IsDBNull(row1("State")) Then
                State.Text = "Unknown"
            Else
                State.Text = (row1("State"))
            End If

            If IsDBNull(row1("Zip")) Then
                Zip.Text = "Unknown"
            Else
                Zip.Text = (row1("Zip"))
            End If

            If IsDBNull(row1("Proposal")) Then
                Proposal.Text = "Unknown"
            Else
                Proposal.Text = (row1("Proposal"))
            End If

            If IsDBNull(row1("Investment")) Then
                Investment.Text = "Unknown"
            Else
                Investment.Text = (row1("Investment"))
            End If

            If IsDBNull(row1("Target")) Then
                Target.Text = "Unknown"
            Else
                Target.Text = (row1("Target"))
            End If

            If IsDBNull(row1("Status")) Then
                Status.Text = "Unknown"
            Else
                Status.Text = (row1("Status"))
            End If

            If IsDBNull(row1("DaysOld")) Then
                DaysOld.Text = "Unknown"
            Else
                DaysOld.Text = (row1("DaysOld"))
            End If

            If IsDBNull(row1("TMV")) Then
                TMV.Text = "Unknown"
            Else
                TMV.Text = (row1("TMV"))
            End If

            If IsDBNull(row1("Notes")) Then
                RichTextBox1.Text = "-None-"
            Else
                RichTextBox1.Text = (row1("Notes"))
            End If

            If IsDBNull(row1("TerritoryID")) Then
                TerritoryID.Text = "NEW"
            Else
                TerritoryID.Text = (row1("TerritoryID"))
            End If

            If IsDBNull(row1("DateAssigned")) Then
                'TerritoryID.Text = "NEW"
            Else
                DateTimePicker2.Text = (row1("DateAssigned"))
            End If

            DateTimePicker1.Text = (row1("DateCreated"))

            cboIStatus.SelectedValue = (row1("StatusID"))

            Call LoadTerritory()
            Call CheckForKnown()

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "ERROR")
        End Try
    End Sub

    Private Sub cboMember_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cboMember.SelectedIndexChanged
        TerritoryID.Text = cboMember.SelectedValue.ToString
        Call LoadTerritory()
    End Sub

    Private Sub cboTerritory_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cboTerritory.SelectedIndexChanged
        TerritoryID.Text = cboTerritory.SelectedValue.ToString
        Call LoadTerritory()
    End Sub

    Public Sub LoadTerritory()
        Dim value As Integer

        If Integer.TryParse(TerritoryID.Text, value) Then

            Try
                Dim conn As New OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & My.Settings.dbloc & "';Jet OLEDB:Database Password='" & My.Settings.dbpass & "';")

                Dim strSQL As String = "SELECT sys_Territory.ID, sys_Department.DepartmentName, sys_Region.RegionName, sys_Territory.TerritoryName, sys_Users.FullName" & _
                " FROM ((sys_Territory INNER JOIN sys_Region ON sys_Territory.RegionID = sys_Region.ID) INNER JOIN sys_Department ON sys_Region.DepartmentID = sys_Department.ID) INNER JOIN sys_Users ON sys_Territory.AssignedMember = sys_Users.ID" & _
                " Where sys_Territory.ID = " & TerritoryID.Text
                Dim queryString As String = String.Format(strSQL)
                Dim cmd As New OleDb.OleDbCommand(queryString, conn)
                Dim da As New OleDb.OleDbDataAdapter(cmd)
                Dim ds As New DataSet

                da.Fill(ds, "User")
                Dim dt As DataTable = ds.Tables("User")

                Dim row1 As DataRow = dt.Rows(0)

                If IsDBNull(row1("DepartmentName")) Then
                    DepartmentName.Text = "Unknown"
                Else
                    DepartmentName.Text = (row1("DepartmentName"))
                End If

                If IsDBNull(row1("RegionName")) Then
                    RegionName.Text = "Unknown"
                Else
                    RegionName.Text = (row1("RegionName"))
                End If

                If IsDBNull(row1("TerritoryName")) Then
                    TerritoryName.Text = "Unknown"
                Else
                    TerritoryName.Text = (row1("TerritoryName"))
                End If

                If IsDBNull(row1("FullName")) Then
                    FullName.Text = "Unknown"
                Else
                    FullName.Text = (row1("FullName"))
                End If

            Catch ex As Exception
                MsgBox(ex.Message, MsgBoxStyle.Critical, "ERROR")
            End Try

        Else

        End If
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        Dim child As New Lead_AdvAssociate
        child.MdiParent = Home
        child.Show()
    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        Call CheckForKnown()
    End Sub

    Public Sub CheckAccounts()
        Try

            Dim conn As New OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & My.Settings.dbloc & "';Jet OLEDB:Database Password='" & My.Settings.dbpass & "';")

            Dim advisor1 As String
            advisor1 = Advisor.Text
            advisor1 = Replace(advisor1, "'", "''")

            Dim strSQL As String = "SELECT dbo_vQbRowDefPortfolio.PortfolioCode As [Portfolio Code], dbo_vQbRowDefPortfolio.TotalMarketValue As [TMV], dbo_vQbRowDefPortfolio.Discipline6 As [Discipline], dbo_vQbRowDefPortfolio.AAMRepName As [AAM Rep]" & _
            " FROM leads_advisors INNER JOIN dbo_vQbRowDefPortfolio ON leads_advisors.APXName = dbo_vQbRowDefPortfolio.ExternalAdvisor" & _
            " WHERE(((dbo_vQbRowDefPortfolio.PortfolioStatus) = 'Open') And ((leads_advisors.LeadName) = '" & advisor1 & "'))" & _
            " GROUP BY dbo_vQbRowDefPortfolio.PortfolioCode, dbo_vQbRowDefPortfolio.TotalMarketValue, dbo_vQbRowDefPortfolio.Discipline6, dbo_vQbRowDefPortfolio.AAMRepName;"
            Dim da As New OleDb.OleDbDataAdapter(strSQL, conn)
            Dim ds As New DataSet
            da.Fill(ds, "Users")

            With DataGridView2
                .DataSource = ds.Tables("Users")
                '.Columns(0).Visible = False
                '.Columns(1).Visible = False
            End With

            Label25.Text = "Found " & DataGridView2.RowCount.ToString & " record(s)."

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "ERROR")
        End Try
    End Sub

    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        Call CheckForKnown()
    End Sub

    Public Sub CheckLeads()
        Try

            Dim conn As New OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & My.Settings.dbloc & "';Jet OLEDB:Database Password='" & My.Settings.dbpass & "';")

            Dim advisor1 As String
            advisor1 = Advisor.Text
            advisor1 = Replace(advisor1, "'", "''")

            'Dim strSQL As String = "SELECT leads_Reps.DaysOld AS [Days Old], leads_Status.StatusName AS [AAM Status], sys_Territory.TerritoryName AS Territory, sys_Users.FullName AS [Assigned To], leads_Reps.Status AS [ENV Status], leads_Reps.Investment" & _
            '" FROM ((leads_advisors INNER JOIN leads_Reps ON leads_advisors.LeadName = leads_Reps.Advisor) INNER JOIN leads_Status ON leads_Reps.StatusID = leads_Status.ID) INNER JOIN (sys_Territory INNER JOIN sys_Users ON sys_Territory.AssignedMember = sys_Users.ID) ON leads_Reps.TerritoryID = sys_Territory.ID" & _
            '" WHERE leads_Reps.Advisor = '" & advisor1 & "'"

            Dim strSql As String = "SELECT Leads_Reps.ID, leads_Reps.DaysOld AS [Days Old], leads_Status.StatusName AS [AAM Status], leads_Reps.DateCreated AS [Date Created], leads_Reps.Investment, leads_Reps.TMV" & _
            " FROM leads_Reps INNER JOIN leads_Status ON leads_Reps.StatusID = leads_Status.ID" & _
            " WHERE leads_Reps.Advisor = '" & advisor1 & "'"

            Dim da As New OleDb.OleDbDataAdapter(strSQL, conn)
            Dim ds As New DataSet
            da.Fill(ds, "Users")

            With DataGridView1
                .DataSource = ds.Tables("Users")
                .Columns(0).Visible = False
                '.Columns(1).Visible = False
            End With

            Label26.Text = "Found " & DataGridView1.RowCount.ToString & " record(s)."

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "ERROR")
        End Try
    End Sub

    Public Sub CheckForKnown()
        Dim Mycn As OleDb.OleDbConnection

        Try
            Mycn = New OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & My.Settings.dbloc & "';Jet OLEDB:Database Password='" & My.Settings.dbpass & "';")
            Mycn.Open()
            Dim sqlstring As String

            Dim advisor1 As String
            advisor1 = Advisor.Text
            advisor1 = Replace(advisor1, "'", "''")

            sqlstring = "SELECT * FROM leads_advisors WHERE LeadName = '" & advisor1 & "'"
            Dim queryString As String = String.Format(sqlstring)
            Dim cmd As New OleDb.OleDbCommand(queryString, Mycn)
            Dim da As New OleDb.OleDbDataAdapter(cmd)
            Dim ds As New DataSet
            da.Fill(ds, "User")
            Dim dt As DataTable = ds.Tables("User")
            If dt.Rows.Count > 0 Then
                Label27.Text = "KNOWN ADVISOR!"
                Label27.BackColor = Color.Green
                Label27.ForeColor = Color.White
                Call CheckLeads()
                Call CheckAccounts()
                Button2.Enabled = False
            Else
                Label27.Text = "NEW ADVISOR!"
                Label27.BackColor = Color.Red
                Label27.ForeColor = Color.White
                Call CheckLeads()
                Button2.Enabled = True
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "ERROR")
        End Try
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        If ID.Text = "NEW" Then

        Else
            Call SaveOld()
        End If
    End Sub

    Public Sub SaveOld()
        Dim Mycn As OleDb.OleDbConnection
        Dim Command As OleDb.OleDbCommand
        Dim SQLstr As String
        Try
            Mycn = New OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & My.Settings.dbloc & "';Jet OLEDB:Database Password='" & My.Settings.dbpass & "';")
            Mycn.Open()

            Dim advisor1 As String
            advisor1 = Advisor.Text
            advisor1 = Replace(advisor1, "'", "''")
            Dim dteassigned As Date
            dteassigned = DateTimePicker2.Text

            SQLstr = "UPDATE leads_Reps SET Advisor = '" & advisor1 & "', Enterprise = '" & Enterprise.Text & "', Firm = '" & Firm.Text & "', [Phone] = '" & Phone.Text & "', [Address] = '" & Address.Text & "', [City] = '" & City.Text & "', [State] = '" & State.Text & "', [Zip] = '" & Zip.Text & "', [Proposal] = '" & Proposal.Text & "', [Investment] = '" & Investment.Text & "', [Target] = '" & Target.Text & "', [Status] = '" & Status.Text & "', [DaysOld] = " & DaysOld.Text & ", [TMV] = " & TMV.Text & ", [TerritoryID] = " & TerritoryID.Text & ", [DateAssigned] = #" & dteassigned & "#, [StatusID] = " & cboIStatus.SelectedValue & ", [Notes] = '" & RichTextBox1.Text & "' WHERE ID = " & ID.Text

            Command = New OleDb.OleDbCommand(SQLstr, Mycn)
            Command.ExecuteNonQuery()

            Mycn.Close()

            MsgBox("Record Updated.")
            Me.Close()

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "ERROR")
        End Try
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
End Class