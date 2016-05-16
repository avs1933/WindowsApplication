Public Class SalesTeams

    Private Sub SalesTeams_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Call LoadDepartmentManager()
        Call LoadDepartments()
        Call LoadRegions()
        Call LoadDepartmentList()
        Call LoadRoles()
        Call LoadTerritories()
        'Call LoadRegionList()

        If (Permissions.AddUser.Checked = False Or Permissions.EditUser.Checked = False) Then
            GroupBox4.Enabled = False
        Else
            GroupBox4.Enabled = True
        End If
    End Sub

    Public Sub LoadDepartmentManager()
        Try

            Dim conn As New OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & My.Settings.roledbloc & "';Jet OLEDB:Database Password='" & My.Settings.dbpass & "';")

            Dim strSQL As String = "SELECT * FROM sys_Users ORDER BY FullName"
            Dim da As New OleDb.OleDbDataAdapter(strSQL, conn)
            Dim ds As New DataSet
            da.Fill(ds, "Users")

            With ComboBox1
                .DataSource = ds.Tables("Users")
                .DisplayMember = "FullName"
                .ValueMember = "ID"
                .SelectedIndex = 0
            End With

            With ComboBox2
                .DataSource = ds.Tables("Users")
                .DisplayMember = "FullName"
                .ValueMember = "ID"
                .SelectedIndex = 0
            End With

            With ComboBox7
                .DataSource = ds.Tables("Users")
                .DisplayMember = "FullName"
                .ValueMember = "ID"
                .SelectedIndex = 0
            End With

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "ERROR")
            'Application.Exit()
        End Try
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        If DID.Text = "NEW" Then
            Call SaveNewDept()
        Else
            Call SaveOldDept()
        End If
    End Sub

    Public Sub SaveNewDept()
        If txtDeptName.Text = "" Or IsDBNull(txtDeptName) Then
            txtDeptName.BackColor = Color.Red
            txtDeptName.ForeColor = Color.White
        Else
            txtDeptName.BackColor = Color.White
            txtDeptName.ForeColor = Color.Black
            If ComboBox1.Text = "" Or IsDBNull(ComboBox1.Text) Then
                ComboBox1.BackColor = Color.Red
                ComboBox1.ForeColor = Color.White
            Else
                ComboBox1.BackColor = Color.White
                ComboBox1.ForeColor = Color.Black
                Dim Mycn As OleDb.OleDbConnection
                Dim Command As OleDb.OleDbCommand
                Dim SQLstr As String

                Try
                    Mycn = New OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & My.Settings.dbloc & "';Jet OLEDB:Database Password='" & My.Settings.dbpass & "';")
                    Mycn.Open()

                    SQLstr = "INSERT INTO sys_Department(DepartmentName, Manager, Active)" & _
                    "VALUES('" & txtDeptName.Text & "', " & ComboBox1.SelectedValue & ", -1)"

                    Command = New OleDb.OleDbCommand(SQLstr, Mycn)
                    Command.ExecuteNonQuery()

                    Mycn.Close()

                    Call LoadDepartments()

                Catch ex As Exception
                    MsgBox(ex.Message, MsgBoxStyle.Critical, "ERROR")
                End Try
            End If
        End If
    End Sub

    Public Sub SaveOldDept()
        If txtDeptName.Text = "" Or IsDBNull(txtDeptName) Then
            txtDeptName.BackColor = Color.Red
            txtDeptName.ForeColor = Color.White
        Else
            txtDeptName.BackColor = Color.White
            txtDeptName.ForeColor = Color.Black
            If ComboBox1.Text = "" Or IsDBNull(ComboBox1.Text) Then
                ComboBox1.BackColor = Color.Red
                ComboBox1.ForeColor = Color.White
            Else
                ComboBox1.BackColor = Color.White
                ComboBox1.ForeColor = Color.Black
                Dim Mycn As OleDb.OleDbConnection
                Dim Command As OleDb.OleDbCommand
                Dim SQLstr As String

                Try
                    Mycn = New OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & My.Settings.dbloc & "';Jet OLEDB:Database Password='" & My.Settings.dbpass & "';")
                    Mycn.Open()

                    SQLstr = "UPDATE sys_Department SET DepartmentName = '" & txtDeptName.Text & "', Manager = " & ComboBox1.SelectedValue & " WHERE ID = " & DID.Text

                    Command = New OleDb.OleDbCommand(SQLstr, Mycn)
                    Command.ExecuteNonQuery()

                    Mycn.Close()

                    Call LoadDepartments()

                Catch ex As Exception
                    MsgBox(ex.Message, MsgBoxStyle.Critical, "ERROR")
                End Try
            End If
        End If
    End Sub

    Public Sub LoadDepartments()
        Try

            Dim conn As New OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & My.Settings.dbloc & "';Jet OLEDB:Database Password='" & My.Settings.dbpass & "';")

            Dim strSQL As String = "SELECT sys_Department.ID, sys_Users.ID, sys_Department.DepartmentName As [Department Name], sys_Users.FullName As [Manager Name]" & _
            " FROM sys_Department INNER JOIN sys_Users ON sys_Department.Manager = sys_Users.ID" & _
            " WHERE(((sys_Department.Active) = True))" & _
            " GROUP BY sys_Department.ID, sys_Users.ID, sys_Department.DepartmentName, sys_Users.FullName;"

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

    Public Sub LoadTerritories()
        Try

            Dim conn As New OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & My.Settings.dbloc & "';Jet OLEDB:Database Password='" & My.Settings.dbpass & "';")

            Dim strSQL As String = "SELECT sys_Territory.ID, sys_Department.ID, sys_Territory.RegionID, sys_Users.ID, sys_Territory.TerritoryName AS Territory, sys_Users.FullName AS [Assigned Member], sys_Region.RegionName AS Region, sys_Department.DepartmentName AS Department, sys_Territory.Runway, sys_Territory.DateInTerritory As [Date in Territory], sys_Territory.Licensed" & _
            " FROM ((sys_Territory INNER JOIN sys_Region ON sys_Territory.RegionID = sys_Region.ID) INNER JOIN sys_Department ON sys_Region.DepartmentID = sys_Department.ID) INNER JOIN sys_Users ON sys_Territory.AssignedMember = sys_Users.ID" & _
            " WHERE(((sys_Territory.Active) = True))" & _
            " GROUP BY sys_Territory.ID, sys_Department.ID, sys_Territory.RegionID, sys_Users.ID, sys_Territory.TerritoryName, sys_Users.FullName, sys_Region.RegionName, sys_Department.DepartmentName, sys_Territory.Runway, sys_Territory.DateInTerritory, sys_Territory.Licensed;"


            Dim da As New OleDb.OleDbDataAdapter(strSQL, conn)
            Dim ds As New DataSet
            da.Fill(ds, "Users")

            With DataGridView3
                .DataSource = ds.Tables("Users")
                .Columns(0).Visible = False
                .Columns(1).Visible = False
                .Columns(2).Visible = False
                .Columns(3).Visible = False
            End With

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "ERROR")
        End Try
    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        DID.Text = "NEW"
        txtDeptName.Text = ""
        Call LoadDepartmentManager()
    End Sub

    Private Sub DataGridView1_CellContentClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick
        DID.Text = DataGridView1.SelectedCells(0).Value
        ComboBox1.SelectedValue = DataGridView1.SelectedCells(1).Value
        txtDeptName.Text = DataGridView1.SelectedCells(2).Value
    End Sub

    Public Sub LoadRegions()
        Try

            Dim conn As New OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & My.Settings.dbloc & "';Jet OLEDB:Database Password='" & My.Settings.dbpass & "';")

            Dim strSQL As String = "SELECT sys_Region.ID, sys_Department.ID, sys_Users.ID, sys_Region.RegionName As [Region Name], sys_Department.DepartmentName AS [Department Name], sys_Users.FullName AS [Manager Name]" & _
            " FROM sys_Department INNER JOIN (sys_Region INNER JOIN sys_Users ON sys_Region.ManagerName = sys_Users.ID) ON sys_Department.ID = sys_Region.DepartmentID" & _
            " GROUP BY sys_Region.ID, sys_Department.ID, sys_Users.ID, sys_Region.RegionName, sys_Department.DepartmentName, sys_Users.FullName"


            Dim da As New OleDb.OleDbDataAdapter(strSQL, conn)
            Dim ds As New DataSet
            da.Fill(ds, "Users")

            With DataGridView2
                .DataSource = ds.Tables("Users")
                .Columns(0).Visible = False
                .Columns(1).Visible = False
                .Columns(2).Visible = False
            End With

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "ERROR")
        End Try
    End Sub

    Public Sub LoadDepartmentList()
        Try

            Dim conn As New OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & My.Settings.roledbloc & "';Jet OLEDB:Database Password='" & My.Settings.dbpass & "';")

            Dim strSQL As String = "SELECT * FROM sys_Department ORDER BY DepartmentName"
            Dim da As New OleDb.OleDbDataAdapter(strSQL, conn)
            Dim ds As New DataSet
            da.Fill(ds, "Users")

            With ComboBox3
                .DataSource = ds.Tables("Users")
                .DisplayMember = "DepartmentName"
                .ValueMember = "ID"
                .SelectedIndex = 0
            End With

            With ComboBox8
                .DataSource = ds.Tables("Users")
                .DisplayMember = "DepartmentName"
                .ValueMember = "ID"
                .SelectedIndex = 0
            End With

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "ERROR")
            'Application.Exit()
        End Try
    End Sub

    Public Sub LoadRegionList()
        Try

            Dim conn As New OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & My.Settings.roledbloc & "';Jet OLEDB:Database Password='" & My.Settings.dbpass & "';")

            Dim strSQL As String = "SELECT * FROM sys_Region WHERE DepartmentID = " & ComboBox8.SelectedValue & " ORDER BY RegionName"
            Dim da As New OleDb.OleDbDataAdapter(strSQL, conn)
            Dim ds As New DataSet
            da.Fill(ds, "Users")

            With ComboBox6
                .DataSource = ds.Tables("Users")
                .DisplayMember = "RegionName"
                .ValueMember = "ID"
                .SelectedIndex = 0
            End With


        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "ERROR")
            'Application.Exit()
        End Try
    End Sub

    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        If RID.Text = "NEW" Then
            Call RegionNew()
        Else
            Call RegionOld()
        End If
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        RID.Text = "NEW"
        txtRegionName.Text = ""
        Call LoadDepartmentManager()
        Call LoadDepartmentList()
    End Sub

    Private Sub DataGridView2_CellContentClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView2.CellContentClick
        RID.Text = DataGridView2.SelectedCells(0).Value
        ComboBox3.SelectedValue = DataGridView2.SelectedCells(1).Value
        ComboBox2.SelectedValue = DataGridView2.SelectedCells(2).Value
        txtRegionName.Text = DataGridView2.SelectedCells(3).Value
    End Sub

    Public Sub RegionNew()
        If txtRegionName.Text = "" Or IsDBNull(txtRegionName) Then
            txtRegionName.BackColor = Color.Red
            txtRegionName.ForeColor = Color.White
        Else
            txtRegionName.BackColor = Color.White
            txtRegionName.ForeColor = Color.Black
            Dim Mycn As OleDb.OleDbConnection
            Dim Command As OleDb.OleDbCommand
            Dim SQLstr As String

            Try
                Mycn = New OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & My.Settings.dbloc & "';Jet OLEDB:Database Password='" & My.Settings.dbpass & "';")
                Mycn.Open()

                SQLstr = "INSERT INTO sys_Region(RegionName, DepartmentID, ManagerName, Active)" & _
                "VALUES('" & txtRegionName.Text & "', " & ComboBox3.SelectedValue & "," & ComboBox2.SelectedValue & ", -1)"

                Command = New OleDb.OleDbCommand(SQLstr, Mycn)
                Command.ExecuteNonQuery()

                Mycn.Close()

                Call LoadRegions()

            Catch ex As Exception
                MsgBox(ex.Message, MsgBoxStyle.Critical, "ERROR")
            End Try
        End If
    End Sub

    Public Sub RegionOld()
        If txtRegionName.Text = "" Or IsDBNull(txtRegionName) Then
            txtRegionName.BackColor = Color.Red
            txtRegionName.ForeColor = Color.White
        Else
            txtRegionName.BackColor = Color.White
            txtRegionName.ForeColor = Color.Black
            Dim Mycn As OleDb.OleDbConnection
            Dim Command As OleDb.OleDbCommand
            Dim SQLstr As String

            Try
                Mycn = New OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & My.Settings.dbloc & "';Jet OLEDB:Database Password='" & My.Settings.dbpass & "';")
                Mycn.Open()

                SQLstr = "UPDATE sys_Region SET RegionName = '" & txtRegionName.Text & "', DepartmentID = " & ComboBox3.SelectedValue & ", ManagerName = " & ComboBox2.SelectedValue & " WHERE ID = " & RID.Text
                Command = New OleDb.OleDbCommand(SQLstr, Mycn)
                Command.ExecuteNonQuery()

                Mycn.Close()

                Call LoadRegions()

            Catch ex As Exception
                MsgBox(ex.Message, MsgBoxStyle.Critical, "ERROR")
            End Try
        End If
    End Sub

    Private Sub Button10_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button10.Click
        Try

            Dim conn As New OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & My.Settings.dbloc & "';Jet OLEDB:Database Password='" & My.Settings.dbpass & "';")

            Dim sqlstring As String

            sqlstring = "SELECT * FROM sys_Users WHERE FullName = '" & txtFullName.Text & "' AND NetworkID = '" & txtNetworkID.Text & "' AND Active = -1"

            Dim queryString As String = String.Format(sqlstring)
            Dim cmd As New OleDb.OleDbCommand(queryString, conn)
            Dim da1 As New OleDb.OleDbDataAdapter(cmd)
            Dim ds1 As New DataSet

            da1.Fill(ds1, "User")
            Dim dt1 As DataTable = ds1.Tables("User")
            If dt1.Rows.Count = 0 Then

                Dim Mycn As OleDb.OleDbConnection
                Dim Command As OleDb.OleDbCommand
                Dim SQLstr As String

                Mycn = New OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & My.Settings.dbloc & "';Jet OLEDB:Database Password='" & My.Settings.dbpass & "';")
                Mycn.Open()

                SQLstr = "INSERT INTO sys_Users (FullName, NetworkID, WI, Active, APXID, RoleID, APXName)" & _
                "VALUES('" & txtFullName.Text & "','" & txtNetworkID.Text & "',-1,-1,0," & cboRole.SelectedValue & ",'" & txtAPXName.Text & "')"

                Command = New OleDb.OleDbCommand(SQLstr, Mycn)
                Command.ExecuteNonQuery()

                Mycn.Close()

                Call LoadDepartmentManager()

            Else
                MsgBox("A record already exsists for that rep.", MsgBoxStyle.Exclamation, "Duplicate Record.")
            End If

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "ERROR")
        End Try
    End Sub

    Private Sub Button9_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button9.Click
        txtAPXName.Text = ""
        txtDeptName.Text = ""
        txtFullName.Text = ""
        txtNetworkID.Text = ""
        Call LoadRoles()
    End Sub

    Public Sub LoadRoles()
        Try

            Dim conn As New OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & My.Settings.roledbloc & "';Jet OLEDB:Database Password='" & My.Settings.dbpass & "';")

            Dim strSQL As String = "SELECT * FROM sys_Roles WHERE Active = -1 ORDER BY RoleName"
            Dim da As New OleDb.OleDbDataAdapter(strSQL, conn)
            Dim ds As New DataSet
            da.Fill(ds, "Users")

            With cboRole
                .DataSource = ds.Tables("Users")
                .DisplayMember = "RoleName"
                .ValueMember = "ID"
                .SelectedIndex = 0
            End With

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "ERROR")
            'Application.Exit()
        End Try
    End Sub

    Private Sub Button7_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button7.Click
        TID.Text = "NEW"
        txtTerritoryName.Text = ""
        Call LoadDepartmentManager()
        Call LoadDepartmentList()
        Call LoadRegionList()

    End Sub

    Private Sub ComboBox8_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles ComboBox8.LostFocus
        Call LoadRegionList()
    End Sub

    Private Sub ComboBox8_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ComboBox8.SelectedIndexChanged
        'Call LoadRegionList()
    End Sub

    Private Sub Button8_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button8.Click
        If TID.Text = "NEW" Then
            Call TerritoryNew()
        Else
            Call territoryold()
        End If
    End Sub

    Public Sub TerritoryNew()
        If txtTerritoryName.Text = "" Or IsDBNull(txtTerritoryName) Then
            txtTerritoryName.BackColor = Color.Red
            txtTerritoryName.ForeColor = Color.White
        Else
            txtTerritoryName.BackColor = Color.White
            txtTerritoryName.ForeColor = Color.Black
            Dim Mycn As OleDb.OleDbConnection
            Dim Command As OleDb.OleDbCommand
            Dim SQLstr As String

            Try
                Mycn = New OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & My.Settings.dbloc & "';Jet OLEDB:Database Password='" & My.Settings.dbpass & "';")
                Mycn.Open()

                SQLstr = "INSERT INTO sys_Territory(TerritoryName, RegionID, AssignedMember, Active, Runway, DateInTerritory, Licensed)" & _
                "VALUES('" & txtTerritoryName.Text & "', " & ComboBox6.SelectedValue & "," & ComboBox7.SelectedValue & ", -1," & CheckBox1.CheckState & ",#" & DateTimePicker1.Text & "#," & CheckBox2.CheckState & ")"

                Command = New OleDb.OleDbCommand(SQLstr, Mycn)
                Command.ExecuteNonQuery()

                Mycn.Close()

                Call LoadTerritories()

            Catch ex As Exception
                MsgBox(ex.Message, MsgBoxStyle.Critical, "ERROR")
            End Try
        End If
    End Sub

    Public Sub territoryold()
        If txtTerritoryName.Text = "" Or IsDBNull(txtTerritoryName) Then
            txtTerritoryName.BackColor = Color.Red
            txtTerritoryName.ForeColor = Color.White
        Else
            txtTerritoryName.BackColor = Color.White
            txtTerritoryName.ForeColor = Color.Black
            Dim Mycn As OleDb.OleDbConnection
            Dim Command As OleDb.OleDbCommand
            Dim SQLstr As String

            Try
                Mycn = New OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & My.Settings.dbloc & "';Jet OLEDB:Database Password='" & My.Settings.dbpass & "';")
                Mycn.Open()

                SQLstr = "UPDATE sys_Territory SET TerritoryName = '" & txtTerritoryName.Text & "', RegionID = " & ComboBox6.SelectedValue & ", AssignedMember = " & ComboBox7.SelectedValue & ", Runway = " & CheckBox1.CheckState & ", DateInTerritory = #" & DateTimePicker1.Text & "#, Licensed = " & CheckBox2.CheckState & " WHERE ID = " & TID.Text

                Command = New OleDb.OleDbCommand(SQLstr, Mycn)
                Command.ExecuteNonQuery()

                Mycn.Close()

                Call LoadTerritories()

            Catch ex As Exception
                MsgBox(ex.Message, MsgBoxStyle.Critical, "ERROR")
            End Try
        End If
    End Sub

    Private Sub DataGridView3_CellContentClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView3.CellContentClick
        TID.Text = DataGridView3.SelectedCells(0).Value
        ComboBox8.SelectedValue = DataGridView3.SelectedCells(1).Value
        Call LoadRegionList()
        ComboBox6.SelectedValue = DataGridView3.SelectedCells(2).Value
        ComboBox7.SelectedValue = DataGridView3.SelectedCells(3).Value
        txtTerritoryName.Text = DataGridView3.SelectedCells(4).Value
        If IsDBNull(DataGridView3.SelectedCells(8).Value) Or DataGridView3.SelectedCells(8).Value = 0 Then
            CheckBox1.Checked = False
        Else
            CheckBox1.Checked = DataGridView3.SelectedCells(8).Value
            DateTimePicker1.Text = DataGridView3.SelectedCells(9).Value
        End If

        CheckBox2.Checked = DataGridView3.SelectedCells(10).Value
    End Sub

    Private Sub Button11_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button11.Click
        Dim Mycn As OleDb.OleDbConnection
        Dim Command As OleDb.OleDbCommand
        Dim Command2 As OleDb.OleDbCommand
        Dim Command3 As OleDb.OleDbCommand
        Dim Command4 As OleDb.OleDbCommand
        Dim SQLstr As String

        Try
            Mycn = New OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & My.Settings.dbloc & "';Jet OLEDB:Database Password='" & My.Settings.dbpass & "';")
            Mycn.Open()

            SQLstr = "DELETE * FROM sys_TEMP_TeamLoad"

            Command2 = New OleDb.OleDbCommand(SQLstr, Mycn)
            Command2.ExecuteNonQuery()

            SQLstr = "INSERT INTO sys_TEMP_TeamLoad (DepartmentName, RegionName, TerritoryName, FullName, APXName, Runway, DateInTerritory, DaysLeft, Licensed)" & _
            "SELECT sys_Department.DepartmentName, sys_Region.RegionName, sys_Territory.TerritoryName, sys_Users.FullName, sys_Users.APXName, sys_Territory.Runway, sys_Territory.DateInTerritory, DateDiff('d', Now(), DateAdd('d',90, sys_Territory.DateInTerritory)), sys_Territory.Licensed" & _
            " FROM ((sys_Territory INNER JOIN sys_Region ON sys_Territory.RegionID = sys_Region.ID) INNER JOIN sys_Department ON sys_Region.DepartmentID = sys_Department.ID) INNER JOIN sys_Users ON sys_Territory.AssignedMember = sys_Users.ID" & _
            " WHERE(((sys_Territory.Active) = True))" & _
            " GROUP BY sys_Department.DepartmentName, sys_Region.RegionName, sys_Territory.TerritoryName, sys_Users.FullName, sys_Users.APXName, sys_Territory.Runway, sys_Territory.DateInTerritory, sys_Territory.Licensed"

            Command = New OleDb.OleDbCommand(SQLstr, Mycn)
            Command.ExecuteNonQuery()

            'sys_TEMP_APXTeamLoad

            SQLstr = "DELETE * FROM sys_TEMP_APXTeamLoad"

            Command3 = New OleDb.OleDbCommand(SQLstr, Mycn)
            Command3.ExecuteNonQuery()

            SQLstr = "INSERT INTO sys_TEMP_APXTeamLoad (APXName)" & _
            "SELECT Distinct AAMRepName FROM dbo_vQbRowDefPortfolio WHERE PortfolioStatus = 'Open'"

            Command4 = New OleDb.OleDbCommand(SQLstr, Mycn)
            Command4.ExecuteNonQuery()

            Mycn.Close()

            Call LoadNonProducing()
            Call LoadOpenTerr()
            Call LoadTeamStats()
            Call LoadRunway()

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "ERROR")
        End Try
    End Sub

    Public Sub LoadNonProducing()
        Try

            Dim conn As New OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & My.Settings.dbloc & "';Jet OLEDB:Database Password='" & My.Settings.dbpass & "';")

            Dim strSQL As String = "SELECT sys_TEMP_TeamLoad.DepartmentName As [Department], sys_TEMP_TeamLoad.RegionName As [Region], sys_TEMP_TeamLoad.TerritoryName As [Territory], sys_TEMP_TeamLoad.FullName As [Rep]" & _
            " FROM sys_TEMP_TeamLoad LEFT JOIN sys_TEMP_APXTeamLoad ON sys_TEMP_TeamLoad.APXName = sys_TEMP_APXTeamLoad.APXName" & _
            " WHERE(((sys_TEMP_APXTeamLoad.ID) Is Null) And ((sys_TEMP_TeamLoad.FullName) <> 'Open'))" & _
            " GROUP BY sys_TEMP_TeamLoad.DepartmentName, sys_TEMP_TeamLoad.RegionName, sys_TEMP_TeamLoad.TerritoryName, sys_TEMP_TeamLoad.FullName;"


            Dim da As New OleDb.OleDbDataAdapter(strSQL, conn)
            Dim ds As New DataSet
            da.Fill(ds, "Users")

            With DataGridView4
                .DataSource = ds.Tables("Users")
            End With

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "ERROR")
        End Try
    End Sub

    Public Sub LoadOpenTerr()
        Try

            Dim conn As New OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & My.Settings.dbloc & "';Jet OLEDB:Database Password='" & My.Settings.dbpass & "';")

            'Dim strSQL As String = "SELECT sys_TEMP_TeamLoad.DepartmentName As [Department], sys_TEMP_TeamLoad.RegionName As [Region], sys_TEMP_TeamLoad.TerritoryName As [Territory], sys_TEMP_TeamLoad.FullName As [Rep], sys_TEMP_TeamLoad.DaysLeft As [Days Left]" & _
            '" FROM sys_TEMP_TeamLoad" & _
            '" WHERE sys_TEMP_TeamLoad.Runway = True AND (sys_TEMP_TeamLoad.DaysLeft > 1)" & _
            '" GROUP BY sys_TEMP_TeamLoad.DepartmentName, sys_TEMP_TeamLoad.RegionName, sys_TEMP_TeamLoad.TerritoryName, sys_TEMP_TeamLoad.FullName, sys_TEMP_TeamLoad.DaysLeft;"

            Dim strSQL As String = "SELECT sys_TEMP_TeamLoad.DepartmentName As [Department], sys_TEMP_TeamLoad.RegionName As [Region], sys_TEMP_TeamLoad.TerritoryName As [Territory], sys_TEMP_TeamLoad.FullName As [Rep]" & _
           " FROM sys_TEMP_TeamLoad LEFT JOIN sys_TEMP_APXTeamLoad ON sys_TEMP_TeamLoad.APXName = sys_TEMP_APXTeamLoad.APXName" & _
           " WHERE(((sys_TEMP_APXTeamLoad.ID) Is Null) And ((sys_TEMP_TeamLoad.FullName) = 'Open'))" & _
           " GROUP BY sys_TEMP_TeamLoad.DepartmentName, sys_TEMP_TeamLoad.RegionName, sys_TEMP_TeamLoad.TerritoryName, sys_TEMP_TeamLoad.FullName;"


            Dim da As New OleDb.OleDbDataAdapter(strSQL, conn)
            Dim ds As New DataSet
            da.Fill(ds, "Users")

            With DataGridView5
                .DataSource = ds.Tables("Users")
            End With

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "ERROR")
        End Try
    End Sub

    Public Sub LoadRunway()
        Try

            Dim conn As New OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & My.Settings.dbloc & "';Jet OLEDB:Database Password='" & My.Settings.dbpass & "';")

           
            Dim strSQL As String = "SELECT sys_TEMP_TeamLoad.DepartmentName As [Department], sys_TEMP_TeamLoad.RegionName As [Region], sys_TEMP_TeamLoad.TerritoryName As [Territory], sys_TEMP_TeamLoad.FullName As [Rep], sys_TEMP_TeamLoad.DaysLeft As [Days Left]" & _
            " FROM sys_TEMP_TeamLoad" & _
            " WHERE sys_TEMP_TeamLoad.Runway = True AND (sys_TEMP_TeamLoad.DaysLeft > 1)" & _
            " GROUP BY sys_TEMP_TeamLoad.DepartmentName, sys_TEMP_TeamLoad.RegionName, sys_TEMP_TeamLoad.TerritoryName, sys_TEMP_TeamLoad.FullName, sys_TEMP_TeamLoad.DaysLeft;"


            Dim da As New OleDb.OleDbDataAdapter(strSQL, conn)
            Dim ds As New DataSet
            da.Fill(ds, "Users")

            With DataGridView7
                .DataSource = ds.Tables("Users")
            End With

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "ERROR")
        End Try
    End Sub

    Public Sub LoadTeamStats()
        'Try

        Dim conn As New OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & My.Settings.dbloc & "';Jet OLEDB:Database Password='" & My.Settings.dbpass & "';")

        Dim strSQL As String = "SELECT sys_Department.DepartmentName As [Department], sys_Region.RegionName As [Region], Count(sys_Users.FullName) AS Producers, (SELECT Count(sys_TEMP_TeamLoad.FullName) AS CountOfFullName2" & _
        " FROM sys_TEMP_TeamLoad LEFT JOIN sys_TEMP_APXTeamLoad ON sys_TEMP_TeamLoad.APXName = sys_TEMP_APXTeamLoad.APXName" & _
        " WHERE(((sys_TEMP_APXTeamLoad.ID) Is Null) AND ((sys_TEMP_TeamLoad.DaysLeft <= 0) OR (sys_TEMP_TeamLoad.DaysLeft Is Null)) And ((sys_TEMP_TeamLoad.FullName) <> 'Open') And DepartmentName = sys_Department.DepartmentName And RegionName = sys_Region.RegionName)" & _
        " GROUP BY sys_TEMP_TeamLoad.DepartmentName, sys_TEMP_TeamLoad.RegionName;) As [Non-Participants], ((Producers - IIF(IsNull([Non-Participants]),0,[Non-Participants])) / Producers) As [Participation %]" & _
        " FROM (sys_Territory INNER JOIN (sys_Region INNER JOIN sys_Department ON sys_Region.DepartmentID = sys_Department.ID) ON sys_Territory.RegionID = sys_Region.ID) INNER JOIN sys_Users ON sys_Territory.AssignedMember = sys_Users.ID" & _
        " WHERE sys_Users.FullName <> 'open' AND sys_Territory.Active = -1" & _
        " GROUP BY sys_Department.DepartmentName, sys_Region.RegionName;"


        Dim da As New OleDb.OleDbDataAdapter(strSQL, conn)
        Dim ds As New DataSet
        da.Fill(ds, "Users")

        With DataGridView6
            .DataSource = ds.Tables("Users")
            .Columns(4).DefaultCellStyle.Format = "p"
        End With

        'Catch ex As Exception
        'MsgBox(ex.Message, MsgBoxStyle.Critical, "ERROR")
        'End Try
    End Sub

    Private Sub Button12_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button12.Click
        Dim ir As MsgBoxResult
        ir = MsgBox("Are you sure you want to delete this record?", MsgBoxStyle.YesNo + MsgBoxStyle.Critical, "Confirm Delete")
        If ir = MsgBoxResult.Yes Then
            Dim Mycn As OleDb.OleDbConnection
            Dim Command As OleDb.OleDbCommand
            Dim SQLstr As String

            Try
                Mycn = New OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & My.Settings.dbloc & "';Jet OLEDB:Database Password='" & My.Settings.dbpass & "';")
                Mycn.Open()

                SQLstr = "UPDATE sys_Territory SET Active = False WHERE ID = " & TID.Text

                Command = New OleDb.OleDbCommand(SQLstr, Mycn)
                Command.ExecuteNonQuery()

            Catch ex As Exception
                MsgBox(ex.Message, MsgBoxStyle.Critical, "ERROR")
            End Try

            TID.Text = "NEW"

            Call LoadTerritories()
        Else

        End If
    End Sub

    Private Sub CheckBox1_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckBox1.CheckedChanged
        If CheckBox1.Checked Then
            DateTimePicker1.Enabled = True
        Else
            DateTimePicker1.Enabled = False
        End If
    End Sub
End Class