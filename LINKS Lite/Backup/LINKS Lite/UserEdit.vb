Public Class UserEdit
    Dim thread1 As System.Threading.Thread

    Public Sub CurrentUser()
        Dim Mycn As OleDb.OleDbConnection

        Try

            Mycn = New OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & My.Settings.roledbloc & "';Jet OLEDB:Database Password='" & My.Settings.dbpass & "';")

            Mycn.Open()

            Dim sqlstring As String

            sqlstring = "SELECT * FROM sys_Users WHERE ID = " & ID.Text & ";"

            Dim queryString As String = String.Format(sqlstring)
            Dim cmd As New OleDb.OleDbCommand(queryString, Mycn)
            Dim da As New OleDb.OleDbDataAdapter(cmd)
            Dim ds As New DataSet

            da.Fill(ds, "User")
            Dim dt As DataTable = ds.Tables("User")
            If dt.Rows.Count > 0 Then

                Dim row As DataRow = dt.Rows(0)

                Active.Checked = (row("Active"))
                WI.Checked = (row("WI"))
                If IsDBNull(row("FullName")) Then
                    FullName.Text = "*UNKNOWN*"
                Else
                    FullName.Text = (row("FullName"))
                End If

                If IsDBNull(row("NetworkID")) Then
                    NetworkID.Text = ""
                Else
                    NetworkID.Text = (row("NetworkID"))
                End If

                If IsDBNull(row("Password")) Then
                    Password.Text = ""
                Else
                    Password.Text = (row("Password"))
                End If

                If IsDBNull(row("Team")) Then
                    Team.Text = ""
                Else
                    Team.Text = (row("Team"))
                End If

                If IsDBNull(row("APXName")) Then
                    APXName.Text = ""
                Else
                    APXName.Text = (row("APXName"))
                End If

                If IsDBNull(row("APXID")) Then
                    APXID.Text = ""
                Else
                    APXID.Text = (row("APXID"))
                End If

                ComboBox1.SelectedValue = (row("RoleID"))

            Else

            End If
            Mycn.Close()
            Mycn.Dispose()

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "ERROR")
            'Application.Exit()
        End Try
    End Sub

    Public Sub NewUser()
        Try

            Dim conn As New OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & My.Settings.roledbloc & "';Jet OLEDB:Database Password='" & My.Settings.dbpass & "';")

            Dim strSQL As String = "SELECT * FROM sys_Roles"
            Dim da As New OleDb.OleDbDataAdapter(strSQL, conn)
            Dim ds As New DataSet
            da.Fill(ds, "Users")

            With ComboBox1
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

    Private Sub UserEdit_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Call NewUser()
    End Sub

    Private Sub cmdCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdCancel.Click
        Me.Close()
    End Sub

    Private Sub cmdSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdSave.Click
        If ID.Text = "NEW" Then
            'Value Check
            If IsDBNull(FullName) Or FullName.Text = "" Then
                FullName.BackColor = Color.Red
                FullName.ForeColor = Color.White
                GoTo Line1
            Else
                FullName.BackColor = Color.White
                FullName.ForeColor = Color.Black
            End If

            If IsDBNull(NetworkID) Or NetworkID.Text = "" Then
                NetworkID.BackColor = Color.Red
                NetworkID.ForeColor = Color.White
                GoTo Line1
            Else
                NetworkID.BackColor = Color.White
                NetworkID.ForeColor = Color.Black
            End If

            If ((IsDBNull(Password) Or Password.Text = "") And (WI.Checked = "FALSE")) Then
                Password.BackColor = Color.Red
                Password.ForeColor = Color.White
                GoTo Line1
            Else
                Password.BackColor = Color.White
                Password.ForeColor = Color.Black
            End If

            If IsDBNull(ComboBox1.SelectedValue) Then
                ComboBox1.BackColor = Color.Red
                ComboBox1.ForeColor = Color.White
            Else
                ComboBox1.BackColor = Color.White
                ComboBox1.ForeColor = Color.Black
            End If

            'Insert new record
            cmdSave.Text = "SAVING..."
            cmdSave.Enabled = False
            Dim Mycn As OleDb.OleDbConnection

            Try

                Mycn = New OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & My.Settings.roledbloc & "';Jet OLEDB:Database Password='" & My.Settings.dbpass & "';")

                Mycn.Open()

                Dim sqlstring As String

                sqlstring = "SELECT * FROM sys_Users WHERE [NetworkID] = '" & NetworkID.Text & "'"


                Dim queryString As String = String.Format(sqlstring)
                Dim cmd As New OleDb.OleDbCommand(queryString, Mycn)
                Dim da As New OleDb.OleDbDataAdapter(cmd)
                Dim ds As New DataSet

                da.Fill(ds, "User")
                Dim dt As DataTable = ds.Tables("User")
                If dt.Rows.Count = 0 Then
                    Control.CheckForIllegalCrossThreadCalls = False
                    thread1 = New System.Threading.Thread(AddressOf InsertRec)
                    thread1.Start()
                Else
                    NetworkID.BackColor = Color.Red
                    NetworkID.ForeColor = Color.White
                    MsgBox("A profile already exsists with that Network ID.  You cannot save a duplicate record.", MsgBoxStyle.Information, "DUPLICATE RECORD DETECTED")
                    cmdSave.Enabled = True
                    cmdSave.Text = "Save"
                    GoTo Line1
                End If

            Catch ex As Exception
                MsgBox(ex.Message)
            End Try
        Else
            'Update record
            cmdSave.Text = "SAVING..."
            cmdSave.Enabled = False
            Control.CheckForIllegalCrossThreadCalls = False
            thread1 = New System.Threading.Thread(AddressOf UpdateRec)
            thread1.Start()
        End If

Line1:


    End Sub

    Public Sub InsertRec()
        Dim Mycn As OleDb.OleDbConnection
        Dim Command As OleDb.OleDbCommand
        Dim SQLstr As String

        Try
            Mycn = New OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & My.Settings.roledbloc & "';Jet OLEDB:Database Password='" & My.Settings.dbpass & "';")

            Dim ds1 As New DataSet
            Dim eds1 As New DataGridView
            Dim dv1 As New DataView

            Mycn.Open()

            Dim pwd As String
            If IsDBNull(Password) Or Password.Text = "" Then
                pwd = ""
            Else
                pwd = Password.Text
            End If

            Dim aid As Integer
            If IsDBNull(APXID.Text) Or APXID.Text = "" Then
                aid = "0"
            Else
                aid = APXID.Text
            End If

            Dim anme As String
            If IsDBNull(APXName) Or APXName.Text = "" Then
                anme = ""
            Else
                anme = APXName.Text
            End If

            Dim tme As String
            If IsDBNull(Team) Or Team.Text = "" Then
                tme = ""
            Else
                tme = Team.Text
            End If

            SQLstr = "INSERT INTO sys_Users([FullName], [NetworkID], [Password], [WI], [Active], [APXID], [APXName], [Team], [RoleID])" & _
            "VALUES('" & FullName.Text & "','" & NetworkID.Text & "','" & pwd & "'," & WI.Checked & "," & Active.Checked & ",'" & aid & "','" & anme & "','" & tme & "'," & ComboBox1.SelectedValue & ")"

            Command = New OleDb.OleDbCommand(SQLstr, Mycn)
            Command.ExecuteNonQuery()

            Mycn.Close()

            Me.Close()

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "ERROR")

        End Try
    End Sub

    Public Sub UpdateRec()
        Dim Mycn As OleDb.OleDbConnection
        Dim Command As OleDb.OleDbCommand
        Dim SQLstr As String

        Try
            Mycn = New OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & My.Settings.roledbloc & "';Jet OLEDB:Database Password='" & My.Settings.dbpass & "';")

            Dim ds1 As New DataSet
            Dim eds1 As New DataGridView
            Dim dv1 As New DataView

            Mycn.Open()

            Dim pwd As String
            If IsDBNull(Password) Or Password.Text = "" Then
                pwd = ""
            Else
                pwd = Password.Text
            End If

            Dim aid As Integer
            If IsDBNull(APXID.Text) Or APXID.Text = "" Then
                aid = "0"
            Else
                aid = APXID.Text
            End If

            Dim anme As String
            If IsDBNull(APXName) Or APXName.Text = "" Then
                anme = ""
            Else
                anme = APXName.Text
            End If

            Dim tme As String
            If IsDBNull(Team) Or Team.Text = "" Then
                tme = ""
            Else
                tme = Team.Text
            End If

            SQLstr = "Update sys_Users SET [FullName] = '" & FullName.Text & "', [NetworkID] = '" & NetworkID.Text & "', [Password] = '" & pwd & "', [WI] = " & WI.Checked & ", [Active] = " & Active.Checked & ", [APXID] = '" & aid & "', [APXName] = '" & anme & "', [Team] = '" & tme & "', [RoleID] = " & ComboBox1.SelectedValue & " WHERE ID = " & ID.Text

            Command = New OleDb.OleDbCommand(SQLstr, Mycn)
            Command.ExecuteNonQuery()

            Mycn.Close()

            Me.Close()

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "ERROR")

        End Try
    End Sub

End Class