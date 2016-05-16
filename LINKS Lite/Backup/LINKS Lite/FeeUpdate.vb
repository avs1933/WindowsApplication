Public Class FeeUpdate
    Dim thread1 As System.Threading.Thread
    Dim thread2 As System.Threading.Thread
    Dim thread3 As System.Threading.Thread

    Private Sub FeeUpdate_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        If Permissions.SetDefaultFee.Checked Then
            CheckBox1.Enabled = True
        Else
            CheckBox1.Enabled = False
        End If

        Control.CheckForIllegalCrossThreadCalls = False
        Call UpdateFirms()

        If ID.Text = "NEW" Then
            'do nothing
        Else
            thread1 = New System.Threading.Thread(AddressOf CurrentFee)
            thread1.Start()
        End If

    End Sub

    Public Sub UpdateRec()
        Dim Mycn As OleDb.OleDbConnection
        Dim Command As OleDb.OleDbCommand
        Dim SQLstr As String

        Try
            Mycn = New OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & My.Settings.dbloc & "';Jet OLEDB:Database Password='" & My.Settings.dbpass & "';")

            Dim ds1 As New DataSet
            Dim eds1 As New DataGridView
            Dim dv1 As New DataView

            Mycn.Open()

            Dim dis As String
            If IsDBNull(Discipline) Or Discipline.Text = "" Then
                dis = ""
            Else
                dis = Discipline.Text
            End If

            Dim f1 As Integer
            If IsDBNull(Fee1) Or Fee1.Text = "" Then
                f1 = ""
            Else
                f1 = Fee1.Text
            End If

            Dim f2 As String
            If IsDBNull(Fee2) Or Fee2.Text = "" Then
                f2 = ""
            Else
                f2 = Fee2.Text
            End If

            SQLstr = "Update env_Fees SET [CompanyID] = " & Company.SelectedValue & ", [DisciplineNme] = '" & dis & "', [Fee1] = " & f1 & ", [Fee2] = " & f2 & ", [Default] = " & CheckBox1.Checked & ", [SubAdvised] = " & CheckBox2.Checked & ", [Advised] = " & CheckBox3.Checked & " WHERE ID = " & ID.Text

            Command = New OleDb.OleDbCommand(SQLstr, Mycn)
            Command.ExecuteNonQuery()

            Mycn.Close()

            Me.Close()

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "ERROR")

        End Try
    End Sub

    Public Sub UpdateRecNew()
        Dim Mycn As OleDb.OleDbConnection
        Dim Command As OleDb.OleDbCommand
        Dim SQLstr As String

        Try
            Mycn = New OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & My.Settings.dbloc & "';Jet OLEDB:Database Password='" & My.Settings.dbpass & "';")

            Dim ds1 As New DataSet
            Dim eds1 As New DataGridView
            Dim dv1 As New DataView

            Mycn.Open()

            Dim dis As String
            If IsDBNull(Discipline) Or Discipline.Text = "" Then
                dis = ""
            Else
                dis = Discipline.Text
            End If

            Dim f1 As Integer
            If IsDBNull(Fee1) Or Fee1.Text = "" Then
                f1 = ""
            Else
                f1 = Fee1.Text
            End If

            Dim f2 As String
            If IsDBNull(Fee2) Or Fee2.Text = "" Then
                f2 = ""
            Else
                f2 = Fee2.Text
            End If

            SQLstr = "Update env_Fees SET [CompanyID] = " & Company.SelectedValue & ", [DisciplineNme] = '" & dis & "', [Fee1] = " & f1 & ", [Fee2] = " & f2 & ", [Default] = " & CheckBox1.Checked & ", [SubAdvised] = " & CheckBox2.Checked & ", [Advised] = " & CheckBox3.Checked & " WHERE ID = " & ID.Text

            Command = New OleDb.OleDbCommand(SQLstr, Mycn)
            Command.ExecuteNonQuery()

            Mycn.Close()

            Call ResetFields()

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "ERROR")

        End Try
    End Sub

    Public Sub InsertRecNew()
        Dim Mycn As OleDb.OleDbConnection
        Dim Command As OleDb.OleDbCommand
        Dim SQLstr As String

        Try
            Mycn = New OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & My.Settings.dbloc & "';Jet OLEDB:Database Password='" & My.Settings.dbpass & "';")

            Dim ds1 As New DataSet
            Dim eds1 As New DataGridView
            Dim dv1 As New DataView

            Mycn.Open()

            Dim dis As String
            If IsDBNull(Discipline) Or Discipline.Text = "" Then
                dis = ""
            Else
                dis = Discipline.Text
            End If

            Dim f1 As Integer
            If IsDBNull(Fee1) Or Fee1.Text = "" Then
                f1 = ""
            Else
                f1 = Fee1.Text
            End If

            Dim f2 As String
            If IsDBNull(Fee2) Or Fee2.Text = "" Then
                f2 = ""
            Else
                f2 = Fee2.Text
            End If

            SQLstr = "INSERT INTO env_Fees ([CompanyID], [DisciplineNme], [Fee1], [Fee2], [Default], [Active], [SubAdvised], [Advised])" & _
            "VALUES(" & Company.SelectedValue & ",'" & dis & "'," & f1 & "," & f2 & "," & CheckBox1.Checked & ", -1, " & CheckBox2.Checked & "," & CheckBox3.Checked & ")"

            Command = New OleDb.OleDbCommand(SQLstr, Mycn)
            Command.ExecuteNonQuery()

            Mycn.Close()

            Call ResetFields()

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "ERROR")

        End Try
    End Sub

    Public Sub ResetFields()
        ID.Text = "NEW"
        'Company.SelectedValue = ""
        Discipline.Text = ""
        Fee1.Text = ""
        Fee2.Text = ""
        cmdSave.Enabled = True
        cmdSave.Text = "Save"
        cmdSaveNew.Enabled = True
        cmdSaveNew.Text = "Save and New"
    End Sub

    Public Sub InsertRec()
        Dim Mycn As OleDb.OleDbConnection
        Dim Command As OleDb.OleDbCommand
        Dim SQLstr As String

        Try
            Mycn = New OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & My.Settings.dbloc & "';Jet OLEDB:Database Password='" & My.Settings.dbpass & "';")

            Dim ds1 As New DataSet
            Dim eds1 As New DataGridView
            Dim dv1 As New DataView

            Mycn.Open()

            Dim dis As String
            If IsDBNull(Discipline) Or Discipline.Text = "" Then
                dis = ""
            Else
                dis = Discipline.Text
            End If

            Dim f1 As Integer
            If IsDBNull(Fee1) Or Fee1.Text = "" Then
                f1 = ""
            Else
                f1 = Fee1.Text
            End If

            Dim f2 As String
            If IsDBNull(Fee2) Or Fee2.Text = "" Then
                f2 = ""
            Else
                f2 = Fee2.Text
            End If

            SQLstr = "INSERT INTO env_Fees ([CompanyID], [DisciplineNme], [Fee1], [Fee2], [Default], [Active], [SubAdvised], [Advised])" & _
            "VALUES(" & Company.SelectedValue & ",'" & dis & "'," & f1 & "," & f2 & "," & CheckBox1.Checked & ", -1, " & CheckBox2.Checked & "," & CheckBox3.Checked & ")"

            Command = New OleDb.OleDbCommand(SQLstr, Mycn)
            Command.ExecuteNonQuery()

            Mycn.Close()

            Me.Close()

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "ERROR")

        End Try
    End Sub

    Public Sub CurrentFee()
        Dim Mycn As OleDb.OleDbConnection

        Try

            Mycn = New OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & My.Settings.dbloc & "';Jet OLEDB:Database Password='" & My.Settings.dbpass & "';")

            Mycn.Open()

            Dim sqlstring As String

            sqlstring = "SELECT * FROM env_Fees WHERE ID = " & ID.Text & ";"

            Dim queryString As String = String.Format(sqlstring)
            Dim cmd As New OleDb.OleDbCommand(queryString, Mycn)
            Dim da As New OleDb.OleDbDataAdapter(cmd)
            Dim ds As New DataSet

            da.Fill(ds, "User")
            Dim dt As DataTable = ds.Tables("User")
            If dt.Rows.Count > 0 Then

                Dim row As DataRow = dt.Rows(0)

                CheckBox1.Checked = (row("Default"))
                CheckBox2.Checked = (row("SubAdvised"))
                CheckBox3.Checked = (row("Advised"))

                If (row("Active")) = False Then
                    Me.Text = "***DELETED***"
                    cmdSave.Enabled = False
                    cmdSaveNew.Enabled = False
                    Button2.Enabled = False
                Else
                End If

                If IsDBNull(row("DisciplineNme")) Then
                    Discipline.Text = "*UNKNOWN*"
                Else
                    Discipline.Text = (row("DisciplineNme"))
                End If

                If IsDBNull(row("Fee1")) Then
                    Fee1.Text = ""
                Else
                    Fee1.Text = (row("Fee1"))
                End If

                If IsDBNull(row("Fee2")) Then
                    Fee2.Text = ""
                Else
                    Fee2.Text = (row("Fee2"))
                End If

                Company.SelectedValue = (row("CompanyID"))

            Else

            End If
            Mycn.Close()
            Mycn.Dispose()

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "ERROR")
            'Application.Exit()
        End Try
    End Sub

    Public Sub UpdateFirms()
        Try

            Dim conn As New OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & My.Settings.dbloc & "';Jet OLEDB:Database Password='" & My.Settings.dbpass & "';")

            Dim strSQL As String = "SELECT * FROM env_Firms"
            Dim da As New OleDb.OleDbDataAdapter(strSQL, conn)
            Dim ds As New DataSet
            da.Fill(ds, "Users")

            With Company
                .DataSource = ds.Tables("Users")
                .DisplayMember = "Company Name"
                .ValueMember = "ContactID"
                .SelectedIndex = 0
            End With

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "ERROR")
            'Application.Exit()
        End Try
    End Sub

    Private Sub cmdSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdSave.Click

        If ((Permissions.SetDefaultFee.Checked = False) And (CheckBox1.Checked = True)) Then
            MsgBox("You do not have permission to set a default fee schedule.", MsgBoxStyle.Information, "Permissions Issue")
            GoTo Line1
        Else

        End If

        If IsDBNull(Discipline) Or Discipline.Text = "" Then
            Discipline.BackColor = Color.Red
            Discipline.ForeColor = Color.White
            GoTo Line1
        Else
            Discipline.BackColor = Color.White
            Discipline.ForeColor = Color.Black
        End If

        If IsDBNull(Fee1) Or Fee1.Text = "" Then
            Fee1.BackColor = Color.Red
            Fee1.ForeColor = Color.White
            GoTo Line1
        Else
            Fee1.BackColor = Color.White
            Fee1.ForeColor = Color.Black
        End If

        If CheckBox2.Checked And CheckBox3.Checked Then
            MsgBox("You must select only one fee type.", MsgBoxStyle.Information, "Cannot Save")
            GoTo Line1
        Else
            If CheckBox2.Checked = False And CheckBox3.Checked = False Then
                MsgBox("You must select at least one fee type.", MsgBoxStyle.Information, "Cannot Save")
            Else
                'do nothing
            End If
        End If

        cmdSave.Text = "SAVING..."
        cmdSave.Enabled = False
        cmdSaveNew.Enabled = False
        Control.CheckForIllegalCrossThreadCalls = False
        If ID.Text = "NEW" Then
            thread2 = New System.Threading.Thread(AddressOf InsertRec)
            thread2.Start()
        Else
            thread2 = New System.Threading.Thread(AddressOf UpdateRec)
            thread2.Start()
        End If

Line1:
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        If ((Permissions.SetDefaultFee.Checked = False) And (CheckBox1.Checked = True)) Then
            MsgBox("You do not have permission to set a default fee schedule.", MsgBoxStyle.Information, "Permissions Issue")
            GoTo Line1
        Else

        End If

        If ID.Text = "NEW" Then
            MsgBox("This record has not been saved.  Nothing to Delete.", MsgBoxStyle.Information, "Cant Delete")
        Else
            Dim str As Integer

            str = MsgBox("You are about to Delete this fee schedule.  Are you sure you want to continue?", MsgBoxStyle.YesNo + MsgBoxStyle.Exclamation, "Confirm Delete")

            If str = MsgBoxResult.Yes Then
                Dim Mycn As OleDb.OleDbConnection
                Dim Command As OleDb.OleDbCommand
                Dim SQLstr As String

                Try
                    Mycn = New OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & My.Settings.dbloc & "';Jet OLEDB:Database Password='" & My.Settings.dbpass & "';")

                    Dim ds1 As New DataSet
                    Dim eds1 As New DataGridView
                    Dim dv1 As New DataView

                    Mycn.Open()

                    SQLstr = "Update env_Fees SET [Active] = NULL WHERE ID = " & ID.Text

                    Command = New OleDb.OleDbCommand(SQLstr, Mycn)
                    Command.ExecuteNonQuery()

                    Mycn.Close()

                    Me.Close()

                Catch ex As Exception
                    MsgBox(ex.Message, MsgBoxStyle.Critical, "ERROR")

                End Try
            Else
                'do nothing
            End If
        End If

line1:

    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Me.Close()
    End Sub

    Private Sub cmdSaveNew_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdSaveNew.Click

        If ((Permissions.SetDefaultFee.Checked = False) And (CheckBox1.Checked = True)) Then
            MsgBox("You do not have permission to set a default fee schedule.", MsgBoxStyle.Information, "Permissions Issue")
            GoTo Line1
        Else

        End If

        If IsDBNull(Discipline) Or Discipline.Text = "" Then
            Discipline.BackColor = Color.Red
            Discipline.ForeColor = Color.White
            GoTo Line1
        Else
            Discipline.BackColor = Color.White
            Discipline.ForeColor = Color.Black
        End If

        If IsDBNull(Fee1) Or Fee1.Text = "" Then
            Fee1.BackColor = Color.Red
            Fee1.ForeColor = Color.White
            GoTo Line1
        Else
            Fee1.BackColor = Color.White
            Fee1.ForeColor = Color.Black
        End If

        cmdSaveNew.Text = "SAVING..."
        cmdSave.Enabled = False
        cmdSaveNew.Enabled = False
        Control.CheckForIllegalCrossThreadCalls = False
        If ID.Text = "NEW" Then
            thread2 = New System.Threading.Thread(AddressOf InsertRecNew)
            thread2.Start()
        Else
            thread2 = New System.Threading.Thread(AddressOf UpdateRecNew)
            thread2.Start()
        End If

Line1:
    End Sub

    Private Sub CheckBox1_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckBox1.CheckedChanged
        If CheckBox1.Checked Then
            Company.SelectedValue = 8915
        Else

        End If
    End Sub
End Class