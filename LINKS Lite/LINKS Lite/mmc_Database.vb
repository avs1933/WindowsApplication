Public Class mmc_Database
    Dim thDatabases As System.Threading.Thread
    Dim thInsert As System.Threading.Thread

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        MsgBox("A database must first be loaded into Advent and 'Status' set to Database.", MsgBoxStyle.Information, "Help")

    End Sub

    Private Sub mmc_Database_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Control.CheckForIllegalCrossThreadCalls = False
        thDatabases = New System.Threading.Thread(AddressOf LoadDatabases)
        thDatabases.Start()
        'Call LoadDatabases()
    End Sub

    Public Sub LoadDatabases()
        Try

            Dim conn As New OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & My.Settings.dbloc & "';Jet OLEDB:Database Password='" & My.Settings.dbpass & "';")

            Dim strSQL As String = "SELECT ContactID, ContactName FROM dbo_vQbRowDefContact WHERE Status = 'Database'"
            Dim da As New OleDb.OleDbDataAdapter(strSQL, conn)
            Dim ds As New DataSet
            da.Fill(ds, "Users")

            With cboDatabase
                .DataSource = ds.Tables("Users")
                .DisplayMember = "ContactName"
                .ValueMember = "ContactID"
                .SelectedIndex = 0
            End With

            Label16.Visible = False
            cboDatabase.Enabled = True
            'Call LoadDBData()

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "ERROR")
        End Try
    End Sub

    Public Sub LoadDBData()

        If IsDBNull(txtContactID.Text) Then

        Else
            Try
                Dim conn As New OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & My.Settings.dbloc & "';Jet OLEDB:Database Password='" & My.Settings.dbpass & "';")

                Dim strSQL As String = "SELECT * FROM dbo_vQbRowDefContact WHERE ContactID = " & txtContactID.Text
                Dim queryString As String = String.Format(strSQL)
                Dim cmd As New OleDb.OleDbCommand(queryString, conn)
                Dim da As New OleDb.OleDbDataAdapter(cmd)
                Dim ds As New DataSet

                da.Fill(ds, "User")
                Dim dt As DataTable = ds.Tables("User")

                Dim row1 As DataRow = dt.Rows(0)


                If IsDBNull(row1("DefaultAddress")) Then
                    rtbAddress.Text = "NULL"
                Else
                    rtbAddress.Text = (row1("DefaultAddress"))
                End If

                If IsDBNull(row1("Notes")) Then
                    rtbAPXNotes.Text = "NULL"
                Else
                    rtbAPXNotes.Text = (row1("Notes"))
                End If

                If IsDBNull(row1("Email")) Then
                    txtEmail.Text = "NULL"
                Else
                    txtEmail.Text = (row1("Email"))
                End If

                If IsDBNull(row1("BusinessPhone")) Then
                    txtPhone.Text = ""
                Else
                    txtPhone.Text = (row1("BusinessPhone"))
                End If

                If IsDBNull(row1("URL")) Then

                Else
                    llblURL.Text = (row1("URL"))
                    llblURL.Visible = True
                End If

                GroupBox1.Visible = True
                GroupBox2.Visible = True
                txtURL.Visible = True
                txtRptDlne.Visible = True
                cmdGo.Visible = True
                cmdAPX.Visible = True
                cmdSave.Visible = True

                cmdLoad.BackColor = Color.White
                cmdLoad.ForeColor = Color.Black

            Catch ex As Exception
                MsgBox(ex.Message, MsgBoxStyle.Critical, "ERROR")
            End Try

            End If
    End Sub

    Private Sub cboDatabase_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cboDatabase.SelectedIndexChanged

    End Sub

    Private Sub cboDatabase_SelectedValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cboDatabase.SelectedValueChanged
        cmdLoad.BackColor = Color.Red
        cmdLoad.ForeColor = Color.White
    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        txtContactID.Text = cboDatabase.SelectedValue
        Call LoadDBData()
    End Sub

    Private Sub llblURL_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles llblURL.LinkClicked
        Dim webAddress As String = llblURL.Text
        Process.Start(webAddress)
    End Sub

    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        MsgBox("A database must first be loaded into Advent and 'Status' set to Database.", MsgBoxStyle.Information, "Help")

    End Sub

    Private Sub Button1_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdLoad.Click
        If IsDBNull(cboDatabase.SelectedValue) Then
            MsgBox("You must SELECT a database", MsgBoxStyle.Information, "Error")
        Else
            txtContactID.Text = cboDatabase.SelectedValue
            Call LoadDBData()
        End If


    End Sub

    Private Sub cmdSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdSave.Click
        If txtID.Text = "NEW" Then
            Dim Mycn As OleDb.OleDbConnection

            Try

                Mycn = New OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & My.Settings.roledbloc & "';Jet OLEDB:Database Password='" & My.Settings.dbpass & "';")

                Mycn.Open()

                Dim sqlstring As String

                sqlstring = "SELECT * FROM mmc_Databases WHERE ContactID = " & cboDatabase.SelectedValue

                Dim queryString As String = String.Format(sqlstring)
                Dim cmd As New OleDb.OleDbCommand(queryString, Mycn)
                Dim da As New OleDb.OleDbDataAdapter(cmd)
                Dim ds As New DataSet

                da.Fill(ds, "User")
                Dim dt As DataTable = ds.Tables("User")
                If dt.Rows.Count = 0 Then
                    Control.CheckForIllegalCrossThreadCalls = False
                    thInsert = New System.Threading.Thread(AddressOf InsertRec)
                    thInsert.Start()
                Else
                    cboDatabase.BackColor = Color.Red
                    cboDatabase.ForeColor = Color.White
                    MsgBox("A record already exsists with that Database.  You cannot save a duplicate record.", MsgBoxStyle.Information, "DUPLICATE RECORD DETECTED")
                    cmdSave.Enabled = True
                    cmdSave.Text = "Save"
                    GoTo Line1
                End If

            Catch ex As Exception
                MsgBox(ex.Message)
            End Try
        Else

        End If
line1:

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

            Dim url As String
            If IsDBNull(txtURL.Text) Or txtURL.Text = "" Then
                url = ""
            Else
                url = txtURL.Text
            End If

            Dim rdlne As Integer
            If IsDBNull(txtRptDlne.Text) Or txtRptDlne.Text = "" Then
                rdlne = "0"
            Else
                rdlne = txtRptDlne.Text
            End If

            Dim unme As String
            If IsDBNull(txtUN.Text) Or txtUN.Text = "" Then
                unme = ""
            Else
                unme = txtUN.Text
            End If

            Dim pswd As String
            If IsDBNull(txtPW.Text) Or txtPW.Text = "" Then
                pswd = ""
            Else
                pswd = txtPW.Text
            End If

            Dim nts As String
            'If IsDBNull(rtbRepNotes.Text) Or rtbRepNotes.Text = "" Then
            'nts = ""
            'Else
            'nts = rtbRepNotes.Text
            'End If

            'Dim question As String
            nts = rtbRepNotes.Text
            nts = Replace(nts, "'", "''")

            'Dim answer As String
            'answer = rtbAnswer.Text
            'answer = Replace(answer, "'", "''")

            SQLstr = "INSERT INTO mmc_Databases([ContactID], [URL], [ReportingDeadline], [ReportingUN], [ReportingPW], [StartDate], [Notes], [IsActive])" & _
            "VALUES(" & cboDatabase.SelectedValue & ",'" & url & "'," & rdlne & ",'" & unme & "','" & pswd & "',#" & dtRepSnce.Text & "#,'" & nts & "',-1)"

            Command = New OleDb.OleDbCommand(SQLstr, Mycn)
            Command.ExecuteNonQuery()

            Mycn.Close()

            Me.Close()

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "ERROR")

        End Try
    End Sub

    Private Sub cmdGo_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdGo.Click
        If IsDBNull(txtURL.Text) Or txtURL.Text = "" Then

        Else
            Dim webAddress As String = txtURL.Text
            Process.Start(webAddress)
        End If
    End Sub

    Private Sub cmdAPX_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdAPX.Click
        Dim webAddress As String = "http://adventapps/APXLogin/ContactDetail.aspx?linkfield=" & cboDatabase.SelectedValue
        Process.Start(webAddress)
    End Sub

    Private Sub cmdCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdCancel.Click
        Me.Close()
    End Sub
End Class