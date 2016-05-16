Public Class Login
    Dim thread1 As System.Threading.Thread

    Private Sub OK_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles OK.Click
        Call Checkdata()

    End Sub

    Public Sub Checkdata()
        Dim Mycn As OleDb.OleDbConnection

        Try

            Mycn = New OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & My.Settings.roledbloc & "';Jet OLEDB:Database Password='" & My.Settings.dbpass & "';")

            Mycn.Open()

            Dim sqlstring As String

            If CheckBox1.Checked Then
                sqlstring = "SELECT * FROM sys_Users WHERE NetworkID = '" & UsernameTextBox.Text & "' AND WI = -1 AND Active = -1;"
            Else
                sqlstring = "SELECT * FROM sys_Users WHERE NetworkID = '" & UsernameTextBox.Text & "' AND Password = '" & PasswordTextBox.Text & "' AND (WI = 0 OR WI Is Null) AND Active = -1;"
            End If

            Dim queryString As String = String.Format(sqlstring)
            Dim cmd As New OleDb.OleDbCommand(queryString, Mycn)
            Dim da As New OleDb.OleDbDataAdapter(cmd)
            Dim ds As New DataSet

            da.Fill(ds, "User")
            Dim dt As DataTable = ds.Tables("User")
            If dt.Rows.Count > 0 Then
                Home.Show()
                Home.BackgroundImage = My.Resources.AAM_logo_Final_Small_BW
                'Home.BackgroundImage = Image.FromFile(Application.StartupPath & "\resources\Background.gif")
                'Home.ToolStripLabel1.Image = Image.FromFile(Application.StartupPath & "\resources\Background.gif")
                For Each ctl As Control In Home.Controls
                    If TypeOf ctl Is MdiClient Then
                        If Environ("USERNAME") = "fcannon" Then
                            ctl.BackColor = Color.DarkSalmon
                        Else
                            If Environ("USERNAME") = "khyde" Then
                                ctl.BackColor = Color.Firebrick
                            Else
                                If My.Settings.dbloc <> "\\monumentco1\data\ToolboxDBs\SalesInterface2.accdb" Then
                                    ctl.BackColor = Color.Red
                                Else
                                    ctl.BackColor = Color.Black
                                End If

                            End If
                    End If
                    Exit For

                        End If
                Next ctl

                Dim row As DataRow = dt.Rows(0)
                Home.Enabled = False

                My.Settings.userid = (row("ID"))
                My.Settings.origuser = (row("ID"))
                'My.Settings.backcolor = (row("BackColor"))



                'NOTE: Team now loads from permissions form
                Permissions.Show()

                Permissions.ProgressBar1.Value = 25

                Call Permissions.Populate()
            Else
                MsgBox("Invalid Login", MsgBoxStyle.Critical, "Permissions Issue")
            End If
            Mycn.Close()
            Mycn.Dispose()

        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub

    Private Sub Cancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Cancel.Click
        End
    End Sub

    Private Sub CheckBox1_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckBox1.CheckedChanged
        If CheckBox1.Checked Then
            UsernameTextBox.Text = Environ("USERNAME")
            UsernameTextBox.Enabled = False
            PasswordTextBox.Enabled = False
        Else
            UsernameTextBox.Text = ""
            UsernameTextBox.Enabled = True
            PasswordTextBox.Enabled = True

        End If
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        DBInfo.Show()
    End Sub

    Private Sub Login_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        'My.Settings.dbloc = "NULL"
        My.Settings.IsSpoof = "NULL"
        Me.Text = Application.ProductName & " Login"
        Me.CheckBox1.Checked = True
        Me.Visible = False
        If My.Settings.dbloc = "NULL" Or My.Settings.roledbloc = "NULL" Then
            MsgBox("It appears this is your first time lauching this program." & vbNewLine & "We will now attempt to map to the database on your local network.  Please be patient during this one time event.  Should the automatic mapping fail, you will be prompted to manually enter the location and DB password.  This can be obtained from the Asset Management Group.", MsgBoxStyle.Information, "Welcome!")
            Me.Visible = False
            MapDB.Show()
            'DBInfo.Show()
        Else
            Me.Visible = True
        End If
    End Sub
End Class
