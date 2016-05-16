Public Class MapDB
    Dim tread1 As System.Threading.Thread

    Public Sub Pause(ByVal dblSecs As Double)
        Const OneSec As Double = 1.0# / (1440.0# * 60.0#)
        Dim dblWaitTil As Date
        Now.AddSeconds(OneSec)
        dblWaitTil = Now.AddSeconds(OneSec).AddSeconds(dblSecs)
        Do Until Now > dblWaitTil
            Application.DoEvents() ' Allow windows messages to be processed
        Loop
    End Sub

    Private Sub MapDB_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Label1.Text = "Attempting to assign local location..."
        ProgressBar1.Value = 25
        Pause(0.01)
        My.Settings.dbloc = "\\monumentco1\data\ToolboxDBs\SalesInterface2.accdb"
        'C:\SalesInterface.accdb"
        ProgressBar1.Value = 35
        Pause(0.01)
        My.Settings.roledbloc = "\\monumentco1\data\ToolboxDBs\SalesInterface2.accdb"
        'C:\SalesInterface.accdb"
        ProgressBar1.Value = 40
        Pause(0.01)
        Label1.Text = "Assigned local location.  Authenticating..."
        ProgressBar1.Value = 45
        Pause(0.01)
        My.Settings.dbpass = "Denver09"
        ProgressBar1.Value = 60
        Pause(0.01)
        Label1.Text = "Assigned cridentials.  Checking for connection..."
        ProgressBar1.Value = 75
        Pause(0.01)
        'MsgBox("Database asssigned: " & My.Settings.dbloc & vbNewLine & "Password Assigned: " & My.Settings.dbpass)

        Control.CheckForIllegalCrossThreadCalls = False
        tread1 = New System.Threading.Thread(AddressOf checkcon)
        tread1.Start()

        
    End Sub

    Private Sub checkcon()
        Dim Mycn As OleDb.OleDbConnection

        Try

            Mycn = New OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & My.Settings.dbloc & "';Jet OLEDB:Database Password='" & My.Settings.dbpass & "';")

            Mycn.Open()

            Dim sqlstring As String

            sqlstring = "SELECT Top 1 ID FROM sys_Users"

            Dim queryString As String = String.Format(sqlstring)
            Dim cmd As New OleDb.OleDbCommand(queryString, Mycn)
            Dim da As New OleDb.OleDbDataAdapter(cmd)
            Dim ds As New DataSet

            da.Fill(ds, "User")
            Dim dt As DataTable = ds.Tables("User")
            If dt.Rows.Count > 0 Then
                ProgressBar1.Value = 100
                'Login.OK.Enabled = True
                Pause(0.01)
            Else

            End If
            Mycn.Close()
            Mycn.Dispose()

            MsgBox("The database has been successfully mapped.", MsgBoxStyle.Information, "Success")

        Catch ex As Exception
            'MsgBox(ex.Message)
            'Timer1.Enabled = False

            Call dbfail()
            
            'Me.Close()
            'MsgBox("No database could be found" & vbNewLine & ex, MsgBoxStyle.Critical, "ERROR")
        End Try
    End Sub

    Private Sub Timer1_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer1.Tick
        If ProgressBar1.Value = 100 Then
            Login.Show()
            Login.OK.Enabled = True
            DBInfo.Close()
            Me.Close()
        Else
            'Login.Visible = False
        End If
    End Sub

    Private Sub dbfail()
        MsgBox("We were unable to successfully map to the local database.  Please enter the path and password of the local DB.", MsgBoxStyle.Critical, "Connection Issue")
        'Timer1.Enabled = False
        'Login.Show()
        Me.Close()
        'DBInfo.Show()
        'DBInfo.Visible = True
    End Sub
End Class