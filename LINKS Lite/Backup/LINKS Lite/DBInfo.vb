Imports System.Runtime.InteropServices

Public Class DBInfo

    '<DllImport("ODBCCP32.DLL")> Shared Function SQLConfigDataSource _
    '(ByVal hwndParent As Integer, ByVal fRequest As Integer, _
    'ByVal lpszDriver As String, _
    'ByVal lpszAttributes As String) As Boolean
    'End Function

    Dim tread1 As System.Threading.Thread
    'Dim thread2 As System.Threading.Thread

    'Private Declare Function SQLConfigDataSource Lib "ODBCCP32.DLL" (ByVal hwndParent As Integer, ByVal ByValfRequest As Integer, ByVal lpszDriver As String, ByVal lpszAttributes As String) As Integer
    'Private Declare Function SQLInstallerError Lib "ODBCCP32.DLL" (ByVal iError As Integer, ByRef pfErrorCode As Integer, ByVal lpszErrorMsg As System.Text.StringBuilder, ByVal cbErrorMsgMax As Integer, ByRef pcbErrorMsg As Integer) As Integer

    'Private Const ODBC_ADD_DSN As Short = 1 ' Add data source
    'Private Const ODBC_ADD_SYS_DSN As Short = 4
    'Private Const vbAPINull As Integer = 0 ' NULL Pointer

    'Constant Declaration
    Private Const ODBC_ADD_DSN = 1        ' Add data source
    Private Const ODBC_CONFIG_DSN = 2     ' Configure (edit) data source
    Private Const ODBC_REMOVE_DSN = 3     ' Remove data source
    Private Const vbAPINull As Long = 0  ' NULL Pointer

    'Function Declare
    '#If WIN32 Then

    Private Declare Function SQLConfigDataSource Lib "ODBCCP32.DLL" _
    (ByVal hwndParent As Long, ByVal fRequest As Long, _
    ByVal lpszDriver As String, ByVal lpszAttributes As String) _
    As Long
    '#Else
    'Private Declare Function SQLConfigDataSource Lib "ODBCINST.DLL" _
    '(ByVal hwndParent As Integer, ByVal fRequest As Integer, ByVal _
    'lpszDriver As String, ByVal lpszAttributes As String) As Integer
    '#End If



    Public Sub Pause(ByVal dblSecs As Double)
        Const OneSec As Double = 1.0# / (1440.0# * 60.0#)
        Dim dblWaitTil As Date
        Now.AddSeconds(OneSec)
        dblWaitTil = Now.AddSeconds(OneSec).AddSeconds(dblSecs)
        Do Until Now > dblWaitTil
            Application.DoEvents() ' Allow windows messages to be processed
        Loop
    End Sub

    Private Sub DBInfo_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Me.Text = Application.ProductName & " DB Info"
        ProgressBar1.Value = 0
        ComboBox1.SelectedItem = My.Settings.dbloc
        ComboBox2.SelectedItem = My.Settings.roledbloc
        TextBox1.Text = My.Settings.dbpass
        Me.Show()
        Login.OK.Enabled = False
    End Sub

    Private Sub Cancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Cancel.Click
        Me.Close()
    End Sub

    Private Sub Test_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Test.Click
        ProgressBar1.Value = 50
        Control.CheckForIllegalCrossThreadCalls = False
        tread1 = New System.Threading.Thread(AddressOf checkcon)
        tread1.Start()
    End Sub

    Private Sub checkcon()
        Dim Mycn As OleDb.OleDbConnection
        Dim permission As OleDb.OleDbConnection

        Try

            Mycn = New OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & ComboBox1.SelectedItem.ToString & "';Jet OLEDB:Database Password='" & TextBox1.Text & "';")
            permission = New OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & ComboBox2.SelectedItem.ToString & "';Jet OLEDB:Database Password='" & TextBox1.Text & "';")

            Mycn.Open()
            permission.Open()

            Dim sqlstring As String

            sqlstring = "SELECT Top 1 ID FROM sys_Users"

            'Search for Primary DB
            Dim queryString As String = String.Format(sqlstring)
            Dim cmd As New OleDb.OleDbCommand(queryString, Mycn)
            Dim da As New OleDb.OleDbDataAdapter(cmd)
            Dim ds As New DataSet

            da.Fill(ds, "User")
            Dim dt As DataTable = ds.Tables("User")

            'Search for Permissions DB
            Dim queryString1 As String = String.Format(sqlstring)
            Dim cmd1 As New OleDb.OleDbCommand(queryString1, permission)
            Dim da1 As New OleDb.OleDbDataAdapter(cmd1)
            Dim ds1 As New DataSet

            da1.Fill(ds1, "User")
            Dim dt1 As DataTable = ds1.Tables("User")

            If dt.Rows.Count > 0 Then
                ProgressBar1.Value = 75
                If dt1.Rows.Count > 0 Then
                    ProgressBar1.Value = 100
                    Login.OK.Enabled = True
                    Pause(0.01)
                Else
                End If
            Else

            End If
            Mycn.Close()
            Mycn.Dispose()

            permission.Close()
            permission.Dispose()

            MsgBox("The database has been successfully found.", MsgBoxStyle.Information, "Success")

        Catch ex As Exception

            ProgressBar1.Value = 0
            MsgBox(ex.Message)

            'MsgBox("No database could be found" & vbNewLine & ex, MsgBoxStyle.Critical, "ERROR")
        End Try
    End Sub

    Private Sub Save_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Save.Click
        ProgressBar1.Value = 50
        Timer1.Enabled = True
        Control.CheckForIllegalCrossThreadCalls = False
        tread1 = New System.Threading.Thread(AddressOf checkcon)
        tread1.Start()
    End Sub

    Private Sub Timer1_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer1.Tick
        If ProgressBar1.Value = 100 Then
            Login.OK.Enabled = True
            ProgressBar1.Value = 99
            My.Settings.dbloc = ComboBox1.SelectedItem.ToString
            My.Settings.roledbloc = ComboBox2.SelectedItem.ToString
            My.Settings.dbpass = TextBox1.Text
            MsgBox("The database has been successfully updated and saved.", MsgBoxStyle.Information, "Success")
            Timer1.Enabled = False
            Me.Close()
        Else

        End If
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        MapDB.Show()
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        Dim ir As MsgBoxResult

        ir = MsgBox("Are you sure you want to create System DSN's for this machine?", MsgBoxStyle.YesNo + MsgBoxStyle.Information, "Confirm")

        If ir = MsgBoxResult.Yes Then

            Call CreateAPXDSN()
            Call CreateMoxyDSN()

        Else

        End If

    End Sub

    Public Sub CreateAPXDSN()
        'Dim intRet As Integer
        'Dim Driver As String
        'Dim Attributes As String

        'Set the driver to SQL Server because it is most common.
        'Driver = "SQL Server"
        'Set the attributes delimited by null.
        'See driver documentation for a complete
        'list of supported attributes.
        'Attributes = "SERVER=adventsql01-1" & Chr(0)
        'Attributes = Attributes & "DESCRIPTION=APX Connection" & Chr(0)
        'Attributes = Attributes & "DSN=Local DSN" & Chr(0)
        'Attributes = Attributes & "DATABASE=Moxy60" & Chr(0)
        'Unsupported by SQL Server
        'Attributes = Attributes & "Uid=AdventAccessDBUser" & Chr(0) & "pwd=A%Ven$d8" & Chr(0)
        'To show dialog, use Form1.Hwnd instead of vbAPINull.
        'intRet = SQLConfigDataSource(vbAPINull, ODBC_ADD_DSN, Driver, Attributes)
        'If intRet <> 0 Then
        'MsgBox("APX DSN Created")
        'Else
        'Dim nErrorCode As Integer
        'Dim strError As New System.Text.StringBuilder(255)
        'Dim nErrorLen As Integer
        'intRet = SQLInstallerError(1, nErrorCode, strError, 255, nErrorLen)
        'MsgBox("APX Create Failed - " & strError.ToString, nErrorLen)
        'Left$
        'End If

#If WIN32 Then
          Dim intRet As Long
#Else
        Dim intRet As Integer
#End If
        Dim strDriver As String
        Dim strAttributes As String

        'Set the driver to SQL Server because it is most common.
        strDriver = "SQL Server"
        'Set the attributes delimited by null.
        'See driver documentation for a complete
        'list of supported attributes.
        strAttributes = "SERVER=adventsql01" & Chr(0)
        strAttributes = strAttributes & "DESCRIPTION=APX Connection" & Chr(0)
        strAttributes = strAttributes & "DSN=adventsql01-1" & Chr(0)
        'strAttributes = strAttributes & "DATABASE=pubs" & Chr(0)
        strAttributes = strAttributes & "Uid=AdventAccessDBUser" & Chr(0) & "pwd=A%Ven$d8" & Chr(0)
        'To show dialog, use Form1.Hwnd instead of vbAPINull.
        intRet = SQLConfigDataSource(vbAPINull, ODBC_ADD_DSN, _
        strDriver, strAttributes)
        If intRet Then
            MsgBox("DSN Created")
        Else
            MsgBox("Create Failed")
        End If


    End Sub

    Public Sub CreateMoxyDSN()
        'Dim intRet As Integer
        'Dim Driver As String
        'Dim Attributes As String

        'Set the driver to SQL Server because it is most common.
        'Driver = "SQL Server"
        'Set the attributes delimited by null.
        'See driver documentation for a complete
        'list of supported attributes.
        'Attributes = "SERVER=adventsql-1" & Chr(0)
        'Attributes = Attributes & "DESCRIPTION=Moxy Connection" & Chr(0)
        'Attributes = Attributes & "DSN=Local DSN" & Chr(0)
        'Attributes = Attributes & "DATABASE=Moxy60" & Chr(0)
        'Unsupported by SQL Server
        'Attributes = Attributes & "Uid=AdventAccessDBUser" & Chr(0) & "pwd=A%Ven$d8" & Chr(0)
        'To show dialog, use Form1.Hwnd instead of vbAPINull.
        'intRet = SQLConfigDataSource(vbAPINull, ODBC_ADD_DSN, Driver, Attributes)
        'If intRet <> 0 Then
        'MsgBox("Moxy DSN Created")
        'Else
        'Dim nErrorCode As Integer
        'Dim strError As New System.Text.StringBuilder(255)
        'Dim nErrorLen As Integer
        'intRet = SQLInstallerError(1, nErrorCode, strError, 255, nErrorLen)
        'MsgBox("Moxy Create Failed - " & strError.ToString, nErrorLen)
        'Left$
        'End If

#If WIN32 Then
          Dim intRet As Long
#Else
        Dim intRet As Integer
#End If
        Dim strDriver As String
        Dim strAttributes As String

        'Set the driver to SQL Server because it is most common.
        strDriver = "SQL Server"
        'Set the attributes delimited by null.
        'See driver documentation for a complete
        'list of supported attributes.
        strAttributes = "SERVER=adventsql" & Chr(0)
        strAttributes = strAttributes & "DESCRIPTION=Moxy Connection" & Chr(0)
        strAttributes = strAttributes & "DSN=adventsql-1" & Chr(0)
        'strAttributes = strAttributes & "DATABASE=pubs" & Chr(0)
        strAttributes = strAttributes & "Uid=AdventAccessDBUser" & Chr(0) & "pwd=A%Ven$d8" & Chr(0)
        'To show dialog, use Form1.Hwnd instead of vbAPINull.
        intRet = SQLConfigDataSource(vbAPINull, ODBC_ADD_DSN, _
        strDriver, strAttributes)
        If intRet Then
            MsgBox("DSN Created")
        Else
            MsgBox("Create Failed")
        End If
    End Sub

End Class