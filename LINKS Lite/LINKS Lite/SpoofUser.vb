Public Class SpoofUser
    Dim thread1 As System.Threading.Thread

    Public Sub LoadUsers()
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

            Label3.Visible = False
            ComboBox1.Visible = True

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "ERROR")
            'Application.Exit()
        End Try
    End Sub

    Private Sub SpoofUser_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        ComboBox1.Visible = False
        Label3.Visible = True

        Control.CheckForIllegalCrossThreadCalls = False
        thread1 = New System.Threading.Thread(AddressOf LoadUsers)
        thread1.Start()

        
    End Sub

    Private Sub OK_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles OK.Click
        If My.Settings.origuser = ComboBox1.SelectedValue Then
            My.Settings.IsSpoof = "FALSE"
            My.Settings.userid = ComboBox1.SelectedValue
            Call Permissions.Populate()
            Me.Close()
        Else
            My.Settings.IsSpoof = "TRUE"
            My.Settings.userid = ComboBox1.SelectedValue
            Permissions.Show()
            Call Permissions.Populate()
            Dim ht As String
            ht = Home.Text
            Home.Text = "***SPOFFING USER*** " & ht
            Me.Close()
        End If
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Me.Close()
    End Sub
End Class