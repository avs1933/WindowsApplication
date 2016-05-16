Public Class Users
    Dim thread1 As System.Threading.Thread
    Dim thread2 As System.Threading.Thread

    Public Sub ReloadRoles()
        Try

            Dim conn As New OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & My.Settings.roledbloc & "';Jet OLEDB:Database Password='" & My.Settings.dbpass & "';")

            Dim strSQL As String = "SELECT ID, RoleName As [Role Name], Active FROM sys_Roles"
            Dim da As New OleDb.OleDbDataAdapter(strSQL, conn)
            Dim ds As New DataSet
            da.Fill(ds, "Users")

            With DataGridView2
                .DataSource = ds.Tables("Users")
                .Columns(0).Visible = False
            End With
            DataGridView2.Visible = True
            Label4.Visible = False

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "ERROR")
            'Application.Exit()
        End Try
    End Sub

    Public Sub ReloadUsers()
        Try

            Dim conn As New OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & My.Settings.roledbloc & "';Jet OLEDB:Database Password='" & My.Settings.dbpass & "';")

            Dim strSQL As String = "SELECT ID, FullName As [User Name], Team As [Team], APXName As [APX Name], Active FROM sys_Users"
            Dim da As New OleDb.OleDbDataAdapter(strSQL, conn)
            Dim ds As New DataSet
            da.Fill(ds, "Users")

            With DataGridView1
                .DataSource = ds.Tables("Users")
                .Columns(0).Visible = False
            End With
            DataGridView1.Visible = True
            Label3.Visible = False

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "ERROR")
            'Application.Exit()
        End Try
    End Sub

    Private Sub Users_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        DataGridView1.Visible = False
        Label3.Visible = True
        DataGridView2.Visible = False
        Label4.Visible = True
        Control.CheckForIllegalCrossThreadCalls = False
        thread1 = New System.Threading.Thread(AddressOf ReloadUsers)
        thread1.Start()

        thread2 = New System.Threading.Thread(AddressOf ReloadRoles)
        thread2.Start()
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        DataGridView1.Visible = False
        Label3.Visible = True
        Control.CheckForIllegalCrossThreadCalls = False
        thread1 = New System.Threading.Thread(AddressOf ReloadUsers)
        thread1.Start()
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        DataGridView2.Visible = False
        Label4.Visible = True
        Control.CheckForIllegalCrossThreadCalls = False
        thread2 = New System.Threading.Thread(AddressOf ReloadRoles)
        thread2.Start()
    End Sub


    Private Sub DataGridView1_CellDoubleClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView1.CellDoubleClick
        Dim child As New UserEdit
        child.MdiParent = Home
        child.Show()
        child.ID.Text = DataGridView1.SelectedCells(0).Value
        Call child.CurrentUser()
    End Sub

    Private Sub OK_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles OK.Click
        Dim child As New UserEdit
        child.MdiParent = Home
        child.Show()
        child.ID.Text = "NEW"
        child.ID.Enabled = False
    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        Dim child As New RoleEdit
        child.MdiParent = Home
        child.Show()
        child.ID.Text = "NEW"
        child.ID.Enabled = False
    End Sub

    Private Sub DataGridView1_CellContentClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick

    End Sub

    Private Sub DataGridView2_CellContentClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView2.CellContentClick

    End Sub

    Private Sub DataGridView2_CellContentDoubleClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView2.CellContentDoubleClick
        Dim child As New RoleEdit
        child.MdiParent = Home
        child.Show()
        child.ID.Text = DataGridView2.SelectedCells(0).Value
        Call child.CurrentRole()
    End Sub
End Class