Public Class Firms
    Dim thread1 As System.Threading.Thread

    Public Sub ReloadFirms()
        'Try

        Dim conn As New OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & My.Settings.dbloc & "';Jet OLEDB:Database Password='" & My.Settings.dbpass & "';")

        Dim strSQL As String = "SELECT * From env_Firms"

        Dim da As New OleDb.OleDbDataAdapter(strSQL, conn)
        Dim ds As New DataSet
        da.Fill(ds, "Users")

        'DataGridView1.DataSource = ds.Tables("Users")
        'DataGridView1.Columns(0).Visible = False

        With DataGridView1
            .DataSource = ds.Tables("Users")
            .Columns(0).Visible = False
        End With
        DataGridView1.Visible = True

        Label3.Visible = False
        DataGridView1.Refresh()
        'Catch ex As Exception
        'MsgBox(ex.Message, MsgBoxStyle.Critical, "ERROR")
        'Label3.Text = "ERROR"
        'End Try
    End Sub

    Public Sub ReloadFirmsSpec()
        Try

            Dim conn As New OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & My.Settings.dbloc & "';Jet OLEDB:Database Password='" & My.Settings.dbpass & "';")

            Dim strSQL As String = "SELECT * FROM env_Firms WHERE [Company Name] LIKE '" & TextBox1.Text & "%'"
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
            Label3.Text = "ERROR"
        End Try
    End Sub

    Private Sub OK_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles OK.Click

        If IsDBNull(TextBox1) Or TextBox1.Text = "" Then
            DataGridView1.Visible = False
            Label3.Visible = True
            Label3.Text = "Loading..."
            Control.CheckForIllegalCrossThreadCalls = False
            thread1 = New System.Threading.Thread(AddressOf ReloadFirms)
            thread1.Start()
        Else
            DataGridView1.Visible = False
            Label3.Visible = True
            Label3.Text = "Loading..."
            Control.CheckForIllegalCrossThreadCalls = False
            thread1 = New System.Threading.Thread(AddressOf ReloadFirmsSpec)
            thread1.Start()
        End If
    End Sub

    Private Sub DataGridView1_CellContentClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick

    End Sub

    Private Sub DataGridView1_CellContentDoubleClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView1.CellContentDoubleClick
        Dim child As New FirmOverview
        child.MdiParent = Home
        child.Show()
        child.FirmName.Text = DataGridView1.SelectedCells(2).Value
        child.ContactID.Text = DataGridView1.SelectedCells(0).Value
        child.ContactCode.Text = DataGridView1.SelectedCells(1).Value
        child.Text = "Viewing Firm: " & DataGridView1.SelectedCells(2).Value
    End Sub

    Private Sub Firms_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Dim p = GetType(DataGridView).GetProperty("DoubleBuffered", Reflection.BindingFlags.Instance Or Reflection.BindingFlags.NonPublic)
        p.SetValue(Me.DataGridView1, True, Nothing)
    End Sub
End Class