Public Class map_UserViewFirms

    Private Sub map_UserViewFirms_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Call LoadTypeOfFirm()
    End Sub

    Public Sub LoadTypeOfFirm()
        Try

            Dim conn As New OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & My.Settings.dbloc & "';Jet OLEDB:Database Password='" & My.Settings.dbpass & "';")

            Dim strSQL As String = "SELECT ID, FirmType FROM map_FirmType WHERE ID <> 7 ORDER BY FirmType"
            Dim da As New OleDb.OleDbDataAdapter(strSQL, conn)
            Dim ds As New DataSet
            da.Fill(ds, "Users")

            With ComboBox1
                .DataSource = ds.Tables("Users")
                .DisplayMember = "FirmType"
                .ValueMember = "ID"
                .SelectedIndex = 0
            End With

            'Call LoadDBData()

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "ERROR")
        End Try
    End Sub

    Public Sub LoadTbl()
        Try
            Dim conn As New OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & My.Settings.dbloc & "';Jet OLEDB:Database Password='" & My.Settings.dbpass & "';")

            Dim strSQL As String

            If CheckBox1.Checked Then
                strSQL = "SELECT Map_Firms.ID, map_FirmType.ID, Map_Firms.AdventContactID, Map_Firms.AdventPortfolioFirm, Map_Firms.FirmName As [Firm], Map_Firms.Address, Map_Firms.City, Map_Firms.State, Map_Firms.Zip, Map_Firms.Phone, map_FirmType.FirmType As [Type of Firm], Map_Firms.NotInAdvent As [Not in APX], Map_Firms.Active" & _
                " FROM (Map_Firms INNER JOIN AdvApp_vContact ON Map_Firms.AdventContactID = AdvApp_vContact.ContactID) INNER JOIN map_FirmType ON Map_Firms.TypeID = map_FirmType.ID" & _
                " WHERE Map_Firms.Active = -1 AND map_FirmType.SystemValue = False AND map_Firms.TypeID = " & ComboBox1.SelectedValue & " AND map_Firms.FirmName Like '%" & TextBox1.Text & "%'"
            Else
                strSQL = "SELECT Map_Firms.ID, map_FirmType.ID, Map_Firms.AdventContactID, Map_Firms.AdventPortfolioFirm, Map_Firms.FirmName As [Firm], Map_Firms.Address, Map_Firms.City, Map_Firms.State, Map_Firms.Zip, Map_Firms.Phone, map_FirmType.FirmType As [Type of Firm], Map_Firms.NotInAdvent As [Not in APX], Map_Firms.Active" & _
                " FROM (Map_Firms INNER JOIN AdvApp_vContact ON Map_Firms.AdventContactID = AdvApp_vContact.ContactID) INNER JOIN map_FirmType ON Map_Firms.TypeID = map_FirmType.ID" & _
                " WHERE Map_Firms.Active = -1 AND map_FirmType.SystemValue = False AND map_Firms.FirmName Like '%" & TextBox1.Text & "%'"
            End If
            
            Dim da As New OleDb.OleDbDataAdapter(strSQL, conn)
            Dim ds As New DataSet
            da.Fill(ds, "Users")

            With DataGridView1
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

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Call LoadTbl()
    End Sub

    Private Sub DataGridView1_CellContentClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick

    End Sub

    Private Sub DataGridView1_CellContentDoubleClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView1.CellContentDoubleClick
        Dim child As New map_Firm
        child.MdiParent = Home
        child.ID.Text = DataGridView1.SelectedCells(0).Value
        child.TextBox1.Text = DataGridView1.SelectedCells(4).Value
        child.Show()
    End Sub

    Private Sub CheckBox1_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckBox1.CheckedChanged
        If CheckBox1.Checked Then
            ComboBox1.Enabled = True
        Else
            ComboBox1.Enabled = False
        End If
    End Sub
End Class