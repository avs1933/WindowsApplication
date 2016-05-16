Public Class RevenueCenterViewFirms

    Private Sub RevenueCenterViewFirms_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Call ReloadGrid()
    End Sub

    Public Sub ReloadGrid()
        Try
            Dim conn As New OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & My.Settings.dbloc & "';Jet OLEDB:Database Password='" & My.Settings.dbpass & "';")

            Dim strSQL As String

            strSQL = "SELECT mdb_BillingFirms.ID, mdb_BillingFirms.MAPFirmID, mdb_BillingFirms.APXFirmName As [APX Firm], mdb_BillingFirms.FirmName As [Firm Name], mdb_BillingFirms.FirmAttn As [Firm Attn], mdb_BillingFirms.Address1, mdb_BillingFirms.Address2, mdb_BillingFirms.City, mdb_BillingFirms.State, mdb_BillingFirms.Zip, mdb_BillingFirms.Phone1, mdb_BillingFirms.Extention, mdb_BillingFirms.Locked" & _
            " FROM mdb_BillingFirms;"


            Dim da As New OleDb.OleDbDataAdapter(strSQL, conn)
            Dim ds As New DataSet
            da.Fill(ds, "Users")

            With DataGridView1
                .DataSource = ds.Tables("Users")
                .Columns(0).Visible = False
                .Columns(1).Visible = False
            End With

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "ERROR")
        End Try
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Call ReloadGrid()
    End Sub

    Private Sub DataGridView1_CellContentDoubleClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView1.CellContentDoubleClick
        Dim child As New RevenueCenter_Address
        child.MdiParent = Home
        child.Show()
        child.ID.Text = DataGridView1.SelectedCells(0).Value
        child.MapID.Text = DataGridView1.SelectedCells(1).Value
        child.cboAPXFirm.SelectedText = DataGridView1.SelectedCells(2).Value
        child.TextBox1.Text = DataGridView1.SelectedCells(3).Value
        child.TextBox2.Text = DataGridView1.SelectedCells(4).Value
        child.TextBox3.Text = DataGridView1.SelectedCells(5).Value
        child.TextBox4.Text = DataGridView1.SelectedCells(6).Value
        child.TextBox5.Text = DataGridView1.SelectedCells(7).Value
        child.TextBox6.Text = DataGridView1.SelectedCells(8).Value
        child.TextBox7.Text = DataGridView1.SelectedCells(9).Value
        child.TextBox8.Text = DataGridView1.SelectedCells(10).Value
        child.TextBox9.Text = DataGridView1.SelectedCells(11).Value
        
        child.ckbLocked.Checked = DataGridView1.SelectedCells(12).Value
        

    End Sub

    Private Sub EditToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles EditToolStripMenuItem.Click
        Dim child As New RevenueCenter_Address
        child.MdiParent = Home
        child.Show()
        child.ID.Text = DataGridView1.SelectedCells(0).Value
        child.MapID.Text = DataGridView1.SelectedCells(1).Value
        child.cboAPXFirm.SelectedText = DataGridView1.SelectedCells(2).Value
        child.TextBox1.Text = DataGridView1.SelectedCells(3).Value
        child.TextBox2.Text = DataGridView1.SelectedCells(4).Value
        child.TextBox3.Text = DataGridView1.SelectedCells(5).Value
        child.TextBox4.Text = DataGridView1.SelectedCells(6).Value
        child.TextBox5.Text = DataGridView1.SelectedCells(7).Value
        child.TextBox6.Text = DataGridView1.SelectedCells(8).Value
        child.TextBox7.Text = DataGridView1.SelectedCells(9).Value
        child.TextBox8.Text = DataGridView1.SelectedCells(10).Value
        child.TextBox9.Text = DataGridView1.SelectedCells(11).Value

        child.ckbLocked.Checked = DataGridView1.SelectedCells(12).Value
    End Sub

    Private Sub DeleteToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles DeleteToolStripMenuItem.Click
        Dim ir As MsgBoxResult
        ir = MsgBox("Are you sure you want to delete this record?", MsgBoxStyle.YesNo + MsgBoxStyle.Critical, "Confirm Delete")
        If ir = MsgBoxResult.Yes Then
            Try
                Dim Mycn As OleDb.OleDbConnection
                Dim Command As OleDb.OleDbCommand
                Dim SQLstr As String

                Mycn = New OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & My.Settings.dbloc & "';Jet OLEDB:Database Password='" & My.Settings.dbpass & "';")

                Mycn.Open()
                
                SQLstr = "DELETE * FROM mdb_BillingFirms" & _
                " WHERE ID = " & DataGridView1.SelectedCells(0).Value
                Command = New OleDb.OleDbCommand(SQLstr, Mycn)
                Command.ExecuteNonQuery()
                Mycn.Close()
                

            Catch ex As Exception
                MsgBox(ex.Message)
            End Try

            Call ReloadGrid()

        Else

        End If
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        Me.Close()
    End Sub
End Class