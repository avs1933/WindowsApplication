Public Class map_ViewPlatforms

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Call LoadPlatforms()
    End Sub

    Public Sub LoadPlatforms()
        Try
            Dim conn As New OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & My.Settings.dbloc & "';Jet OLEDB:Database Password='" & My.Settings.dbpass & "';")
            Dim SqlString As String

            SqlString = "SELECT map_Platforms.ID, Map_Firms.FirmName As [Platform Name], map_Platforms.ApprovalName As [Approval Name], map_Platforms.Active" & _
            " FROM Map_Firms INNER JOIN map_Platforms ON Map_Firms.ID = map_Platforms.PlatformID" & _
            " WHERE(((map_Platforms.Active) = True) AND Map_Firms.FirmName Like '%" & TextBox1.Text & "%')" & _
            " GROUP BY map_Platforms.ID, Map_Firms.FirmName, map_Platforms.ApprovalName, map_Platforms.Active;"

            Dim da As New OleDb.OleDbDataAdapter(SqlString, conn)
            Dim ds As New DataSet
            da.Fill(ds, "Users")

            With DataGridView1
                .DataSource = ds.Tables("Users")
                .Columns(0).Visible = False
            End With

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "ERROR")
        End Try
    End Sub

    Private Sub map_ViewPlatforms_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

    End Sub

    Public Sub LoadEditForm()
        If DataGridView1.RowCount = "0" Then
            'do nothing
        Else
            Dim child As New map_EditPlatform
            child.MdiParent = Home
            child.Show()

            Try
                Dim conn As New OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & My.Settings.dbloc & "';Jet OLEDB:Database Password='" & My.Settings.dbpass & "';")

                Dim strSQL As String = "SELECT * FROM map_Platforms WHERE ID = " & DataGridView1.SelectedCells(0).Value
                Dim queryString As String = String.Format(strSQL)
                Dim cmd As New OleDb.OleDbCommand(queryString, conn)
                Dim da As New OleDb.OleDbDataAdapter(cmd)
                Dim ds As New DataSet

                da.Fill(ds, "User")
                Dim dt As DataTable = ds.Tables("User")
                Dim row1 As DataRow = dt.Rows(0)

                child.ID.Text = (row1("ID"))

                If IsDBNull(row1("ApprovalName")) Then
                    child.TextBox1.Text = ""
                Else
                    child.TextBox1.Text = (row1("ApprovalName"))
                End If

                If IsDBNull(row1("DBDriverID")) Then
                    'do nothing
                Else
                    child.cboDatabase.SelectedValue = (row1("DBDriverID"))
                End If

                If IsDBNull(row1("PlatformID")) Then
                    'do nothing
                Else
                    child.cboFirms.SelectedValue = (row1("PlatformID"))
                End If

                If IsDBNull(row1("ApprovalProcess")) Then
                    child.RichTextBox1.Text = ""
                Else
                    child.RichTextBox1.Text = (row1("ApprovalProcess"))
                End If

                child.ckbActive.Checked = (row1("Active"))
                child.ckbMF.Checked = (row1("OfferMF"))
                child.ckbSMA.Checked = (row1("OfferSMA"))
                child.ckbUIT.Checked = (row1("OfferUIT"))
                child.ckbUMA.Checked = (row1("OfferUMA"))

            Catch ex As Exception
                MsgBox(ex.Message, MsgBoxStyle.Critical, "ERROR")
            End Try
        End If
    End Sub

    Private Sub DataGridView1_CellContentClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick

    End Sub

    Private Sub DataGridView1_CellContentDoubleClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView1.CellContentDoubleClick
        If Permissions.MAPEditPlatform.Checked Then
            Call LoadEditForm()
        Else
        End If
    End Sub
End Class