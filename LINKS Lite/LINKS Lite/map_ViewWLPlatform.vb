Public Class map_ViewWLPlatform

    Private Sub CheckBox1_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckBox1.CheckedChanged
        If CheckBox1.Checked Then
            cboPlatforms.Enabled = True
        Else
            cboPlatforms.Enabled = False
        End If
    End Sub

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

            If CheckBox1.Checked Then
                SqlString = "SELECT map_PlatformsWL.ID, map_PlatformsWL.WLName As [Platform Name], Map_Firms.FirmName As [Driving Platform]" & _
                " FROM (map_PlatformsWL INNER JOIN map_Platforms ON map_PlatformsWL.PlatformDriverID = map_Platforms.ID) INNER JOIN Map_Firms ON map_Platforms.PlatformID = Map_Firms.ID" & _
                " WHERE(((map_PlatformsWL.Active) = True) AND map_PlatformsWL.PlatformDriverID = " & cboPlatforms.SelectedValue & ")" & _
                " GROUP BY map_PlatformsWL.ID, map_PlatformsWL.WLName, Map_Firms.FirmName;"
            Else
                SqlString = "SELECT map_PlatformsWL.ID, map_PlatformsWL.WLName As [Platform Name], Map_Firms.FirmName As [Driving Platform]" & _
                " FROM (map_PlatformsWL INNER JOIN map_Platforms ON map_PlatformsWL.PlatformDriverID = map_Platforms.ID) INNER JOIN Map_Firms ON map_Platforms.PlatformID = Map_Firms.ID" & _
                " WHERE(((map_PlatformsWL.Active) = True))" & _
                " GROUP BY map_PlatformsWL.ID, map_PlatformsWL.WLName, Map_Firms.FirmName;"
            End If

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

    Private Sub map_ViewWLPlatform_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Call LoadAPXFirms()
    End Sub
    Public Sub LoadAPXFirms()
        Try

            Dim conn As New OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & My.Settings.dbloc & "';Jet OLEDB:Database Password='" & My.Settings.dbpass & "';")

            Dim strSQL As String = "SELECT map_Platforms.ID, Map_Firms.FirmName" & _
            " FROM map_Platforms INNER JOIN Map_Firms ON map_Platforms.PlatformID = Map_Firms.ID" & _
            " WHERE(((map_Platforms.Active) = True))" & _
            " GROUP BY map_Platforms.ID, Map_Firms.FirmName;"

            Dim da As New OleDb.OleDbDataAdapter(strSQL, conn)
            Dim ds As New DataSet
            da.Fill(ds, "Users")

            With cboPlatforms
                .DataSource = ds.Tables("Users")
                .DisplayMember = "FirmName"
                .ValueMember = "ID"
                .SelectedIndex = 0
            End With

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "ERROR")
        End Try
    End Sub

    Private Sub DataGridView1_CellContentClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick

    End Sub

    Public Sub LoadWLEdit()
        If DataGridView1.RowCount = "0" Then
            'do nothing
        Else
            Dim child As New map_EditWLPlatform
            child.MdiParent = Home
            child.Show()

            Try
                Dim conn As New OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & My.Settings.dbloc & "';Jet OLEDB:Database Password='" & My.Settings.dbpass & "';")

                Dim strSQL As String = "SELECT * FROM map_PlatformsWL WHERE ID = " & DataGridView1.SelectedCells(0).Value
                Dim queryString As String = String.Format(strSQL)
                Dim cmd As New OleDb.OleDbCommand(queryString, conn)
                Dim da As New OleDb.OleDbDataAdapter(cmd)
                Dim ds As New DataSet

                da.Fill(ds, "User")
                Dim dt As DataTable = ds.Tables("User")
                Dim row1 As DataRow = dt.Rows(0)

                child.ID.Text = (row1("ID"))

                If IsDBNull(row1("WLName")) Then
                    child.WLName.Text = ""
                Else
                    child.WLName.Text = (row1("WLName"))
                End If

                If IsDBNull(row1("PlatformDriverID")) Then
                    'do nothing
                Else
                    child.cboPlatform.SelectedValue = (row1("PlatformDriverID"))
                End If

                If IsDBNull(row1("AdditionalApprovalProcess")) Then
                    child.RichTextBox1.Text = ""
                Else
                    child.RichTextBox1.Text = (row1("AdditionalApprovalProcess"))
                End If

                child.ckbActive.Checked = (row1("Active"))
                child.CheckBox1.Checked = (row1("RequirePlatformApproval"))
                child.CheckBox2.Checked = (row1("AdditionalApproval"))
                child.CheckBox3.Checked = (row1("OfferSMA"))
                child.CheckBox4.Checked = (row1("OfferUMA"))
                child.CheckBox5.Checked = (row1("OfferMF"))
                child.CheckBox6.Checked = (row1("OfferUIT"))
                child.CheckBox7.Checked = (row1("MirrorPlatformDriver"))
                child.ckbCustomFee.Checked = (row1("CustomFee"))

                Call child.LoadChangesDetection()

            Catch ex As Exception
                MsgBox(ex.Message, MsgBoxStyle.Critical, "ERROR")
            End Try
        End If
    End Sub

    Private Sub DataGridView1_CellContentDoubleClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView1.CellContentDoubleClick
        If Permissions.MAPEditWLPlatform.Checked Then
            Call LoadWLEdit()
        Else
            MsgBox("You do not have permissions to edit this record.", MsgBoxStyle.Critical, "Insufficient Permissions")
        End If
    End Sub
End Class