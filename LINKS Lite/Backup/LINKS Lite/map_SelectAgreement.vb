Public Class map_SelectAgreement

    Private Sub map_SelectAgreement_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Call LoadAgreementType()
    End Sub

    Public Sub LoadAgreementType()
        Try

            Dim conn As New OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & My.Settings.dbloc & "';Jet OLEDB:Database Password='" & My.Settings.dbpass & "';")

            Dim strSQL As String = "SELECT ID, TypeName FROM map_AgreementType ORDER BY TypeName"
            Dim da As New OleDb.OleDbDataAdapter(strSQL, conn)
            Dim ds As New DataSet
            da.Fill(ds, "Users")

            With cboAType
                .DataSource = ds.Tables("Users")
                .DisplayMember = "TypeName"
                .ValueMember = "ID"
                .SelectedIndex = 0
            End With

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "ERROR")
        End Try
    End Sub

    Private Sub CheckBox1_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckBox1.CheckedChanged
        If CheckBox1.Checked Then
            cboAType.Enabled = True
        Else
            cboAType.Enabled = False
        End If
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Call LoadTbl()
    End Sub

    Public Sub LoadTbl()
        Try
            Dim conn As New OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & My.Settings.dbloc & "';Jet OLEDB:Database Password='" & My.Settings.dbpass & "';")

            Dim SqlString As String

            If CheckBox2.Checked Then
                'Look for inactive records
                If CheckBox1.Checked Then
                    'Look for product types
                    SqlString = "SELECT map_Agreements.ID, Map_Firms.FirmName As [Firm Name], map_AgreementType.TypeName As [Agreement Type], map_Agreements.VersionNumber As [Version Number], map_Agreements.DateExecuted As [Date Executed], map_Agreements.Active" & _
                    " FROM (map_Agreements INNER JOIN map_AgreementType ON map_Agreements.TypeID = map_AgreementType.ID) INNER JOIN Map_Firms ON map_Agreements.FirmID = Map_Firms.ID" & _
                    " WHERE map_Agreements.Active <> -1 AND Map_Firms.FirmName Like '%" & TextBox1.Text & "%' AND map_Agreements.VersionNumber Like '%" & TextBox2.Text & "%' AND map_Agreements.TypeID = " & cboAType.SelectedValue
                Else
                    SqlString = "SELECT map_Agreements.ID, Map_Firms.FirmName As [Firm Name], map_AgreementType.TypeName As [Agreement Type], map_Agreements.VersionNumber As [Version Number], map_Agreements.DateExecuted As [Date Executed], map_Agreements.Active" & _
                    " FROM (map_Agreements INNER JOIN map_AgreementType ON map_Agreements.TypeID = map_AgreementType.ID) INNER JOIN Map_Firms ON map_Agreements.FirmID = Map_Firms.ID" & _
                    " WHERE map_Agreements.Active <> -1 AND Map_Firms.FirmName Like '%" & TextBox1.Text & "%' AND map_Agreements.VersionNumber Like '%" & TextBox2.Text & "%'"
                End If
            Else
                'Look for active records
                If CheckBox1.Checked Then
                    'Look for product types
                    SqlString = "SELECT map_Agreements.ID, Map_Firms.FirmName As [Firm Name], map_AgreementType.TypeName As [Agreement Type], map_Agreements.VersionNumber As [Version Number], map_Agreements.DateExecuted As [Date Executed], map_Agreements.Active" & _
                    " FROM (map_Agreements INNER JOIN map_AgreementType ON map_Agreements.TypeID = map_AgreementType.ID) INNER JOIN Map_Firms ON map_Agreements.FirmID = Map_Firms.ID" & _
                    " WHERE map_Agreements.Active = -1 AND Map_Firms.FirmName Like '%" & TextBox1.Text & "%' AND map_Agreements.VersionNumber Like '%" & TextBox2.Text & "%' AND map_Agreements.TypeID = " & cboAType.SelectedValue
                Else
                    SqlString = "SELECT map_Agreements.ID, Map_Firms.FirmName As [Firm Name], map_AgreementType.TypeName As [Agreement Type], map_Agreements.VersionNumber As [Version Number], map_Agreements.DateExecuted As [Date Executed], map_Agreements.Active" & _
                    " FROM (map_Agreements INNER JOIN map_AgreementType ON map_Agreements.TypeID = map_AgreementType.ID) INNER JOIN Map_Firms ON map_Agreements.FirmID = Map_Firms.ID" & _
                    " WHERE map_Agreements.Active = -1 AND Map_Firms.FirmName Like '%" & TextBox1.Text & "%' AND map_Agreements.VersionNumber Like '%" & TextBox2.Text & "%'"
                End If
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

    Public Sub LoadEditForm()
        If DataGridView1.RowCount = "0" Then
            'do nothing
        Else
            Dim child As New map_EditAgreement
            child.MdiParent = Home
            child.Show()

            Try
                Dim conn As New OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & My.Settings.dbloc & "';Jet OLEDB:Database Password='" & My.Settings.dbpass & "';")

                Dim strSQL As String = "SELECT * FROM map_Agreements WHERE ID = " & DataGridView1.SelectedCells(0).Value
                Dim queryString As String = String.Format(strSQL)
                Dim cmd As New OleDb.OleDbCommand(queryString, conn)
                Dim da As New OleDb.OleDbDataAdapter(cmd)
                Dim ds As New DataSet

                da.Fill(ds, "User")
                Dim dt As DataTable = ds.Tables("User")
                Dim row1 As DataRow = dt.Rows(0)

                child.ID.Text = (row1("ID"))

                If IsDBNull(row1("Notes")) Then
                    child.RichTextBox1.Text = ""
                Else
                    child.RichTextBox1.Text = (row1("Notes"))
                End If

                If IsDBNull(row1("TypeID")) Then
                    'do nothing
                Else
                    child.cboAType.SelectedValue = (row1("TypeID"))
                End If

                If IsDBNull(row1("FirmID")) Then
                    'do nothing
                Else
                    child.cboFirms.SelectedValue = (row1("FirmID"))
                End If

                If IsDBNull(row1("VersionNumber")) Then
                    child.TextBox2.Text = ""
                Else
                    child.TextBox2.Text = (row1("VersionNumber"))
                End If

                If IsDBNull(row1("DateExecuted")) Then
                    'do nothing
                Else
                    child.DateTimePicker1.Text = (row1("DateExecuted"))
                End If

                child.CheckBox1.Checked = (row1("UseDefaultFee"))
                child.CheckBox2.Checked = (row1("AllInHouseSMAApproved"))
                child.CheckBox3.Checked = (row1("AllOutsideSMAApproved"))
                child.CheckBox4.Checked = (row1("AllMFApproved"))
                child.CheckBox5.Checked = (row1("AllUITApproved"))
                child.CheckBox6.Checked = (row1("Active"))

            Catch ex As Exception
                MsgBox(ex.Message, MsgBoxStyle.Critical, "ERROR")
            End Try
        End If
    End Sub

    Private Sub DataGridView1_CellContentClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick

    End Sub

    Private Sub DataGridView1_CellContentDoubleClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView1.CellContentDoubleClick
        Call LoadEditForm()
    End Sub

    Private Sub EditToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles EditToolStripMenuItem.Click
        Call LoadEditForm()
    End Sub
End Class