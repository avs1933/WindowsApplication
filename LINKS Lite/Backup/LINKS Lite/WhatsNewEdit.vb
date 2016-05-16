Imports Microsoft.Office.Interop

Public Class WhatsNewEdit

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSave.Click
        If txtID.Text = "NEW" Then
            Call SaveNew()
        Else
            Call SaveOld()
        End If
    End Sub

    Public Sub SaveNew()
        Dim Mycn As OleDb.OleDbConnection
        Dim Command As OleDb.OleDbCommand
        Dim SQLstr As String

        Try
            Mycn = New OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & My.Settings.roledbloc & "';Jet OLEDB:Database Password='" & My.Settings.dbpass & "';")
            Mycn.Open()

            Dim CMade As String
            CMade = rtbChanges.Text
            CMade = Replace(CMade, "'", "''")

            Dim DNote As String
            DNote = rtbDevNotes.Text
            DNote = Replace(DNote, "'", "''")

            Dim ver As String
            If txtVersion.Text = "" Or IsDBNull(txtVersion) Then
                ver = "NA"
            Else
                ver = txtVersion.Text
            End If

            Dim frm As String
            If txtForm.Text = "" Or IsDBNull(txtForm) Then
                frm = "NA"
            Else
                frm = txtForm.Text
            End If

            Dim tbl As String
            If txtTable.Text = "" Or IsDBNull(txtForm) Then
                tbl = "NA"
            Else
                tbl = txtTable.Text
            End If

            Dim fnct As String
            If txtFunction.Text = "" Or IsDBNull(txtFunction) Then
                fnct = "NA"
            Else
                fnct = txtFunction.Text
                fnct = Replace(fnct, "'", "''")
            End If

            SQLstr = "INSERT INTO sys_SoftwareUpdates([VersionNumber],[FormName], [Table/QueryName],[MainFunction],[ChangesMade],[DevNotes],[MadeBy],[DateOfRelease])" & _
            "VALUES('" & ver & "','" & frm & "','" & tbl & "','" & fnct & "','" & CMade & "','" & DNote & "'," & My.Settings.userid & ",#" & dtpRelease.Text & "#)"

            Command = New OleDb.OleDbCommand(SQLstr, Mycn)
            Command.ExecuteNonQuery()

            Mycn.Close()

            Call ResetFields()
            Call LoadGrid()

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "ERROR")
        End Try
    End Sub

    Public Sub SaveOld()
        Dim Mycn As OleDb.OleDbConnection
        Dim Command As OleDb.OleDbCommand
        Dim SQLstr As String

        Try
            Mycn = New OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & My.Settings.roledbloc & "';Jet OLEDB:Database Password='" & My.Settings.dbpass & "';")
            Mycn.Open()

            Dim CMade As String
            CMade = rtbChanges.Text
            CMade = Replace(CMade, "'", "''")

            Dim DNote As String
            DNote = rtbDevNotes.Text
            DNote = Replace(DNote, "'", "''")

            Dim ver As String
            If txtVersion.Text = "" Or IsDBNull(txtVersion) Then
                ver = "NA"
            Else
                ver = txtVersion.Text
            End If

            Dim frm As String
            If txtForm.Text = "" Or IsDBNull(txtForm) Then
                frm = "NA"
            Else
                frm = txtForm.Text
            End If

            Dim tbl As String
            If txtTable.Text = "" Or IsDBNull(txtForm) Then
                tbl = "NA"
            Else
                tbl = txtTable.Text
            End If

            Dim fnct As String
            If txtFunction.Text = "" Or IsDBNull(txtFunction) Then
                fnct = "NA"
            Else
                fnct = txtFunction.Text
                fnct = Replace(fnct, "'", "''")
            End If

            SQLstr = "UPDATE sys_SoftwareUpdates SET [VersionNumber] = '" & ver & "',[FormName] = '" & frm & "', [Table/QueryName] = '" & tbl & "',[MainFunction] = '" & fnct & "',[ChangesMade] = '" & CMade & "',[DevNotes] = '" & DNote & "',[MadeBy] = " & My.Settings.userid & ",[DateOfRelease] = #" & dtpRelease.Text & "# WHERE ID = " & txtID.Text

            Command = New OleDb.OleDbCommand(SQLstr, Mycn)
            Command.ExecuteNonQuery()

            Mycn.Close()

            Call LoadGrid()

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "ERROR")
        End Try
    End Sub

    Private Sub btnReset_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnReset.Click
        Call ResetFields()
    End Sub

    Public Sub ResetFields()
        txtVersion.Text = My.Application.Info.Version.Major & "." & My.Application.Info.Version.Minor & "." & My.Application.Info.Version.Build & "." & My.Application.Info.Version.Revision
        txtID.Text = "NEW"
        txtFunction.Text = ""
        txtForm.Text = ""
        txtTable.Text = ""
        rtbChanges.Text = ""
        rtbDevNotes.Text = ""
    End Sub

    Private Sub WhatsNewEdit_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Call LoadVersionCbo()
        txtVersion.Text = My.Application.Info.Version.Major & "." & My.Application.Info.Version.Minor & "." & My.Application.Info.Version.Build & "." & My.Application.Info.Version.Revision
    End Sub

    Private Sub btnDelete_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnDelete.Click
        If txtID.Text = "NEW" Then
            MsgBox("You cannot delete an unsaved record.", MsgBoxStyle.Information, "Cant Delete")
        Else
            Dim ir As MsgBoxResult
            ir = MsgBox("Are you sure you want to delete this record?", MsgBoxStyle.YesNo + MsgBoxStyle.Critical, "Confirm Delete")
            If ir = MsgBoxResult.Yes Then
                Call DeleteRecord()
                Call ResetFields()
            Else

            End If
        End If
    End Sub

    Public Sub DeleteRecord()
        Dim Mycn As OleDb.OleDbConnection
        Dim Command As OleDb.OleDbCommand
        Dim SQLstr As String

        Try
            Mycn = New OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & My.Settings.roledbloc & "';Jet OLEDB:Database Password='" & My.Settings.dbpass & "';")
            Mycn.Open()


            SQLstr = "DELETE * FROM sys_SoftwareUpdates WHERE ID = " & txtID.Text

            Command = New OleDb.OleDbCommand(SQLstr, Mycn)
            Command.ExecuteNonQuery()

            Mycn.Close()

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "ERROR")
        End Try
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Call LoadGrid()
    End Sub

    Public Sub LoadVersionCbo()
        Try
            Dim conn As New OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & My.Settings.dbloc & "';Jet OLEDB:Database Password='" & My.Settings.dbpass & "';")
            Dim strSQL As String = "SELECT Distinct VersionNumber FROM sys_SoftwareUpdates"
            Dim da As New OleDb.OleDbDataAdapter(strSQL, conn)
            Dim ds As New DataSet
            da.Fill(ds, "Users")
            Dim dt As DataTable = ds.Tables("Users")

            If dt.Rows.Count > 0 Then
                With cboSVersion
                    .DataSource = ds.Tables("Users")
                    .DisplayMember = "VersionNumber"
                    .ValueMember = "VersionNumber"
                    .SelectedIndex = 0
                End With
                'btnMapSave.Enabled = True
            Else
                ckbVersion.Checked = False
                cboSVersion.Text = "NOTHING TO LOAD."
                cboSVersion.Enabled = False
                'btnMapSave.Enabled = False
            End If

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "ERROR")
        End Try
    End Sub

    Public Sub LoadGrid()
        Try

            Dim strSQL As String
            Dim conn As New OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & My.Settings.dbloc & "';Jet OLEDB:Database Password='" & My.Settings.dbpass & "';")

            strSQL = "SELECT sys_SoftwareUpdates.ID, sys_SoftwareUpdates.VersionNumber AS [Version Number], sys_SoftwareUpdates.FormName AS [Form Name], sys_SoftwareUpdates.[Table/QueryName] AS [Table], sys_SoftwareUpdates.MainFunction AS Function, sys_SoftwareUpdates.ChangesMade AS [Changes Made], sys_SoftwareUpdates.DevNotes AS [Dev Notes], sys_Users.FullName AS [Made By], sys_SoftwareUpdates.DateOfRelease AS [Date of Relese]" & _
            " FROM sys_SoftwareUpdates INNER JOIN sys_Users ON sys_SoftwareUpdates.MadeBy = sys_Users.ID"

            If ckbFormName.Checked Or ckbFunction.Checked Or ckbTable.Checked Or ckbVersion.Checked Then
                strSQL = strSQL & " WHERE"
            Else

            End If

            If ckbVersion.Checked Then
                strSQL = strSQL & " VersionNumber = '" & cboSVersion.Text & "'"
            Else

            End If

            If ckbFormName.Checked Then
                If ckbVersion.Checked Then
                    strSQL = strSQL & ", FormName = '" & txtSForm.Text & "'"
                Else
                    strSQL = strSQL & " FormName = '" & txtSForm.Text & "'"
                End If
            Else

            End If

            If ckbFunction.Checked Then
                If ckbVersion.Checked Or ckbFormName.Checked Then
                    strSQL = strSQL & ", MainFunction = '" & txtSFunction.Text & "'"
                Else
                    strSQL = strSQL & " MainFunction = '" & txtSFunction.Text & "'"
                End If
            Else

            End If

            If ckbTable.Checked Then
                If ckbVersion.Checked Or ckbFormName.Checked Or ckbFunction.Checked Then
                    strSQL = strSQL & ", Table/QueryName = '" & txtSTable.Text & "'"
                Else
                    strSQL = strSQL & " Table/QueryName = '" & txtSTable.Text & "'"
                End If
            Else

            End If

            strSQL = strSQL & " GROUP BY sys_SoftwareUpdates.ID, sys_SoftwareUpdates.VersionNumber, sys_SoftwareUpdates.FormName, sys_SoftwareUpdates.[Table/QueryName], sys_SoftwareUpdates.MainFunction, sys_SoftwareUpdates.ChangesMade, sys_SoftwareUpdates.DevNotes, sys_Users.FullName, sys_SoftwareUpdates.DateOfRelease;"

            Dim da As New OleDb.OleDbDataAdapter(strSQL, conn)
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

    Private Sub DataGridView1_CellDoubleClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView1.CellDoubleClick
        If DataGridView1.RowCount = 0 Then
            'do nothing
        Else
            txtID.Text = DataGridView1.SelectedCells(0).Value
            txtVersion.Text = DataGridView1.SelectedCells(1).Value
            txtForm.Text = DataGridView1.SelectedCells(2).Value
            txtTable.Text = DataGridView1.SelectedCells(3).Value
            txtFunction.Text = DataGridView1.SelectedCells(4).Value
            rtbChanges.Text = DataGridView1.SelectedCells(5).Value
            rtbDevNotes.Text = DataGridView1.SelectedCells(6).Value
            dtpRelease.Text = DataGridView1.SelectedCells(8).Value
        End If
    End Sub

    Private Sub ckbVersion_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ckbVersion.CheckedChanged
        If ckbVersion.Checked Then
            cboSVersion.Enabled = True
            Call LoadVersionCbo()
        Else
            cboSVersion.Enabled = False
        End If
    End Sub

    Private Sub ckbFormName_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ckbFormName.CheckedChanged
        If ckbFormName.Checked Then
            txtSForm.Enabled = True
        Else
            txtSForm.Text = ""
            txtSForm.Enabled = False
        End If
    End Sub

    Private Sub ckbTable_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ckbTable.CheckedChanged
        If ckbTable.Checked Then
            txtSTable.Enabled = True
        Else
            txtSTable.Text = ""
            txtSTable.Enabled = False
        End If
    End Sub

    Private Sub ckbFunction_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ckbFunction.CheckedChanged
        If ckbFunction.Checked Then
            txtSFunction.Enabled = True
        Else
            txtSFunction.Text = ""
            txtSFunction.Enabled = False
        End If
    End Sub

    Private Sub Button5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button5.Click
        If rtbChanges.Text.Length > 0 Then
            Dim wordapp As New Word.Application
            wordapp.Visible = False
            Dim doc As Word.Document = wordapp.Documents.Add
            Dim range As Word.Range
            range = doc.Range
            range.Text = rtbChanges.Text
            doc.Activate()
            doc.CheckSpelling()
            Dim chars() As Char = {CType(vbCr, Char), CType(vbLf, Char)}
            rtbChanges.Text = doc.Range().Text.Trim(chars)
            doc.Close(SaveChanges:=False)
            wordapp.Quit()
            MsgBox("Spell Check Completed.", MsgBoxStyle.Information, "Done")
        End If
    End Sub

    Private Sub Button2_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        If rtbDevNotes.Text.Length > 0 Then
            Dim wordapp As New Word.Application
            wordapp.Visible = False
            Dim doc As Word.Document = wordapp.Documents.Add
            Dim range As Word.Range
            range = doc.Range
            range.Text = rtbDevNotes.Text
            doc.Activate()
            doc.CheckSpelling()
            Dim chars() As Char = {CType(vbCr, Char), CType(vbLf, Char)}
            rtbDevNotes.Text = doc.Range().Text.Trim(chars)
            doc.Close(SaveChanges:=False)
            wordapp.Quit()
            MsgBox("Spell Check Completed.", MsgBoxStyle.Information, "Done")
        End If
    End Sub
End Class