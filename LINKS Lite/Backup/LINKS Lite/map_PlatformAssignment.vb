Public Class map_PlatformAssignment

    Private Sub Label2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label2.Click

    End Sub
    Private Sub ComboBox1_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cboPlatform.SelectedIndexChanged

    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        Me.Close()
    End Sub

    Private Sub map_PlatformAssignment_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Call LoadFirms()
        Call LoadPlatforms()
        Call LoadWLPlatforms()
    End Sub

    Public Sub LoadFirms()
        Try
            Dim conn As New OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & My.Settings.dbloc & "';Jet OLEDB:Database Password='" & My.Settings.dbpass & "';")

            Dim strSQL As String = "SELECT Map_Firms.ID, Map_Firms.FirmName" & _
            " FROM(Map_Firms)" & _
            " WHERE(((Map_Firms.Active) = True) And ((Map_Firms.TypeID) = 1 Or (Map_Firms.TypeID) = 2 Or (Map_Firms.TypeID) = 3))" & _
            " GROUP BY Map_Firms.ID, Map_Firms.FirmName;"

            Dim da As New OleDb.OleDbDataAdapter(strSQL, conn)
            Dim ds As New DataSet
            da.Fill(ds, "Users")

            With cboFirm
                .DataSource = ds.Tables("Users")
                .DisplayMember = "FirmName"
                .ValueMember = "ID"
                .SelectedIndex = 0
            End With

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "ERROR")
        End Try
    End Sub

    Public Sub LoadPlatforms()
        Try
            Dim conn As New OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & My.Settings.dbloc & "';Jet OLEDB:Database Password='" & My.Settings.dbpass & "';")

            Dim strSQL As String = "SELECT map_Platforms.ID, Map_Firms.FirmName" & _
            " FROM map_Platforms INNER JOIN Map_Firms ON map_Platforms.PlatformID = Map_Firms.ID" & _
            " WHERE(((map_Platforms.Active) = True))" & _
            " GROUP BY map_Platforms.ID, Map_Firms.FirmName;"

            Dim da As New OleDb.OleDbDataAdapter(strSQL, conn)
            Dim ds As New DataSet
            da.Fill(ds, "Users")

            With cboPlatform
                .DataSource = ds.Tables("Users")
                .DisplayMember = "FirmName"
                .ValueMember = "ID"
                .SelectedIndex = 0
            End With

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "ERROR")
        End Try
    End Sub

    Public Sub LoadWLPlatforms()
        Try
            Dim conn As New OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & My.Settings.dbloc & "';Jet OLEDB:Database Password='" & My.Settings.dbpass & "';")

            Dim strSQL As String = "SELECT map_PlatformsWL.ID, map_PlatformsWL.WLName" & _
            " FROM(map_PlatformsWL)" & _
            " WHERE(((map_PlatformsWL.Active) = True))" & _
            "GROUP BY map_PlatformsWL.ID, map_PlatformsWL.WLName;"

            Dim da As New OleDb.OleDbDataAdapter(strSQL, conn)
            Dim ds As New DataSet
            da.Fill(ds, "Users")

            With cboWLPlatform
                .DataSource = ds.Tables("Users")
                .DisplayMember = "WLName"
                .ValueMember = "ID"
                .SelectedIndex = 0
            End With

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "ERROR")
        End Try
    End Sub

    Private Sub ckbPlatform_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ckbPlatform.CheckedChanged
        Call platformtypecheck()
    End Sub

    Public Sub platformtypecheck()
        If ckbPlatform.Checked = True Then
            ckbWLPlatform.Checked = False
            cboPlatform.Enabled = True
            cboWLPlatform.Enabled = False
        Else
            ckbPlatform.Checked = False
            ckbWLPlatform.Checked = True
            cboWLPlatform.Enabled = True
            cboPlatform.Enabled = False
        End If
    End Sub

    Private Sub ckbWLPlatform_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ckbWLPlatform.CheckedChanged
        Call platformtypecheck()
    End Sub

    Private Sub Timer1_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer1.Tick
        If ckbActive.Checked Then
            Label7.Visible = False
            Button1.Enabled = True
            Button2.Text = "Delete"
        Else
            Label7.Visible = True
            Button1.Enabled = False
            Button2.Text = "Un-Delete"
        End If
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        If ID.Text = "NEW" Then
            MsgBox("You cannot delete an un-saved record.", MsgBoxStyle.Exclamation, "Not Saved")
        Else
            Dim Mycn As OleDb.OleDbConnection
            Dim Command As OleDb.OleDbCommand
            Dim SQLstr As String
            Try
                Mycn = New OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & My.Settings.dbloc & "';Jet OLEDB:Database Password='" & My.Settings.dbpass & "';")
                Mycn.Open()

                If ckbActive.Checked Then
                    SQLstr = "Update map_PlatformAssignments SET Active = False WHERE ID = " & ID.Text
                    ckbActive.Checked = False
                Else
                    SQLstr = "Update map_PlatformAssignments SET Active = -1 WHERE ID = " & ID.Text
                    ckbActive.Checked = True
                End If

                Command = New OleDb.OleDbCommand(SQLstr, Mycn)
                Command.ExecuteNonQuery()

                'Call LoadChangesDetection()

                Mycn.Close()

            Catch ex As Exception
                MsgBox(ex.Message, MsgBoxStyle.Critical, "ERROR")
            End Try
        End If
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        If ID.Text = "NEW" Then
            Call CheckForDupe()
        Else
            Call SaveOld()
        End If
    End Sub

    Public Sub CheckForDupe()

        Dim Mycn As OleDb.OleDbConnection
        Dim SQLstring As String
        Try
            Mycn = New OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & My.Settings.dbloc & "';Jet OLEDB:Database Password='" & My.Settings.dbpass & "';")
            Mycn.Open()
            If ckbPlatform.Checked Then
                SQLstring = "SELECT * FROM Map_PlatformAssignments WHERE Active = True AND FirmID = " & cboFirm.SelectedValue & " AND UsePlatform = True AND PlatformID = " & cboPlatform.SelectedValue
            Else
                SQLstring = "SELECT * FROM Map_PlatformAssignments WHERE Active = True AND FirmID = " & cboFirm.SelectedValue & " AND UseWL = True AND WLID = " & cboWLPlatform.SelectedValue
            End If

            Dim queryString As String = String.Format(SQLstring)
            Dim cmd As New OleDb.OleDbCommand(queryString, Mycn)
            Dim da As New OleDb.OleDbDataAdapter(cmd)
            Dim ds As New DataSet

            da.Fill(ds, "User")
            Dim dt As DataTable = ds.Tables("User")
            If dt.Rows.Count = 0 Then
                Call SaveNew()
            Else
                MsgBox("Platform is already associated with the selected firm.", MsgBoxStyle.Information, "DUPLICATE RECORD DETECTED")
            End If

        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub

    Public Sub SaveNew()
        Dim Mycn As OleDb.OleDbConnection
        Dim Command As OleDb.OleDbCommand
        Dim SQLstr As String
        Try
            Mycn = New OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & My.Settings.dbloc & "';Jet OLEDB:Database Password='" & My.Settings.dbpass & "';")
            Mycn.Open()

            If ckbPlatform.Checked Then
                SQLstr = "INSERT INTO map_PlatformAssignments(FirmID, UsePlatform, PlatformID, Active)" & _
                "VALUES(" & cboFirm.SelectedValue & "," & ckbPlatform.CheckState & "," & cboPlatform.SelectedValue & ", -1)"
            Else
                SQLstr = "INSERT INTO map_PlatformAssignments(FirmID, UseWL, WLID, Active)" & _
                "VALUES(" & cboFirm.SelectedValue & "," & ckbWLPlatform.CheckState & "," & cboWLPlatform.SelectedValue & ", -1)"
            End If


            Command = New OleDb.OleDbCommand(SQLstr, Mycn)
            Command.ExecuteNonQuery()

            Mycn.Close()

            ckbPlatform.Enabled = False
            ckbWLPlatform.Enabled = False

            MsgBox("Record Saved.")

            Call PullID()
            'Call LoadChangesDetection()

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "ERROR")
        End Try
    End Sub

    Public Sub PullID()
        Try
            Dim conn As New OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & My.Settings.dbloc & "';Jet OLEDB:Database Password='" & My.Settings.dbpass & "';")
            Dim strSQL As String
            If ckbPlatform.Checked Then
                strSQL = "SELECT Top 1 ID FROM map_PlatformAssignments WHERE FirmID = " & cboFirm.SelectedValue & " AND PlatformID = " & cboPlatform.SelectedValue & " AND Active = -1 ORDER BY ID DESC"
            Else
                strSQL = "SELECT Top 1 ID FROM map_PlatformAssignments WHERE FirmID = " & cboFirm.SelectedValue & " AND WLID = " & cboWLPlatform.SelectedValue & " AND Active = -1 ORDER BY ID DESC"
            End If

            Dim queryString As String = String.Format(strSQL)
            Dim cmd As New OleDb.OleDbCommand(queryString, conn)
            Dim da As New OleDb.OleDbDataAdapter(cmd)
            Dim ds As New DataSet

            da.Fill(ds, "User")
            Dim dt As DataTable = ds.Tables("User")

            Dim row1 As DataRow = dt.Rows(0)

            If IsDBNull(row1("ID")) Then
                ID.Text = "NEW"
            Else
                ID.Text = (row1("ID"))
            End If

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "ERROR")
        End Try

    End Sub

    Public Sub SaveOld()
        Dim Mycn As OleDb.OleDbConnection
        Dim Command As OleDb.OleDbCommand
        Dim SQLstr As String
        Try
            Mycn = New OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & My.Settings.dbloc & "';Jet OLEDB:Database Password='" & My.Settings.dbpass & "';")
            Mycn.Open()

            If ckbPlatform.Checked Then
                SQLstr = "Update map_PlatformAssignments SET FirmID = " & cboFirm.SelectedValue & ", UsePlatform = " & ckbPlatform.CheckState & ", PlatformID = " & cboPlatform.SelectedValue & ", Active = " & ckbActive.CheckState & " WHERE ID = " & ID.Text
            Else
                SQLstr = "Update map_PlatformAssignments SET FirmID = " & cboFirm.SelectedValue & ", UseWL = " & ckbWLPlatform.CheckState & ", WLID = " & cboWLPlatform.SelectedValue & ", Active = " & ckbActive.CheckState & " WHERE ID = " & ID.Text
            End If

            Command = New OleDb.OleDbCommand(SQLstr, Mycn)
            Command.ExecuteNonQuery()

            Mycn.Close()

            ckbPlatform.Enabled = False
            ckbWLPlatform.Enabled = False

            MsgBox("Record Updated.")
            'Call LoadChangesDetection()

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "ERROR")
        End Try
    End Sub
End Class