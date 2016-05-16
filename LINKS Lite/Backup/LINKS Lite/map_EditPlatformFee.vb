﻿Public Class map_EditPlatformFee

    Private Sub map_EditPlatformFee_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Call LoadPlatforms()
        Call LoadProducts()
        Call LoadWLPlatforms()
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

            Dim strSQL As String = "SELECT ID, WLName FROM map_PlatformsWL WHERE Active = -1"
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

    Public Sub LoadProducts()
        Try

            Dim conn As New OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & My.Settings.dbloc & "';Jet OLEDB:Database Password='" & My.Settings.dbpass & "';")

            Dim strSQL As String = "SELECT ID, ProductName FROM map_Products ORDER BY ProductName"
            Dim da As New OleDb.OleDbDataAdapter(strSQL, conn)
            Dim ds As New DataSet
            da.Fill(ds, "Users")

            With cboProduct
                .DataSource = ds.Tables("Users")
                .DisplayMember = "ProductName"
                .ValueMember = "ID"
                .SelectedIndex = 0
            End With

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "ERROR")
        End Try
    End Sub

    Dim strCurrency As String = ""
    Dim acceptableKey As Boolean = False

    Dim strCurrency1 As String = ""
    Dim acceptableKey1 As Boolean = False

    Private Sub bp1_KeyDown(ByVal sender1 As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles BP1.KeyDown
        If (e.KeyCode >= Keys.D0 And e.KeyCode <= Keys.D9) OrElse (e.KeyCode >= Keys.NumPad0 And e.KeyCode <= Keys.NumPad9) OrElse e.KeyCode = Keys.Back Then
            acceptableKey = True
        Else
            acceptableKey = False
        End If
    End Sub

    Private Sub bp1_KeyPress(ByVal sender1 As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles BP1.KeyPress
        ' Check for the flag being set in the KeyDown event.
        If acceptableKey = False Then
            ' Stop the character from being entered into the control since it is non-numerical.
            e.Handled = True
            Return
        Else
            If e.KeyChar = Convert.ToChar(Keys.Back) Then
                If strCurrency.Length > 0 Then
                    strCurrency = strCurrency.Substring(0, strCurrency.Length - 1)
                End If
            Else
                strCurrency = strCurrency & e.KeyChar
            End If

            If strCurrency.Length = 0 Then
                BP1.Text = ""
            ElseIf strCurrency.Length = 1 Then
                BP1.Text = "0.0" & strCurrency
            ElseIf strCurrency.Length = 2 Then
                BP1.Text = "0." & strCurrency
            ElseIf strCurrency.Length > 2 Then
                BP1.Text = strCurrency.Substring(0, strCurrency.Length - 2) & "." & strCurrency.Substring(strCurrency.Length - 2)
            End If
            BP1.Select(BP1.Text.Length, 0)

        End If
        e.Handled = True
    End Sub

    Private Sub bp2_KeyDown(ByVal sender As Object, ByVal e1 As System.Windows.Forms.KeyEventArgs) Handles BP2.KeyDown
        If (e1.KeyCode >= Keys.D0 And e1.KeyCode <= Keys.D9) OrElse (e1.KeyCode >= Keys.NumPad0 And e1.KeyCode <= Keys.NumPad9) OrElse e1.KeyCode = Keys.Back Then
            acceptableKey1 = True
        Else
            acceptableKey1 = False
        End If
    End Sub

    Private Sub bp2_KeyPress(ByVal sender As Object, ByVal e1 As System.Windows.Forms.KeyPressEventArgs) Handles BP2.KeyPress
        ' Check for the flag being set in the KeyDown event.
        If acceptableKey1 = False Then
            ' Stop the character from being entered into the control since it is non-numerical.
            e1.Handled = True
            Return
        Else
            If e1.KeyChar = Convert.ToChar(Keys.Back) Then
                If strCurrency1.Length > 0 Then
                    strCurrency1 = strCurrency1.Substring(0, strCurrency1.Length - 1)
                End If
            Else
                strCurrency1 = strCurrency1 & e1.KeyChar
            End If

            If strCurrency1.Length = 0 Then
                BP2.Text = ""
            ElseIf strCurrency1.Length = 1 Then
                BP2.Text = "0.0" & strCurrency1
            ElseIf strCurrency1.Length = 2 Then
                BP2.Text = "0." & strCurrency1
            ElseIf strCurrency1.Length > 2 Then
                BP2.Text = strCurrency1.Substring(0, strCurrency1.Length - 2) & "." & strCurrency1.Substring(strCurrency1.Length - 2)
            End If
            BP2.Select(BP2.Text.Length, 0)

        End If
        e1.Handled = True
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        If ID.Text = "NEW" Then
            Call CheckForDupe()
        Else
            Call SaveOld()
        End If
    End Sub

    Public Sub CheckForDupe()
        'check for duplicate records
        Dim Mycn As OleDb.OleDbConnection
        Dim SQLstring As String

        Try
            Mycn = New OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & My.Settings.dbloc & "';Jet OLEDB:Database Password='" & My.Settings.dbpass & "';")
            Mycn.Open()
            If ckbWL.Checked Then
                SQLstring = "SELECT * FROM map_PlatformFees WHERE [ProductID] = " & cboProduct.SelectedValue & " AND [WLID] = " & cboWLPlatform.SelectedValue & " AND [PlatformID] = " & cboPlatform.SelectedValue & " AND BreakPointFrom = " & BP1.Text & " AND BreakPointTo = " & BP2.Text & " AND Active = -1"
            Else
                SQLstring = "SELECT * FROM map_PlatformFees WHERE [ProductID] = " & cboProduct.SelectedValue & " AND [PlatformID] = " & cboPlatform.SelectedValue & " AND BreakPointFrom = " & BP1.Text & " AND BreakPointTo = " & BP2.Text & " AND Active = -1"
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
                MsgBox("A record already exsists for that fee schedule.", MsgBoxStyle.Information, "DUPLICATE RECORD DETECTED")
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

            Dim rfee As Double
            If IsDBNull(RepFee.Text) Or RepFee.Text = "" Then
                rfee = 0
            Else
                rfee = RepFee.Text
            End If

            If ckbWL.Checked Then
                SQLstr = "INSERT INTO map_PlatformFees([BreakPointFrom], [ProductID], [PlatformID], [WLID], [BreakPointTo], [Fee], [MaxRepFee], [Active])" & _
                " VALUES('" & BP1.Text & "'," & cboProduct.SelectedValue & "," & cboPlatform.SelectedValue & "," & cboWLPlatform.SelectedValue & ",'" & BP2.Text & "'," & Fee.Text & "," & rfee & ",-1)"
            Else
                SQLstr = "INSERT INTO map_PlatformFees([BreakPointFrom], [ProductID], [PlatformID], [BreakPointTo], [Fee], [MaxRepFee], [Active])" & _
                " VALUES('" & BP1.Text & "'," & cboProduct.SelectedValue & "," & cboPlatform.SelectedValue & ",'" & BP2.Text & "'," & Fee.Text & "," & rfee & ",-1)"
            End If

            Command = New OleDb.OleDbCommand(SQLstr, Mycn)
            Command.ExecuteNonQuery()

            Mycn.Close()

            MsgBox("Record Saved.")
            Me.Close()

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

            Dim rfee As Double
            If IsDBNull(RepFee.Text) Or RepFee.Text = "" Then
                rfee = 0
            Else
                rfee = RepFee.Text
            End If

            If ckbWL.Checked Then
                SQLstr = "Update map_PlatformFees SET [ProductID] = " & cboProduct.SelectedValue & ", [PlatformID] = " & cboPlatform.SelectedValue & ", [WLID] = " & cboWLPlatform.SelectedValue & ", [BreakPointFrom] = '" & BP1.Text & "', [BreakPointTo] = '" & BP2.Text & "', [Fee] = " & Fee.Text & ", [MaxRepFee] = " & rfee & " WHERE ID = " & ID.Text
            Else
                SQLstr = "Update map_PlatformFees SET [ProductID] = " & cboProduct.SelectedValue & ", [PlatformID] = " & cboPlatform.SelectedValue & ", [BreakPointFrom] = '" & BP1.Text & "', [BreakPointTo] = '" & BP2.Text & "', [Fee] = " & Fee.Text & ", [MaxRepFee] = " & rfee & " WHERE ID = " & ID.Text
            End If

            Command = New OleDb.OleDbCommand(SQLstr, Mycn)
            Command.ExecuteNonQuery()

            Mycn.Close()

            MsgBox("Record Updated.")

            Me.Close()

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "ERROR")
        End Try
    End Sub

    Private Sub Timer1_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer1.Tick
        If ckbActive.Checked Then
            Label9.Visible = False
            Button2.Text = "Delete"
            Button1.Enabled = True
        Else
            Label9.Visible = True
            Button2.Text = "Un-Delete"
            Button1.Enabled = False
        End If

        If ckbWL.Checked Then
            cboWLPlatform.Visible = True
        Else
            cboWLPlatform.Visible = False
        End If
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        If ID.Text = "NEW" Then
            'do nothing
        Else
            Call DeleteFeeRecord()
        End If
    End Sub

    Public Sub DeleteFeeRecord()
        Dim Mycn As OleDb.OleDbConnection
        Dim Command As OleDb.OleDbCommand
        Dim SQLstr As String
        Try
            Mycn = New OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & My.Settings.dbloc & "';Jet OLEDB:Database Password='" & My.Settings.dbpass & "';")
            Mycn.Open()

            If ckbActive.Checked Then
                SQLstr = "Update map_PlatformFees SET [Active] = False WHERE ID = " & ID.Text
                ckbActive.Checked = False
            Else
                SQLstr = "Update map_PlatformFees SET [Active] = True WHERE ID = " & ID.Text
                ckbActive.Checked = True
            End If
            Command = New OleDb.OleDbCommand(SQLstr, Mycn)
            Command.ExecuteNonQuery()

            Mycn.Close()

            MsgBox("Record Updated.")

            Me.Close()

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "ERROR")
        End Try
    End Sub

    Private Sub ckbWL_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ckbWL.CheckedChanged

    End Sub
End Class