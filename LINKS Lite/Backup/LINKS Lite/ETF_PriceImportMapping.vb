Public Class ETF_PriceImportMapping

    Private Sub ETF_PriceImportMapping_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Call LoadSymbolFrom()
        Call LoadSymbolTo()
    End Sub

    Public Sub LoadSymbolFrom()
        Try

            Dim conn As New OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & My.Settings.dbloc & "';Jet OLEDB:Database Password='" & My.Settings.dbpass & "';")

            Dim strSQL As String = "SELECT * FROM mdb_SecurityWithSecType"
            Dim da As New OleDb.OleDbDataAdapter(strSQL, conn)
            Dim ds As New DataSet
            da.Fill(ds, "Users")

            With cboSymbolFrom
                .DataSource = ds.Tables("Users")
                .DisplayMember = "Symbol"
                .ValueMember = "SecurityID"
                .SelectedIndex = 0
            End With

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "ERROR")
        End Try
    End Sub

    Public Sub LoadSymbolTo()
        Try

            Dim conn As New OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & My.Settings.dbloc & "';Jet OLEDB:Database Password='" & My.Settings.dbpass & "';")

            Dim strSQL As String = "SELECT * FROM mdb_SecurityWithSecType"
            Dim da As New OleDb.OleDbDataAdapter(strSQL, conn)
            Dim ds As New DataSet
            da.Fill(ds, "Users")

            With cboSymbolTo
                .DataSource = ds.Tables("Users")
                .DisplayMember = "Symbol"
                .ValueMember = "SecurityID"
                .SelectedIndex = 0
            End With

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "ERROR")
        End Try
    End Sub

    Private Sub cboSymbolFrom_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cboSymbolFrom.SelectedIndexChanged

        If cboSymbolFrom.ValueMember.Count = 0 Then

        Else
            Try
                Dim conn As New OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & My.Settings.dbloc & "';Jet OLEDB:Database Password='" & My.Settings.dbpass & "';")

                Dim strSQL As String = "SELECT * FROM dbo_vQbRowDefSecurity WHERE SecurityID = " & cboSymbolFrom.SelectedValue
                Dim queryString As String = String.Format(strSQL)
                Dim cmd As New OleDb.OleDbCommand(queryString, conn)
                Dim da As New OleDb.OleDbDataAdapter(cmd)
                Dim ds As New DataSet

                da.Fill(ds, "User")
                Dim dt As DataTable = ds.Tables("User")

                Dim row1 As DataRow = dt.Rows(0)


                SecTypeFrom.Text = (row1("SecType"))
                lblFrom.Text = (row1("FullName"))

            Catch ex As Exception
                MsgBox(ex.Message, MsgBoxStyle.Critical, "ERROR")
            End Try

        End If
    End Sub

    Private Sub cboSymbolTo_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cboSymbolTo.SelectedIndexChanged
        'If IsDBNull(cboSymbolTo.SelectedValue) Then
        If cboSymbolTo.ValueMember.Count = 0 Then

        Else
            Try
                Dim conn As New OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & My.Settings.dbloc & "';Jet OLEDB:Database Password='" & My.Settings.dbpass & "';")

                Dim strSQL As String = "SELECT * FROM dbo_vQbRowDefSecurity WHERE SecurityID = " & cboSymbolTo.SelectedValue
                Dim queryString As String = String.Format(strSQL)
                Dim cmd As New OleDb.OleDbCommand(queryString, conn)
                Dim da As New OleDb.OleDbDataAdapter(cmd)
                Dim ds As New DataSet

                da.Fill(ds, "User")
                Dim dt As DataTable = ds.Tables("User")

                Dim row1 As DataRow = dt.Rows(0)


                SecTypeTo.Text = (row1("SecType"))
                lblTo.Text = (row1("FullName"))

            Catch ex As Exception
                MsgBox(ex.Message, MsgBoxStyle.Critical, "ERROR")
            End Try

        End If
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        If lblFrom.Text = lblTo.Text Then
            If ID.Text = "NEW" Then
                Call CheckForDupe()
            Else
                Call SaveOld()
            End If
        Else
            Dim ir As MsgBoxResult
            ir = MsgBox("The security names do not match.  Would you like to save anyway?", MsgBoxStyle.YesNo + MsgBoxStyle.Information, "Data does not match")
            If ir = MsgBoxResult.Yes Then
                If ID.Text = "NEW" Then
                    Call CheckForDupe()
                Else
                    Call SaveOld()
                End If
            Else
                'do nothing
            End If
        End If
    End Sub

    Public Sub SaveNew()
        Dim Mycn As OleDb.OleDbConnection
        Dim Command As OleDb.OleDbCommand
        Dim SQLstr As String
        Try
            Mycn = New OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & My.Settings.dbloc & "';Jet OLEDB:Database Password='" & My.Settings.dbpass & "';")
            Mycn.Open()

            SQLstr = "INSERT INTO mdb_SymbolMapping(APXSecType, APXSymbol, NewAPXSecType, NewAPXSymbol, Active)" & _
            "VALUES('" & SecTypeFrom.Text & "'," & cboSymbolFrom.SelectedValue & ",'" & SecTypeTo.Text & "'," & cboSymbolTo.SelectedValue & ", -1)"

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

            SQLstr = "Update mdb_SymbolMapping SET APXSecType = '" & SecTypeFrom.Text & "', APXSymbol = " & cboSymbolFrom.SelectedValue & ", NewAPXSecType = '" & SecTypeTo.Text & "', NewAPXSymbol = " & cboSymbolTo.SelectedValue & " WHERE ID = " & ID.Text

            Command = New OleDb.OleDbCommand(SQLstr, Mycn)
            Command.ExecuteNonQuery()

            Mycn.Close()

            MsgBox("Record Updated.")
            Me.Close()

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "ERROR")
        End Try
    End Sub

    Public Sub CheckForDupe()
        Dim Mycn As OleDb.OleDbConnection
            Dim SQLstring As String
            Try
                Mycn = New OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & My.Settings.dbloc & "';Jet OLEDB:Database Password='" & My.Settings.dbpass & "';")
                Mycn.Open()

            SQLstring = "SELECT * FROM mdb_SymbolMapping WHERE APXSymbol = " & cboSymbolFrom.SelectedValue & " AND NewAPXSymbol = " & cboSymbolTo.SelectedValue

                Dim queryString As String = String.Format(SQLstring)
                Dim cmd As New OleDb.OleDbCommand(queryString, Mycn)
                Dim da As New OleDb.OleDbDataAdapter(cmd)
                Dim ds As New DataSet

                da.Fill(ds, "User")
                Dim dt As DataTable = ds.Tables("User")
                If dt.Rows.Count = 0 Then
                    Call SaveNew()
                Else
                MsgBox("A record already exsists for that Mapping.", MsgBoxStyle.Information, "DUPLICATE RECORD DETECTED")
                End If

            Catch ex As Exception
                MsgBox(ex.Message)
            End Try
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        Dim ir As New MsgBoxResult
        ir = MsgBox("Are you sure you want to delete this record?", MsgBoxStyle.YesNo + MsgBoxStyle.Information, "Confirm Delete?")
        If ir = MsgBoxResult.Yes Then
            Call DeleteRecord()
        Else
            'do nothing
        End If
    End Sub

    Public Sub DeleteRecord()
        Dim Mycn As OleDb.OleDbConnection
        Dim Command As OleDb.OleDbCommand
        Dim SQLstr As String
        Try
            Mycn = New OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & My.Settings.dbloc & "';Jet OLEDB:Database Password='" & My.Settings.dbpass & "';")
            Mycn.Open()

            If Active.Checked Then
                SQLstr = "Update Map_Firms SET Active = False WHERE ID = " & ID.Text
                Active.Checked = False

            Else
                SQLstr = "Update Map_Firms SET Active = True WHERE ID = " & ID.Text
                Active.Checked = True
            End If

            Command = New OleDb.OleDbCommand(SQLstr, Mycn)
            Command.ExecuteNonQuery()

            Mycn.Close()

            'MsgBox("Record Updated.")
            'Me.Close()

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "ERROR")
        End Try
    End Sub

    Private Sub Timer1_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer1.Tick
        If Active.Checked Then
            Button1.Enabled = True
            Label9.Visible = False
            Button2.Text = "Delete"
        Else
            Button1.Enabled = False
            Label9.Visible = True
            Button2.Text = "Un-Delete"
        End If
    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        Me.Close()
    End Sub
End Class