Public Class map_EditPlatform

    Private Sub map_EditPlatform_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Call LoadAPXFirms()
        Call LoadDatabases()
    End Sub

    Public Sub LoadAPXFirms()
        Try

            Dim conn As New OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & My.Settings.dbloc & "';Jet OLEDB:Database Password='" & My.Settings.dbpass & "';")

            Dim strSQL As String = "SELECT ID, FirmName FROM Map_Firms WHERE TypeID = 6 AND Active = -1 AND SystemValue = False"
            Dim da As New OleDb.OleDbDataAdapter(strSQL, conn)
            Dim ds As New DataSet
            da.Fill(ds, "Users")

            With cboFirms
                .DataSource = ds.Tables("Users")
                .DisplayMember = "FirmName"
                .ValueMember = "ID"
                .SelectedIndex = 0
            End With

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "ERROR")
        End Try
    End Sub

    Public Sub LoadDatabases()
        Try

            Dim conn As New OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & My.Settings.dbloc & "';Jet OLEDB:Database Password='" & My.Settings.dbpass & "';")

            Dim strSQL As String = "SELECT ContactID, ContactName FROM dbo_vQbRowDefContact WHERE Status = 'Database'"
            Dim da As New OleDb.OleDbDataAdapter(strSQL, conn)
            Dim ds As New DataSet
            da.Fill(ds, "Users")

            With cboDatabase
                .DataSource = ds.Tables("Users")
                .DisplayMember = "ContactName"
                .ValueMember = "ContactID"
                .SelectedIndex = 0
            End With

            'Label16.Visible = False
            'cboDatabase.Enabled = True
            'Call LoadDBData()

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "ERROR")
        End Try
    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        If ID.Text = "NEW" Then
            Call CheckForDupe()
        Else
            Call SaveOld()
        End If
    End Sub
    Public Sub PullID()
        Try
            Dim conn As New OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & My.Settings.dbloc & "';Jet OLEDB:Database Password='" & My.Settings.dbpass & "';")

            Dim strSQL As String = "SELECT Top 1 ID FROM map_Platforms WHERE PlatformID = " & cboFirms.SelectedValue & " AND Active = -1 ORDER BY ID DESC"
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

    Public Sub CheckForDupe()

        Dim Mycn As OleDb.OleDbConnection
        Dim SQLstring As String
        Try
            Mycn = New OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & My.Settings.dbloc & "';Jet OLEDB:Database Password='" & My.Settings.dbpass & "';")
            Mycn.Open()

            SQLstring = "SELECT * FROM Map_Platforms WHERE [PlatformID] = " & cboFirms.SelectedValue

            Dim queryString As String = String.Format(SQLstring)
            Dim cmd As New OleDb.OleDbCommand(queryString, Mycn)
            Dim da As New OleDb.OleDbDataAdapter(cmd)
            Dim ds As New DataSet

            da.Fill(ds, "User")
            Dim dt As DataTable = ds.Tables("User")
            If dt.Rows.Count = 0 Then
                Call SaveNew()
            Else
                MsgBox("A record already exsists for that Platform.", MsgBoxStyle.Information, "DUPLICATE RECORD DETECTED")
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

            SQLstr = "INSERT INTO map_Platforms(PlatformID, ApprovalName, ApprovalProcess, DBDriverID, OfferSMA, OfferUMA, OfferMF, OfferUIT, Active)" & _
            "VALUES(" & cboFirms.SelectedValue & ",'" & TextBox1.Text & "','" & RichTextBox1.Text & "'," & cboDatabase.SelectedValue & "," & ckbSMA.CheckState & "," & ckbUMA.CheckState & "," & ckbMF.CheckState & "," & ckbUIT.CheckState & ",-1)"

            Command = New OleDb.OleDbCommand(SQLstr, Mycn)
            Command.ExecuteNonQuery()

            Mycn.Close()

            MsgBox("Record Saved.")

            Call PullID()

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

            SQLstr = "Update map_Platforms SET PlatformID = " & cboFirms.SelectedValue & ", ApprovalName = '" & TextBox1.Text & "', ApprovalProcess = '" & RichTextBox1.Text & "', DBDriverID = " & cboDatabase.SelectedValue & ", OfferSMA = " & ckbSMA.CheckState & ", OfferUMA = " & ckbUMA.CheckState & ", OfferMF = " & ckbMF.CheckState & ", OfferUIT = " & ckbUIT.CheckState & ", Active = " & ckbActive.CheckState & " WHERE ID = " & ID.Text

            Command = New OleDb.OleDbCommand(SQLstr, Mycn)
            Command.ExecuteNonQuery()

            Mycn.Close()

            MsgBox("Record Updated.")

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "ERROR")
        End Try
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        If ID.Text = "NEW" Then
            Dim ir As MsgBoxResult
            ir = MsgBox("You must save this record before adding fees.  Would you like to save now?", MsgBoxStyle.YesNo + MsgBoxStyle.Information, "Save Record?")
            If ir = MsgBoxResult.Yes Then
                Call CheckForDupe()
                Call PullID()
                Dim child As New map_PlatformProducts
                child.MdiParent = Home
                child.Show()
                child.ID.Text = ID.Text
            Else
                'do nothing
            End If
        Else
            Dim child As New map_PlatformProducts
            child.MdiParent = Home
            child.Show()
            child.ID.Text = ID.Text
        End If
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Call LoadAssigned()
    End Sub

    Public Sub LoadAssigned()
        Try
            Dim conn As New OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & My.Settings.dbloc & "';Jet OLEDB:Database Password='" & My.Settings.dbpass & "';")

            Dim SqlString As String

            SqlString = "SELECT map_PlatformListings.ID, map_platformListings.ProductID, map_Products.ProductName As [Product Name], map_ProductType.TypeName As [Product Type], Map_Firms.FirmName As [Managing Firm], map_PlatformListings.PlatformApproved As [Platform Approved], map_PlatformListings.SMAOffered As [Offered as SMA], map_PlatformListings.UMAOffered As [Offered as UMA]" & _
            " FROM ((map_PlatformListings INNER JOIN map_Products ON map_PlatformListings.ProductID = map_Products.ID) INNER JOIN map_ProductType ON map_Products.TypeOfProductID = map_ProductType.ID) INNER JOIN Map_Firms ON map_Products.ManagingFirmID = Map_Firms.ID" & _
            " WHERE(((map_PlatformListings.Active) = True) And ((map_PlatformListings.PlatformID) = " & ID.Text & "))" & _
            " GROUP BY map_PlatformListings.ID, map_platformListings.ProductID, map_Products.ProductName, map_ProductType.TypeName, Map_Firms.FirmName, map_PlatformListings.PlatformApproved, map_PlatformListings.SMAOffered, map_PlatformListings.UMAOffered;"

            Dim da As New OleDb.OleDbDataAdapter(SqlString, conn)
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

    Private Sub Timer1_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer1.Tick
        If ckbActive.Checked Then
            Label7.Visible = False
            Button2.Enabled = True
            Button3.Enabled = True
            Button4.Text = "Delete"
        Else
            Label7.Visible = True
            Button2.Enabled = False
            Button3.Enabled = False
            Button4.Text = "Un-Delete"
        End If
    End Sub

    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
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
                    SQLstr = "Update map_Platforms SET Active = False WHERE ID = " & ID.Text
                    ckbActive.Checked = False
                Else
                    SQLstr = "Update map_Platforms SET Active = -1 WHERE ID = " & ID.Text
                    ckbActive.Checked = True
                End If

                Command = New OleDb.OleDbCommand(SQLstr, Mycn)
                Command.ExecuteNonQuery()

                Mycn.Close()

            Catch ex As Exception
                MsgBox(ex.Message, MsgBoxStyle.Critical, "ERROR")
            End Try
        End If
    End Sub

    Private Sub AddFeeToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles AddFeeToolStripMenuItem.Click
        If DataGridView1.RowCount = 0 Then

        Else
            Dim child As New map_EditPlatformFee
            child.MdiParent = Home
            child.Show()
            child.cboPlatform.SelectedValue = ID.Text
            child.cboProduct.SelectedValue = DataGridView1.SelectedCells(1).Value
            child.ID.Text = "NEW"

        End If
    End Sub

    Private Sub Button6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button6.Click
        If ID.Text = "NEW" Then
            'do nothing
        Else
            Call LoadFees()
        End If

    End Sub

    Public Sub LoadFees()
        Try
            Dim conn As New OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & My.Settings.dbloc & "';Jet OLEDB:Database Password='" & My.Settings.dbpass & "';")

            Dim SqlString As String

            SqlString = "SELECT map_PlatformFees.ID, map_Products.ProductName, map_ProductType.TypeName, Map_Firms.FirmName, map_PlatformFees.BreakpointFrom, map_PlatformFees.BreakpointTo, map_PlatformFees.Fee, map_PlatformFees.MaxRepFee" & _
            " FROM ((map_PlatformFees INNER JOIN map_Products ON map_PlatformFees.ProductID = map_Products.ID) INNER JOIN map_ProductType ON map_Products.TypeOfProductID = map_ProductType.ID) INNER JOIN Map_Firms ON map_Products.ManagingFirmID = Map_Firms.ID" & _
            " WHERE(((map_PlatformFees.Active) = True) And ((map_PlatformFees.WLID) Is Null) AND map_PlatformFees.PlatformID = " & ID.Text & ")" & _
            " GROUP BY map_PlatformFees.ID, map_Products.ProductName, map_ProductType.TypeName, Map_Firms.FirmName, map_PlatformFees.BreakpointFrom, map_PlatformFees.BreakpointTo, map_PlatformFees.Fee, map_PlatformFees.MaxRepFee;"

            Dim da As New OleDb.OleDbDataAdapter(SqlString, conn)
            Dim ds As New DataSet
            da.Fill(ds, "Users")

            With DataGridView2
                .DataSource = ds.Tables("Users")
                .Columns(0).Visible = False
                .Columns(4).DefaultCellStyle.Format = "c"
                .Columns(5).DefaultCellStyle.Format = "c"
                .Columns(4).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                .Columns(5).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
            End With

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "ERROR")
        End Try
    End Sub

    Public Sub LoadFeeEdit()
        If DataGridView2.RowCount = "0" Then
            'do nothing
        Else
            Dim child As New map_EditPlatformFee
            child.MdiParent = Home
            child.Show()

            Try
                Dim conn As New OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & My.Settings.dbloc & "';Jet OLEDB:Database Password='" & My.Settings.dbpass & "';")

                Dim strSQL As String = "SELECT * FROM map_PlatformFees WHERE ID = " & DataGridView2.SelectedCells(0).Value
                Dim queryString As String = String.Format(strSQL)
                Dim cmd As New OleDb.OleDbCommand(queryString, conn)
                Dim da As New OleDb.OleDbDataAdapter(cmd)
                Dim ds As New DataSet

                da.Fill(ds, "User")
                Dim dt As DataTable = ds.Tables("User")
                Dim row1 As DataRow = dt.Rows(0)

                child.ID.Text = (row1("ID"))

                If IsDBNull(row1("PlatformID")) Then
                    'do nothing
                Else
                    child.cboPlatform.SelectedValue = (row1("PlatformID"))
                End If

                child.ckbWL.Checked = False
                child.ckbActive.Checked = (row1("Active"))

                If IsDBNull(row1("ProductID")) Then
                    'do nothing
                Else
                    child.cboProduct.SelectedValue = (row1("ProductID"))
                End If

                If IsDBNull(row1("BreakPointFrom")) Then
                    'do nothing
                Else
                    child.BP1.Text = (row1("BreakPointFrom"))
                End If

                If IsDBNull(row1("BreakPointTo")) Then
                    'do nothing
                Else
                    child.BP2.Text = (row1("BreakPointTo"))
                End If

                If IsDBNull(row1("Fee")) Then
                    'do nothing
                Else
                    child.Fee.Text = (row1("Fee"))
                End If

                If IsDBNull(row1("MaxRepFee")) Then
                    'do nothing
                Else
                    child.RepFee.Text = (row1("MaxRepFee"))
                End If

            Catch ex As Exception
                MsgBox(ex.Message, MsgBoxStyle.Critical, "ERROR")
            End Try
        End If
    End Sub

    Private Sub EditFeeToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles EditFeeToolStripMenuItem.Click
        Call LoadFeeEdit()
    End Sub
End Class