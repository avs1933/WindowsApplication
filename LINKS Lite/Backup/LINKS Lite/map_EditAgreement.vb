Public Class map_EditAgreement

    Private Sub map_EditAgreement_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Call LoadMFirms()
        Call LoadAgreementType()
        'Cannot load assigned products because the product ID wont be populated when the form loads.
        'Call LoadAssigned()
    End Sub

    Public Sub LoadMFirms()
        Try

            Dim conn As New OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & My.Settings.dbloc & "';Jet OLEDB:Database Password='" & My.Settings.dbpass & "';")

            Dim strSQL As String = "SELECT ID, FirmName FROM map_Firms WHERE Active = -1 AND SystemValue = False ORDER BY FirmName"
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

    Public Sub LoadAgreementType()
        Try

            Dim conn As New OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & My.Settings.dbloc & "';Jet OLEDB:Database Password='" & My.Settings.dbpass & "';")

            Dim strSQL As String = "SELECT ID, TypeName FROM map_AgreementType WHERE SystemValue = False"
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

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        If ID.Text = "NEW" Then
            Dim ir As MsgBoxResult
            ir = MsgBox("You must save this record before adding products.  Would you like to save now?", MsgBoxStyle.YesNo + MsgBoxStyle.Information, "Record not saved")
            If ir = MsgBoxResult.Yes Then
                If CheckBox2.Checked Or CheckBox3.Checked Or CheckBox4.Checked Or CheckBox5.Checked Then
                    MsgBox("This agreement approves all selected products.  Remove 'Approve All' flag above to select specific products.", MsgBoxStyle.Information, "All Approved")
                Else
                    Call SaveNew()
                    Call PullID()
                    Dim child As New map_AgreementProducts
                    child.MdiParent = Home
                    child.ID.Text = Me.ID.Text
                    child.Show()
                End If
            Else
                'do nothing
            End If
        Else
            If CheckBox2.Checked Or CheckBox3.Checked Or CheckBox4.Checked Or CheckBox5.Checked Then
                MsgBox("This agreement approves all selected products.  Remove 'Approve All' flag above to select specific products.", MsgBoxStyle.Information, "All Approved")
            Else
                Dim child As New map_AgreementProducts
                child.MdiParent = Home
                child.ID.Text = Me.ID.Text
                child.Show()
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

            SQLstr = "INSERT INTO map_Agreements([Notes], [TypeID], [FirmID], [VersionNumber], [DateExecuted], [UseDefaultFee], [AllInHouseSMAApproved], [AllMFApproved], [AllUITApproved], [AllOutsideSMAApproved], [Active])" & _
            "VALUES('" & RichTextBox1.Text & "'," & cboAType.SelectedValue & "," & cboFirms.SelectedValue & ",'" & TextBox2.Text & "',#" & DateTimePicker1.Text & "#," & CheckBox1.CheckState & "," & CheckBox2.CheckState & "," & CheckBox4.CheckState & "," & CheckBox5.CheckState & "," & CheckBox3.CheckState & ", -1)"

            Command = New OleDb.OleDbCommand(SQLstr, Mycn)
            Command.ExecuteNonQuery()

            Mycn.Close()

            MsgBox("Record Saved.")

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "ERROR")
        End Try
    End Sub

    Public Sub PullID()
        Try
            Dim conn As New OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & My.Settings.dbloc & "';Jet OLEDB:Database Password='" & My.Settings.dbpass & "';")

            Dim strSQL As String = "SELECT Top 1 ID FROM map_Agreements WHERE TypeID = " & cboAType.SelectedValue & " AND FirmID = " & cboFirms.SelectedValue & " AND Active = -1 ORDER BY ID DESC"
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

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        If ID.Text = "NEW" Then
            Call SaveNew()
            Call PullID()
        Else
            Call SaveOld()
        End If
    End Sub

    Private Sub Button5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button5.Click
        'Call LoadAssigned()

        If CheckBox2.Checked Then
            'Agreement Approves all In-House SMA
            LoadAllInHouseSMA()
            GoTo line1
        Else
            Call LoadAssigned()
            GoTo line1
        End If

line1:
    End Sub

    Public Sub SaveOld()
        Dim Mycn As OleDb.OleDbConnection
        Dim Command As OleDb.OleDbCommand
        Dim SQLstr As String
        Try
            Mycn = New OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & My.Settings.dbloc & "';Jet OLEDB:Database Password='" & My.Settings.dbpass & "';")
            Mycn.Open()

            SQLstr = "Update map_Agreements SET TypeID = " & cboAType.SelectedValue & ", FirmID = " & cboFirms.SelectedValue & ", VersionNumber = '" & TextBox2.Text & "', DateExecuted = #" & DateTimePicker1.Text & "#, Notes = '" & RichTextBox1.Text & "', UseDefaultFee = " & CheckBox1.CheckState & ", AllInHouseSMAApproved = " & CheckBox2.CheckState & ", AllMFApproved = " & CheckBox4.CheckState & ", AllUITApproved = " & CheckBox5.CheckState & ", AllOutsideSMAApproved = " & CheckBox3.CheckState & " WHERE ID = " & ID.Text

            Command = New OleDb.OleDbCommand(SQLstr, Mycn)
            Command.ExecuteNonQuery()

            Mycn.Close()

            MsgBox("Record Updated.")

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "ERROR")
        End Try
    End Sub

    Public Sub LoadAllInHouseSMA()
        Try
            Dim conn As New OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & My.Settings.dbloc & "';Jet OLEDB:Database Password='" & My.Settings.dbpass & "';")

            Dim SqlString As String


            SqlString = "SELECT map_Products.ID, map_Products.ProductName As [Product Name], map_ProductType.TypeName As [Product Type], Map_Firms.FirmName As [Managing Firm]" & _
            " FROM (map_Products INNER JOIN map_ProductType ON map_Products.TypeOfProductID = map_ProductType.ID) INNER JOIN Map_Firms ON map_Products.ManagingFirmID = Map_Firms.ID" & _
            " WHERE(((map_Products.Active) = True) AND map_Products.ManagingFirmID = 1)" & _
            " GROUP BY map_Products.ID, map_Products.ProductName, map_ProductType.TypeName, Map_Firms.FirmName"

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

    Public Sub LoadAssigned()
        Try
            Dim conn As New OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & My.Settings.dbloc & "';Jet OLEDB:Database Password='" & My.Settings.dbpass & "';")

            Dim SqlString As String


            SqlString = "SELECT map_Products.ID, map_ProductListing.ID, map_Products.ProductName As [Product Name], map_ProductType.TypeName As [Product Type], Map_Firms.FirmName As [Managing Firm]" & _
            " FROM map_ProductListing INNER JOIN ((map_Products INNER JOIN map_ProductType ON map_Products.TypeOfProductID = map_ProductType.ID) INNER JOIN Map_Firms ON map_Products.ManagingFirmID = Map_Firms.ID) ON map_ProductListing.ProductID = map_Products.ID" & _
            " WHERE(((map_Products.Active) = True) And ((map_ProductListing.AgreementID) = " & ID.Text & ") And ((map_ProductListing.Agreement) = True))" & _
            " GROUP BY map_Products.ID, map_Products.ProductName, map_ProductType.TypeName, Map_Firms.FirmName, map_ProductListing.AgreementID, map_ProductListing.ID;"

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
        If CheckBox6.Checked Then
            Label7.Visible = False
            Button2.Enabled = True
            Button1.Enabled = True
            Button3.Text = "Delete"
        Else
            Label7.Visible = True
            Button2.Enabled = False
            Button1.Enabled = False
            Button3.Text = "Un-Delete"
        End If
    End Sub

    Public Sub CheckApprovals()
        If CheckBox2.Checked And CheckBox2.Enabled = True Then
            CheckBox3.Enabled = False
            CheckBox4.Enabled = False
            CheckBox5.Enabled = False
            'Button1.Enabled = False
            GoTo line1
        Else
            CheckBox3.Enabled = True
            CheckBox4.Enabled = True
            CheckBox5.Enabled = True
            'Button1.Enabled = True
        End If

        If CheckBox3.Checked And CheckBox3.Enabled = True Then
            CheckBox2.Enabled = False
            CheckBox4.Enabled = False
            CheckBox5.Enabled = False
            'Button1.Enabled = False
            GoTo line1
        Else
            CheckBox2.Enabled = True
            CheckBox4.Enabled = True
            CheckBox5.Enabled = True
            'Button1.Enabled = True
        End If

        If CheckBox4.Checked And CheckBox4.Enabled = True Then
            CheckBox2.Enabled = False
            CheckBox3.Enabled = False
            CheckBox5.Enabled = False
            'Button1.Enabled = False
            GoTo line1
        Else
            CheckBox2.Enabled = True
            CheckBox3.Enabled = True
            CheckBox5.Enabled = True
            'Button1.Enabled = True
        End If

        If CheckBox5.Checked And CheckBox5.Enabled = True Then
            CheckBox2.Enabled = False
            CheckBox3.Enabled = False
            CheckBox4.Enabled = False
            'Button1.Enabled = False
            GoTo line1
        Else
            CheckBox2.Enabled = True
            CheckBox3.Enabled = True
            CheckBox4.Enabled = True
            'Button1.Enabled = True
        End If

line1:

    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        If ID.Text = "NEW" Then
            MsgBox("You cannot delete an un-saved record.", MsgBoxStyle.Exclamation, "Not Saved")
        Else
            Dim Mycn As OleDb.OleDbConnection
            Dim Command As OleDb.OleDbCommand
            Dim SQLstr As String
            Try
                Mycn = New OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & My.Settings.dbloc & "';Jet OLEDB:Database Password='" & My.Settings.dbpass & "';")
                Mycn.Open()

                If CheckBox6.Checked Then
                    SQLstr = "Update map_Agreements SET Active = False WHERE ID = " & ID.Text
                    CheckBox6.Checked = False
                Else
                    SQLstr = "Update map_Agreements SET Active = -1 WHERE ID = " & ID.Text
                    CheckBox6.Checked = True
                End If

                Command = New OleDb.OleDbCommand(SQLstr, Mycn)
                Command.ExecuteNonQuery()

                Mycn.Close()

            Catch ex As Exception
                MsgBox(ex.Message, MsgBoxStyle.Critical, "ERROR")
            End Try
        End If
    End Sub

    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        Me.Close()
    End Sub

    Private Sub Button6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button6.Click
        If CheckBox1.Checked Then
            'Agreement Uses Default Fee Schedule
            If CheckBox2.Checked Then
                'Agreement Approves all In-House SMA
                LoadAllDefaultSMAFees()
                GoTo line1
            Else
                Call LoadAllDefaultSMAFeesForApproved()
                GoTo line1
            End If
        Else
            If CheckBox2.Checked Then
                Call LoadCustomSMAFeeForLimitedApproved()
            Else
                Call LoadCustomSMAFeeForApproved()
            End If
        End If

line1:

    End Sub

    Public Sub LoadCustomSMAFeeForLimitedApproved()
        Try
            Dim conn As New OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & My.Settings.dbloc & "';Jet OLEDB:Database Password='" & My.Settings.dbpass & "';")

            Dim SqlString As String

            SqlString = "SELECT map_Fees.ID, map_Products.ProductName, map_AgreementType.TypeName, map_Fees.BreakpointFrom, map_Fees.BreakpointTo, map_Fees.Fee, map_Fees.MaxRepFee" & _
            " FROM (map_Fees INNER JOIN map_Products ON map_Fees.ProductID = map_Products.ID) INNER JOIN map_AgreementType ON map_Fees.AgreementTypeID = map_AgreementType.ID" & _
            " WHERE(((map_Fees.Active) = True) And ((map_Fees.Default) = False) AND map_Fees.AgreementID = " & ID.Text & ")" & _
            " GROUP BY map_Fees.ID, map_Products.ProductName, map_AgreementType.TypeName, map_Fees.BreakpointFrom, map_Fees.BreakpointTo, map_Fees.Fee, map_Fees.MaxRepFee, map_Fees.AgreementID"

            Dim da As New OleDb.OleDbDataAdapter(SqlString, conn)
            Dim ds As New DataSet
            da.Fill(ds, "Users")

            With DataGridView2
                .DataSource = ds.Tables("Users")
                .Columns(0).Visible = False
                .Columns(3).DefaultCellStyle.Format = "c"
                .Columns(4).DefaultCellStyle.Format = "c"
                .Columns(3).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                .Columns(4).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
            End With

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "ERROR")
        End Try
    End Sub

    Public Sub LoadCustomSMAFeeForApproved()
        Try
            Dim conn As New OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & My.Settings.dbloc & "';Jet OLEDB:Database Password='" & My.Settings.dbpass & "';")

            Dim SqlString As String

            SqlString = "SELECT map_Fees.ID, map_Products.ProductName, map_AgreementType.TypeName, map_Fees.BreakpointFrom, map_Fees.BreakpointTo, map_Fees.Fee, map_Fees.MaxRepFee" & _
            " FROM ((map_Fees INNER JOIN map_AgreementType ON map_Fees.AgreementTypeID = map_AgreementType.ID) INNER JOIN map_Products ON map_Fees.ProductID = map_Products.ID) INNER JOIN map_ProductListing ON map_Fees.ProductListingID = map_ProductListing.ID" & _
            " WHERE(((map_Fees.Active) = True) And ((map_Fees.Default) = False) And ((map_ProductListing.AgreementID) = " & ID.Text & ") And ((map_ProductListing.Agreement) = True))" & _
            " GROUP BY map_Fees.ID, map_Products.ProductName, map_AgreementType.TypeName, map_Fees.BreakpointFrom, map_Fees.BreakpointTo, map_Fees.Fee, map_Fees.MaxRepFee;"

            Dim da As New OleDb.OleDbDataAdapter(SqlString, conn)
            Dim ds As New DataSet
            da.Fill(ds, "Users")

            With DataGridView2
                .DataSource = ds.Tables("Users")
                .Columns(0).Visible = False
                .Columns(3).DefaultCellStyle.Format = "c"
                .Columns(4).DefaultCellStyle.Format = "c"
                .Columns(3).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                .Columns(4).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
            End With

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "ERROR")
        End Try
    End Sub

    Public Sub LoadAllDefaultSMAFeesForApproved()
        Try
            Dim conn As New OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & My.Settings.dbloc & "';Jet OLEDB:Database Password='" & My.Settings.dbpass & "';")

            Dim SqlString As String

            SqlString = "SELECT map_Fees.ID, map_ProductListing.AgreementID, map_Products.ProductName, map_AgreementType.TypeName, map_Fees.BreakpointFrom, map_Fees.BreakpointTo, map_Fees.Fee, map_Fees.MaxRepFee" & _
            " FROM (map_Fees INNER JOIN (map_ProductListing INNER JOIN map_Products ON map_ProductListing.ProductID = map_Products.ID) ON map_Fees.ProductID = map_ProductListing.ProductID) INNER JOIN map_AgreementType ON map_Fees.AgreementTypeID = map_AgreementType.ID" & _
            " WHERE(((map_Fees.Active) = True) And ((map_Fees.Default) = True) AND map_ProductListing.AgreementID = " & ID.Text & " AND map_Fees.AgreementTypeID = " & cboAType.SelectedValue & ")" & _
            " GROUP BY map_Fees.ID, map_ProductListing.AgreementID, map_Products.ProductName, map_AgreementType.TypeName, map_Fees.BreakpointFrom, map_Fees.BreakpointTo, map_Fees.Fee, map_Fees.MaxRepFee;"

            Dim da As New OleDb.OleDbDataAdapter(SqlString, conn)
            Dim ds As New DataSet
            da.Fill(ds, "Users")

            With DataGridView2
                .DataSource = ds.Tables("Users")
                .Columns(0).Visible = False
                .Columns(1).Visible = False
                .Columns(5).DefaultCellStyle.Format = "c"
                .Columns(4).DefaultCellStyle.Format = "c"
                .Columns(5).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                .Columns(4).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
            End With

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "ERROR")
        End Try
    End Sub

    Public Sub LoadAllDefaultSMAFees()
        Try
            Dim conn As New OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & My.Settings.dbloc & "';Jet OLEDB:Database Password='" & My.Settings.dbpass & "';")

            Dim SqlString As String

            SqlString = "SELECT map_Fees.ID, map_Products.ProductName, map_AgreementType.TypeName, map_Fees.BreakpointFrom, map_Fees.BreakpointTo, map_Fees.Fee, map_Fees.MaxRepFee" & _
            " FROM (map_Fees INNER JOIN map_AgreementType ON map_Fees.AgreementTypeID = map_AgreementType.ID) INNER JOIN map_Products ON map_Fees.ProductID = map_Products.ID" & _
            " WHERE(((map_Fees.Active) = True) And ((map_Fees.Default) = True) And ((map_Products.ManagingFirmID) = 1) AND map_Fees.AgreementTypeID = " & cboAType.SelectedValue & ")" & _
            " GROUP BY map_Fees.ID, map_Products.ProductName, map_AgreementType.TypeName, map_Fees.BreakpointFrom, map_Fees.BreakpointTo, map_Fees.Fee, map_Fees.MaxRepFee;"

            Dim da As New OleDb.OleDbDataAdapter(SqlString, conn)
            Dim ds As New DataSet
            da.Fill(ds, "Users")

            With DataGridView2
                .DataSource = ds.Tables("Users")
                .Columns(0).Visible = False
                .Columns(3).DefaultCellStyle.Format = "c"
                .Columns(4).DefaultCellStyle.Format = "c"
                .Columns(3).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                .Columns(4).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
            End With

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "ERROR")
        End Try
    End Sub

    Private Sub CheckBox2_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckBox2.CheckedChanged
        Call CheckApprovals()
    End Sub

    Private Sub CheckBox3_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckBox3.CheckedChanged
        Call CheckApprovals()
    End Sub

    Private Sub CheckBox4_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckBox4.CheckedChanged
        Call CheckApprovals()
    End Sub

    Private Sub CheckBox5_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckBox5.CheckedChanged
        Call CheckApprovals()
    End Sub

    Private Sub EditFeeToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles EditFeeToolStripMenuItem.Click
        Call LoadFeeEdit()
    End Sub

    Public Sub LoadFeeEdit()
        If DataGridView2.RowCount = "0" Then
            'do nothing
        Else
            Dim child As New map_EditFee
            child.MdiParent = Home
            child.Show()

            Try
                Dim conn As New OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & My.Settings.dbloc & "';Jet OLEDB:Database Password='" & My.Settings.dbpass & "';")

                Dim strSQL As String = "SELECT * FROM map_Fees WHERE ID = " & DataGridView2.SelectedCells(0).Value
                Dim queryString As String = String.Format(strSQL)
                Dim cmd As New OleDb.OleDbCommand(queryString, conn)
                Dim da As New OleDb.OleDbDataAdapter(cmd)
                Dim ds As New DataSet

                da.Fill(ds, "User")
                Dim dt As DataTable = ds.Tables("User")
                Dim row1 As DataRow = dt.Rows(0)

                child.ID.Text = (row1("ID"))

                If IsDBNull(row1("ProductListingID")) Then
                    'do nothing
                Else
                    child.ProductListingID.Text = (row1("ProductListingID"))
                End If

                child.AgreementID.Text = ID.Text

                If IsDBNull(row1("ProductID")) Then
                    'do nothing
                Else
                    child.cboProduct.SelectedValue = (row1("ProductID"))
                End If

                If IsDBNull(row1("AgreementTypeID")) Then
                    'do nothing
                Else
                    child.cboAType.SelectedValue = (row1("AgreementTypeID"))
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

                If (row1("Default")) = True Then

                Else
                    child.AgreementID.Text = ID.Text
                    'child.AgreementText.Text = cboFirms.SelectedText
                End If

                child.ckbDefaultFee.Checked = (row1("Default"))
                child.ckb40Act.Checked = (row1("40Act"))
                child.ckbActive.Checked = (row1("Active"))

            Catch ex As Exception
                MsgBox(ex.Message, MsgBoxStyle.Critical, "ERROR")
            End Try
        End If
    End Sub

    Private Sub DataGridView2_CellContentDoubleClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView2.CellContentDoubleClick
        Call LoadFeeEdit()
    End Sub

    Private Sub AddFeeToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles AddFeeToolStripMenuItem.Click
        If DataGridView1.RowCount = "0" Then
            'do nothing
        Else
            If CheckBox1.Checked Then

            Else
                If CheckBox2.Checked Then
                    Call NewCustomFee()
                Else
                    Call NewCustomFeeNotAll()
                End If
            End If
        End If
    End Sub

    Public Sub NewCustomFee()
        Dim child As New map_EditFee
        child.MdiParent = Home
        child.Show()
        child.ID.Text = "NEW"
        child.AgreementID.Text = ID.Text
        'child.ProductListingID.Text = DataGridView1.SelectedCells(1).Value
        child.cboProduct.SelectedValue = DataGridView1.SelectedCells(0).Value
        child.cboAType.SelectedValue = cboAType.SelectedValue
        child.ckbDefaultFee.Enabled = False
        child.ckb40Act.Enabled = False
        child.ckbActive.Enabled = True
    End Sub

    Public Sub NewCustomFeeNotAll()
        Dim child As New map_EditFee
        child.MdiParent = Home
        child.Show()
        child.ID.Text = "NEW"
        child.AgreementID.Text = ID.Text
        child.ProductListingID.Text = DataGridView1.SelectedCells(1).Value
        child.cboProduct.SelectedValue = DataGridView1.SelectedCells(0).Value
        child.cboAType.SelectedValue = cboAType.SelectedValue
        child.ckbDefaultFee.Enabled = False
        child.ckb40Act.Enabled = False
        child.ckbActive.Enabled = True
    End Sub

End Class