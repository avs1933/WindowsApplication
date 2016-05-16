Public Class map_SMA
    Dim loadtype As System.Threading.Thread
    Dim load40act As System.Threading.Thread
    Dim loadfirms As System.Threading.Thread

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Dim child As New map_EditProductType
        child.MdiParent = Me.MdiParent
        child.Show()
        child.TypeID.Text = "NEW"
    End Sub

    Private Sub Button10_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button10.Click
        Me.Close()
    End Sub

    Private Sub map_SMA_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        'Control.CheckForIllegalCrossThreadCalls = False
        'loadtype = New System.Threading.Thread(AddressOf LoadProductType)
        'loadtype.Start()

        'loadfirms = New System.Threading.Thread(AddressOf LoadMFirms)
        'loadfirms.Start()
        Call LoadPermissions()
        Call LoadProductType()
        Call LoadMFirms()
        Call LoadAssetClass()
        Call LoadObjectives1()
        Call LoadObjectives2()
    End Sub

    Public Sub LoadPermissions()
        If Permissions.MAPAddFirm.Checked Then
            Button2.Enabled = True
        Else
            Button2.Enabled = False
        End If

        If Permissions.MAPAddProductType.Checked Then
            Button1.Enabled = True
        Else
            Button1.Enabled = False
        End If

        If Permissions.MAPAddObjective.Checked Then
            Button4.Enabled = True
        Else
            Button4.Enabled = False
        End If
    End Sub

    Public Sub LoadAssetClass()
        Try

            Dim conn As New OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & My.Settings.dbloc & "';Jet OLEDB:Database Password='" & My.Settings.dbpass & "';")

            Dim strSQL As String = "SELECT AssetClassCode, AssetClassName FROM AdvApp_vAssetClass"
            Dim da As New OleDb.OleDbDataAdapter(strSQL, conn)
            Dim ds As New DataSet
            da.Fill(ds, "Users")

            With cboAssetClass
                .DataSource = ds.Tables("Users")
                .DisplayMember = "AssetClassName"
                .ValueMember = "AssetClassCode"
                .SelectedIndex = 0
            End With

            'Label14.Visible = False
            cboAssetClass.Enabled = True
            'Call LoadDBData()

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "ERROR")
        End Try
    End Sub

    Public Sub LoadProductType()
        Try

            Dim conn As New OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & My.Settings.dbloc & "';Jet OLEDB:Database Password='" & My.Settings.dbpass & "';")

            Dim strSQL As String = "SELECT ID, TypeName FROM map_ProductType WHERE Active = -1"
            Dim da As New OleDb.OleDbDataAdapter(strSQL, conn)
            Dim ds As New DataSet
            da.Fill(ds, "Users")

            With cboProductType
                .DataSource = ds.Tables("Users")
                .DisplayMember = "TypeName"
                .ValueMember = "ID"
                .SelectedIndex = 0
            End With

            Label14.Visible = False
            cboProductType.Enabled = True
            'Call LoadDBData()

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "ERROR")
        End Try
    End Sub

    Public Sub LoadMFirms()
        Try

            Dim conn As New OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & My.Settings.dbloc & "';Jet OLEDB:Database Password='" & My.Settings.dbpass & "';")

            Dim strSQL As String = "SELECT ID, FirmName FROM map_Firms WHERE Active = -1 AND TypeID = 4 ORDER BY FirmName"
            Dim da As New OleDb.OleDbDataAdapter(strSQL, conn)
            Dim ds As New DataSet
            da.Fill(ds, "Users")

            With cboFirms
                .DataSource = ds.Tables("Users")
                .DisplayMember = "FirmName"
                .ValueMember = "ID"
                .SelectedIndex = 0
            End With

            Label15.Visible = False
            cboFirms.Enabled = True
            'Call LoadDBData()

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "ERROR")
        End Try
    End Sub

    Public Sub LoadObjectives1()
        Try

            Dim conn As New OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & My.Settings.dbloc & "';Jet OLEDB:Database Password='" & My.Settings.dbpass & "';")

            Dim strSQL As String = "SELECT * FROM map_Objective ORDER BY ObjectiveName"
            Dim da As New OleDb.OleDbDataAdapter(strSQL, conn)
            Dim ds As New DataSet
            da.Fill(ds, "Users")

            With cboObjective1
                .DataSource = ds.Tables("Users")
                .DisplayMember = "ObjectiveName"
                .ValueMember = "ID"
                .SelectedIndex = 0
            End With

            'Label15.Visible = False
            'cboFirms.Enabled = True
            'Call LoadDBData()

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "ERROR")
        End Try
    End Sub

    Public Sub LoadObjectives2()
        Try

            Dim conn As New OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & My.Settings.dbloc & "';Jet OLEDB:Database Password='" & My.Settings.dbpass & "';")

            Dim strSQL As String = "SELECT * FROM map_Objective ORDER BY ObjectiveName"
            Dim da As New OleDb.OleDbDataAdapter(strSQL, conn)
            Dim ds As New DataSet
            da.Fill(ds, "Users")

            With cboObjective2
                .DataSource = ds.Tables("Users")
                .DisplayMember = "ObjectiveName"
                .ValueMember = "ID"
                .SelectedIndex = 0
            End With

            'Label15.Visible = False
            'cboFirms.Enabled = True
            'Call LoadDBData()

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "ERROR")
        End Try
    End Sub

    Public Sub Load40ActDriver()
        Try
            Dim conn As New OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & My.Settings.dbloc & "';Jet OLEDB:Database Password='" & My.Settings.dbpass & "';")

            Dim strSQL As String = "SELECT * FROM map_ProductType WHERE ID = " & cboProductType.SelectedValue
            Dim queryString As String = String.Format(strSQL)
            Dim cmd As New OleDb.OleDbCommand(queryString, conn)
            Dim da As New OleDb.OleDbDataAdapter(cmd)
            Dim ds As New DataSet

            da.Fill(ds, "User")
            Dim dt As DataTable = ds.Tables("User")

            Dim row1 As DataRow = dt.Rows(0)

            CheckBox1.Checked = (row1("40Act"))

            If CheckBox1.Checked = False Then

                GroupBox1.Visible = False
            Else

                GroupBox1.Visible = True
            End If

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "ERROR")
        End Try
    End Sub

    Private Sub cboProductType_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles cboProductType.LostFocus
        'Control.CheckForIllegalCrossThreadCalls = False
        'load40act = New System.Threading.Thread(AddressOf Load40ActDriver)
        'load40act.Start()
        Call Load40ActDriver()
    End Sub

    Private Sub cboProductType_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cboProductType.SelectedIndexChanged
        
    End Sub

    Private Sub cboProductType_SelectedValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cboProductType.SelectedValueChanged
        'Control.CheckForIllegalCrossThreadCalls = False
        'load40act = New System.Threading.Thread(AddressOf Load40ActDriver)
        'load40act.Start()

    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        Dim child As New map_EditMFirms
        child.MdiParent = Home
        child.Show()
        child.ID.Text = "NEW"
    End Sub

    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        Dim child As New map_EditObjective
        child.MdiParent = Home
        child.Show()
        child.ID.Text = "NEW"
    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        Call LoadProductType()
    End Sub

    Private Sub Button5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button5.Click
        Call LoadMFirms()
    End Sub

    Private Sub Button12_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button12.Click
        Call LoadObjectives1()
        Call LoadObjectives2()
    End Sub

    Private Sub Button13_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button13.Click
        Call LoadAssetClass()
    End Sub

    Private Sub Button14_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button14.Click
        MsgBox("Asset Classes are loaded from the Advent Database.  To add or edit asset classes click on Information, Reference and Asset Class in APX.", MsgBoxStyle.Information, "APX Driven Field")
    End Sub

    Private Sub Button11_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button11.Click
        Dim ir As MsgBoxResult
        ir = MsgBox("You are about to reload all fields.  This will reset all drop down boxes and refresh with updated field.  Are you sure you want to continue?", MsgBoxStyle.YesNo + MsgBoxStyle.Critical, "Confirm Reload")

        If ir = MsgBoxResult.Yes Then
            Call LoadProductType()
            Call LoadMFirms()
            Call LoadAssetClass()
            Call LoadObjectives1()
            Call LoadObjectives2()
        Else

        End If
    End Sub

    Private Sub Button7_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button7.Click
        If ProductID.Text = "NEW" Then
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

            SQLstring = "SELECT * FROM map_Products WHERE [ProductName] = '" & TextBox2.Text & "' AND TypeOfProductID = " & cboProductType.SelectedValue & " AND Active = -1"

            Dim queryString As String = String.Format(SQLstring)
            Dim cmd As New OleDb.OleDbCommand(queryString, Mycn)
            Dim da As New OleDb.OleDbDataAdapter(cmd)
            Dim ds As New DataSet

            da.Fill(ds, "User")
            Dim dt As DataTable = ds.Tables("User")
            If dt.Rows.Count = 0 Then
                Call SaveNew()
            Else
                MsgBox("A record already exsists for that Product Name.", MsgBoxStyle.Information, "DUPLICATE RECORD DETECTED")
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

            SQLstr = "INSERT INTO map_Products(ProductName, TypeOfProductID, ManagingFirmID, AssetClassCode, PrimaryObjectiveID, SecondaryObjectiveID, ProductDesc, SelectionProcess, Symbol, CUSIP, Series, 40ActDriver, Active)" & _
            "VALUES('" & TextBox2.Text & "', " & cboProductType.SelectedValue & "," & cboFirms.SelectedValue & ",'" & cboAssetClass.SelectedValue & "'," & cboObjective1.SelectedValue & "," & cboObjective2.SelectedValue & ",'" & RichTextBox1.Text & "','" & RichTextBox2.Text & "','" & TextBox3.Text & "','" & TextBox4.Text & "','" & TextBox5.Text & "'," & CheckBox1.CheckState & ", -1)"

            Command = New OleDb.OleDbCommand(SQLstr, Mycn)
            Command.ExecuteNonQuery()

            Mycn.Close()

            MsgBox("Record Saved.")
            Call PullProductID()

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "ERROR")
        End Try
    End Sub

    Public Sub PullProductID()
        Try
            Dim conn As New OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & My.Settings.dbloc & "';Jet OLEDB:Database Password='" & My.Settings.dbpass & "';")

            Dim strSQL As String = "SELECT ID FROM map_Products WHERE ProductName = '" & TextBox2.Text & "' AND Active = -1"
            Dim queryString As String = String.Format(strSQL)
            Dim cmd As New OleDb.OleDbCommand(queryString, conn)
            Dim da As New OleDb.OleDbDataAdapter(cmd)
            Dim ds As New DataSet

            da.Fill(ds, "User")
            Dim dt As DataTable = ds.Tables("User")

            Dim row1 As DataRow = dt.Rows(0)

            If IsDBNull(row1("ID")) Then
                ProductID.Text = "NEW"
            Else
                ProductID.Text = (row1("ID"))
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

            SQLstr = "Update map_Products SET ProductName = '" & TextBox2.Text & "', TypeOfProductID = " & cboProductType.SelectedValue & ", ManagingFirmID = " & cboFirms.SelectedValue & ", AssetClassCode = '" & cboAssetClass.SelectedValue & "', PrimaryObjectiveID = " & cboObjective1.SelectedValue & ", SecondaryObjectiveID = " & cboObjective2.SelectedValue & ", ProductDesc = '" & RichTextBox1.Text & "', SelectionProcess = '" & RichTextBox2.Text & "', Symbol = '" & TextBox3.Text & "', CUSIP = '" & TextBox4.Text & "', Series = '" & TextBox5.Text & "', 40ActDriver = " & CheckBox1.CheckState & " WHERE ID = " & ProductID.Text

            Command = New OleDb.OleDbCommand(SQLstr, Mycn)
            Command.ExecuteNonQuery()

            Mycn.Close()

            MsgBox("Record Updated.")

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "ERROR")
        End Try
    End Sub

    Private Sub Button9_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button9.Click
        If ProductID.Text = "NEW" Then
            MsgBox("This record has not been saved.", MsgBoxStyle.Information, "Nothing to delete")
        Else
            Dim Mycn As OleDb.OleDbConnection
            Dim Command As OleDb.OleDbCommand
            Dim SQLstr As String
            Try
                Mycn = New OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & My.Settings.dbloc & "';Jet OLEDB:Database Password='" & My.Settings.dbpass & "';")
                Mycn.Open()

                If CheckBox2.Checked Then
                    SQLstr = "Update map_Products SET Active = NULL WHERE ID = " & ProductID.Text
                    CheckBox2.Checked = False
                    Button9.Text = "Un-Delete"
                Else
                    SQLstr = "Update map_Products SET Active = -1 WHERE ID = " & ProductID.Text
                    CheckBox2.Checked = True
                    Button9.Text = "Delete"
                End If

                Command = New OleDb.OleDbCommand(SQLstr, Mycn)
                Command.ExecuteNonQuery()

                Mycn.Close()

            Catch ex As Exception
                MsgBox(ex.Message, MsgBoxStyle.Critical, "ERROR")
            End Try
        End If
    End Sub

    Private Sub Timer1_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer1.Tick
        If CheckBox2.Checked Then
            Label16.Visible = False
            Button7.Enabled = True
            Button9.Text = "Delete"
        Else
            Label16.Visible = True
            Button7.Enabled = False
            Button9.Text = "Un-Delete"
        End If
    End Sub

    Private Sub Button6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button6.Click
        If ProductID.Text = "NEW" Then
            Dim ir As MsgBoxResult
            ir = MsgBox("You must save this record before adding fees.  Would you like to save now?", MsgBoxStyle.YesNo + MsgBoxStyle.Information, "Save Record?")
            If ir = MsgBoxResult.Yes Then
                Call CheckForDupe()
                Call PullProductID()
                Call LoadFeeEdit()
            Else
                'do nothing
            End If
        Else
            Call LoadFeeEdit()
        End If
    End Sub

    Public Sub LoadFeeEdit()
        Dim child As New map_EditFee
        child.MdiParent = Home
        child.Show()
        child.ID.Text = "NEW"
        child.cboProduct.SelectedValue = ProductID.Text
        child.ckbDefaultFee.Checked = True
        child.ckbDefaultFee.Enabled = False
        If CheckBox1.Checked Then
            child.ckb40Act.Checked = True
            child.ckb40Act.Enabled = False
        Else
            child.ckb40Act.Checked = False
            child.ckb40Act.Enabled = False
        End If
    End Sub

    Private Sub Button15_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button15.Click
        Call LoadDefaultFees()
    End Sub

    Public Sub LoadDefaultFees()
        If ProductID.Text = "NEW" Then
            'do nothing
        Else
            Try
                Dim conn As New OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & My.Settings.dbloc & "';Jet OLEDB:Database Password='" & My.Settings.dbpass & "';")

                Dim SqlString As String

                'SqlString = "SELECT map_Fees.ID, map_Fees.BreakpointFrom As [Starting Breakpoint], map_Fees.BreakpointTo As [Ending Breakpoint], map_Fees.Fee, map_Fees.MaxRepFee As [Max Rep Fee], map_Fees.Active, map_Fees.Default" & _
                '" FROM map_Fees" & _
                '" WHERE map_Fees.ProductID = " & ProductID.Text & " AND Active = -1 AND Default = -1"

                SqlString = "SELECT map_Fees.ID, map_AgreementType.TypeName As [Agreement Type], map_Fees.BreakpointFrom As [Starting Breakpoint], map_Fees.BreakpointTo As [Ending Breakpoint], map_Fees.Fee, map_Fees.MaxRepFee As [Max Rep Fee], map_Fees.Active, map_Fees.Default" & _
                " FROM map_Fees INNER JOIN map_AgreementType ON map_Fees.AgreementTypeID = map_AgreementType.ID" & _
                " WHERE(((map_Fees.Active) = True) And ((map_Fees.Default) = True) And ((map_Fees.ProductID) = " & ProductID.Text & "))" & _
                " GROUP BY map_Fees.ID, map_AgreementType.TypeName, map_Fees.BreakpointFrom, map_Fees.BreakpointTo, map_Fees.Fee, map_Fees.MaxRepFee, map_Fees.Active, map_Fees.Default;"

                Dim da As New OleDb.OleDbDataAdapter(SqlString, conn)
                Dim ds As New DataSet
                da.Fill(ds, "Users")

                With DataGridView1
                    .DataSource = ds.Tables("Users")
                    .Columns(0).Visible = False
                    .Columns(2).DefaultCellStyle.Format = "c"
                    .Columns(3).DefaultCellStyle.Format = "c"
                    .Columns(2).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                    .Columns(3).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                End With

            Catch ex As Exception
                MsgBox(ex.Message, MsgBoxStyle.Critical, "ERROR")
            End Try
        End If
    End Sub

    Private Sub EditToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles EditToolStripMenuItem.Click
        Call LoadEditFee()
    End Sub

    Public Sub LoadEditFee()
        If DataGridView1.RowCount = "0" Then
            'do nothing
        Else
            Dim child As New map_EditFee
            child.MdiParent = Home
            child.Show()

            Try
                Dim conn As New OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & My.Settings.dbloc & "';Jet OLEDB:Database Password='" & My.Settings.dbpass & "';")

                Dim strSQL As String = "SELECT * FROM map_Fees WHERE ID = " & DataGridView1.SelectedCells(0).Value
                Dim queryString As String = String.Format(strSQL)
                Dim cmd As New OleDb.OleDbCommand(queryString, conn)
                Dim da As New OleDb.OleDbDataAdapter(cmd)
                Dim ds As New DataSet

                da.Fill(ds, "User")
                Dim dt As DataTable = ds.Tables("User")
                Dim row1 As DataRow = dt.Rows(0)

                child.ID.Text = (row1("ID"))

                If IsDBNull(row1("ProductID")) Then
                    'do nothing
                Else
                    child.cboProduct.SelectedValue = (row1("ProductID"))
                End If

                If IsDBNull(row1("BreakpointFrom")) Then
                    child.BP1.Text = ""
                Else
                    child.BP1.Text = (row1("BreakpointFrom"))
                End If

                If IsDBNull(row1("BreakpointTo")) Then
                    child.BP2.Text = ""
                Else
                    child.BP2.Text = (row1("BreakpointTo"))
                End If

                If IsDBNull(row1("Fee")) Then
                    child.Fee.Text = ""
                Else
                    child.Fee.Text = (row1("Fee"))
                End If

                If IsDBNull(row1("MaxRepFee")) Then
                    child.RepFee.Text = ""
                Else
                    child.RepFee.Text = (row1("MaxRepFee"))
                End If

                If IsDBNull(row1("AgreementTypeID")) Then
                    'do nothing
                Else
                    child.cboAType.SelectedValue = (row1("AgreementTypeID"))
                End If

                child.ckb40Act.Checked = (row1("40Act"))
                child.ckbDefaultFee.Checked = (row1("Default"))
                child.ckb40Act.Enabled = False
                child.ckbDefaultFee.Enabled = False

            Catch ex As Exception
                MsgBox(ex.Message, MsgBoxStyle.Critical, "ERROR")
            End Try
        End If
    End Sub
End Class