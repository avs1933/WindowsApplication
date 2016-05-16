Public Class map_AgreementProducts

    Private Sub map_AgreementProducts_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Call LoadTypes()
        Call LoadAssigned()
    End Sub

    Public Sub LoadTypes()
        Try

            Dim conn As New OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & My.Settings.dbloc & "';Jet OLEDB:Database Password='" & My.Settings.dbpass & "';")

            Dim strSQL As String = "SELECT ID, TypeName FROM map_ProductType WHERE Active = -1"
            Dim da As New OleDb.OleDbDataAdapter(strSQL, conn)
            Dim ds As New DataSet
            da.Fill(ds, "Users")

            With cboProductTypes
                .DataSource = ds.Tables("Users")
                .DisplayMember = "TypeName"
                .ValueMember = "ID"
                .SelectedIndex = 0
            End With

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "ERROR")
        End Try
    End Sub

    Public Sub MovetoAssignedOne()
        If DataGridView1.RowCount = "0" Then
            'do nothing
        Else
            Dim Mycn As OleDb.OleDbConnection
            Dim Command As OleDb.OleDbCommand
            Dim SQLstr As String
            Try
                Mycn = New OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & My.Settings.dbloc & "';Jet OLEDB:Database Password='" & My.Settings.dbpass & "';")
                Mycn.Open()

                SQLstr = "INSERT INTO map_ProductListing([ProductID], [AgreementID], [Agreement])" & _
                "VALUES(" & DataGridView1.SelectedCells(0).Value & "," & ID.Text & ", -1)"

                Command = New OleDb.OleDbCommand(SQLstr, Mycn)
                Command.ExecuteNonQuery()

                Mycn.Close()

                Call LoadAssigned()
                Call LoadProducts()

            Catch ex As Exception
                MsgBox(ex.Message, MsgBoxStyle.Critical, "ERROR")
            End Try
        End If
    End Sub

    Public Sub RemoveFromAssignedOne()
        If DataGridView2.RowCount = "0" Then
            'do nothing
        Else
            Dim Mycn As OleDb.OleDbConnection
            Dim Command As OleDb.OleDbCommand
            Dim SQLstr As String
            Try
                Mycn = New OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & My.Settings.dbloc & "';Jet OLEDB:Database Password='" & My.Settings.dbpass & "';")
                Mycn.Open()

                SQLstr = "DELETE * FROM map_ProductListing WHERE ID = " & DataGridView2.SelectedCells(0).Value

                Command = New OleDb.OleDbCommand(SQLstr, Mycn)
                Command.ExecuteNonQuery()

                Mycn.Close()

                Call LoadAssigned()
                Call LoadProducts()

            Catch ex As Exception
                MsgBox(ex.Message, MsgBoxStyle.Critical, "ERROR")
            End Try
        End If
    End Sub

    Public Sub LoadAssigned()
        Try
            Dim conn As New OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & My.Settings.dbloc & "';Jet OLEDB:Database Password='" & My.Settings.dbpass & "';")

            Dim SqlString As String


            SqlString = "SELECT map_ProductListing.ID, map_Products.ID, map_Products.ProductName As [Product Name], map_ProductType.TypeName As [Product Type], Map_Firms.FirmName As [Managing Firm]" & _
            " FROM map_ProductListing INNER JOIN ((map_Products INNER JOIN map_ProductType ON map_Products.TypeOfProductID = map_ProductType.ID) INNER JOIN Map_Firms ON map_Products.ManagingFirmID = Map_Firms.ID) ON map_ProductListing.ProductID = map_Products.ID" & _
            " WHERE(((map_Products.Active) = True) And ((map_ProductListing.AgreementID) = " & ID.Text & ") And ((map_ProductListing.Agreement) = True))" & _
            " GROUP BY map_Products.ID, map_Products.ProductName, map_ProductType.TypeName, Map_Firms.FirmName, map_ProductListing.ID;"

            Dim da As New OleDb.OleDbDataAdapter(SqlString, conn)
            Dim ds As New DataSet
            da.Fill(ds, "Users")

            With DataGridView2
                .DataSource = ds.Tables("Users")
                .Columns(0).Visible = False
                .Columns(1).Visible = False
            End With

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "ERROR")
        End Try
    End Sub

    Public Sub LoadProducts()
        Try
            Dim conn As New OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & My.Settings.dbloc & "';Jet OLEDB:Database Password='" & My.Settings.dbpass & "';")

            Dim SqlString As String

            'Look for active records
            If CheckBox1.Checked Then
                'look for product types
                SqlString = "SELECT map_Products.ID, map_Products.ProductName As [Product Name], map_ProductType.TypeName As [Product Type], Map_Firms.FirmName As [Managing Firm]" & _
                " FROM ((map_Products INNER JOIN AdvApp_vAssetClass ON map_Products.AssetClassCode = AdvApp_vAssetClass.AssetClassCode) INNER JOIN map_ProductType ON map_Products.TypeOfProductID = map_ProductType.ID) INNER JOIN Map_Firms ON map_Products.ManagingFirmID = Map_Firms.ID" & _
                " WHERE map_Products.Active = -1 AND map_Products.TypeOfProductID = " & cboProductTypes.SelectedValue & " AND map_Products.ProductName Like '%" & TextBox1.Text & "%' AND map_Products.ID Not In (SELECT map_ProductListing.ProductID FROM map_ProductListing WHERE map_ProductListing.AgreementID = " & ID.Text & ")"
            Else
                SqlString = "SELECT map_Products.ID, map_Products.ProductName As [Product Name], map_ProductType.TypeName As [Product Type], Map_Firms.FirmName As [Managing Firm]" & _
                " FROM ((map_Products INNER JOIN AdvApp_vAssetClass ON map_Products.AssetClassCode = AdvApp_vAssetClass.AssetClassCode) INNER JOIN map_ProductType ON map_Products.TypeOfProductID = map_ProductType.ID) INNER JOIN Map_Firms ON map_Products.ManagingFirmID = Map_Firms.ID" & _
                " WHERE map_Products.Active = -1 AND map_Products.ProductName Like '%" & TextBox1.Text & "%' AND map_Products.ID Not In (SELECT map_ProductListing.ProductID FROM map_ProductListing WHERE map_ProductListing.AgreementID = " & ID.Text & ")"
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

    Private Sub Button6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button6.Click
        Call LoadAssigned()
    End Sub

    Private Sub Button5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button5.Click
        Call LoadProducts()
    End Sub

    Private Sub CheckBox1_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckBox1.CheckedChanged
        If CheckBox1.Checked Then
            cboProductTypes.Enabled = True
        Else
            cboProductTypes.Enabled = False
        End If
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Call MovetoAssignedOne()
    End Sub

    Private Sub DataGridView1_CellContentClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick

    End Sub

    Private Sub DataGridView1_CellContentDoubleClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView1.CellContentDoubleClick
        Call MovetoAssignedOne()
    End Sub

    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        Call RemoveFromAssignedOne()
    End Sub

    Private Sub DataGridView2_CellContentClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView2.CellContentClick

    End Sub

    Private Sub DataGridView2_CellContentDoubleClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView2.CellContentDoubleClick
        Call RemoveFromAssignedOne()
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        Call MoveAllToAssigned()
    End Sub

    Public Sub MoveAllToAssigned()
        Dim Mycn As OleDb.OleDbConnection
        Dim Command As OleDb.OleDbCommand
        Dim SQLstr As String
        Try
            Mycn = New OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & My.Settings.dbloc & "';Jet OLEDB:Database Password='" & My.Settings.dbpass & "';")
            Mycn.Open()

            If CheckBox1.Checked Then
                SQLstr = "INSERT INTO map_ProductListing([ProductID], [AgreementID], [Agreement])" & _
                "SELECT ID, " & ID.Text & ", -1" & _
                " FROM map_Products" & _
                " WHERE map_Products.Active = -1 AND map_Products.TypeOfProductID = " & cboProductTypes.SelectedValue & " AND map_Products.ProductName Like '%" & TextBox1.Text & "%' AND map_Products.ID Not In (SELECT map_ProductListing.ProductID FROM map_ProductListing WHERE map_ProductListing.AgreementID = " & ID.Text & ")"
            Else
                SQLstr = "INSERT INTO map_ProductListing([ProductID], [AgreementID], [Agreement])" & _
                "SELECT ID, " & ID.Text & ", -1" & _
                " FROM map_Products" & _
                " WHERE map_Products.Active = -1 AND map_Products.ProductName Like '%" & TextBox1.Text & "%' AND map_Products.ID Not In (SELECT map_ProductListing.ProductID FROM map_ProductListing WHERE map_ProductListing.AgreementID = " & ID.Text & ")"
            End If

            Command = New OleDb.OleDbCommand(SQLstr, Mycn)
            Command.ExecuteNonQuery()

            Mycn.Close()

            Call LoadAssigned()
            Call LoadProducts()

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "ERROR")
        End Try
    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        Call RemoveAllFromAssigned()
    End Sub

    Public Sub RemoveAllFromAssigned()
        Dim Mycn As OleDb.OleDbConnection
        Dim Command As OleDb.OleDbCommand
        Dim SQLstr As String
        Try
            Mycn = New OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & My.Settings.dbloc & "';Jet OLEDB:Database Password='" & My.Settings.dbpass & "';")
            Mycn.Open()

            SQLstr = "DELETE * FROM map_ProductListing WHERE AgreementID = " & ID.Text

            Command = New OleDb.OleDbCommand(SQLstr, Mycn)
            Command.ExecuteNonQuery()

            Mycn.Close()

            Call LoadAssigned()
            Call LoadProducts()

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "ERROR")
        End Try
    End Sub
End Class