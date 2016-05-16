Public Class map_PlatformProducts

    Private Sub Button5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button5.Click
        Call LoadProducts()
    End Sub

    Public Sub LoadProducts()
        Try
            Dim conn As New OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & My.Settings.dbloc & "';Jet OLEDB:Database Password='" & My.Settings.dbpass & "';")

            Dim SqlString As String

            If ckbWLPlatform.Checked Then
                If CheckBox1.Checked Then
                    SqlString = "SELECT map_Products.ID, map_Products.ProductName As [Product Name], Map_Firms.FirmName As [Managing Firm], map_ProductType.TypeName As [Product Type]" & _
                    " FROM (map_Products INNER JOIN Map_Firms ON map_Products.ManagingFirmID = Map_Firms.ID) INNER JOIN map_ProductType ON map_Products.TypeOfProductID = map_ProductType.ID" & _
                    " WHERE map_Products.Active = -1 AND map_Products.TypeOfProductID = " & cboProductTypes.SelectedValue & " AND map_Products.ProductName Like '%" & TextBox1.Text & "%' AND map_Products.ID In (SELECT map_PlatformListings.ProductID FROM map_PlatformListings WHERE map_PlatformListings.PlatformID = " & ID.Text & ") AND map_Products.ID Not In (SELECT map_PlatformListings.ProductID FROM map_PlatformListings WHERE map_PlatformListings.WLID = " & WLID.Text & ")" & _
                    " GROUP BY map_Products.ID, map_Products.ProductName, Map_Firms.FirmName, map_ProductType.TypeName;"

                Else
                    SqlString = "SELECT map_Products.ID, map_Products.ProductName As [Product Name], Map_Firms.FirmName As [Managing Firm], map_ProductType.TypeName As [Product Type]" & _
                    " FROM (map_Products INNER JOIN Map_Firms ON map_Products.ManagingFirmID = Map_Firms.ID) INNER JOIN map_ProductType ON map_Products.TypeOfProductID = map_ProductType.ID" & _
                    " WHERE map_Products.Active = -1 AND map_Products.ProductName Like '%" & TextBox1.Text & "%' AND map_Products.ID In (SELECT map_PlatformListings.ProductID FROM map_PlatformListings WHERE map_PlatformListings.PlatformID = " & ID.Text & ") AND map_Products.ID Not In (SELECT map_PlatformListings.ProductID FROM map_PlatformListings WHERE map_PlatformListings.WLID = " & WLID.Text & ")" & _
                    " GROUP BY map_Products.ID, map_Products.ProductName, Map_Firms.FirmName, map_ProductType.TypeName;"

                    'SqlString = "SELECT map_Products.ID, map_Products.ProductName As [Product Name], map_ProductType.TypeName As [Product Type], Map_Firms.FirmName As [Managing Firm]" & _
                    '" FROM ((map_Products INNER JOIN AdvApp_vAssetClass ON map_Products.AssetClassCode = AdvApp_vAssetClass.AssetClassCode) INNER JOIN map_ProductType ON map_Products.TypeOfProductID = map_ProductType.ID) INNER JOIN Map_Firms ON map_Products.ManagingFirmID = Map_Firms.ID" & _
                    '" WHERE map_Products.Active = -1 AND map_Products.ProductName Like '%" & TextBox1.Text & "%' AND map_Products.ID Not In (SELECT map_ProductListing.ProductID FROM map_ProductListing WHERE map_ProductListing.AgreementID = " & ID.Text & ")"
                End If
            Else
                If CheckBox1.Checked Then
                    'look for product types

                    SqlString = "SELECT map_Products.ID, map_Products.ProductName As [Product Name], Map_Firms.FirmName As [Managing Firm], map_ProductType.TypeName As [Product Type]" & _
                    " FROM (map_Products INNER JOIN Map_Firms ON map_Products.ManagingFirmID = Map_Firms.ID) INNER JOIN map_ProductType ON map_Products.TypeOfProductID = map_ProductType.ID" & _
                    " WHERE map_Products.Active = -1 AND map_Products.TypeOfProductID = " & cboProductTypes.SelectedValue & " AND map_Products.ProductName Like '%" & TextBox1.Text & "%' AND map_Products.ID Not In (SELECT map_PlatformListings.ProductID FROM map_PlatformListings WHERE map_PlatformListings.PlatformID = " & ID.Text & ")" & _
                    " GROUP BY map_Products.ID, map_Products.ProductName, Map_Firms.FirmName, map_ProductType.TypeName;"

                    'SqlString = "SELECT map_Products.ID, map_Products.ProductName As [Product Name], map_ProductType.TypeName As [Product Type], Map_Firms.FirmName As [Managing Firm]" & _
                    '" FROM ((map_Products INNER JOIN AdvApp_vAssetClass ON map_Products.AssetClassCode = AdvApp_vAssetClass.AssetClassCode) INNER JOIN map_ProductType ON map_Products.TypeOfProductID = map_ProductType.ID) INNER JOIN Map_Firms ON map_Products.ManagingFirmID = Map_Firms.ID" & _
                    '" WHERE map_Products.Active = -1 AND map_Products.TypeOfProductID = " & cboProductTypes.SelectedValue & " AND map_Products.ProductName Like '%" & TextBox1.Text & "%' AND map_Products.ID Not In (SELECT map_ProductListing.ProductID FROM map_ProductListing WHERE map_ProductListing.AgreementID = " & ID.Text & ")"
                Else
                    SqlString = "SELECT map_Products.ID, map_Products.ProductName As [Product Name], Map_Firms.FirmName As [Managing Firm], map_ProductType.TypeName As [Product Type]" & _
                    " FROM (map_Products INNER JOIN Map_Firms ON map_Products.ManagingFirmID = Map_Firms.ID) INNER JOIN map_ProductType ON map_Products.TypeOfProductID = map_ProductType.ID" & _
                    " WHERE map_Products.Active = -1 AND map_Products.ProductName Like '%" & TextBox1.Text & "%' AND map_Products.ID Not In (SELECT map_PlatformListings.ProductID FROM map_PlatformListings WHERE map_PlatformListings.PlatformID = " & ID.Text & ")" & _
                    " GROUP BY map_Products.ID, map_Products.ProductName, Map_Firms.FirmName, map_ProductType.TypeName;"

                    'SqlString = "SELECT map_Products.ID, map_Products.ProductName As [Product Name], map_ProductType.TypeName As [Product Type], Map_Firms.FirmName As [Managing Firm]" & _
                    '" FROM ((map_Products INNER JOIN AdvApp_vAssetClass ON map_Products.AssetClassCode = AdvApp_vAssetClass.AssetClassCode) INNER JOIN map_ProductType ON map_Products.TypeOfProductID = map_ProductType.ID) INNER JOIN Map_Firms ON map_Products.ManagingFirmID = Map_Firms.ID" & _
                    '" WHERE map_Products.Active = -1 AND map_Products.ProductName Like '%" & TextBox1.Text & "%' AND map_Products.ID Not In (SELECT map_ProductListing.ProductID FROM map_ProductListing WHERE map_ProductListing.AgreementID = " & ID.Text & ")"
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

    Private Sub Button6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button6.Click
        Call LoadAssigned()
    End Sub

    Public Sub LoadAssigned()
        Try
            Dim conn As New OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & My.Settings.dbloc & "';Jet OLEDB:Database Password='" & My.Settings.dbpass & "';")

            Dim SqlString As String

            If ckbWLPlatform.Checked Then
                'SqlString = "SELECT map_PlatformListings.ID, map_Products.ProductName, map_ProductType.TypeName, Map_Firms.FirmName, map_PlatformListings.PlatformApproved, map_PlatformListings.WLApproved, map_PlatformListings.SMAOffered, map_PlatformListings.UMAOffered" & _
                '" FROM ((map_PlatformListings INNER JOIN map_Products ON map_PlatformListings.ProductID = map_Products.ID) INNER JOIN map_ProductType ON map_Products.TypeOfProductID = map_ProductType.ID) INNER JOIN Map_Firms ON map_Products.ManagingFirmID = Map_Firms.ID" & _
                '" WHERE(((map_PlatformListings.Active) = True) And ((map_PlatformListings.WLID) = " & WLID.Text & "))" & _
                '" GROUP BY map_PlatformListings.ID, map_Products.ProductName, map_ProductType.TypeName, Map_Firms.FirmName, map_PlatformListings.PlatformApproved, map_PlatformListings.WLApproved, map_PlatformListings.SMAOffered, map_PlatformListings.UMAOffered;"

                SqlString = "SELECT map_PlatformListings.ID, map_Products.ProductName As [Product Name], map_ProductType.TypeName As [Product Type], Map_Firms.FirmName As [Managing Firm], map_PlatformListings.PlatformApproved As [Platform Approved], map_PlatformListings.WLApproved As [Whitelabel Approved], map_PlatformListings.SMAOffered As [Offered as SMA], map_PlatformListings.UMAOffered As [Offered as UMA]" & _
                " FROM ((map_PlatformListings INNER JOIN map_Products ON map_PlatformListings.ProductID = map_Products.ID) INNER JOIN map_ProductType ON map_Products.TypeOfProductID = map_ProductType.ID) INNER JOIN Map_Firms ON map_Products.ManagingFirmID = Map_Firms.ID" & _
                " WHERE(((map_PlatformListings.Active) = True) And ((map_PlatformListings.WLID) = " & WLID.Text & "))" & _
                " GROUP BY map_PlatformListings.ID, map_Products.ProductName, map_ProductType.TypeName, Map_Firms.FirmName, map_PlatformListings.PlatformApproved, map_PlatformListings.WLApproved, map_PlatformListings.SMAOffered, map_PlatformListings.UMAOffered;"

            Else
                'SqlString = "SELECT map_PlatformListings.ID, map_Products.ProductName, map_ProductType.TypeName, Map_Firms.FirmName, map_PlatformListings.PlatformApproved, map_PlatformListings.SMAOffered, map_PlatformListings.UMAOffered" & _
                '" FROM ((map_PlatformListings INNER JOIN map_Products ON map_PlatformListings.ProductID = map_Products.ID) INNER JOIN map_ProductType ON map_Products.TypeOfProductID = map_ProductType.ID) INNER JOIN Map_Firms ON map_Products.ManagingFirmID = Map_Firms.ID" & _
                '" WHERE(((map_PlatformListings.Active) = True) And ((map_PlatformListings.PlatformID) = " & ID.Text & "))" & _
                '" GROUP BY map_PlatformListings.ID, map_Products.ProductName, map_ProductType.TypeName, Map_Firms.FirmName, map_PlatformListings.PlatformApproved, map_PlatformListings.SMAOffered, map_PlatformListings.UMAOffered;"

                SqlString = "SELECT map_PlatformListings.ID, map_Products.ProductName As [Product Name], map_ProductType.TypeName As [Product Type], Map_Firms.FirmName As [Managing Firm], map_PlatformListings.PlatformApproved As [Platform Approved], map_PlatformListings.SMAOffered As [Offered as SMA], map_PlatformListings.UMAOffered As [Offered as UMA]" & _
                " FROM ((map_PlatformListings INNER JOIN map_Products ON map_PlatformListings.ProductID = map_Products.ID) INNER JOIN map_ProductType ON map_Products.TypeOfProductID = map_ProductType.ID) INNER JOIN Map_Firms ON map_Products.ManagingFirmID = Map_Firms.ID" & _
                " WHERE(((map_PlatformListings.Active) = True) AND map_PlatformListings.WLID Is Null And ((map_PlatformListings.PlatformID) = " & ID.Text & "))" & _
                " GROUP BY map_PlatformListings.ID, map_Products.ProductName, map_ProductType.TypeName, Map_Firms.FirmName, map_PlatformListings.PlatformApproved, map_PlatformListings.SMAOffered, map_PlatformListings.UMAOffered;"


            End If

            Dim da As New OleDb.OleDbDataAdapter(SqlString, conn)
            Dim ds As New DataSet
            da.Fill(ds, "Users")

            With DataGridView2
                .DataSource = ds.Tables("Users")
                .Columns(0).Visible = False
            End With

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "ERROR")
        End Try
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

                If ckbWLPlatform.Checked Then
                    SQLstr = "INSERT INTO map_PlatformListings([PlatformID], [WLID], [ProductID], [WLApproved], [Active])" & _
                    "VALUES(" & ID.Text & "," & WLID.Text & "," & DataGridView1.SelectedCells(0).Value & ",-1,-1)"
                Else
                    SQLstr = "INSERT INTO map_PlatformListings([PlatformID], [ProductID], [Active])" & _
                    "VALUES(" & ID.Text & "," & DataGridView1.SelectedCells(0).Value & ",-1)"
                End If

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

                SQLstr = "DELETE * FROM map_PlatformListings WHERE ID = " & DataGridView2.SelectedCells(0).Value

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

    Public Sub MoveAllToAssigned()
        Dim Mycn As OleDb.OleDbConnection
        Dim Command As OleDb.OleDbCommand
        Dim SQLstr As String
        Try
            Mycn = New OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & My.Settings.dbloc & "';Jet OLEDB:Database Password='" & My.Settings.dbpass & "';")
            Mycn.Open()

            If ckbWLPlatform.Checked Then
                If CheckBox1.Checked Then
                    SQLstr = "INSERT INTO map_PlatformListings([PlatformID], [WLID], [ProductID], [WLApproved], [Active])" & _
                    "SELECT " & ID.Text & "," & WLID.Text & ", ID,-1,-1" & _
                    " FROM map_Products" & _
                    " WHERE map_Products.Active = -1 AND map_Products.TypeOfProductID = " & cboProductTypes.SelectedValue & " AND map_Products.ProductName Like '%" & TextBox1.Text & "%' AND map_Products.ID In (SELECT map_PlatformListings.ProductID FROM map_PlatformListings WHERE map_PlatformListings.PlatformID = " & ID.Text & ") AND map_Products.ID Not In (SELECT map_PlatformListings.ProductID FROM map_PlatformListings WHERE map_PlatformListings.WLID = " & WLID.Text & ")"
                Else
                    SQLstr = "INSERT INTO map_PlatformListings([PlatformID], [WLID], [ProductID], [WLApproved], [Active])" & _
                    "SELECT " & ID.Text & "," & WLID.Text & ", ID,-1,-1" & _
                    " FROM map_Products" & _
                    " WHERE map_Products.Active = -1 AND map_Products.ProductName Like '%" & TextBox1.Text & "%' AND map_Products.ID In (SELECT map_PlatformListings.ProductID FROM map_PlatformListings WHERE map_PlatformListings.PlatformID = " & ID.Text & ") AND map_Products.ID Not In (SELECT map_PlatformListings.ProductID FROM map_PlatformListings WHERE map_PlatformListings.WLID = " & WLID.Text & ")"
                End If

            Else
                If CheckBox1.Checked Then
                    SQLstr = "INSERT INTO map_PlatformListings([PlatformID], [ProductID], [Active])" & _
                    "SELECT " & ID.Text & ", ID, -1" & _
                    " FROM map_Products" & _
                    " WHERE map_Products.Active = -1 AND map_Products.TypeOfProductID = " & cboProductTypes.SelectedValue & " AND map_Products.ProductName Like '%" & TextBox1.Text & "%' AND map_Products.ID Not In (SELECT map_PlatformListings.ProductID FROM map_PlatformListings WHERE map_PlatformListings.PlatformID = " & ID.Text & ")"
                Else
                    SQLstr = "INSERT INTO map_PlatformListings([PlatformID], [ProductID], [Active])" & _
                    "SELECT " & ID.Text & ", ID, -1" & _
                    " FROM map_Products" & _
                    " WHERE map_Products.Active = -1 AND map_Products.ProductName Like '%" & TextBox1.Text & "%' AND map_Products.ID Not In (SELECT map_PlatformListings.ProductID FROM map_PlatformListings WHERE map_PlatformListings.PlatformID = " & ID.Text & ")"
                End If

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

    Public Sub RemoveAllFromAssigned()
        Dim Mycn As OleDb.OleDbConnection
        Dim Command As OleDb.OleDbCommand
        Dim SQLstr As String
        Try
            Mycn = New OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & My.Settings.dbloc & "';Jet OLEDB:Database Password='" & My.Settings.dbpass & "';")
            Mycn.Open()

            If ckbWLPlatform.Checked Then
                SQLstr = "DELETE * FROM map_PlatformListings WHERE WLID = " & WLID.Text
            Else
                SQLstr = "DELETE * FROM map_PlatformListings WHERE PlatformID = " & ID.Text
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

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        Call MoveAllToAssigned()
    End Sub

    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        Call RemoveFromAssignedOne()
    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        Call RemoveAllFromAssigned()
    End Sub

    Private Sub DataGridView1_CellContentClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick

    End Sub

    Private Sub DataGridView1_CellContentDoubleClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView1.CellContentDoubleClick
        Call MovetoAssignedOne()
    End Sub

    Private Sub ChangeApprovalStatusToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ChangeApprovalStatusToolStripMenuItem.Click
        Call ChangeApprovalStatus()
    End Sub

    Public Sub ChangeApprovalStatus()
        If DataGridView2.RowCount = "0" Then
            'do nothing
        Else
            Dim Mycn As OleDb.OleDbConnection
            Dim Command As OleDb.OleDbCommand
            Dim SQLstr As String
            Try
                Mycn = New OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & My.Settings.dbloc & "';Jet OLEDB:Database Password='" & My.Settings.dbpass & "';")
                Mycn.Open()

                If ckbWLPlatform.Checked Then
                    If DataGridView2.SelectedCells(5).Value = True Then
                        SQLstr = "UPDATE map_PlatformListings SET WLApproved = False WHERE ID = " & DataGridView2.SelectedCells(0).Value
                    Else
                        SQLstr = "UPDATE map_PlatformListings SET WLApproved = True WHERE ID = " & DataGridView2.SelectedCells(0).Value
                    End If
                Else
                    If DataGridView2.SelectedCells(4).Value = True Then
                        SQLstr = "UPDATE map_PlatformListings SET PlatformApproved = False WHERE ID = " & DataGridView2.SelectedCells(0).Value
                    Else
                        SQLstr = "UPDATE map_PlatformListings SET PlatformApproved = TRUE WHERE ID = " & DataGridView2.SelectedCells(0).Value
                    End If
                End If


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

    Public Sub ChangeSMAStatus()
        If DataGridView2.RowCount = "0" Then
            'do nothing
        Else
            Dim Mycn As OleDb.OleDbConnection
            Dim Command As OleDb.OleDbCommand
            Dim SQLstr As String
            Try
                Mycn = New OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & My.Settings.dbloc & "';Jet OLEDB:Database Password='" & My.Settings.dbpass & "';")
                Mycn.Open()

                If ckbWLPlatform.Checked Then
                    If DataGridView2.SelectedCells(6).Value = True Then
                        SQLstr = "UPDATE map_PlatformListings SET SMAOffered = False WHERE ID = " & DataGridView2.SelectedCells(0).Value
                    Else
                        SQLstr = "UPDATE map_PlatformListings SET SMAOffered = True WHERE ID = " & DataGridView2.SelectedCells(0).Value
                    End If
                Else
                    If DataGridView2.SelectedCells(5).Value = True Then
                        SQLstr = "UPDATE map_PlatformListings SET SMAOffered = False WHERE ID = " & DataGridView2.SelectedCells(0).Value
                    Else
                        SQLstr = "UPDATE map_PlatformListings SET SMAOffered = TRUE WHERE ID = " & DataGridView2.SelectedCells(0).Value
                    End If
                End If

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

    Public Sub ChangeUMAStatus()
        If DataGridView2.RowCount = "0" Then
            'do nothing
        Else
            Dim Mycn As OleDb.OleDbConnection
            Dim Command As OleDb.OleDbCommand
            Dim SQLstr As String
            Try
                Mycn = New OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & My.Settings.dbloc & "';Jet OLEDB:Database Password='" & My.Settings.dbpass & "';")
                Mycn.Open()

                If ckbWLPlatform.Checked Then
                    If DataGridView2.SelectedCells(7).Value = True Then
                        SQLstr = "UPDATE map_PlatformListings SET UMAOffered = False WHERE ID = " & DataGridView2.SelectedCells(0).Value
                    Else
                        SQLstr = "UPDATE map_PlatformListings SET UMAOffered = True WHERE ID = " & DataGridView2.SelectedCells(0).Value
                    End If
                Else
                    If DataGridView2.SelectedCells(6).Value = True Then
                        SQLstr = "UPDATE map_PlatformListings SET UMAOffered = False WHERE ID = " & DataGridView2.SelectedCells(0).Value
                    Else
                        SQLstr = "UPDATE map_PlatformListings SET UMAOffered = TRUE WHERE ID = " & DataGridView2.SelectedCells(0).Value
                    End If
                End If

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

    Private Sub ChangeSMAStatusToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ChangeSMAStatusToolStripMenuItem.Click
        Call ChangeSMAStatus()
    End Sub

    Private Sub ChangeUMAStatusToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ChangeUMAStatusToolStripMenuItem.Click
        Call ChangeUMAStatus()
    End Sub
End Class