Public Class map_ViewProducts

    Private Sub map_ViewProducts_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Call LoadProductType()
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

            'Call LoadDBData()

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "ERROR")
        End Try
    End Sub

    Private Sub CheckBox2_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckBox2.CheckedChanged
        If CheckBox2.Checked Then
            cboProductType.Enabled = True
        Else
            cboProductType.Enabled = False
        End If
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Try
            Dim conn As New OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & My.Settings.dbloc & "';Jet OLEDB:Database Password='" & My.Settings.dbpass & "';")

            Dim SqlString As String

            If CheckBox1.Checked Then
                'Look for inactive records
                If CheckBox2.Checked Then
                    'Look for product types
                    SqlString = "SELECT map_Products.ID, map_Products.TypeOfProductID, map_Products.ManagingFirmID, map_Products.AssetClassCode, map_Products.PrimaryObjectiveID, map_Products.SecondaryObjectiveID, map_Products.ProductName, map_ProductType.TypeName, Map_Firms.FirmName, AdvApp_vAssetClass.AssetClassName, map_Products.Symbol, map_Products.CUSIP, map_Products.Active" & _
                    " FROM ((map_Products INNER JOIN AdvApp_vAssetClass ON map_Products.AssetClassCode = AdvApp_vAssetClass.AssetClassCode) INNER JOIN map_ProductType ON map_Products.TypeOfProductID = map_ProductType.ID) INNER JOIN Map_Firms ON map_Products.ManagingFirmID = Map_Firms.ID" & _
                    " WHERE map_Products.Active <> -1 AND map_Products.TypeOfProductID = " & cboProductType.SelectedValue & " AND map_Products.ProductName Like '%" & TextBox1.Text & "%'"
                Else
                    SqlString = "SELECT map_Products.ID, map_Products.TypeOfProductID, map_Products.ManagingFirmID, map_Products.AssetClassCode, map_Products.PrimaryObjectiveID, map_Products.SecondaryObjectiveID, map_Products.ProductName, map_ProductType.TypeName, Map_Firms.FirmName, AdvApp_vAssetClass.AssetClassName, map_Products.Symbol, map_Products.CUSIP, map_Products.Active" & _
                    " FROM ((map_Products INNER JOIN AdvApp_vAssetClass ON map_Products.AssetClassCode = AdvApp_vAssetClass.AssetClassCode) INNER JOIN map_ProductType ON map_Products.TypeOfProductID = map_ProductType.ID) INNER JOIN Map_Firms ON map_Products.ManagingFirmID = Map_Firms.ID" & _
                    " WHERE map_Products.Active <> -1 AND map_Products.ProductName Like '%" & TextBox1.Text & "%'"
                End If
            Else
                'Look for active records
                If CheckBox2.Checked Then
                    'look for product types
                    SqlString = "SELECT map_Products.ID, map_Products.TypeOfProductID, map_Products.ManagingFirmID, map_Products.AssetClassCode, map_Products.PrimaryObjectiveID, map_Products.SecondaryObjectiveID, map_Products.ProductName, map_ProductType.TypeName, Map_Firms.FirmName, AdvApp_vAssetClass.AssetClassName, map_Products.Symbol, map_Products.CUSIP, map_Products.Active" & _
                    " FROM ((map_Products INNER JOIN AdvApp_vAssetClass ON map_Products.AssetClassCode = AdvApp_vAssetClass.AssetClassCode) INNER JOIN map_ProductType ON map_Products.TypeOfProductID = map_ProductType.ID) INNER JOIN Map_Firms ON map_Products.ManagingFirmID = Map_Firms.ID" & _
                    " WHERE map_Products.Active = -1 AND map_Products.TypeOfProductID = " & cboProductType.SelectedValue & " AND map_Products.ProductName Like '%" & TextBox1.Text & "%'"
                Else
                    SqlString = "SELECT map_Products.ID, map_Products.TypeOfProductID, map_Products.ManagingFirmID, map_Products.AssetClassCode, map_Products.PrimaryObjectiveID, map_Products.SecondaryObjectiveID, map_Products.ProductName, map_ProductType.TypeName, Map_Firms.FirmName, AdvApp_vAssetClass.AssetClassName, map_Products.Symbol, map_Products.CUSIP, map_Products.Active" & _
                    " FROM ((map_Products INNER JOIN AdvApp_vAssetClass ON map_Products.AssetClassCode = AdvApp_vAssetClass.AssetClassCode) INNER JOIN map_ProductType ON map_Products.TypeOfProductID = map_ProductType.ID) INNER JOIN Map_Firms ON map_Products.ManagingFirmID = Map_Firms.ID" & _
                    " WHERE map_Products.Active = -1 AND map_Products.ProductName Like '%" & TextBox1.Text & "%'"
                End If
            End If

            'Dim strSQL As String = "SELECT * FROM map_Objective"

            Dim da As New OleDb.OleDbDataAdapter(SqlString, conn)
            Dim ds As New DataSet
            da.Fill(ds, "Users")

            With DataGridView1
                .DataSource = ds.Tables("Users")
                .Columns(0).Visible = False
                .Columns(1).Visible = False
                .Columns(2).Visible = False
                .Columns(3).Visible = False
                .Columns(4).Visible = False
                .Columns(5).Visible = False
            End With

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "ERROR")
        End Try

    End Sub

    Public Sub LoadEditForm()
        If DataGridView1.RowCount = "0" Then
            'do nothing
        Else
            Dim child As New map_SMA
            child.MdiParent = Home
            child.Show()

            Try
                Dim conn As New OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & My.Settings.dbloc & "';Jet OLEDB:Database Password='" & My.Settings.dbpass & "';")

                Dim strSQL As String = "SELECT * FROM map_Products WHERE ID = " & DataGridView1.SelectedCells(0).Value
                Dim queryString As String = String.Format(strSQL)
                Dim cmd As New OleDb.OleDbCommand(queryString, conn)
                Dim da As New OleDb.OleDbDataAdapter(cmd)
                Dim ds As New DataSet

                da.Fill(ds, "User")
                Dim dt As DataTable = ds.Tables("User")
                Dim row1 As DataRow = dt.Rows(0)

                child.ProductID.Text = (row1("ID"))

                If IsDBNull(row1("ProductName")) Then
                    child.TextBox2.Text = ""
                Else
                    child.TextBox2.Text = (row1("ProductName"))
                End If

                If IsDBNull(row1("TypeOfProductID")) Then
                    'do nothing
                Else
                    child.cboProductType.SelectedValue = (row1("TypeOfProductID"))
                End If

                If IsDBNull(row1("ManagingFirmID")) Then
                    'do nothing
                Else
                    child.cboFirms.SelectedValue = (row1("ManagingFirmID"))
                End If

                If IsDBNull(row1("AssetClassCode")) Then
                    'do nothing
                Else
                    child.cboAssetClass.SelectedValue = (row1("AssetClassCode"))
                End If

                If IsDBNull(row1("PrimaryObjectiveID")) Then
                    'do nothing
                Else
                    child.cboObjective1.SelectedValue = (row1("PrimaryObjectiveID"))
                End If

                If IsDBNull(row1("SecondaryObjectiveID")) Then
                    'do nothing
                Else
                    child.cboObjective2.SelectedValue = (row1("SecondaryObjectiveID"))
                End If

                If IsDBNull(row1("ProductDesc")) Then
                    child.RichTextBox1.Text = ""
                Else
                    child.RichTextBox1.Text = (row1("ProductDesc"))
                End If

                If IsDBNull(row1("SelectionProcess")) Then
                    child.RichTextBox2.Text = ""
                Else
                    child.RichTextBox2.Text = (row1("SelectionProcess"))
                End If

                If IsDBNull(row1("Symbol")) Then
                    child.TextBox3.Text = ""
                Else
                    child.TextBox3.Text = (row1("Symbol"))
                End If

                If IsDBNull(row1("CUSIP")) Then
                    child.TextBox4.Text = ""
                Else
                    child.TextBox4.Text = (row1("CUSIP"))
                End If

                If IsDBNull(row1("Series")) Then
                    child.TextBox5.Text = ""
                Else
                    child.TextBox5.Text = (row1("Series"))
                End If

                child.CheckBox1.Checked = (row1("40ActDriver"))
                child.CheckBox2.Checked = (row1("Active"))

                If child.CheckBox1.Checked Then
                    child.GroupBox1.Visible = True
                Else
                    child.GroupBox1.Visible = False
                End If

            Catch ex As Exception
                MsgBox(ex.Message, MsgBoxStyle.Critical, "ERROR")
            End Try
        End If
    End Sub


    Private Sub DataGridView1_CellContentDoubleClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView1.CellContentDoubleClick
        If Permissions.MAPEditProducts.Checked Then
            Call LoadEditForm()
        Else
            MsgBox("You do not have permission to edit products.", MsgBoxStyle.Critical, "Insufficient Permissions")
        End If
    End Sub

    Private Sub EditToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles EditToolStripMenuItem.Click
        If Permissions.MAPEditProducts.Checked Then
            Call LoadEditForm()
        Else
            MsgBox("You do not have permission to edit products.", MsgBoxStyle.Critical, "Insufficient Permissions")
        End If
    End Sub

    Private Sub DataGridView1_CellContentClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick

    End Sub
End Class