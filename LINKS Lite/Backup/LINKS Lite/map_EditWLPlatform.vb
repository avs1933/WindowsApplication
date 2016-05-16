Public Class map_EditWLPlatform

    Private Sub map_EditWLPlatform_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        Call CheckForChanges()
        If ckbChanges.Checked Then
            Dim ir As MsgBoxResult
            ir = MsgBox("Changes have been detected.  Would you like to save them before closing?", MsgBoxStyle.YesNo + MsgBoxStyle.Information, "Changes Detected")
            If ir = MsgBoxResult.Yes Then
                If ID.Text = "NEW" Then
                    Call CheckForDupe()
                Else
                    Call SaveOld()
                End If
            Else
            End If
        End If
    End Sub

    Private Sub map_EditWLPlatform_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Call LoadAPXFirms()
    End Sub

    Public Sub CheckForChanges()
        If cb1copy.Checked = CheckBox1.Checked Then
        Else
            ckbChanges.Checked = True
        End If

        If cb2copy.Checked = CheckBox2.Checked Then
        Else
            ckbChanges.Checked = True
        End If

        If cb3copy.Checked = CheckBox3.Checked Then
        Else
            ckbChanges.Checked = True
        End If

        If cb4copy.Checked = CheckBox4.Checked Then
        Else
            ckbChanges.Checked = True
        End If

        If cb5copy.Checked = CheckBox5.Checked Then
        Else
            ckbChanges.Checked = True
        End If

        If cb6copy.Checked = CheckBox6.Checked Then
        Else
            ckbChanges.Checked = True
        End If
        If cb7copy.Checked = CheckBox7.Checked Then
        Else
            ckbChanges.Checked = True
        End If

        If cbcfcopy.Checked = ckbCustomFee.Checked Then
        Else
            ckbChanges.Checked = True
        End If

        If RichTextBox2.Text = RichTextBox1.Text Then
        Else
            ckbChanges.Checked = True
        End If

        If WLNameCopy.Text = WLName.Text Then
        Else
            ckbChanges.Checked = True
        End If

        If PlatformIDCopy.Text = cboPlatform.SelectedValue.ToString Then
        Else
            ckbChanges.Checked = True
        End If
    End Sub

    Public Sub LoadChangesDetection()
        cb1copy.Checked = CheckBox1.CheckState
        cb2copy.Checked = CheckBox2.CheckState
        cb3copy.Checked = CheckBox3.CheckState
        cb4copy.Checked = CheckBox4.CheckState
        cb5copy.Checked = CheckBox5.CheckState
        cb6copy.Checked = CheckBox6.CheckState
        cb7copy.Checked = CheckBox7.CheckState
        cbcfcopy.Checked = ckbCustomFee.CheckState
        RichTextBox2.Text = RichTextBox1.Text
        WLNameCopy.Text = WLName.Text
        PlatformIDCopy.Text = cboPlatform.SelectedValue.ToString
        
    End Sub

    Public Sub LoadAPXFirms()
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

    Private Sub CheckBox2_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckBox2.CheckedChanged
        Call CheckForAutoApprove()

        If CheckBox2.Checked = False And CheckBox1.Checked = True Then
            Call RemovedAutoApprovedRecords()
        Else

        End If

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

    Private Sub CheckBox7_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckBox7.CheckedChanged
        If CheckBox7.Checked Then
            CheckBox1.Checked = False
            CheckBox1.Enabled = False
            CheckBox2.Checked = False
            CheckBox2.Enabled = False
            'CheckBox3.Checked = False
            'CheckBox3.Enabled = False
            'CheckBox4.Checked = False
            'CheckBox4.Enabled = False
            'CheckBox5.Checked = False
            'CheckBox5.Enabled = False
            'CheckBox6.Checked = False
            'CheckBox6.Enabled = False
            GroupBox2.Enabled = False
            Call CheckPlatformProducts()
        Else
            CheckBox1.Enabled = True
            CheckBox2.Enabled = True
            CheckBox3.Enabled = True
            CheckBox4.Enabled = True
            CheckBox5.Enabled = True
            CheckBox6.Enabled = True
        End If
    End Sub

    Public Sub CheckPlatformProducts()
        If ID.Text = "NEW" Then
            Try
                Dim conn As New OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & My.Settings.dbloc & "';Jet OLEDB:Database Password='" & My.Settings.dbpass & "';")

                Dim strSQL As String = "SELECT * FROM map_Platforms WHERE Active = True AND ID = " & cboPlatform.SelectedValue
                Dim queryString As String = String.Format(strSQL)
                Dim cmd As New OleDb.OleDbCommand(queryString, conn)
                Dim da As New OleDb.OleDbDataAdapter(cmd)
                Dim ds As New DataSet

                da.Fill(ds, "User")
                Dim dt As DataTable = ds.Tables("User")
                Dim row1 As DataRow = dt.Rows(0)

                CheckBox3.Checked = (row1("OfferSMA"))
                CheckBox4.Checked = (row1("OfferUMA"))
                CheckBox5.Checked = (row1("OfferMF"))
                CheckBox6.Checked = (row1("OfferUIT"))

            Catch ex As Exception
                MsgBox(ex.Message, MsgBoxStyle.Critical, "ERROR")
            End Try
        Else

        End If


    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        If ID.Text = "NEW" Then
            Dim ir As MsgBoxResult
            ir = MsgBox("You must save this record before adding fees.  Would you like to save now?", MsgBoxStyle.YesNo + MsgBoxStyle.Information, "Save Record?")
            If ir = MsgBoxResult.Yes Then
                Call CheckForDupe()
                Call PullID()
                If CheckBox2.Checked Then
                    Dim child As New map_PlatformProducts
                    child.MdiParent = Home
                    child.Show()
                    child.ID.Text = cboPlatform.SelectedValue.ToString
                    child.ckbWLPlatform.Checked = True
                    child.WLID.Text = ID.Text
                Else
                    MsgBox("This Platform is not setup for custom approval.  You must select Additional Approval from the attributes before adding custom products.", MsgBoxStyle.Information, "Cannot customize products")
                End If
            Else
                'do nothing
            End If
        Else
            If CheckBox2.Checked Then
                Dim child As New map_PlatformProducts
                child.MdiParent = Home
                child.Show()
                child.ID.Text = cboPlatform.SelectedValue.ToString
                child.ckbWLPlatform.Checked = True
                child.WLID.Text = ID.Text
            Else
                MsgBox("This Platform is not setup for custom approval.  You must select Additional Approval from the attributes before adding custom products.", MsgBoxStyle.Information, "Cannot customize products")
            End If
        End If
    End Sub

    Public Sub LoadProductForm()

    End Sub

    Public Sub PullID()
        Try
            Dim conn As New OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & My.Settings.dbloc & "';Jet OLEDB:Database Password='" & My.Settings.dbpass & "';")

            Dim strSQL As String = "SELECT Top 1 ID FROM map_PlatformsWL WHERE WLName = '" & WLName.Text & "' AND PlatformDriverID = " & cboPlatform.SelectedValue & " AND Active = -1 ORDER BY ID DESC"
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

            SQLstring = "SELECT * FROM Map_PlatformsWL WHERE Active = True AND WLName = '" & WLName.Text & "' AND PlatformDriverID = " & cboPlatform.SelectedValue

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

            SQLstr = "INSERT INTO map_PlatformsWL(WLName, PlatformDriverID, RequirePlatformApproval, AdditionalApproval, AdditionalApprovalProcess, OfferSMA, OfferUMA, OfferMF, OfferUIT, MirrorPlatformDriver, Active, CustomFee)" & _
            "VALUES('" & WLName.Text & "'," & cboPlatform.SelectedValue & "," & CheckBox1.CheckState & "," & CheckBox2.CheckState & ",'" & RichTextBox1.Text & "'," & CheckBox3.CheckState & "," & CheckBox4.CheckState & "," & CheckBox5.CheckState & "," & CheckBox6.CheckState & "," & CheckBox7.CheckState & ",-1," & ckbCustomFee.CheckState & ")"

            Command = New OleDb.OleDbCommand(SQLstr, Mycn)
            Command.ExecuteNonQuery()

            Mycn.Close()

            MsgBox("Record Saved.")

            Call PullID()
            Call LoadChangesDetection()
            ckbChanges.Checked = False

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

            SQLstr = "Update map_PlatformsWL SET WLName = '" & WLName.Text & "', PlatformDriverID = " & cboPlatform.SelectedValue & ", RequirePlatformApproval = " & CheckBox1.CheckState & ", AdditionalApproval = " & CheckBox2.CheckState & ", AdditionalApprovalProcess = '" & RichTextBox1.Text & "', OfferSMA = " & CheckBox3.CheckState & ", OfferUMA = " & CheckBox4.CheckState & ", OfferMF = " & CheckBox5.CheckState & ", OfferUIT = " & CheckBox6.CheckState & ", MirrorPlatformDriver = " & CheckBox7.CheckState & ", CustomFee = " & ckbCustomFee.CheckState & ", Active = " & ckbActive.CheckState & " WHERE ID = " & ID.Text

            Command = New OleDb.OleDbCommand(SQLstr, Mycn)
            Command.ExecuteNonQuery()

            Mycn.Close()

            MsgBox("Record Updated.")
            Call LoadChangesDetection()
            ckbChanges.Checked = False

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
                    SQLstr = "Update map_PlatformsWL SET Active = False WHERE ID = " & ID.Text
                    ckbActive.Checked = False
                Else
                    SQLstr = "Update map_PlatformsWL SET Active = -1 WHERE ID = " & ID.Text
                    ckbActive.Checked = True
                End If

                Command = New OleDb.OleDbCommand(SQLstr, Mycn)
                Command.ExecuteNonQuery()

                Call LoadChangesDetection()

                Mycn.Close()

            Catch ex As Exception
                MsgBox(ex.Message, MsgBoxStyle.Critical, "ERROR")
            End Try
        End If
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        If ID.Text = "NEW" Then
            MsgBox("You must save this record before loading products.", MsgBoxStyle.Information, "Not Saved")
        Else
            Call LoadAssigned()
        End If

    End Sub

    Public Sub LoadAssigned()
        Try
            Dim conn As New OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & My.Settings.dbloc & "';Jet OLEDB:Database Password='" & My.Settings.dbpass & "';")

            Dim SqlString As String

            If CheckBox7.Checked Then
                'Products Mirror Platform
                If CheckBox3.Checked = False And CheckBox4.Checked = True Then
                    'only offer UMA
                    SqlString = "SELECT map_PlatformListings.ID, map_platformListings.ProductID, map_Products.ProductName As [Product Name], map_ProductType.TypeName As [Product Type], Map_Firms.FirmName As [Managing Firm], map_PlatformListings.PlatformApproved As [Platform Approved], map_PlatformListings.SMAOffered As [Offered as SMA], map_PlatformListings.UMAOffered As [Offered as UMA]" & _
                    " FROM ((map_PlatformListings INNER JOIN map_Products ON map_PlatformListings.ProductID = map_Products.ID) INNER JOIN map_ProductType ON map_Products.TypeOfProductID = map_ProductType.ID) INNER JOIN Map_Firms ON map_Products.ManagingFirmID = Map_Firms.ID" & _
                    " WHERE(((map_PlatformListings.Active) = True) AND map_PlatformListings.UMAOffered = True And ((map_PlatformListings.PlatformID) = " & cboPlatform.SelectedValue & "))" & _
                    " GROUP BY map_PlatformListings.ID, map_platformListings.ProductID, map_Products.ProductName, map_ProductType.TypeName, Map_Firms.FirmName, map_PlatformListings.PlatformApproved, map_PlatformListings.SMAOffered, map_PlatformListings.UMAOffered;"
                Else
                    If CheckBox3.Checked = True And CheckBox4.Checked = False Then
                        'only offer SMA
                        SqlString = "SELECT map_PlatformListings.ID, map_platformListings.ProductID, map_Products.ProductName As [Product Name], map_ProductType.TypeName As [Product Type], Map_Firms.FirmName As [Managing Firm], map_PlatformListings.PlatformApproved As [Platform Approved], map_PlatformListings.SMAOffered As [Offered as SMA], map_PlatformListings.UMAOffered As [Offered as UMA]" & _
                        " FROM ((map_PlatformListings INNER JOIN map_Products ON map_PlatformListings.ProductID = map_Products.ID) INNER JOIN map_ProductType ON map_Products.TypeOfProductID = map_ProductType.ID) INNER JOIN Map_Firms ON map_Products.ManagingFirmID = Map_Firms.ID" & _
                        " WHERE(((map_PlatformListings.Active) = True) AND map_PlatformListings.SMAOffered = True And ((map_PlatformListings.PlatformID) = " & cboPlatform.SelectedValue & "))" & _
                        " GROUP BY map_PlatformListings.ID, map_platformListings.ProductID, map_Products.ProductName, map_ProductType.TypeName, Map_Firms.FirmName, map_PlatformListings.PlatformApproved, map_PlatformListings.SMAOffered, map_PlatformListings.UMAOffered;"
                    Else
                        'default to offering both SMA and UMA
                        SqlString = "SELECT map_PlatformListings.ID, map_platformListings.ProductID, map_Products.ProductName As [Product Name], map_ProductType.TypeName As [Product Type], Map_Firms.FirmName As [Managing Firm], map_PlatformListings.PlatformApproved As [Platform Approved], map_PlatformListings.SMAOffered As [Offered as SMA], map_PlatformListings.UMAOffered As [Offered as UMA]" & _
                        " FROM ((map_PlatformListings INNER JOIN map_Products ON map_PlatformListings.ProductID = map_Products.ID) INNER JOIN map_ProductType ON map_Products.TypeOfProductID = map_ProductType.ID) INNER JOIN Map_Firms ON map_Products.ManagingFirmID = Map_Firms.ID" & _
                        " WHERE(((map_PlatformListings.Active) = True) And ((map_PlatformListings.PlatformID) = " & cboPlatform.SelectedValue & "))" & _
                        " GROUP BY map_PlatformListings.ID, map_platformListings.ProductID, map_Products.ProductName, map_ProductType.TypeName, Map_Firms.FirmName, map_PlatformListings.PlatformApproved, map_PlatformListings.SMAOffered, map_PlatformListings.UMAOffered;"
                    End If
                End If
            Else
                If CheckBox1.Checked = True And CheckBox2.Checked = False Then
                    If CheckBox3.Checked = False And CheckBox4.Checked = True Then
                        'only offer UMA
                        SqlString = "SELECT map_PlatformListings.ID, map_platformListings.ProductID, map_Products.ProductName As [Product Name], map_ProductType.TypeName As [Product Type], Map_Firms.FirmName As [Managing Firm], map_PlatformListings.PlatformApproved As [Platform Approved], map_PlatformListings.SMAOffered As [Offered as SMA], map_PlatformListings.UMAOffered As [Offered as UMA]" & _
                            " FROM ((map_PlatformListings INNER JOIN map_Products ON map_PlatformListings.ProductID = map_Products.ID) INNER JOIN map_ProductType ON map_Products.TypeOfProductID = map_ProductType.ID) INNER JOIN Map_Firms ON map_Products.ManagingFirmID = Map_Firms.ID" & _
                            " WHERE(((map_PlatformListings.Active) = True) And (map_PlatformListings.PlatformApproved = True) AND map_PlatformListings.UMAOffered = True AND map_PlatformListings.WLID Is Null AND ((map_PlatformListings.PlatformID) = " & cboPlatform.SelectedValue & "))" & _
                            " GROUP BY map_PlatformListings.ID, map_platformListings.ProductID, map_Products.ProductName, map_ProductType.TypeName, Map_Firms.FirmName, map_PlatformListings.PlatformApproved, map_PlatformListings.SMAOffered, map_PlatformListings.UMAOffered;"
                    Else
                        If CheckBox3.Checked = True And CheckBox4.Checked = False Then
                            'only offer SMA
                            SqlString = "SELECT map_PlatformListings.ID, map_platformListings.ProductID, map_Products.ProductName As [Product Name], map_ProductType.TypeName As [Product Type], Map_Firms.FirmName As [Managing Firm], map_PlatformListings.PlatformApproved As [Platform Approved], map_PlatformListings.SMAOffered As [Offered as SMA], map_PlatformListings.UMAOffered As [Offered as UMA]" & _
                            " FROM ((map_PlatformListings INNER JOIN map_Products ON map_PlatformListings.ProductID = map_Products.ID) INNER JOIN map_ProductType ON map_Products.TypeOfProductID = map_ProductType.ID) INNER JOIN Map_Firms ON map_Products.ManagingFirmID = Map_Firms.ID" & _
                            " WHERE(((map_PlatformListings.Active) = True) And (map_PlatformListings.PlatformApproved = True) AND map_PlatformListings.SMAOffered = True AND map_PlatformListings.WLID Is Null AND ((map_PlatformListings.PlatformID) = " & cboPlatform.SelectedValue & "))" & _
                            " GROUP BY map_PlatformListings.ID, map_platformListings.ProductID, map_Products.ProductName, map_ProductType.TypeName, Map_Firms.FirmName, map_PlatformListings.PlatformApproved, map_PlatformListings.SMAOffered, map_PlatformListings.UMAOffered;"
                        Else
                            'default to offering both SMA and UMA
                            SqlString = "SELECT map_PlatformListings.ID, map_platformListings.ProductID, map_Products.ProductName As [Product Name], map_ProductType.TypeName As [Product Type], Map_Firms.FirmName As [Managing Firm], map_PlatformListings.PlatformApproved As [Platform Approved], map_PlatformListings.SMAOffered As [Offered as SMA], map_PlatformListings.UMAOffered As [Offered as UMA]" & _
                            " FROM ((map_PlatformListings INNER JOIN map_Products ON map_PlatformListings.ProductID = map_Products.ID) INNER JOIN map_ProductType ON map_Products.TypeOfProductID = map_ProductType.ID) INNER JOIN Map_Firms ON map_Products.ManagingFirmID = Map_Firms.ID" & _
                            " WHERE(((map_PlatformListings.Active) = True) And (map_PlatformListings.PlatformApproved = True) AND map_PlatformListings.WLID Is Null AND ((map_PlatformListings.PlatformID) = " & cboPlatform.SelectedValue & "))" & _
                            " GROUP BY map_PlatformListings.ID, map_platformListings.ProductID, map_Products.ProductName, map_ProductType.TypeName, Map_Firms.FirmName, map_PlatformListings.PlatformApproved, map_PlatformListings.SMAOffered, map_PlatformListings.UMAOffered;"
                        End If
                    End If
                Else
                    If CheckBox3.Checked = False And CheckBox4.Checked = True Then
                        'only offer UMA
                        SqlString = "SELECT map_PlatformListings.ID, map_platformListings.ProductID, map_Products.ProductName As [Product Name], map_ProductType.TypeName As [Product Type], Map_Firms.FirmName As [Managing Firm], map_PlatformListings.PlatformApproved As [Platform Approved], map_PlatformListings.WLApproved As [Whitelabel Approved], map_PlatformListings.SMAOffered As [Offered as SMA], map_PlatformListings.UMAOffered As [Offered as UMA]" & _
                            " FROM ((map_PlatformListings INNER JOIN map_Products ON map_PlatformListings.ProductID = map_Products.ID) INNER JOIN map_ProductType ON map_Products.TypeOfProductID = map_ProductType.ID) INNER JOIN Map_Firms ON map_Products.ManagingFirmID = Map_Firms.ID" & _
                            " WHERE(((map_PlatformListings.Active) = True) AND map_PlatformListings.UMAOffered = True And ((map_PlatformListings.WLID) = " & ID.Text & "))" & _
                            " GROUP BY map_PlatformListings.ID, map_Products.ProductName, map_platformListings.ProductID, map_ProductType.TypeName, Map_Firms.FirmName, map_PlatformListings.PlatformApproved, map_PlatformListings.WLApproved, map_PlatformListings.SMAOffered, map_PlatformListings.UMAOffered;"
                    Else
                        If CheckBox3.Checked = True And CheckBox4.Checked = False Then
                            'only offer SMA
                            SqlString = "SELECT map_PlatformListings.ID, map_platformListings.ProductID, map_Products.ProductName As [Product Name], map_ProductType.TypeName As [Product Type], Map_Firms.FirmName As [Managing Firm], map_PlatformListings.PlatformApproved As [Platform Approved], map_PlatformListings.WLApproved As [Whitelabel Approved], map_PlatformListings.SMAOffered As [Offered as SMA], map_PlatformListings.UMAOffered As [Offered as UMA]" & _
                             " FROM ((map_PlatformListings INNER JOIN map_Products ON map_PlatformListings.ProductID = map_Products.ID) INNER JOIN map_ProductType ON map_Products.TypeOfProductID = map_ProductType.ID) INNER JOIN Map_Firms ON map_Products.ManagingFirmID = Map_Firms.ID" & _
                             " WHERE(((map_PlatformListings.Active) = True) AND map_PlatformListings.SMAOffered = True And ((map_PlatformListings.WLID) = " & ID.Text & "))" & _
                             " GROUP BY map_PlatformListings.ID, map_Products.ProductName, map_platformListings.ProductID, map_ProductType.TypeName, Map_Firms.FirmName, map_PlatformListings.PlatformApproved, map_PlatformListings.WLApproved, map_PlatformListings.SMAOffered, map_PlatformListings.UMAOffered;"
                        Else
                            'default to offering both SMA and UMA
                            SqlString = "SELECT map_PlatformListings.ID, map_platformListings.ProductID, map_Products.ProductName As [Product Name], map_ProductType.TypeName As [Product Type], Map_Firms.FirmName As [Managing Firm], map_PlatformListings.PlatformApproved As [Platform Approved], map_PlatformListings.WLApproved As [Whitelabel Approved], map_PlatformListings.SMAOffered As [Offered as SMA], map_PlatformListings.UMAOffered As [Offered as UMA]" & _
                            " FROM ((map_PlatformListings INNER JOIN map_Products ON map_PlatformListings.ProductID = map_Products.ID) INNER JOIN map_ProductType ON map_Products.TypeOfProductID = map_ProductType.ID) INNER JOIN Map_Firms ON map_Products.ManagingFirmID = Map_Firms.ID" & _
                            " WHERE(((map_PlatformListings.Active) = True) And ((map_PlatformListings.WLID) = " & ID.Text & "))" & _
                            " GROUP BY map_PlatformListings.ID, map_Products.ProductName, map_platformListings.ProductID, map_ProductType.TypeName, Map_Firms.FirmName, map_PlatformListings.PlatformApproved, map_PlatformListings.WLApproved, map_PlatformListings.SMAOffered, map_PlatformListings.UMAOffered;"
                        End If
                    End If
                End If
            End If

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

    Public Sub InsertPlatformApproved()
        Dim Mycn As OleDb.OleDbConnection
        Dim Command As OleDb.OleDbCommand
        Dim SQLstr As String
        Try
            Mycn = New OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & My.Settings.dbloc & "';Jet OLEDB:Database Password='" & My.Settings.dbpass & "';")
            Mycn.Open()

            SQLstr = "INSERT INTO map_PlatformListings(PlatformID, WLID, ProductID, PlatformApproved, WLApproved, Active, SMAOffered, UMAOffered, AutoAdded)" & _
            "SELECT map_PlatformListings.PlatformID," & ID.Text & ", map_PlatformListings.ProductID, map_PlatformListings.PlatformApproved, -1, map_PlatformListings.Active, map_PlatformListings.SMAOffered, map_PlatformListings.UMAOffered, -1" & _
            " FROM map_PlatformListings" & _
            " WHERE PlatformID = " & cboPlatform.SelectedValue & " AND WLID Is Null AND PlatformApproved = True AND Active = True AND ProductID Not In (Select productID From Map_PlatformListings WHERE WLID = " & ID.Text & " AND Active = True)"

            Command = New OleDb.OleDbCommand(SQLstr, Mycn)
            Command.ExecuteNonQuery()

            Mycn.Close()

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "ERROR")
        End Try
    End Sub

    Public Sub LoadFees()
        'Try
        Dim conn As New OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & My.Settings.dbloc & "';Jet OLEDB:Database Password='" & My.Settings.dbpass & "';")

        Dim SqlString As String
        If ckbCustomFee.Checked Then
            If CheckBox7.Checked Then
                'Load Custom Fees for all Platform Products
                If CheckBox3.Checked = False And CheckBox4.Checked = True Then
                    'only offer UMA
                    SqlString = "SELECT map_PlatformFees.ID, map_Products.ProductName, map_ProductType.TypeName, Map_Firms.FirmName, map_PlatformFees.BreakpointFrom, map_PlatformFees.BreakpointTo, map_PlatformFees.Fee, map_PlatformFees.MaxRepFee, '1' As [QID]" & _
                    " FROM ((map_PlatformFees INNER JOIN map_Products ON map_PlatformFees.ProductID = map_Products.ID) INNER JOIN map_ProductType ON map_Products.TypeOfProductID = map_ProductType.ID) INNER JOIN Map_Firms ON map_Products.ManagingFirmID = Map_Firms.ID" & _
                    " WHERE(((map_PlatformFees.Active) = True) And ((map_PlatformFees.WLID) = " & ID.Text & ") AND map_PlatformFees.PlatformID = " & cboPlatform.SelectedValue & " AND map_PlatformFees.ProductID In (Select ProductID FROM map_PlatformListings WHERE UMAOffered = True AND PlatformID = " & cboPlatform.SelectedValue & " AND WLID Is Null))" & _
                    " GROUP BY map_PlatformFees.ID, map_Products.ProductName, map_ProductType.TypeName, Map_Firms.FirmName, map_PlatformFees.BreakpointFrom, map_PlatformFees.BreakpointTo, map_PlatformFees.Fee, map_PlatformFees.MaxRepFee;"
                Else
                    If CheckBox3.Checked = True And CheckBox4.Checked = False Then
                        'only offer SMA
                        SqlString = "SELECT map_PlatformFees.ID, map_Products.ProductName, map_ProductType.TypeName, Map_Firms.FirmName, map_PlatformFees.BreakpointFrom, map_PlatformFees.BreakpointTo, map_PlatformFees.Fee, map_PlatformFees.MaxRepFee, '2' As [QID]" & _
                        " FROM ((map_PlatformFees INNER JOIN map_Products ON map_PlatformFees.ProductID = map_Products.ID) INNER JOIN map_ProductType ON map_Products.TypeOfProductID = map_ProductType.ID) INNER JOIN Map_Firms ON map_Products.ManagingFirmID = Map_Firms.ID" & _
                        " WHERE(((map_PlatformFees.Active) = True) And ((map_PlatformFees.WLID) = " & ID.Text & ") AND map_PlatformFees.PlatformID = " & cboPlatform.SelectedValue & " AND map_PlatformFees.ProductID In (Select ProductID FROM map_PlatformListings WHERE SMAOffered = True AND PlatformID = " & cboPlatform.SelectedValue & " AND WLID Is Null))" & _
                        " GROUP BY map_PlatformFees.ID, map_Products.ProductName, map_ProductType.TypeName, Map_Firms.FirmName, map_PlatformFees.BreakpointFrom, map_PlatformFees.BreakpointTo, map_PlatformFees.Fee, map_PlatformFees.MaxRepFee;"
                    Else
                        'default to offering both SMA and UMA
                        SqlString = "SELECT map_PlatformFees.ID, map_Products.ProductName, map_ProductType.TypeName, Map_Firms.FirmName, map_PlatformFees.BreakpointFrom, map_PlatformFees.BreakpointTo, map_PlatformFees.Fee, map_PlatformFees.MaxRepFee, '3' As [QID]" & _
                        " FROM ((map_PlatformFees INNER JOIN map_Products ON map_PlatformFees.ProductID = map_Products.ID) INNER JOIN map_ProductType ON map_Products.TypeOfProductID = map_ProductType.ID) INNER JOIN Map_Firms ON map_Products.ManagingFirmID = Map_Firms.ID" & _
                        " WHERE(((map_PlatformFees.Active) = True) And ((map_PlatformFees.WLID) = " & ID.Text & ") AND map_PlatformFees.PlatformID = " & cboPlatform.SelectedValue & ")" & _
                        " GROUP BY map_PlatformFees.ID, map_Products.ProductName, map_ProductType.TypeName, Map_Firms.FirmName, map_PlatformFees.BreakpointFrom, map_PlatformFees.BreakpointTo, map_PlatformFees.Fee, map_PlatformFees.MaxRepFee;"
                    End If
                End If
            Else
                If CheckBox1.Checked = True And CheckBox2.Checked = False Then
                    'Only Load Fees for Platform Approved Products
                    If CheckBox3.Checked = False And CheckBox4.Checked = True Then
                        'only offer UMA
                        SqlString = "SELECT map_PlatformFees.ID, map_Products.ProductName, map_ProductType.TypeName, Map_Firms.FirmName, map_PlatformFees.BreakpointFrom, map_PlatformFees.BreakpointTo, map_PlatformFees.Fee, map_PlatformFees.MaxRepFee, '4' As [QID]" & _
                        " FROM map_PlatformFees INNER JOIN ((map_Products INNER JOIN map_ProductType ON map_Products.TypeOfProductID = map_ProductType.ID) INNER JOIN Map_Firms ON map_Products.ManagingFirmID = Map_Firms.ID) ON map_PlatformFees.ProductID = map_Products.ID" & _
                        " WHERE map_PlatformFees.Active = True AND map_PlatformFees.WLID = " & ID.Text & " AND map_PlatformFees.ProductID In (Select ProductID FROM map_PlatformListings WHERE UMAOffered = True AND Active = True AND PlatformApproved = True AND WLID Is Null AND PlatformID = " & cboPlatform.SelectedValue & ")" & _
                        " GROUP BY map_PlatformFees.ID, map_Products.ProductName, map_ProductType.TypeName, map_PlatformFees.BreakpointFrom, map_PlatformFees.BreakpointTo, map_PlatformFees.Fee, map_PlatformFees.MaxRepFee, Map_Firms.FirmName;"
                    Else
                        If CheckBox3.Checked = True And CheckBox4.Checked = False Then
                            'only offer SMA
                            SqlString = "SELECT map_PlatformFees.ID, map_Products.ProductName, map_ProductType.TypeName, Map_Firms.FirmName, map_PlatformFees.BreakpointFrom, map_PlatformFees.BreakpointTo, map_PlatformFees.Fee, map_PlatformFees.MaxRepFee, '5' As [QID]" & _
                            " FROM map_PlatformFees INNER JOIN ((map_Products INNER JOIN map_ProductType ON map_Products.TypeOfProductID = map_ProductType.ID) INNER JOIN Map_Firms ON map_Products.ManagingFirmID = Map_Firms.ID) ON map_PlatformFees.ProductID = map_Products.ID" & _
                            " WHERE map_PlatformFees.Active = True AND map_PlatformFees.WLID = " & ID.Text & " AND map_PlatformFees.ProductID In (Select ProductID FROM map_PlatformListings WHERE SMAOffered = True AND Active = True AND PlatformApproved = True AND WLID Is Null AND PlatformID = " & cboPlatform.SelectedValue & ")" & _
                            " GROUP BY map_PlatformFees.ID, map_Products.ProductName, map_ProductType.TypeName, map_PlatformFees.BreakpointFrom, map_PlatformFees.BreakpointTo, map_PlatformFees.Fee, map_PlatformFees.MaxRepFee, Map_Firms.FirmName;"
                        Else
                            'default to offering both SMA and UMA
                            SqlString = "SELECT map_PlatformFees.ID, map_Products.ProductName, map_ProductType.TypeName, Map_Firms.FirmName, map_PlatformFees.BreakpointFrom, map_PlatformFees.BreakpointTo, map_PlatformFees.Fee, map_PlatformFees.MaxRepFee, '6' As [QID]" & _
                            " FROM map_PlatformFees INNER JOIN ((map_Products INNER JOIN map_ProductType ON map_Products.TypeOfProductID = map_ProductType.ID) INNER JOIN Map_Firms ON map_Products.ManagingFirmID = Map_Firms.ID) ON map_PlatformFees.ProductID = map_Products.ID" & _
                            " WHERE map_PlatformFees.Active = True AND map_PlatformFees.WLID = " & ID.Text & " AND map_PlatformFees.ProductID In (Select ProductID FROM map_PlatformListings WHERE Active = True AND PlatformApproved = True AND WLID Is Null AND PlatformID = " & cboPlatform.SelectedValue & ")" & _
                            " GROUP BY map_PlatformFees.ID, map_Products.ProductName, map_ProductType.TypeName, map_PlatformFees.BreakpointFrom, map_PlatformFees.BreakpointTo, map_PlatformFees.Fee, map_PlatformFees.MaxRepFee, Map_Firms.FirmName;"
                        End If
                    End If

                    'SqlString = "SELECT map_PlatformFees.ID, map_Products.ProductName, map_ProductType.TypeName, Map_Firms.FirmName, map_PlatformFees.BreakpointFrom, map_PlatformFees.BreakpointTo, map_PlatformFees.Fee, map_PlatformFees.MaxRepFee" & _
                    '" FROM map_PlatformFees INNER JOIN ((map_Products INNER JOIN map_ProductType ON map_Products.TypeOfProductID = map_ProductType.ID) INNER JOIN Map_Firms ON map_Products.ManagingFirmID = Map_Firms.ID) ON map_PlatformFees.ProductID = map_Products.ID" & _
                    '" WHERE map_PlatformFees.Active = True AND map_PlatformFees.WLID = " & ID.Text & " AND map_PlatformFees.ProductID In (Select ProductID FROM map_PlatformListings WHERE Active = True AND PlatformApproved = True AND WLID Is Null AND PlatformID = " & cboPlatform.SelectedValue & ")" & _
                    '" GROUP BY map_PlatformFees.ID, map_Products.ProductName, map_ProductType.TypeName, map_PlatformFees.BreakpointFrom, map_PlatformFees.BreakpointTo, map_PlatformFees.Fee, map_PlatformFees.MaxRepFee, Map_Firms.FirmName;"
                Else
                    'Only load fees for Custom Approved
                    If CheckBox3.Checked = False And CheckBox4.Checked = True Then
                        'only offer UMA
                        SqlString = "SELECT map_PlatformFees.ID, map_Products.ProductName, map_ProductType.TypeName, Map_Firms.FirmName, map_PlatformFees.BreakpointFrom, map_PlatformFees.BreakpointTo, map_PlatformFees.Fee, map_PlatformFees.MaxRepFee, '7' As [QID]" & _
                        " FROM map_PlatformFees INNER JOIN ((map_Products INNER JOIN map_ProductType ON map_Products.TypeOfProductID = map_ProductType.ID) INNER JOIN Map_Firms ON map_Products.ManagingFirmID = Map_Firms.ID) ON map_PlatformFees.ProductID = map_Products.ID" & _
                        " WHERE(((map_PlatformFees.Active) = True) AND map_PlatformFees.WLID = " & ID.Text & " AND map_PlatformFees.ProductID In (Select ProductID FROM map_PlatformListings WHERE UMAOffered = True AND WLID = " & ID.Text & "))" & _
                        " GROUP BY map_PlatformFees.ID, map_Products.ProductName, map_ProductType.TypeName, Map_Firms.FirmName, map_PlatformFees.BreakpointFrom, map_PlatformFees.BreakpointTo, map_PlatformFees.Fee, map_PlatformFees.MaxRepFee;"
                    Else
                        If CheckBox3.Checked = True And CheckBox4.Checked = False Then
                            'only offer SMA
                            SqlString = "SELECT map_PlatformFees.ID, map_Products.ProductName, map_ProductType.TypeName, Map_Firms.FirmName, map_PlatformFees.BreakpointFrom, map_PlatformFees.BreakpointTo, map_PlatformFees.Fee, map_PlatformFees.MaxRepFee, '8' As [QID]" & _
                            " FROM map_PlatformFees INNER JOIN ((map_Products INNER JOIN map_ProductType ON map_Products.TypeOfProductID = map_ProductType.ID) INNER JOIN Map_Firms ON map_Products.ManagingFirmID = Map_Firms.ID) ON map_PlatformFees.ProductID = map_Products.ID" & _
                            " WHERE(((map_PlatformFees.Active) = True) AND map_PlatformFees.WLID = " & ID.Text & " AND map_PlatformFees.ProductID In (Select ProductID FROM map_PlatformListings WHERE SMAOffered = True AND WLID = " & ID.Text & "))" & _
                            " GROUP BY map_PlatformFees.ID, map_Products.ProductName, map_ProductType.TypeName, Map_Firms.FirmName, map_PlatformFees.BreakpointFrom, map_PlatformFees.BreakpointTo, map_PlatformFees.Fee, map_PlatformFees.MaxRepFee;"
                        Else
                            'default to offering both SMA and UMA
                            SqlString = "SELECT map_PlatformFees.ID, map_Products.ProductName, map_ProductType.TypeName, Map_Firms.FirmName, map_PlatformFees.BreakpointFrom, map_PlatformFees.BreakpointTo, map_PlatformFees.Fee, map_PlatformFees.MaxRepFee, '9' As [QID]" & _
                            " FROM map_PlatformFees INNER JOIN ((map_Products INNER JOIN map_ProductType ON map_Products.TypeOfProductID = map_ProductType.ID) INNER JOIN Map_Firms ON map_Products.ManagingFirmID = Map_Firms.ID) ON map_PlatformFees.ProductID = map_Products.ID" & _
                            " WHERE(((map_PlatformFees.Active) = True) AND map_PlatformFees.WLID = " & ID.Text & ")" & _
                            " GROUP BY map_PlatformFees.ID, map_Products.ProductName, map_ProductType.TypeName, Map_Firms.FirmName, map_PlatformFees.BreakpointFrom, map_PlatformFees.BreakpointTo, map_PlatformFees.Fee, map_PlatformFees.MaxRepFee;"
                        End If
                    End If
                    'SqlString = "SELECT map_PlatformFees.ID, map_Products.ProductName, map_ProductType.TypeName, Map_Firms.FirmName, map_PlatformFees.BreakpointFrom, map_PlatformFees.BreakpointTo, map_PlatformFees.Fee, map_PlatformFees.MaxRepFee" & _
                    '" FROM map_PlatformFees INNER JOIN ((map_Products INNER JOIN map_ProductType ON map_Products.TypeOfProductID = map_ProductType.ID) INNER JOIN Map_Firms ON map_Products.ManagingFirmID = Map_Firms.ID) ON map_PlatformFees.ProductID = map_Products.ID" & _
                    '" WHERE(((map_PlatformFees.Active) = True) AND map_PlatformFees.WLID = " & ID.Text & ")" & _
                    '" GROUP BY map_PlatformFees.ID, map_Products.ProductName, map_ProductType.TypeName, Map_Firms.FirmName, map_PlatformFees.BreakpointFrom, map_PlatformFees.BreakpointTo, map_PlatformFees.Fee, map_PlatformFees.MaxRepFee;"
                End If
            End If
        Else
            If CheckBox7.Checked Then
                'Load Default Fees for all Platform Products
                If CheckBox3.Checked = False And CheckBox4.Checked = True Then
                    'only offer UMA
                    SqlString = "SELECT map_PlatformFees.ID, map_Products.ProductName, map_ProductType.TypeName, Map_Firms.FirmName, map_PlatformFees.BreakpointFrom, map_PlatformFees.BreakpointTo, map_PlatformFees.Fee, map_PlatformFees.MaxRepFee, '10' As [QID]" & _
                    " FROM map_PlatformFees INNER JOIN ((map_Products INNER JOIN map_ProductType ON map_Products.TypeOfProductID = map_ProductType.ID) INNER JOIN Map_Firms ON map_Products.ManagingFirmID = Map_Firms.ID) ON map_PlatformFees.ProductID = map_Products.ID" & _
                    " WHERE map_PlatformFees.Active = True AND map_PlatformFees.WLID Is Null AND map_PlatformFees.PlatformID = " & cboPlatform.SelectedValue & " AND map_PlatformFees.ProductID In (Select ProductID FROM map_PlatformListings WHERE UMAOffered = True AND Active = True AND WLID Is Null AND PlatformID = " & cboPlatform.SelectedValue & ")" & _
                    " GROUP BY map_PlatformFees.ID, map_Products.ProductName, map_ProductType.TypeName, map_PlatformFees.BreakpointFrom, map_PlatformFees.BreakpointTo, map_PlatformFees.Fee, map_PlatformFees.MaxRepFee, Map_Firms.FirmName;"
                Else
                    If CheckBox3.Checked = True And CheckBox4.Checked = False Then
                        'only offer SMA
                        SqlString = "SELECT map_PlatformFees.ID, map_Products.ProductName, map_ProductType.TypeName, Map_Firms.FirmName, map_PlatformFees.BreakpointFrom, map_PlatformFees.BreakpointTo, map_PlatformFees.Fee, map_PlatformFees.MaxRepFee, '11' As [QID]" & _
                        " FROM map_PlatformFees INNER JOIN ((map_Products INNER JOIN map_ProductType ON map_Products.TypeOfProductID = map_ProductType.ID) INNER JOIN Map_Firms ON map_Products.ManagingFirmID = Map_Firms.ID) ON map_PlatformFees.ProductID = map_Products.ID" & _
                        " WHERE map_PlatformFees.Active = True AND map_PlatformFees.WLID Is Null AND map_PlatformFees.PlatformID = " & cboPlatform.SelectedValue & " AND map_PlatformFees.ProductID In (Select ProductID FROM map_PlatformListings WHERE SMAOffered = True AND Active = True AND WLID Is Null AND PlatformID = " & cboPlatform.SelectedValue & ")" & _
                        " GROUP BY map_PlatformFees.ID, map_Products.ProductName, map_ProductType.TypeName, map_PlatformFees.BreakpointFrom, map_PlatformFees.BreakpointTo, map_PlatformFees.Fee, map_PlatformFees.MaxRepFee, Map_Firms.FirmName;"
                    Else
                        'default to offering both SMA and UMA
                        SqlString = "SELECT map_PlatformFees.ID, map_Products.ProductName, map_ProductType.TypeName, Map_Firms.FirmName, map_PlatformFees.BreakpointFrom, map_PlatformFees.BreakpointTo, map_PlatformFees.Fee, map_PlatformFees.MaxRepFee, '12' As [QID]" & _
                        " FROM map_PlatformFees INNER JOIN ((map_Products INNER JOIN map_ProductType ON map_Products.TypeOfProductID = map_ProductType.ID) INNER JOIN Map_Firms ON map_Products.ManagingFirmID = Map_Firms.ID) ON map_PlatformFees.ProductID = map_Products.ID" & _
                        " WHERE map_PlatformFees.Active = True AND map_PlatformFees.WLID Is Null AND map_PlatformFees.PlatformID = " & cboPlatform.SelectedValue & " AND map_PlatformFees.ProductID In (Select ProductID FROM map_PlatformListings WHERE Active = True AND WLID Is Null AND PlatformID = " & cboPlatform.SelectedValue & ")" & _
                        " GROUP BY map_PlatformFees.ID, map_Products.ProductName, map_ProductType.TypeName, map_PlatformFees.BreakpointFrom, map_PlatformFees.BreakpointTo, map_PlatformFees.Fee, map_PlatformFees.MaxRepFee, Map_Firms.FirmName;"
                    End If
                End If
                'SqlString = "SELECT map_PlatformFees.ID, map_Products.ProductName, map_ProductType.TypeName, Map_Firms.FirmName, map_PlatformFees.BreakpointFrom, map_PlatformFees.BreakpointTo, map_PlatformFees.Fee, map_PlatformFees.MaxRepFee" & _
                '" FROM map_PlatformFees INNER JOIN ((map_Products INNER JOIN map_ProductType ON map_Products.TypeOfProductID = map_ProductType.ID) INNER JOIN Map_Firms ON map_Products.ManagingFirmID = Map_Firms.ID) ON map_PlatformFees.ProductID = map_Products.ID" & _
                '" WHERE map_PlatformFees.Active = True AND map_PlatformFees.WLID Is Null AND map_PlatformFees.PlatformID = " & cboPlatform.SelectedValue & " AND map_PlatformFees.ProductID In (Select ProductID FROM map_PlatformListings WHERE Active = True AND WLID Is Null AND PlatformID = " & cboPlatform.SelectedValue & ")" & _
                '" GROUP BY map_PlatformFees.ID, map_Products.ProductName, map_ProductType.TypeName, map_PlatformFees.BreakpointFrom, map_PlatformFees.BreakpointTo, map_PlatformFees.Fee, map_PlatformFees.MaxRepFee, Map_Firms.FirmName;"
            Else
                If CheckBox1.Checked = True And CheckBox2.Checked = False Then
                    If CheckBox3.Checked = False And CheckBox4.Checked = True Then
                        'only offer UMA
                        SqlString = "SELECT map_PlatformFees.ID, map_Products.ProductName, map_ProductType.TypeName, Map_Firms.FirmName, map_PlatformFees.BreakpointFrom, map_PlatformFees.BreakpointTo, map_PlatformFees.Fee, map_PlatformFees.MaxRepFee, '13' As [QID]" & _
                        " FROM map_PlatformFees INNER JOIN ((map_Products INNER JOIN map_ProductType ON map_Products.TypeOfProductID = map_ProductType.ID) INNER JOIN Map_Firms ON map_Products.ManagingFirmID = Map_Firms.ID) ON map_PlatformFees.ProductID = map_Products.ID" & _
                        " WHERE map_PlatformFees.Active = True AND map_PlatformFees.WLID Is Null AND map_PlatformFees.PlatformID = " & cboPlatform.SelectedValue & " AND map_PlatformFees.ProductID In (Select ProductID FROM map_PlatformListings WHERE UMAOffered = True AND Active = True AND PlatformApproved = True AND WLID Is Null AND PlatformID = " & cboPlatform.SelectedValue & ")" & _
                        " GROUP BY map_PlatformFees.ID, map_Products.ProductName, map_ProductType.TypeName, map_PlatformFees.BreakpointFrom, map_PlatformFees.BreakpointTo, map_PlatformFees.Fee, map_PlatformFees.MaxRepFee, Map_Firms.FirmName;"
                    Else
                        If CheckBox3.Checked = True And CheckBox4.Checked = False Then
                            'only offer SMA
                            SqlString = "SELECT map_PlatformFees.ID, map_Products.ProductName, map_ProductType.TypeName, Map_Firms.FirmName, map_PlatformFees.BreakpointFrom, map_PlatformFees.BreakpointTo, map_PlatformFees.Fee, map_PlatformFees.MaxRepFee, '14' As [QID]" & _
                            " FROM map_PlatformFees INNER JOIN ((map_Products INNER JOIN map_ProductType ON map_Products.TypeOfProductID = map_ProductType.ID) INNER JOIN Map_Firms ON map_Products.ManagingFirmID = Map_Firms.ID) ON map_PlatformFees.ProductID = map_Products.ID" & _
                            " WHERE map_PlatformFees.Active = True AND map_PlatformFees.WLID Is Null AND map_PlatformFees.PlatformID = " & cboPlatform.SelectedValue & " AND map_PlatformFees.ProductID In (Select ProductID FROM map_PlatformListings WHERE SMAOffered = True AND Active = True AND PlatformApproved = True AND WLID Is Null AND PlatformID = " & cboPlatform.SelectedValue & ")" & _
                            " GROUP BY map_PlatformFees.ID, map_Products.ProductName, map_ProductType.TypeName, map_PlatformFees.BreakpointFrom, map_PlatformFees.BreakpointTo, map_PlatformFees.Fee, map_PlatformFees.MaxRepFee, Map_Firms.FirmName;"
                        Else
                            'default to offering both SMA and UMA
                            SqlString = "SELECT map_PlatformFees.ID, map_Products.ProductName, map_ProductType.TypeName, Map_Firms.FirmName, map_PlatformFees.BreakpointFrom, map_PlatformFees.BreakpointTo, map_PlatformFees.Fee, map_PlatformFees.MaxRepFee, '15' As [QID]" & _
                            " FROM map_PlatformFees INNER JOIN ((map_Products INNER JOIN map_ProductType ON map_Products.TypeOfProductID = map_ProductType.ID) INNER JOIN Map_Firms ON map_Products.ManagingFirmID = Map_Firms.ID) ON map_PlatformFees.ProductID = map_Products.ID" & _
                            " WHERE map_PlatformFees.Active = True AND map_PlatformFees.WLID Is Null AND map_PlatformFees.PlatformID = " & cboPlatform.SelectedValue & " AND map_PlatformFees.ProductID In (Select ProductID FROM map_PlatformListings WHERE Active = True AND PlatformApproved = True AND WLID Is Null AND PlatformID = " & cboPlatform.SelectedValue & ")" & _
                            " GROUP BY map_PlatformFees.ID, map_Products.ProductName, map_ProductType.TypeName, map_PlatformFees.BreakpointFrom, map_PlatformFees.BreakpointTo, map_PlatformFees.Fee, map_PlatformFees.MaxRepFee, Map_Firms.FirmName;"
                        End If
                    End If
                    'SqlString = "SELECT map_PlatformFees.ID, map_Products.ProductName, map_ProductType.TypeName, Map_Firms.FirmName, map_PlatformFees.BreakpointFrom, map_PlatformFees.BreakpointTo, map_PlatformFees.Fee, map_PlatformFees.MaxRepFee" & _
                    '" FROM map_PlatformFees INNER JOIN ((map_Products INNER JOIN map_ProductType ON map_Products.TypeOfProductID = map_ProductType.ID) INNER JOIN Map_Firms ON map_Products.ManagingFirmID = Map_Firms.ID) ON map_PlatformFees.ProductID = map_Products.ID" & _
                    '" WHERE map_PlatformFees.Active = True AND map_PlatformFees.WLID Is Null AND map_PlatformFees.PlatformID = " & cboPlatform.SelectedValue & " AND map_PlatformFees.ProductID In (Select ProductID FROM map_PlatformListings WHERE Active = True AND PlatformApproved = True AND WLID Is Null AND PlatformID = " & cboPlatform.SelectedValue & ")" & _
                    '" GROUP BY map_PlatformFees.ID, map_Products.ProductName, map_ProductType.TypeName, map_PlatformFees.BreakpointFrom, map_PlatformFees.BreakpointTo, map_PlatformFees.Fee, map_PlatformFees.MaxRepFee, Map_Firms.FirmName;"
                Else
                    'Only load fees for Custom Approved
                    If CheckBox3.Checked = False And CheckBox4.Checked = True Then
                        'only offer UMA
                        SqlString = "SELECT map_PlatformFees.ID, map_Products.ProductName, map_ProductType.TypeName, Map_Firms.FirmName, map_PlatformFees.BreakpointFrom, map_PlatformFees.BreakpointTo, map_PlatformFees.Fee, map_PlatformFees.MaxRepFee, '16' As [QID]" & _
                        " FROM map_PlatformFees INNER JOIN ((map_Products INNER JOIN map_ProductType ON map_Products.TypeOfProductID = map_ProductType.ID) INNER JOIN Map_Firms ON map_Products.ManagingFirmID = Map_Firms.ID) ON map_PlatformFees.ProductID = map_Products.ID" & _
                        " WHERE(((map_PlatformFees.Active) = True) AND map_PlatformFees.WLID Is Null AND map_PlatformFees.PlatformID = " & cboPlatform.SelectedValue & " AND ProductID In (Select ProductID From map_PlatformListings WHERE UMAOffered = True AND WLID = " & ID.Text & " AND Active = -1))" & _
                        " GROUP BY map_PlatformFees.ID, map_Products.ProductName, map_ProductType.TypeName, Map_Firms.FirmName, map_PlatformFees.BreakpointFrom, map_PlatformFees.BreakpointTo, map_PlatformFees.Fee, map_PlatformFees.MaxRepFee;"
                    Else
                        If CheckBox3.Checked = True And CheckBox4.Checked = False Then
                            'only offer SMA
                            SqlString = "SELECT map_PlatformFees.ID, map_Products.ProductName, map_ProductType.TypeName, Map_Firms.FirmName, map_PlatformFees.BreakpointFrom, map_PlatformFees.BreakpointTo, map_PlatformFees.Fee, map_PlatformFees.MaxRepFee, '17' As [QID]" & _
                            " FROM map_PlatformFees INNER JOIN ((map_Products INNER JOIN map_ProductType ON map_Products.TypeOfProductID = map_ProductType.ID) INNER JOIN Map_Firms ON map_Products.ManagingFirmID = Map_Firms.ID) ON map_PlatformFees.ProductID = map_Products.ID" & _
                            " WHERE(((map_PlatformFees.Active) = True) AND map_PlatformFees.WLID Is Null AND map_PlatformFees.PlatformID = " & cboPlatform.SelectedValue & " AND ProductID In (Select ProductID From map_PlatformListings WHERE SMAOffered = True AND WLID = " & ID.Text & " AND Active = -1))" & _
                            " GROUP BY map_PlatformFees.ID, map_Products.ProductName, map_ProductType.TypeName, Map_Firms.FirmName, map_PlatformFees.BreakpointFrom, map_PlatformFees.BreakpointTo, map_PlatformFees.Fee, map_PlatformFees.MaxRepFee;"
                        Else
                            'default to offering both SMA and UMA
                            SqlString = "SELECT map_PlatformFees.ID, map_Products.ProductName, map_ProductType.TypeName, Map_Firms.FirmName, map_PlatformFees.BreakpointFrom, map_PlatformFees.BreakpointTo, map_PlatformFees.Fee, map_PlatformFees.MaxRepFee, '18' As [QID]" & _
                            " FROM map_PlatformFees INNER JOIN ((map_Products INNER JOIN map_ProductType ON map_Products.TypeOfProductID = map_ProductType.ID) INNER JOIN Map_Firms ON map_Products.ManagingFirmID = Map_Firms.ID) ON map_PlatformFees.ProductID = map_Products.ID" & _
                            " WHERE(((map_PlatformFees.Active) = True) AND map_PlatformFees.WLID Is Null AND map_PlatformFees.PlatformID = " & cboPlatform.SelectedValue & " AND ProductID In (Select ProductID From map_PlatformListings WHERE WLID = " & ID.Text & " AND Active = -1))" & _
                            " GROUP BY map_PlatformFees.ID, map_Products.ProductName, map_ProductType.TypeName, Map_Firms.FirmName, map_PlatformFees.BreakpointFrom, map_PlatformFees.BreakpointTo, map_PlatformFees.Fee, map_PlatformFees.MaxRepFee;"
                        End If
                    End If
                    'SqlString = "SELECT map_PlatformFees.ID, map_Products.ProductName, map_ProductType.TypeName, Map_Firms.FirmName, map_PlatformFees.BreakpointFrom, map_PlatformFees.BreakpointTo, map_PlatformFees.Fee, map_PlatformFees.MaxRepFee" & _
                    '" FROM map_PlatformFees INNER JOIN ((map_Products INNER JOIN map_ProductType ON map_Products.TypeOfProductID = map_ProductType.ID) INNER JOIN Map_Firms ON map_Products.ManagingFirmID = Map_Firms.ID) ON map_PlatformFees.ProductID = map_Products.ID" & _
                    '" WHERE(((map_PlatformFees.Active) = True) AND map_PlatformFees.WLID Is Null AND map_PlatformFees.PlatformID = " & cboPlatform.SelectedValue & " AND ProductID In (Select ProductID From map_PlatformListings WHERE WLID = " & ID.Text & " AND Active = -1))" & _
                    '" GROUP BY map_PlatformFees.ID, map_Products.ProductName, map_ProductType.TypeName, Map_Firms.FirmName, map_PlatformFees.BreakpointFrom, map_PlatformFees.BreakpointTo, map_PlatformFees.Fee, map_PlatformFees.MaxRepFee;"
                End If
            End If
        End If

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

        'Catch ex As Exception
        'MsgBox(ex.Message, MsgBoxStyle.Critical, "ERROR")
        'End Try
    End Sub

    Private Sub Button6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button6.Click
        If ID.Text = "NEW" Then
            MsgBox("You must save this record before loading Fees.", MsgBoxStyle.Information, "Not Saved")
        Else
            Call LoadFees()
        End If
    End Sub

    Private Sub AddFeeToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles AddFeeToolStripMenuItem.Click
        If DataGridView1.RowCount = 0 Then

        Else
            Dim child As New map_EditPlatformFee
            child.MdiParent = Home
            child.Show()
            child.cboPlatform.SelectedValue = cboPlatform.SelectedValue
            child.cboProduct.SelectedValue = DataGridView1.SelectedCells(1).Value
            child.ckbWL.Checked = True
            child.cboWLPlatform.SelectedValue = ID.Text
            child.ID.Text = "NEW"

        End If
    End Sub

    Private Sub ckbCustomFee_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ckbCustomFee.CheckedChanged
        If ckbCustomFee.Checked Then
            AddFeeToolStripMenuItem.Visible = True
            EditFeeToolStripMenuItem.Visible = True
        Else
            AddFeeToolStripMenuItem.Visible = False
            EditFeeToolStripMenuItem.Visible = False
        End If
    End Sub

    Private Sub EditFeeToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles EditFeeToolStripMenuItem.Click
        Call LoadFeeEdit()
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
                child.ckbWL.Checked = True

                If IsDBNull(row1("PlatformID")) Then
                    'do nothing
                Else
                    child.cboPlatform.SelectedValue = (row1("PlatformID"))
                End If

                If IsDBNull(row1("WLID")) Then

                Else
                    child.cboWLPlatform.SelectedValue = (row1("PlatformID"))
                End If

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

    Public Sub CheckForAutoApprove()
        'If ckbSupressAutoAdd.Checked Then

        'Else
        If CheckBox2.Checked = True And CheckBox1.Checked = True Then
            Dim ir As MsgBoxResult
            ir = MsgBox("You are about to automatically add all platform approved products to this record.  You will also be able to add custom approved products.  Are you sure you want to continue?", MsgBoxStyle.YesNo + MsgBoxStyle.Information, "Confirm Auto Add")
            If ir = MsgBoxResult.Yes Then
                If ID.Text = "NEW" Then
                    Call CheckForDupe()
                    If ID.Text = "NEW" Then
                        MsgBox("Record could not be saved.  Please correct issues and try again.", MsgBoxStyle.Critical, "Not Saved")
                        CheckBox2.Checked = False
                    Else
                        Call RemovedAutoApprovedRecords()
                        Call InsertPlatformApproved()
                        CheckBox1.Enabled = False
                    End If
                Else
                    Call RemovedAutoApprovedRecords()
                    Call InsertPlatformApproved()
                    CheckBox1.Enabled = False
                End If
            Else
                CheckBox2.Checked = False
                'Call RemovedAutoApprovedRecords()
            End If
        Else
            If CheckBox7.Checked Then
                'do nothing.  The other code will handle.
            Else
                CheckBox1.Enabled = True
            End If
        End If
        'End If
    End Sub

    Public Sub RemovedAutoApprovedRecords()
        Dim Mycn As OleDb.OleDbConnection
        Dim Command As OleDb.OleDbCommand
        Dim SQLstr As String
        Try
            Mycn = New OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & My.Settings.dbloc & "';Jet OLEDB:Database Password='" & My.Settings.dbpass & "';")
            Mycn.Open()

            SQLstr = "DELETE * FROM map_PlatformListings WHERE WLID = " & ID.Text & " AND ProductID In (Select ProductID FROM Map_PlatformListings WHERE PlatformID = " & cboPlatform.SelectedValue & " AND AutoAdded = True AND Active = True)"

            Command = New OleDb.OleDbCommand(SQLstr, Mycn)
            Command.ExecuteNonQuery()

            Mycn.Close()

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "ERROR")
        End Try
    End Sub

    Private Sub CheckBox1_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckBox1.CheckedChanged
        Call CheckForAutoApprove()
        If CheckBox1.Checked = True Then

        Else
            Call RemovedAutoApprovedRecords()
        End If
    End Sub

    Private Sub cboPlatform_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cboPlatform.SelectedIndexChanged
        'Call CheckPlatformProducts()
    End Sub
End Class