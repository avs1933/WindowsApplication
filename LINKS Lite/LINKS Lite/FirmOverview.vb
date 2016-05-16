Public Class FirmOverview
    '***ALL DATA QUERIED FROM env_FirmsAll Query
    'Thread1 controls approvals
    Dim thread1 As System.Threading.Thread
    'Thread2 controls Firm Contact Info AND Notes
    Dim thread2 As System.Threading.Thread
    'Thread3 controls Fee schedule
    Dim thread3 As System.Threading.Thread
    'Thread4 controls marketing support
    Dim thread4 As System.Threading.Thread

    Private Sub FirmOverview_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        'NOTE FORM LEVEL PERMISSION!
        If Permissions.EditFirmFee.Checked Then
            OK.Enabled = True
        Else
            OK.Enabled = False
        End If

        If Permissions.ViewAPX.Checked Then
            ViewAPX.Visible = True
        Else
            ViewAPX.Visible = False
        End If

        Control.CheckForIllegalCrossThreadCalls = False
        'hide data for loading
        Label11.Visible = True
        ConservativeTaxable.Visible = False
        CorePlus.Visible = False
        MortgageInvestmentPortfolio.Visible = False
        CreditOpportunities.Visible = False
        ShortTermTaxExempt.Visible = False
        CoreTaxExempt.Visible = False
        High50Dividend.Visible = False
        PeroniMethod.Visible = False
        TAMTaxable.Visible = False
        TAMTaxExempt.Visible = False

        Label12.Visible = True
        Label14.Visible = True
        Label1.Visible = False
        Label4.Visible = False
        Label5.Visible = False
        Label6.Visible = False
        Label7.Visible = False
        Label8.Visible = False
        Label9.Visible = False
        Label10.Visible = False
        PlatformOnly.Visible = False
        AdditionalPlatformApproval.Visible = False
        UMA.Visible = False
        AgreementSigned.Visible = False
        ERISA.Visible = False
        EPaperwork.Visible = False
        SubAdvised.Visible = False
        Solicited.Visible = False
        Advised.Visible = False
        AgreementVersion.Visible = False
        DateExecuted.Visible = False
        Platforms.Visible = False
        PlatformsWhiteLabel.Visible = False
        Notes.Visible = False
        Address.Visible = False
        Phone.Visible = False
        Fax.Visible = False
        URL.Visible = False

        Label13.Visible = True
        DataGridView1.Visible = False
        OK.Visible = False
        Button1.Visible = False

        thread1 = New System.Threading.Thread(AddressOf LoadApprovals)
        thread1.Start()

        thread2 = New System.Threading.Thread(AddressOf LoadFirmData)
        thread2.Start()

    End Sub

    Public Sub LoadApprovals()
        Dim Mycn As OleDb.OleDbConnection

        Try

            Mycn = New OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & My.Settings.dbloc & "';Jet OLEDB:Database Password='" & My.Settings.dbpass & "';")

            Mycn.Open()

            Dim sqlstring As String

            sqlstring = "SELECT * FROM env_FirmsAll WHERE AdvApp_vContact.ContactID = " & ContactID.Text & ";"

            Dim queryString As String = String.Format(sqlstring)
            Dim cmd As New OleDb.OleDbCommand(queryString, Mycn)
            Dim da As New OleDb.OleDbDataAdapter(cmd)
            Dim ds As New DataSet

            da.Fill(ds, "User")
            Dim dt As DataTable = ds.Tables("User")
            If dt.Rows.Count > 0 Then

                Dim row As DataRow = dt.Rows(0)

                ConservativeTaxable.Checked = (row("AdvApp_vContactCustom.HasPref10"))
                CorePlus.Checked = (row("AdvApp_vContactCustom.HasPref11"))
                MortgageInvestmentPortfolio.Checked = (row("AdvApp_vContactCustom.HasPref12"))
                CreditOpportunities.Checked = (row("AdvApp_vContactCustom.HasPref13"))
                ShortTermTaxExempt.Checked = (row("AdvApp_vContactCustom.HasPref14"))
                CoreTaxExempt.Checked = (row("AdvApp_vContactCustom.HasPref15"))
                High50Dividend.Checked = (row("AdvApp_vContactCustom.HasPref16"))
                PeroniMethod.Checked = (row("AdvApp_vContactCustom.HasPref17"))
                TAMTaxable.Checked = (row("AdvApp_vContactCustom.HasPref18"))
                TAMTaxExempt.Checked = (row("AdvApp_vContactCustom.HasPref19"))

                'show data
                Label11.Visible = False
                ConservativeTaxable.Visible = True
                ConservativeTaxable.Enabled = False
                CorePlus.Visible = True
                CorePlus.Enabled = False
                MortgageInvestmentPortfolio.Visible = True
                MortgageInvestmentPortfolio.Enabled = False
                CreditOpportunities.Visible = True
                CreditOpportunities.Enabled = False
                ShortTermTaxExempt.Visible = True
                ShortTermTaxExempt.Enabled = False
                CoreTaxExempt.Visible = True
                CoreTaxExempt.Enabled = False
                High50Dividend.Visible = True
                High50Dividend.Enabled = False
                PeroniMethod.Visible = True
                PeroniMethod.Enabled = False
                TAMTaxable.Visible = True
                TAMTaxable.Enabled = False
                TAMTaxExempt.Visible = True
                TAMTaxExempt.Enabled = False

            Else

            End If
            Mycn.Close()
            Mycn.Dispose()

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "ERROR")

        End Try
    End Sub

    Public Sub LoadFirmData()
        Dim Mycn As OleDb.OleDbConnection

        Try

            Mycn = New OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & My.Settings.dbloc & "';Jet OLEDB:Database Password='" & My.Settings.dbpass & "';")

            Mycn.Open()

            Dim sqlstring As String

            sqlstring = "SELECT * FROM env_FirmsAll WHERE AdvApp_vContact.ContactID = " & ContactID.Text & ";"

            Dim queryString As String = String.Format(sqlstring)
            Dim cmd As New OleDb.OleDbCommand(queryString, Mycn)
            Dim da As New OleDb.OleDbDataAdapter(cmd)
            Dim ds As New DataSet

            da.Fill(ds, "User")
            Dim dt As DataTable = ds.Tables("User")
            If dt.Rows.Count > 0 Then

                Dim row As DataRow = dt.Rows(0)

                PlatformOnly.Checked = (row("AdvApp_vContactCustom.HasPref01"))
                AdditionalPlatformApproval.Checked = (row("AdvApp_vContactCustom.HasPref02"))
                UMA.Checked = (row("AdvApp_vContactCustom.HasPref03"))
                AgreementSigned.Checked = (row("AdvApp_vContactCustom.HasPref04"))
                ERISA.Checked = (row("AdvApp_vContactCustom.HasPref20"))
                EPaperwork.Checked = (row("AdvApp_vContactCustom.HasPref05"))
                SubAdvised.Checked = (row("AdvApp_vContactCustom.HasPref06"))
                Solicited.Checked = (row("AdvApp_vContactCustom.HasPref07"))
                Advised.Checked = (row("AdvApp_vContactCustom.HasPref08"))

                If IsDBNull(row("AdvApp_vContactCustom.Custom14")) Then
                    AgreementVersion.Text = "NO AGREEMENT"
                Else
                    AgreementVersion.Text = (row("Custom14"))
                End If

                If IsDBNull(row("AdvApp_vContactCustom.Custom15")) Then
                    DateExecuted.Text = "N/A"
                Else
                    DateExecuted.Text = (row("AdvApp_vContactCustom.Custom15"))
                End If

                If IsDBNull(row("AdvApp_vContactCustom.Custom16")) Then
                    Platforms.Text = "**UNKNOWN**"
                Else
                    Platforms.Text = (row("AdvApp_vContactCustom.Custom16"))
                End If

                If IsDBNull(row("AdvApp_vContactCustom.Custom17")) Then
                    PlatformsWhiteLabel.Text = "**UNKNOWN**"
                Else
                    PlatformsWhiteLabel.Text = (row("AdvApp_vContactCustom.Custom17"))
                End If

                If IsDBNull(row("Notes")) Then
                    Notes.Text = "**No Notes Found**"
                Else
                    Notes.Text = (row("Notes"))
                End If

                If IsDBNull(row("AddressFull")) Then
                    Address.Text = "No Address on file"
                Else
                    Address.Text = (row("AddressFull"))
                End If

                If IsDBNull(row("BusinessPhone")) Then
                    Phone.Text = "(000) 000-0000"
                Else
                    Phone.Text = (row("BusinessPhone"))
                End If

                If IsDBNull(row("OtherPhone")) Then
                    Fax.Text = "(000) 000-0000"
                Else
                    Fax.Text = (row("OtherPhone"))
                End If

                If IsDBNull(row("URL")) Then
                    URL.Text = "**UNKNOWN**"
                Else
                    URL.Text = (row("URL"))
                End If



                'check for custom fee schedule
                CheckBox1.Checked = (row("AdvApp_vContactCustom.HasPref09"))

                'show data
                Label12.Visible = False
                Label14.Visible = False
                Label1.Visible = True
                Label4.Visible = True
                Label5.Visible = True
                Label6.Visible = True
                Label7.Visible = True
                Label8.Visible = True
                Label9.Visible = True
                Label10.Visible = True
                PlatformOnly.Visible = True
                PlatformOnly.Enabled = False
                AdditionalPlatformApproval.Visible = True
                AdditionalPlatformApproval.Enabled = False
                UMA.Visible = True
                UMA.Enabled = False
                AgreementSigned.Visible = True
                AgreementSigned.Enabled = False
                ERISA.Visible = True
                ERISA.Enabled = False
                EPaperwork.Visible = True
                EPaperwork.Enabled = False
                SubAdvised.Visible = True
                SubAdvised.Enabled = False
                Solicited.Visible = True
                Solicited.Enabled = False
                Advised.Visible = True
                Advised.Enabled = False
                AgreementVersion.Visible = True
                DateExecuted.Visible = True
                Platforms.Visible = True
                PlatformsWhiteLabel.Visible = True
                Notes.Visible = True
                Address.Visible = True
                Phone.Visible = True
                Fax.Visible = True
                URL.Visible = True

            Else

            End If
            Mycn.Close()
            Mycn.Dispose()

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "ERROR")

        End Try
    End Sub

    Public Sub LoadFeeSchedule()
        Try

            Dim conn As New OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & My.Settings.dbloc & "';Jet OLEDB:Database Password='" & My.Settings.dbpass & "';")

            Dim strSQL As String = "SELECT ID, DisciplineNme As [Discipline], Fee1 As [Fee 1], Fee2 As [Fee 2], SubAdvised As [Sub-Advised], Advised As [Advised] FROM env_Fees WHERE CompanyID = " & ContactID.Text & " AND Active = -1"
            Dim da As New OleDb.OleDbDataAdapter(strSQL, conn)
            Dim ds As New DataSet
            da.Fill(ds, "Users")

            With DataGridView1
                .DataSource = ds.Tables("Users")
                .Columns(0).Visible = False
            End With

            Label13.Visible = False
            DataGridView1.Visible = True
            OK.Visible = True
            Button1.Visible = True

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "ERROR")

        End Try
    End Sub

    Public Sub LoadDefaultFees()
        Try

            Dim conn As New OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & My.Settings.dbloc & "';Jet OLEDB:Database Password='" & My.Settings.dbpass & "';")
            Dim strSQL As String

            If ((SubAdvised.Checked) And (Solicited.Checked = False) And (Advised.Checked = False)) Then
                strSQL = "SELECT ID, DisciplineNme As [Discipline], Fee1 As [Fee 1], Fee2 As [Fee 2] FROM env_Fees WHERE Default = -1 AND Active = -1 AND SubAdvised = -1"
            Else
                If ((Solicited.Checked Or Advised.Checked) And SubAdvised.Checked = False) Then
                    strSQL = "SELECT ID, DisciplineNme As [Discipline], Fee1 As [Fee 1], Fee2 As [Fee 2] FROM env_Fees WHERE Default = -1 AND Active = -1 AND Advised = -1"
                Else
                    strSQL = "SELECT ID, DisciplineNme As [Discipline], Fee1 As [Fee 1], Fee2 As [Fee 2], SubAdvised As [Sub-Advised], Advised As [Advised] FROM env_Fees WHERE Default = -1 AND Active = -1"
                End If
            End If

            Dim da As New OleDb.OleDbDataAdapter(strSQL, conn)
            Dim ds As New DataSet
            da.Fill(ds, "Users")

            With DataGridView1
                .DataSource = ds.Tables("Users")
                .Columns(0).Visible = False
            End With

            Label13.Visible = False
            DataGridView1.Visible = True
            OK.Visible = True
            Button1.Visible = True

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "ERROR")

        End Try
    End Sub

    Public Sub LoadMarketing()

    End Sub

    Private Sub Timer1_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer1.Tick
        If CheckBox1.Checked Then
            If CheckBox2.Checked Then
                'do nothing
            Else
                CheckBox2.Checked = True
                thread3 = New System.Threading.Thread(AddressOf LoadFeeSchedule)
                thread3.Start()
            End If

            If CustomFeeSchedule.Visible = True Then
                CustomFeeSchedule.Visible = False
            Else
                CustomFeeSchedule.Visible = False
                CustomFeeSchedule.Visible = True
            End If
        Else
            If AgreementSigned.Checked Then
                If CheckBox2.Checked Then
                    'do nothing
                Else
                    CheckBox2.Checked = True
                    thread3 = New System.Threading.Thread(AddressOf LoadDefaultFees)
                    thread3.Start()
                End If
            Else
                'Label13.Text = "UNKNOWN FEE SCHEDULE"
            End If
        End If
    End Sub

    Private Sub ViewAPX_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ViewAPX.Click
        APXBrowser.Show()
        Dim str As String
        str = "http://adventapps/APXLogin/ContactDetail.aspx?linkfield=" & ContactID.Text
        'MsgBox(str) 'Added to verify data
        APXBrowser.wbAdvent.Navigate(str)
    End Sub

    Private Sub EditToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles EditToolStripMenuItem.Click
        If ((AgreementSigned.Checked) And (CheckBox1.Checked = False) And (Permissions.SetDefaultFee.Checked = False)) Then
            MsgBox("You do not have permission to alter default fee schedules", MsgBoxStyle.Information, "Permissions Issue")
            GoTo Line1
        Else

        End If

        If Permissions.EditFirmFee.Checked Then

            If DataGridView1.RowCount = "0" Then
                'do nothing
            Else
                Dim child As New FeeUpdate
                child.MdiParent = Home

                child.ID.Text = DataGridView1.SelectedCells(0).Value
                child.Show()

                If Permissions.SetDefaultFee.Checked Then
                    child.Company.Enabled = True
                Else
                    child.Company.Enabled = False
                End If

            End If

        Else
            MsgBox("You do not have permissions to perform this function.", MsgBoxStyle.Information, "Permissions")
        End If

line1:
    End Sub

    Private Sub OK_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles OK.Click
        If ((AgreementSigned.Checked) And (CheckBox1.Checked = False) And (Permissions.SetDefaultFee.Checked = False)) Then
            MsgBox("You do not have permission to alter default fee schedules", MsgBoxStyle.Information, "Permissions Issue")
            GoTo Line1
        Else

        End If

        Dim child As New FeeUpdate
        child.MdiParent = Home
        child.ID.Text = "NEW"
        child.Show()
        child.Company.SelectedValue = ContactID.Text

        If Permissions.SetDefaultFee.Checked Then
            child.Company.Enabled = True
        Else
            child.Company.Enabled = False
        End If

line1:
    End Sub

    Public Sub ReloadFees()

        Label13.Visible = True
        DataGridView1.Visible = False
        OK.Visible = False
        Button1.Visible = False

        If CheckBox1.Checked Then
            Try
                Dim conn As New OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & My.Settings.dbloc & "';Jet OLEDB:Database Password='" & My.Settings.dbpass & "';")
                Dim strSQL As String = "SELECT ID, DisciplineNme As [Discipline], Fee1 As [Fee 1], Fee2 As [Fee 2], SubAdvised As [Sub-Advised], Advised As [Advised] FROM env_Fees WHERE CompanyID = " & ContactID.Text & " AND Active = -1"
                Dim da As New OleDb.OleDbDataAdapter(strSQL, conn)
                Dim ds As New DataSet
                da.Fill(ds, "Users")

                With DataGridView1
                    .DataSource = ds.Tables("Users")
                    .Columns(0).Visible = False
                End With

            Catch ex As Exception
                MsgBox(ex.Message, MsgBoxStyle.Critical, "ERROR")

            End Try
        Else
            Try
                Dim conn As New OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & My.Settings.dbloc & "';Jet OLEDB:Database Password='" & My.Settings.dbpass & "';")
                Dim strSQL As String

                If ((SubAdvised.Checked) And (Solicited.Checked = False) And (Advised.Checked = False)) Then
                    strSQL = "SELECT ID, DisciplineNme As [Discipline], Fee1 As [Fee 1], Fee2 As [Fee 2] FROM env_Fees WHERE Default = -1 AND Active = -1 AND SubAdvised = -1"
                Else
                    If ((Solicited.Checked Or Advised.Checked) And SubAdvised.Checked = False) Then
                        strSQL = "SELECT ID, DisciplineNme As [Discipline], Fee1 As [Fee 1], Fee2 As [Fee 2] FROM env_Fees WHERE Default = -1 AND Active = -1 AND Advised = -1"
                    Else
                        strSQL = "SELECT ID, DisciplineNme As [Discipline], Fee1 As [Fee 1], Fee2 As [Fee 2], SubAdvised As [Sub-Advised], Advised As [Advised] FROM env_Fees WHERE Default = -1 AND Active = -1"
                    End If
                End If

                Dim da As New OleDb.OleDbDataAdapter(strSQL, conn)
                Dim ds As New DataSet
                da.Fill(ds, "Users")

                With DataGridView1
                    .DataSource = ds.Tables("Users")
                    .Columns(0).Visible = False
                End With

            Catch ex As Exception
                MsgBox(ex.Message, MsgBoxStyle.Critical, "ERROR")

            End Try
        End If

        Label13.Visible = False
        DataGridView1.Visible = True
        OK.Visible = True
        Button1.Visible = True

    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        thread3 = New System.Threading.Thread(AddressOf ReloadFees)
        thread3.Start()
    End Sub

    Private Sub DeleteToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles DeleteToolStripMenuItem.Click
        If ((AgreementSigned.Checked) And (CheckBox1.Checked = False) And (Permissions.SetDefaultFee.Checked = False)) Then
            MsgBox("You do not have permission to alter default fee schedules", MsgBoxStyle.Information, "Permissions Issue")
            GoTo Line1
        Else

        End If

        If Permissions.EditFirmFee.Checked Then

            If DataGridView1.RowCount = "0" Then
                'do nothing
            Else
                Dim str As Integer

                str = MsgBox("You are about to Delete this fee schedule.  Are you sure you want to continue?", MsgBoxStyle.YesNo + MsgBoxStyle.Exclamation, "Confirm Delete")

                If str = MsgBoxResult.Yes Then
                    Dim Mycn As OleDb.OleDbConnection
                    Dim Command As OleDb.OleDbCommand
                    Dim SQLstr As String

                    Try
                        Mycn = New OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & My.Settings.dbloc & "';Jet OLEDB:Database Password='" & My.Settings.dbpass & "';")

                        Dim ds1 As New DataSet
                        Dim eds1 As New DataGridView
                        Dim dv1 As New DataView

                        Mycn.Open()

                        SQLstr = "Update env_Fees SET [Active] = NULL WHERE ID = " & DataGridView1.SelectedCells(0).Value

                        Command = New OleDb.OleDbCommand(SQLstr, Mycn)
                        Command.ExecuteNonQuery()

                        Mycn.Close()

                        thread3 = New System.Threading.Thread(AddressOf ReloadFees)
                        thread3.Start()

                    Catch ex As Exception
                        MsgBox(ex.Message, MsgBoxStyle.Critical, "ERROR")

                    End Try
                Else
                    'do nothing
                End If
            End If

        Else
            MsgBox("You do not have permissions to perform this function.", MsgBoxStyle.Information, "Permissions")
        End If
line1:

    End Sub
End Class