Public Class map_EditMFirms
    Dim loadfirms As System.Threading.Thread
    Dim loadtypes As System.Threading.Thread
    Dim loadpfirms As System.Threading.Thread


    Private Sub map_EditMFirms_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        'Control.CheckForIllegalCrossThreadCalls = False
        'loadfirms = New System.Threading.Thread(AddressOf LoadAPXFirms)
        'loadfirms.Start()

        'loadtypes = New System.Threading.Thread(AddressOf LoadContactType)
        'loadtypes.Start()

        'loadpfirms = New System.Threading.Thread(AddressOf loadpfirms1)
        'loadpfirms.Start()
        Call LoadAPXFirms()
        Call LoadContactType()
        'Call LoadPFirms1()

    End Sub

    Public Sub LoadAPXFirms()
        Try

            Dim conn As New OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & My.Settings.dbloc & "';Jet OLEDB:Database Password='" & My.Settings.dbpass & "';")

            Dim strSQL As String = "SELECT ContactID, DeliveryName FROM map_APXFirms ORDER BY DeliveryName"
            Dim da As New OleDb.OleDbDataAdapter(strSQL, conn)
            Dim ds As New DataSet
            da.Fill(ds, "Users")

            With cboAPXFirms
                .DataSource = ds.Tables("Users")
                .DisplayMember = "DeliveryName"
                .ValueMember = "ContactID"
                .SelectedIndex = 0
            End With

            Label14.Visible = False
            cboAPXFirms.Enabled = True
            'Call LoadDBData()

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "ERROR")
        End Try
    End Sub

    Public Sub LoadPFirms1()
        Try

            Dim conn As New OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & My.Settings.dbloc & "';Jet OLEDB:Database Password='" & My.Settings.dbpass & "';")

            Dim strSQL As String = "SELECT DataValue, DisplayValue FROM map_FirmPortfolioLookup ORDER BY DisplayValue"
            Dim da As New OleDb.OleDbDataAdapter(strSQL, conn)
            Dim ds As New DataSet
            da.Fill(ds, "Users")

            With cboPortfolioFirm
                .DataSource = ds.Tables("Users")
                .DisplayMember = "DisplayValue"
                .ValueMember = "DataValue"
                .SelectedIndex = 0
            End With

            Label11.Visible = False
            cboPortfolioFirm.Enabled = True
            'Call LoadDBData()

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "ERROR")
        End Try
    End Sub

    Public Sub LoadContactType()
        Try

            Dim conn As New OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & My.Settings.dbloc & "';Jet OLEDB:Database Password='" & My.Settings.dbpass & "';")

            Dim strSQL As String = "SELECT ID, FirmType FROM map_FirmType WHERE ID <> 7 ORDER BY FirmType"
            Dim da As New OleDb.OleDbDataAdapter(strSQL, conn)
            Dim ds As New DataSet
            da.Fill(ds, "Users")

            With cboFirmType
                .DataSource = ds.Tables("Users")
                .DisplayMember = "FirmType"
                .ValueMember = "ID"
                .SelectedIndex = 0
            End With

            Label10.Visible = False
            cboFirmType.Enabled = True
            'Call LoadDBData()

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "ERROR")
        End Try
    End Sub

    Private Sub CheckBox1_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckBox1.CheckedChanged
        If CheckBox1.Checked Then
            TextBox1.Enabled = True
            TextBox2.Enabled = True
            TextBox3.Enabled = True
            TextBox4.Enabled = True
            TextBox5.Enabled = True
            TextBox6.Enabled = True
            cboAPXFirms.Enabled = False
            cboPortfolioFirm.Enabled = False
        Else
            cboAPXFirms.Enabled = True
            cboPortfolioFirm.Enabled = True
            Call LoadFirmData()
        End If
    End Sub

    Public Sub LoadFirmData()
        If IsDBNull(cboAPXFirms.SelectedValue) Then

        Else
            Try
                Dim conn As New OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & My.Settings.dbloc & "';Jet OLEDB:Database Password='" & My.Settings.dbpass & "';")

                Dim strSQL As String = "SELECT * FROM map_APXFirms WHERE ContactID = " & cboAPXFirms.SelectedValue
                Dim queryString As String = String.Format(strSQL)
                Dim cmd As New OleDb.OleDbCommand(queryString, conn)
                Dim da As New OleDb.OleDbDataAdapter(cmd)
                Dim ds As New DataSet

                da.Fill(ds, "User")
                Dim dt As DataTable = ds.Tables("User")

                Dim row1 As DataRow = dt.Rows(0)


                If IsDBNull(row1("AddressLine1")) Then
                    TextBox2.Text = "UNKNOWN"
                Else
                    TextBox2.Text = (row1("AddressLine1"))
                End If

                If IsDBNull(row1("AddressCity")) Then
                    TextBox3.Text = "UNKNOWN"
                Else
                    TextBox3.Text = (row1("AddressCity"))
                End If

                If IsDBNull(row1("AddressStateCode")) Then
                    TextBox4.Text = "UNKNOWN"
                Else
                    TextBox4.Text = (row1("AddressStateCode"))
                End If

                If IsDBNull(row1("AddressPostalCode")) Then
                    TextBox5.Text = "UNKNOWN"
                Else
                    TextBox5.Text = (row1("AddressPostalCode"))
                End If

                If IsDBNull(row1("BusinessPhone")) Then
                    TextBox6.Text = "UNKNOWN"
                Else
                    TextBox6.Text = (row1("BusinessPhone"))
                End If


                TextBox2.Enabled = False
                TextBox3.Enabled = False
                TextBox4.Enabled = False
                TextBox5.Enabled = False
                TextBox6.Enabled = False

                'Dim msgbox1 As MsgBoxResult

                If (row1("DeliveryName")) = TextBox1.Text Then
                    TextBox1.Enabled = False
                Else
                    Dim ir As MsgBoxResult

                    ir = MsgBox("The firm name you entered does not match the Advent contact record.  Would you like to update the firm name to match the advent record?", MsgBoxStyle.YesNo + MsgBoxStyle.Information, "Match APX Contact")

                    If ir = MsgBoxResult.Yes Then
                        If IsDBNull(row1("DeliveryName")) Then
                            TextBox1.Text = "UNKNOWN"
                        Else
                            TextBox1.Text = (row1("DeliveryName"))
                        End If
                        TextBox1.Enabled = False
                    Else
                        TextBox1.Enabled = True
                    End If
                End If

            Catch ex As Exception
                MsgBox(ex.Message, MsgBoxStyle.Critical, "ERROR")
            End Try

        End If
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        If ID.Text = "NEW" Then
            Call CheckForDupe()
        Else
            Call SaveOld()
        End If
    End Sub

    Public Sub CheckForDupe()

        If CheckBox1.Checked Then
            'do nothing
            Call SaveNew()
        Else
            'check for dupes
            Dim Mycn As OleDb.OleDbConnection
            Dim SQLstring As String
            Try
                Mycn = New OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & My.Settings.dbloc & "';Jet OLEDB:Database Password='" & My.Settings.dbpass & "';")
                Mycn.Open()

                SQLstring = "SELECT * FROM Map_Firms WHERE [AdventContactID] = " & cboAPXFirms.SelectedValue & " AND [AdventPortfolioFirm] = '" & cboAPXFirms.SelectedValue & "' AND NotInAdvent <> -1"

                Dim queryString As String = String.Format(SQLstring)
                Dim cmd As New OleDb.OleDbCommand(queryString, Mycn)
                Dim da As New OleDb.OleDbDataAdapter(cmd)
                Dim ds As New DataSet

                da.Fill(ds, "User")
                Dim dt As DataTable = ds.Tables("User")
                If dt.Rows.Count = 0 Then
                    Call SaveNew()
                Else
                    MsgBox("A record already exsists for that Firm.", MsgBoxStyle.Information, "DUPLICATE RECORD DETECTED")
                End If

            Catch ex As Exception
                MsgBox(ex.Message)
            End Try
        End If
    End Sub

    Public Sub SaveNew()
        Dim Mycn As OleDb.OleDbConnection
        Dim Command As OleDb.OleDbCommand
        Dim SQLstr As String
        Try
            Mycn = New OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & My.Settings.dbloc & "';Jet OLEDB:Database Password='" & My.Settings.dbpass & "';")
            Mycn.Open()

            SQLstr = "INSERT INTO map_Firms(FirmName, Address, City, State, Zip, Phone, TypeID, NotInAdvent, AdventContactID, AdventPortfolioFirm, Active)" & _
            "VALUES('" & TextBox1.Text & "','" & TextBox2.Text & "','" & TextBox3.Text & "','" & TextBox4.Text & "','" & TextBox5.Text & "','" & TextBox6.Text & "'," & cboFirmType.SelectedValue & "," & CheckBox1.CheckState & "," & cboAPXFirms.SelectedValue & ",'" & cboPortfolioFirm.SelectedValue & "', -1)"

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

            SQLstr = "Update Map_Firms SET FirmName = '" & TextBox1.Text & "', Address = '" & TextBox2.Text & "', City = '" & TextBox3.Text & "', State = '" & TextBox4.Text & "', Zip = '" & TextBox5.Text & "', Phone = '" & TextBox6.Text & "', TypeID = " & cboFirmType.SelectedValue & ", NotInAdvent = " & CheckBox1.CheckState & ", AdventContactID = " & cboAPXFirms.SelectedValue & ", AdventPortfolioFirm = '" & cboPortfolioFirm.SelectedValue & "' WHERE ID = " & ID.Text

            Command = New OleDb.OleDbCommand(SQLstr, Mycn)
            Command.ExecuteNonQuery()

            Mycn.Close()

            MsgBox("Record Updated.")
            Me.Close()

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "ERROR")
        End Try
    End Sub

    Private Sub cboAPXFirms_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles cboAPXFirms.LostFocus
        Call LoadFirmData()
    End Sub

    Private Sub cboAPXFirms_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cboAPXFirms.SelectedIndexChanged

    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        Me.Close()
    End Sub

    Private Sub cboFirmType_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cboFirmType.SelectedIndexChanged
        If IsDBNull(cboFirmType.SelectedValue) Then
            GoTo line1
        Else
            If ((cboFirmType.SelectedValue.ToString = "1") Or (cboFirmType.SelectedValue.ToString = "2") Or (cboFirmType.SelectedValue.ToString = "3")) Then
                Call LoadPFirms1()
            Else
                If cboFirmType.SelectedValue.ToString = "6" Then
                    Call LoadPlatformFirms()
                Else
                    cboPortfolioFirm.Enabled = False
                End If
            End If
        End If
line1:

    End Sub

    Public Sub LoadPlatformFirms()
        Try

            Dim conn As New OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & My.Settings.dbloc & "';Jet OLEDB:Database Password='" & My.Settings.dbpass & "';")

            Dim strSQL As String = "SELECT DataValue, DisplayValue FROM map_PlatformPortfolioLookup ORDER BY DisplayValue"
            Dim da As New OleDb.OleDbDataAdapter(strSQL, conn)
            Dim ds As New DataSet
            da.Fill(ds, "Users")

            With cboPortfolioFirm
                .DataSource = ds.Tables("Users")
                .DisplayMember = "DisplayValue"
                .ValueMember = "DataValue"
                .SelectedIndex = 0
            End With

            Label11.Visible = False
            cboPortfolioFirm.Enabled = True
            'Call LoadDBData()

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "ERROR")
        End Try
    End Sub
End Class