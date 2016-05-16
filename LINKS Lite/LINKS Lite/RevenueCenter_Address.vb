Public Class RevenueCenter_Address
    Dim loadfirms As System.Threading.Thread

    Private Sub RevenueCenter_Address_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        WebBrowser1.Navigate(Application.StartupPath & "\resources\159.gif")
        Call LoadAPXFirms()
    End Sub

    Public Sub LoadAPXFirms()
        Try

            Dim conn As New OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & My.Settings.dbloc & "';Jet OLEDB:Database Password='" & My.Settings.dbpass & "';")

            Dim strSQL As String = "SELECT DISTINCT dbo_vQbRowDefPortfolio.ExternalFirm" & _
            " FROM(dbo_vQbRowDefPortfolio)" & _
            " WHERE(((dbo_vQbRowDefPortfolio.ManagingFirm) Is Not Null))" & _
            " GROUP BY dbo_vQbRowDefPortfolio.ExternalFirm" & _
            " ORDER BY dbo_vQbRowDefPortfolio.ExternalFirm;"

            Dim da As New OleDb.OleDbDataAdapter(strSQL, conn)
            Dim ds As New DataSet
            da.Fill(ds, "Users")

            With cboAPXFirm
                .DataSource = ds.Tables("Users")
                .DisplayMember = "ExternalFirm"
                .ValueMember = "ExternalFirm"
                .SelectedIndex = 0
            End With

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "ERROR")
        End Try
    End Sub

    Private Sub cboAPXFirm_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles cboAPXFirm.LostFocus
        
    End Sub

    Public Sub LoadAddressData()
        Label13.Text = "Looking for address inforamtion..."
        Label13.Visible = True
        Try
            Dim conn As New OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & My.Settings.dbloc & "';Jet OLEDB:Database Password='" & My.Settings.dbpass & "';")

            Dim strSQL As String = "SELECT * FROM map_Firms WHERE Active = -1 AND TypeID = 2 AND AdventPortfolioFirm = '" & cboAPXFirm.Text & "'"
            Dim queryString As String = String.Format(strSQL)
            Dim cmd As New OleDb.OleDbCommand(queryString, conn)
            Dim da As New OleDb.OleDbDataAdapter(cmd)
            Dim ds As New DataSet

            da.Fill(ds, "User")
            Dim dt As DataTable = ds.Tables("User")
            If dt.Rows.Count > 0 Then
                Dim row1 As DataRow = dt.Rows(0)

                MapID.Text = (row1("ID"))

                If IsDBNull(row1("FirmName")) Then
                    TextBox1.Text = "UNKNOWN"
                Else
                    TextBox1.Text = (row1("FirmName"))
                End If

                If IsDBNull(row1("Address")) Then
                    TextBox3.Text = "UNKNOWN"
                Else
                    TextBox3.Text = (row1("Address"))
                End If

                If IsDBNull(row1("City")) Then
                    TextBox5.Text = "UNKNOWN"
                Else
                    TextBox5.Text = (row1("City"))
                End If

                If IsDBNull(row1("State")) Then
                    TextBox6.Text = "UNKNOWN"
                Else
                    TextBox6.Text = (row1("State"))
                End If

                If IsDBNull(row1("Zip")) Then
                    TextBox7.Text = "UNKNOWN"
                Else
                    TextBox7.Text = (row1("Zip"))
                End If

                If IsDBNull(row1("Phone")) Then
                    TextBox8.Text = "UNKNOWN"
                Else
                    TextBox8.Text = (row1("Phone"))
                End If
                WebBrowser1.Visible = False
                Label13.Visible = False
                ckbLocked.Checked = True
            Else
                Label13.Text = "Couldnt find address inforamtion..."
                MapID.Text = ""
                WebBrowser1.Visible = False
            End If

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "ERROR")
        End Try

    End Sub

    Private Sub cboAPXFirm_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cboAPXFirm.SelectedIndexChanged
        WebBrowser1.Visible = True

        If IsDBNull(cboAPXFirm.SelectedValue) Then

        Else
            If ckbLocked.Checked = True Then
                Dim ir As MsgBoxResult
                ir = MsgBox("This Record has been locked.  Do you want to auto refresh the data?", MsgBoxStyle.YesNo + MsgBoxStyle.Information, "Refresh Data?")
                If ir = MsgBoxResult.Yes Then
                    Control.CheckForIllegalCrossThreadCalls = False
                    loadfirms = New System.Threading.Thread(AddressOf LoadAddressData)
                    loadfirms.Start()
                    'Call LoadAddressData()
                Else
                    WebBrowser1.Visible = False
                    MapID.Text = ""
                End If
            Else
                Control.CheckForIllegalCrossThreadCalls = False
                loadfirms = New System.Threading.Thread(AddressOf LoadAddressData)
                loadfirms.Start()
            End If
        End If
    End Sub

    Private Sub Button12_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button12.Click
        If ID.Text = "NEW" Then
            Call CheckForDupe()
        Else
            Call SaveOld()
        End If
    End Sub

    Public Sub CheckForDupe()
        'check for dupes
        Dim Mycn As OleDb.OleDbConnection
        Dim SQLstring As String
        Try
            Mycn = New OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & My.Settings.dbloc & "';Jet OLEDB:Database Password='" & My.Settings.dbpass & "';")
            Mycn.Open()

            SQLstring = "SELECT * FROM mdb_BillingFirms WHERE [APXFirmName] = '" & cboAPXFirm.Text & "'"

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
    End Sub

    Public Sub SaveNew()
        Dim Mycn As OleDb.OleDbConnection
        Dim Command As OleDb.OleDbCommand
        Dim SQLstr As String
        Try
            Mycn = New OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & My.Settings.dbloc & "';Jet OLEDB:Database Password='" & My.Settings.dbpass & "';")
            Mycn.Open()

            SQLstr = "INSERT INTO mdb_BillingFirms(APXFirmName, FirmName, FirmAttn, Address1, Address2, City, State, Zip, Phone1, Extention, MAPFirmID, Locked)" & _
            "VALUES('" & cboAPXFirm.Text & "','" & TextBox1.Text & "','" & TextBox2.Text & "','" & TextBox3.Text & "','" & TextBox4.Text & "','" & TextBox5.Text & "','" & TextBox6.Text & "','" & TextBox7.Text & "', '" & TextBox8.Text & "', '" & TextBox9.Text & "'," & MapID.Text & ",-1)"

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
        'Try
        Mycn = New OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & My.Settings.dbloc & "';Jet OLEDB:Database Password='" & My.Settings.dbpass & "';")
        Mycn.Open()

        SQLstr = "Update mdb_BillingFirms SET APXFirmName = '" & cboAPXFirm.Text & "' AND FirmName = '" & TextBox1.Text & "', FirmAttn = '" & TextBox2.Text & "', Address1 = '" & TextBox3.Text & "', Address2 = '" & TextBox4.Text & "', City = '" & TextBox5.Text & "', State = '" & TextBox6.Text & "', Zip = '" & TextBox7.Text & "', Phone1 = '" & TextBox8.Text & "', Extention = '" & TextBox9.Text & "', MAPFirmID = " & MapID.Text & ", Locked = " & ckbLocked.Checked & " WHERE ID = " & ID.Text

        Command = New OleDb.OleDbCommand(SQLstr, Mycn)
        Command.ExecuteNonQuery()

        Mycn.Close()

        MsgBox("Record Updated.")
        Me.Close()

        'Catch ex As Exception
        'MsgBox(ex.Message, MsgBoxStyle.Critical, "ERROR")
        'End Try
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Me.Close()
    End Sub
End Class