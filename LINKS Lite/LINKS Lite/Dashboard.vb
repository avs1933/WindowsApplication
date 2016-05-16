Public Class Dashboard
    'Thread1 controls At-a-Glance
    Dim thread1 As System.Threading.Thread
    'Thread2 controls At-a-Glance External
    Dim thread2 As System.Threading.Thread
    'Thread3 controls At-a-Glance External
    Dim thread3 As System.Threading.Thread
    Private Sub Dashboard_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        Call AdvisorBreakdown()

        'thread3 = New System.Threading.Thread(AddressOf AdvisorBreakdown)
        'thread3.Start()

    End Sub

    Public Sub AtAGlance()
        Dim Mycn As OleDb.OleDbConnection

        Try

            Mycn = New OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & My.Settings.dbloc & "';Jet OLEDB:Database Password='" & My.Settings.dbpass & "';")

            Mycn.Open()

            Dim sqlstring As String
            sqlstring = "SELECT Count(env_BSC.PortfolioCode) AS Accounts1, Sum(TMV) AS TMV1, Sum(AAM1) AS AAM1" & _
            " FROM(env_BSC)" & _
            " WHERE env_BSC.AAMRepName = (SELECT APXName FROM sys_Users WHERE ID = " & My.Settings.userid & ")"

            Dim queryString As String = String.Format(sqlstring)
            Dim cmd As New OleDb.OleDbCommand(queryString, Mycn)
            Dim da As New OleDb.OleDbDataAdapter(cmd)
            Dim ds As New DataSet

            da.Fill(ds, "User")
            Dim dt As DataTable = ds.Tables("User")
            If dt.Rows.Count > 0 Then

                Dim row As DataRow = dt.Rows(0)


                Dim Billing As Double
                If IsDBNull((row("AAM1"))) Then
                    Billing = 0.0
                Else
                    Billing = (row("AAM1"))
                End If

                Dim TMV As Integer
                If IsDBNull((row("TMV1"))) Then
                    TMV = 0.0
                Else
                    TMV = (row("TMV1"))
                End If

                Dim velocity As Double
                If TMV = 0.0 Or Billing = 0.0 Then
                    velocity = 0
                Else
                    velocity = ((Billing / TMV) * 100)
                End If

                Dim accts As Integer
                If IsDBNull((row("Accounts1"))) Then
                    accts = 0
                Else
                    accts = (row("Accounts1"))
                End If

                txtAccounts.Text = accts
                txtAUM.Text = Format(TMV, "$#,###.00")
                txtBSC.Text = Format(Billing, "$#,###.00")
                txtVelocity.Text = Format(velocity, "#.00")
            Else

            End If
            Mycn.Close()
            Mycn.Dispose()

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "ERROR")

        End Try
    End Sub

    Public Sub AtAGlance_External()
        Dim Mycn As OleDb.OleDbConnection

        Try

            Mycn = New OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & My.Settings.dbloc & "';Jet OLEDB:Database Password='" & My.Settings.dbpass & "';")

            Mycn.Open()

            Dim sqlstring As String
            sqlstring = "SELECT Count(env_BSC_External.PortfolioCode) AS Accounts1, Sum(TMV) AS TMV1, Sum(BSC1) AS AAM1" & _
            " FROM(env_BSC_External)" & _
            " WHERE env_BSC_External.AAMRepName = (SELECT APXName FROM sys_Users WHERE ID = " & My.Settings.userid & ")"

            Dim queryString As String = String.Format(sqlstring)
            Dim cmd As New OleDb.OleDbCommand(queryString, Mycn)
            Dim da As New OleDb.OleDbDataAdapter(cmd)
            Dim ds As New DataSet

            da.Fill(ds, "User")
            Dim dt As DataTable = ds.Tables("User")
            If dt.Rows.Count > 0 Then

                Dim row As DataRow = dt.Rows(0)


                Dim Billing As Double
                If IsDBNull((row("AAM1"))) Then
                    Billing = 0.0
                Else
                    Billing = (row("AAM1"))
                End If

                Dim TMV As Integer
                If IsDBNull((row("TMV1"))) Then
                    TMV = 0.0
                Else
                    TMV = (row("TMV1"))
                End If

                Dim velocity As Double
                If TMV = 0.0 Or Billing = 0.0 Then
                    velocity = 0
                Else
                    velocity = ((Billing / TMV) * 100)
                End If

                Dim accts As Integer
                If IsDBNull((row("Accounts1"))) Then
                    accts = 0
                Else
                    accts = (row("Accounts1"))
                End If

                txtAUSAccts.Text = accts
                txtAUS.Text = Format(TMV, "$#,###.00")
                txtAUSBSC.Text = Format(Billing, "$#,###.00")
                txtAUSVelocity.Text = Format(velocity, "#.00")
            Else

            End If
            Mycn.Close()
            Mycn.Dispose()

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "ERROR")

        End Try
    End Sub

    Private Sub Timer1_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer1.Tick
        If txtAccounts.Text = "Loading..." Or txtAUSAccts.Text = "Loading..." Then

        Else
            Dim accts As Integer
            accts = txtAccounts.Text
            Dim ausaccts As Integer
            ausaccts = txtAUSAccts.Text
            txtTAccts.Text = accts + ausaccts

            Dim aus As Double
            aus = txtAUS.Text
            Dim aum As Double
            aum = txtAUM.Text
            Dim taum As Double
            taum = aus + aum
            TxtTAssets.Text = Format(taum, "$#,###.00")

            Dim bsc As Double
            bsc = txtBSC.Text
            Dim ausbsc As Double
            ausbsc = txtAUSBSC.Text
            Dim tbsc As Double
            tbsc = bsc + ausbsc
            TxtTBSC.Text = Format(tbsc, "$#,###.00")

            Dim tvel As Double
            tvel = ((tbsc / taum) * 100)
            TxtTVel.Text = Format(tvel, "#.00")

            Timer1.Enabled = False

        End If
    End Sub

    Public Sub AdvisorBreakdown()
        'Try

        Dim conn As New OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & My.Settings.dbloc & "';Jet OLEDB:Database Password='" & My.Settings.dbpass & "';")

        Dim strSQL As String = "SELECT env_BSC_AAM.AAMRepName, env_BSC_AAM.ExternalAdvisor As Advisor, env_BSC_AAM.ExternalFirm As Firm, Count(env_BSC_AAM.PortfolioCode) AS [# of accounts], Sum(env_BSC_AAM.TotalMarketValue) AS TMV, Sum(env_BSC_AAM.BSC) AS BSC" & _
        " FROM(env_BSC_AAM)" & _
        " WHERE env_BSC_AAM.AAMRepName = (SELECT APXName FROM sys_Users WHERE ID = " & My.Settings.userid & ")" & _
        " GROUP BY env_BSC_AAM.AAMRepName, env_BSC_AAM.ExternalAdvisor, env_BSC_AAM.ExternalFirm;"

        Dim da As New OleDb.OleDbDataAdapter(strSQL, conn)
        Dim ds As New DataSet
        da.Fill(ds, "Users")

        With DataGridView1
            .DataSource = ds.Tables("Users")
            .Columns(0).Visible = False
            .Columns(4).DefaultCellStyle.Format = "c"
            .Columns(5).DefaultCellStyle.Format = "c"
        End With

        'txtAdvisorLoading.Visible = False
        'DataGridView1.Visible = True
        'Call scrollhelp()

        'Catch ex As Exception
        'MsgBox(ex.Message, MsgBoxStyle.Critical, "ERROR")

        'End Try

        Control.CheckForIllegalCrossThreadCalls = False
        thread1 = New System.Threading.Thread(AddressOf AtAGlance)
        thread1.Start()

        thread2 = New System.Threading.Thread(AddressOf AtAGlance_External)
        thread2.Start()

    End Sub

    Public Sub scrollhelp()
        Me.InitializeComponent()

        Dim hscroll As HScrollBar = Me.DataGridView1.Controls.OfType(Of HScrollBar)() _
        .FirstOrDefault()
        Dim vscroll As VScrollBar = Me.DataGridView1.Controls.OfType(Of VScrollBar)() _
        .FirstOrDefault()

        If hscroll IsNot Nothing Then
            AddHandler hscroll.Scroll, _
            AddressOf Me.OnDataGridViewScroll
        End If

        If vscroll IsNot Nothing Then
            AddHandler vscroll.Scroll, _
            AddressOf Me.OnDataGridViewScroll
        End If

    End Sub

    Protected Sub OnDataGridViewScroll(ByVal sender As Object, _
    ByVal e As ScrollEventArgs)
        If e.Type = ScrollEventType.EndScroll Then
            Me.DataGridView1.Refresh()
        End If
    End Sub

    Private Sub Timer2_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer2.Tick
        'DataGridView1.Refresh()
        'Timer2.Enabled = False
    End Sub
End Class