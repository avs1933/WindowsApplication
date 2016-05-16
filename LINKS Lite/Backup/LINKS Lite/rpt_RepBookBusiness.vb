Public Class rpt_RepBookBusiness

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Dim Mycn As OleDb.OleDbConnection
        Dim Command As OleDb.OleDbCommand
        Dim Command1 As OleDb.OleDbCommand
        Dim Command2 As OleDb.OleDbCommand

        Dim SQLstr As String

        Try
            Mycn = New OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & My.Settings.dbloc & "';Jet OLEDB:Database Password='" & My.Settings.dbpass & "';")

            Dim ds1 As New DataSet
            Dim eds1 As New DataGridView
            Dim dv1 As New DataView

            Mycn.Open()

            SQLstr = "DELETE * FROM mdb_RepRunReport"

            Command1 = New OleDb.OleDbCommand(SQLstr, Mycn)
            Command1.ExecuteNonQuery()

            Mycn.Close()

            Mycn.Open()

            SQLstr = "INSERT INTO mdb_RepRunReport (PortfolioCode, AccountName, ExtAdvisor, ExtFirm, AAMRep, AAMTeam, AAMRegion, AAMDepartment, Discipline, BillingRate, TMV, Cash, StartDate, EstBSC)" & _
            "SELECT Right(dbo_vQbRowDefPortfolio.PortfolioCode,4) AS PortfolioCode, dbo_vQbRowDefPortfolio.ReportHeading1, dbo_vQbRowDefPortfolio.ExternalAdvisor, dbo_vQbRowDefPortfolio.ExternalFirm, dbo_vQbRowDefPortfolio.AAMRepName, dbo_vQbRowDefPortfolio.AAMTeamName, dbo_vQbRowDefPortfolio.SalesRegion, dbo_vQbRowDefPortfolio.SalesDepartment, dbo_vQbRowDefPortfolio.Discipline6, dbo_vQbRowDefPortfolio.TieredBillingRate1, dbo_vQbRowDefPortfolio.TotalMarketValue, dbo_vQbRowDefPortfolio.TotalTradableCash, dbo_vQbRowDefPortfolio.StartDate, (dbo_vQbRowDefPortfolio.TotalMarketValue*dbo_vQbRowDefPortfolio.TieredBillingRate1)/100 AS EstBSC" & _
            " FROM(dbo_vQbRowDefPortfolio)" & _
            " WHERE(((dbo_vQbRowDefPortfolio.ManagerCode) = 'AAM') And ((dbo_vQbRowDefPortfolio.PortfolioStatus) = 'open'))"
            ' Or (dbo_vQbRowDefPortfolio.ManagerCode) = 'U AAM'
            If CheckBox1.Checked Then
                SQLstr = SQLstr & " AND dbo_vQbRowDefPortfolio.AAMRepName = '" & cboIntRep.Text & "'"
            End If

            If CheckBox2.Checked Then
                SQLstr = SQLstr & " AND dbo_vQbRowDefPortfolio.AAMTeamName = '" & cboTeam.Text & "'"
            End If

            If CheckBox3.Checked Then
                SQLstr = SQLstr & " AND dbo_vQbRowDefPortfolio.SalesRegion = '" & cboRegion.Text & "'"
            End If

            If CheckBox4.Checked Then
                SQLstr = SQLstr & " AND dbo_vQbRowDefPortfolio.SalesDepartment = '" & cboDepartment.Text & "'"
            End If

            SQLstr = SQLstr & " GROUP BY dbo_vQbRowDefPortfolio.ReportHeading1, dbo_vQbRowDefPortfolio.ExternalAdvisor, dbo_vQbRowDefPortfolio.ExternalFirm, dbo_vQbRowDefPortfolio.AAMRepName, dbo_vQbRowDefPortfolio.AAMTeamName, dbo_vQbRowDefPortfolio.SalesRegion, dbo_vQbRowDefPortfolio.SalesDepartment, dbo_vQbRowDefPortfolio.Discipline6, dbo_vQbRowDefPortfolio.TieredBillingRate1, dbo_vQbRowDefPortfolio.TotalMarketValue, dbo_vQbRowDefPortfolio.TotalTradableCash, dbo_vQbRowDefPortfolio.StartDate, dbo_vQbRowDefPortfolio.PortfolioCode"

            Command = New OleDb.OleDbCommand(SQLstr, Mycn)
            Command.ExecuteNonQuery()

            Mycn.Close()

            Mycn.Open()

            SQLstr = "INSERT INTO mdb_RepRunReport (PortfolioCode, AccountName, ExtAdvisor, ExtFirm, AAMRep, AAMTeam, AAMRegion, AAMDepartment, Discipline, BillingRate, TMV, Cash, StartDate, EstBSC)" & _
            "SELECT Right(dbo_vQbRowDefPortfolio.PortfolioCode,4) AS PortfolioCode, dbo_vQbRowDefPortfolio.ReportHeading1, dbo_vQbRowDefPortfolio.ExternalAdvisor, dbo_vQbRowDefPortfolio.ExternalFirm, dbo_vQbRowDefPortfolio.AAMRepName, dbo_vQbRowDefPortfolio.AAMTeamName, dbo_vQbRowDefPortfolio.SalesRegion, dbo_vQbRowDefPortfolio.SalesDepartment, dbo_vQbRowDefPortfolio.Discipline6, env_ExternalRate.Rate, dbo_vQbRowDefPortfolio.TotalMarketValue, dbo_vQbRowDefPortfolio.TotalTradableCash, dbo_vQbRowDefPortfolio.StartDate, (dbo_vQbRowDefPortfolio.TotalMarketValue*env_ExternalRate.Rate)/100 AS EstBSC" & _
            " FROM dbo_vQbRowDefPortfolio INNER JOIN env_ExternalRate ON dbo_vQbRowDefPortfolio.ManagerCode = env_ExternalRate.MCode" & _
            " WHERE(((dbo_vQbRowDefPortfolio.ManagerCode) = 'U AAM') And ((dbo_vQbRowDefPortfolio.PortfolioStatus) = 'open') And ((env_ExternalRate.Active) = True))"
            ' Or (dbo_vQbRowDefPortfolio.ManagerCode) = 'U AAM'
            If CheckBox1.Checked Then
                SQLstr = SQLstr & " AND dbo_vQbRowDefPortfolio.AAMRepName = '" & cboIntRep.Text & "'"
            End If

            If CheckBox2.Checked Then
                SQLstr = SQLstr & " AND dbo_vQbRowDefPortfolio.AAMTeamName = '" & cboTeam.Text & "'"
            End If

            If CheckBox3.Checked Then
                SQLstr = SQLstr & " AND dbo_vQbRowDefPortfolio.SalesRegion = '" & cboRegion.Text & "'"
            End If

            If CheckBox4.Checked Then
                SQLstr = SQLstr & " AND dbo_vQbRowDefPortfolio.SalesDepartment = '" & cboDepartment.Text & "'"
            End If

            SQLstr = SQLstr & " GROUP BY dbo_vQbRowDefPortfolio.ReportHeading1, dbo_vQbRowDefPortfolio.ExternalAdvisor, dbo_vQbRowDefPortfolio.ExternalFirm, dbo_vQbRowDefPortfolio.AAMRepName, dbo_vQbRowDefPortfolio.AAMTeamName, dbo_vQbRowDefPortfolio.SalesRegion, dbo_vQbRowDefPortfolio.SalesDepartment, dbo_vQbRowDefPortfolio.Discipline6, env_ExternalRate.Rate, dbo_vQbRowDefPortfolio.TotalMarketValue, dbo_vQbRowDefPortfolio.TotalTradableCash, dbo_vQbRowDefPortfolio.StartDate, dbo_vQbRowDefPortfolio.PortfolioCode;"

            Command2 = New OleDb.OleDbCommand(SQLstr, Mycn)
            Command2.ExecuteNonQuery()

            Mycn.Close()
            If RadioButton1.Checked Then
                Dim child As New rpt_RepBookBusinessViewer
                child.MdiParent = Home
                child.Show()

                'If CheckBox1.Checked Then
                ''SQLstr = SQLstr & " AND dbo_vQbRowDefPortfolio.AAMRepName = '" & cboIntRep.Text & "'"
                'child.CrystalReportViewer1.SelectionFormula = "{command.AAMRep} = '" & cboIntRep.Text & "'"
                'End If

                'If CheckBox2.Checked Then
                ''SQLstr = SQLstr & " AND dbo_vQbRowDefPortfolio.AAMTeamName = '" & cboTeam.Text & "'"
                'child.CrystalReportViewer1.SelectionFormula = "{command.AAMTeam} = '" & cboTeam.Text & "'"
                'End If

                'If CheckBox3.Checked Then
                ''SQLstr = SQLstr & " AND dbo_vQbRowDefPortfolio.SalesRegion = '" & cboRegion.Text & "'"
                'child.CrystalReportViewer1.SelectionFormula = "{command.AAMRegion} = '" & cboRegion.Text & "'"
                'End If

                'If CheckBox4.Checked Then
                ''SQLstr = SQLstr & " AND dbo_vQbRowDefPortfolio.SalesDepartment = '" & cboDepartment.Text & "'"
                'child.CrystalReportViewer1.SelectionFormula = "{command.AAMDepartment} = '" & cboDepartment.Text & "'"
                'End If

                child.CrystalReportViewer1.Refresh()
            Else
                Dim child As New rpt_RepFirmLevelBreakdownViewer
                child.MdiParent = Home
                child.Show()

                'If CheckBox1.Checked Then
                ''SQLstr = SQLstr & " AND dbo_vQbRowDefPortfolio.AAMRepName = '" & cboIntRep.Text & "'"
                'child.CrystalReportViewer1.SelectionFormula = "{command.AAMRep} = '" & cboIntRep.Text & "'"
                'End If

                'If CheckBox2.Checked Then
                ''SQLstr = SQLstr & " AND dbo_vQbRowDefPortfolio.AAMTeamName = '" & cboTeam.Text & "'"
                'child.CrystalReportViewer1.SelectionFormula = "{command.AAMTeam} = '" & cboTeam.Text & "'"
                'End If

                'If CheckBox3.Checked Then
                ''SQLstr = SQLstr & " AND dbo_vQbRowDefPortfolio.SalesRegion = '" & cboRegion.Text & "'"
                'child.CrystalReportViewer1.SelectionFormula = "{command.AAMRegion} = '" & cboRegion.Text & "'"
                'End If

                'If CheckBox4.Checked Then
                ''SQLstr = SQLstr & " AND dbo_vQbRowDefPortfolio.SalesDepartment = '" & cboDepartment.Text & "'"
                'child.CrystalReportViewer1.SelectionFormula = "{command.AAMDepartment} = '" & cboDepartment.Text & "'"
                'End If

                child.CrystalReportViewer1.Refresh()
            End If

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "ERROR")

        End Try
    End Sub

    Private Sub rpt_RepBookBusiness_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Call loadinternalreps()
        Call loadregions()
        Call loaddepartment()
        Call loadteams()
    End Sub

    Public Sub loadinternalreps()
        Try

            Dim conn As New OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & My.Settings.dbloc & "';Jet OLEDB:Database Password='" & My.Settings.dbpass & "';")

            Dim strSQL As String = "SELECT DISTINCT dbo_vQbRowDefPortfolio.AAMRepName FROM dbo_vQbRowDefPortfolio ORDER BY AAMRepName"

            Dim da As New OleDb.OleDbDataAdapter(strSQL, conn)
            Dim ds As New DataSet
            da.Fill(ds, "Users")

            With cboIntRep
                .DataSource = ds.Tables("Users")
                .DisplayMember = "AAMRepName"
                .ValueMember = "AAMRepName"
                .SelectedIndex = 0
            End With

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "ERROR")
        End Try
    End Sub

    Public Sub loadregions()
        Try

            Dim conn As New OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & My.Settings.dbloc & "';Jet OLEDB:Database Password='" & My.Settings.dbpass & "';")

            Dim strSQL As String = "SELECT DISTINCT dbo_vQbRowDefPortfolio.SalesRegion FROM dbo_vQbRowDefPortfolio ORDER BY SalesRegion"

            Dim da As New OleDb.OleDbDataAdapter(strSQL, conn)
            Dim ds As New DataSet
            da.Fill(ds, "Users")

            With cboRegion
                .DataSource = ds.Tables("Users")
                .DisplayMember = "SalesRegion"
                .ValueMember = "SalesRegion"
                .SelectedIndex = 0
            End With

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "ERROR")
        End Try
    End Sub

    Public Sub loadteams()
        Try

            Dim conn As New OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & My.Settings.dbloc & "';Jet OLEDB:Database Password='" & My.Settings.dbpass & "';")

            Dim strSQL As String = "SELECT DISTINCT AAMTeamName FROM dbo_vQbRowDefPortfolio ORDER BY AAMTeamName"

            Dim da As New OleDb.OleDbDataAdapter(strSQL, conn)
            Dim ds As New DataSet
            da.Fill(ds, "Users")

            With cboTeam
                .DataSource = ds.Tables("Users")
                .DisplayMember = "AAMTeamName"
                .ValueMember = "AAMTeamName"
                .SelectedIndex = 0
            End With

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "ERROR")
        End Try
    End Sub
    Public Sub loaddepartment()
        Try

            Dim conn As New OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & My.Settings.dbloc & "';Jet OLEDB:Database Password='" & My.Settings.dbpass & "';")

            Dim strSQL As String = "SELECT DISTINCT SalesDepartment FROM dbo_vQbRowDefPortfolio ORDER BY SalesDepartment"

            Dim da As New OleDb.OleDbDataAdapter(strSQL, conn)
            Dim ds As New DataSet
            da.Fill(ds, "Users")

            With cboDepartment
                .DataSource = ds.Tables("Users")
                .DisplayMember = "SalesDepartment"
                .ValueMember = "SalesDepartment"
                .SelectedIndex = 0
            End With

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "ERROR")
        End Try
    End Sub

    Private Sub CheckBox1_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckBox1.CheckedChanged
        If CheckBox1.Checked Then
            cboIntRep.Enabled = True
            CheckBox5.Checked = False
        Else
            cboIntRep.Enabled = False
        End If
    End Sub

    Private Sub CheckBox2_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckBox2.CheckedChanged
        If CheckBox2.Checked Then
            cboTeam.Enabled = True
            CheckBox5.Checked = False
        Else
            cboTeam.Enabled = False
        End If
    End Sub

    Private Sub CheckBox3_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckBox3.CheckedChanged
        If CheckBox3.Checked Then
            cboRegion.Enabled = True
            CheckBox5.Checked = False
        Else
            cboRegion.Enabled = False
        End If
    End Sub

    Private Sub CheckBox4_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckBox4.CheckedChanged
        If CheckBox4.Checked Then
            cboDepartment.Enabled = True
            CheckBox5.Checked = False
        Else
            cboDepartment.Enabled = False
        End If
    End Sub

    Private Sub CheckBox5_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckBox5.CheckedChanged
        If CheckBox5.Checked Then
            CheckBox1.Checked = False
            CheckBox2.Checked = False
            CheckBox3.Checked = False
            CheckBox4.Checked = False
        Else

        End If
    End Sub
End Class