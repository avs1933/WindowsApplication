Public Class RevenueCenter_Adjustments

    Private Sub GroupBox2_Enter(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles GroupBox2.Enter

    End Sub

    Private Sub RadioButton1_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RadioButton1.CheckedChanged
        If RadioButton1.Checked Then
            RIntFee.Enabled = True
            RSFFee.Enabled = True
            RSRFee.Enabled = True
            ckbIFDollar.Enabled = True
            ckbIFPer.Enabled = True
            ckbSFDollar.Enabled = True
            ckbSFPer.Enabled = True
            ckbSRDollar.Enabled = True
            ckbSRPer.Enabled = True
            ORDaysBill.Enabled = False
            ORIRate.Enabled = False
            ORIValue.Enabled = False
            ORMarketVal.Enabled = False
            ORSFRate.Enabled = False
            ORSFValue.Enabled = False
            ORSRRate.Enabled = False
            ORSRValue.Enabled = False
            Button3.Enabled = False
        Else
            RIntFee.Enabled = False
            RSFFee.Enabled = False
            RSRFee.Enabled = False
            ckbIFDollar.Enabled = False
            ckbIFPer.Enabled = False
            ckbSFDollar.Enabled = False
            ckbSFPer.Enabled = False
            ckbSRDollar.Enabled = False
            ckbSRPer.Enabled = False
            ORDaysBill.Enabled = True
            ORIRate.Enabled = True
            ORIValue.Enabled = True
            ORMarketVal.Enabled = True
            ORSFRate.Enabled = True
            ORSFValue.Enabled = True
            ORSRRate.Enabled = True
            ORSRValue.Enabled = True
            Button3.Enabled = True
        End If
    End Sub

    Private Sub ckbIFPer_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ckbIFPer.CheckedChanged
        If ckbIFPer.Checked = True Then
            ckbIFDollar.Checked = False
        Else
            ckbIFDollar.Checked = True
            ckbIFPer.Checked = False
        End If
    End Sub

    Private Sub ckbIFDollar_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ckbIFDollar.CheckedChanged
        If ckbIFDollar.Checked = True Then
            ckbIFPer.Checked = False
        Else
            ckbIFPer.Checked = True
            ckbIFDollar.Checked = False
        End If
    End Sub

    Private Sub ckbSFPer_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ckbSFPer.CheckedChanged
        If ckbSFPer.Checked Then
            ckbSFDollar.Checked = False
        Else
            ckbSFDollar.Checked = True
            ckbSFPer.Checked = False
        End If
    End Sub

    Private Sub ckbSFDollar_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ckbSFDollar.CheckedChanged
        If ckbSFDollar.Checked Then
            ckbSFPer.Checked = False
        Else
            ckbSFPer.Checked = True
            ckbSFDollar.Checked = False
        End If
    End Sub

    Private Sub ckbSRPer_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ckbSRPer.CheckedChanged
        If ckbSRPer.Checked Then
            ckbSRDollar.Checked = False
        Else
            ckbSRDollar.Checked = True
            ckbSRPer.Checked = False
        End If
    End Sub

    Private Sub ckbSRDollar_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ckbSRDollar.CheckedChanged
        If ckbSRDollar.Checked Then
            ckbSRPer.Checked = False
        Else
            ckbSRPer.Checked = True
            ckbSRDollar.Checked = False
        End If
    End Sub

    Public Sub InitialLoadAccountData()
        Dim Mycn As OleDb.OleDbConnection
        Dim ds As New DataSet
        Dim dv As New DataView
        Mycn = New OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & My.Settings.dbloc & "';Jet OLEDB:Database Password='" & My.Settings.dbpass & "';")

        Mycn.Open()

        Dim ad As New OleDb.OleDbDataAdapter("SELECT PortfolioCode, ReportHeading1 FROM dbo_vQbRowDefPortfolio WHERE PortfolioID = " & PortfolioID.Text, Mycn)
        ad.Fill(ds, "Production")
        dv.Table = ds.Tables("Production")

        Mycn.Close()

        Dim dt As DataTable = ds.Tables("Production")

        Dim row As DataRow = dt.Rows(0)
        If IsDBNull(row("ReportHeading1")) Then
            PortfolioName.Text = "NULL"
        Else
            PortfolioName.Text = (row("ReportHeading1"))
        End If

        PortfolioCode.Text = (row("PortfolioCode"))

        Mycn.Close()

    End Sub

    Public Sub InitialLoadOldAdjustment()
        Dim Mycn As OleDb.OleDbConnection
        Dim ds As New DataSet
        Dim dv As New DataView
        Mycn = New OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & My.Settings.dbloc & "';Jet OLEDB:Database Password='" & My.Settings.dbpass & "';")

        Mycn.Open()

        Dim ad As New OleDb.OleDbDataAdapter("SELECT * FROM mdb_BillingBaseAdj WHERE ID = " & AdjID.Text, Mycn)
        ad.Fill(ds, "Production")
        dv.Table = ds.Tables("Production")

        Mycn.Close()

        Dim dt As DataTable = ds.Tables("Production")

        Dim tmv As Double
        Dim otmv As Double
        Dim dbill As Integer
        Dim odbill As Integer
        Dim ir As Double
        Dim oir As Double
        Dim sf As Double
        Dim osf As Double
        Dim sr As Double
        Dim osr As Double
        Dim iv As Double
        Dim oiv As Double
        Dim sfv As Double
        Dim osfv As Double
        Dim srv As Double
        Dim osrv As Double

        Dim row As DataRow = dt.Rows(0)
        If IsDBNull(row("NewTMV")) Then
            tmv = 0
        Else
            tmv = (row("NewTMV"))
        End If

        If IsDBNull(row("OldTMV")) Then
            otmv = 0
        Else
            otmv = (row("OldTMV"))
        End If

        If IsDBNull(row("NewDaysBillable")) Then
            dbill = 0
        Else
            dbill = (row("NewDaysBillable"))
        End If

        If IsDBNull(row("OldDaysBillable")) Then
            odbill = 0
        Else
            odbill = (row("OldDaysBillable"))
        End If

        If IsDBNull(row("NewIRate")) Then
            ir = 0
        Else
            ir = (row("NewIRate"))
        End If

        If IsDBNull(row("OldIRate")) Then
            oir = 0
        Else
            oir = (row("OldIRate"))
        End If

        If IsDBNull(row("NewSFRate")) Then
            sf = 0
        Else
            sf = (row("NewSFRate"))
        End If

        If IsDBNull(row("OldSFRate")) Then
            osf = 0
        Else
            osf = (row("OldSFRate"))
        End If

        If IsDBNull(row("NewSRRate")) Then
            sr = 0
        Else
            sr = (row("NewSRRate"))
        End If

        If IsDBNull(row("OldSRRate")) Then
            osr = 0
        Else
            osr = (row("OldSRRate"))
        End If

        If IsDBNull(row("NewAAMValue")) Then
            iv = 0
        Else
            iv = (row("NewAAMValue"))
        End If

        If IsDBNull(row("OldAAMValue")) Then
            oiv = 0
        Else
            oiv = (row("OldAAMValue"))
        End If

        If IsDBNull(row("NewSolicitorFirmValue")) Then
            sfv = 0
        Else
            sfv = (row("NewSolicitorFirmValue"))
        End If

        If IsDBNull(row("OldSolicitorFirmValue")) Then
            osfv = 0
        Else
            osfv = (row("OldSolicitorFirmValue"))
        End If

        If IsDBNull(row("NewSolicitorRepValue")) Then
            srv = 0
        Else
            srv = (row("NewSolicitorRepValue"))
        End If

        If IsDBNull(row("OldSolicitorRepValue")) Then
            osrv = 0
        Else
            osrv = (row("OldSolicitorRepValue"))
        End If

        If (row("AdjustmentType")) = 1 Then
            RadioButton1.Checked = True
            RadioButton2.Checked = False
        Else
            RadioButton1.Checked = False
            RadioButton2.Checked = tmv
        End If

        cboReason.SelectedValue = (row("ReasonCode"))
        TextBox1.Text = (row("ReasonText"))
        RIntFee.Text = (row("RIntFee"))
        If (row("RIntFeeType")) = 1 Then
            ckbIFPer.Checked = True
        Else
            ckbIFDollar.Checked = True
        End If

        RSFFee.Text = (row("RSFFee"))
        If (row("RSFFeeType")) = 1 Then
            ckbSFPer.Checked = True
        Else
            ckbSFDollar.Checked = True
        End If

        RSRFee.Text = (row("RSRFee"))
        If (row("RSRFeeType")) = 1 Then
            ckbSRPer.Checked = True
        Else
            ckbSRDollar.Checked = True
        End If

        ckbActive.Checked = (row("Active"))

        OMarketValue.Text = Format(otmv, "#.00")
        ODays.Text = odbill
        OIRate.Text = Format(oir, ".##")
        OSFRate.Text = Format(osf, ".##")
        OSRRate.Text = Format(osr, ".##")
        OIValue.Text = Format(oiv, "#.00")
        OSFValue.Text = Format(osfv, "#.00")
        OSRValue.Text = Format(osrv, "#.00")

        ORMarketVal.Text = Format(tmv, "#.00")
        ORDaysBill.Text = dbill
        ORIRate.Text = Format(ir, ".##")
        ORSFRate.Text = Format(sf, ".##")
        ORSRRate.Text = Format(sr, ".##")
        ORIValue.Text = Format(iv, "#.00")
        ORSFValue.Text = Format(sfv, "#.00")
        ORSRValue.Text = Format(srv, "#.00")

        NMarketValue.Text = OMarketValue.Text
        NDays.Text = ODays.Text
        NIR.Text = OIRate.Text
        NSF.Text = OSFRate.Text
        NSR.Text = OSRRate.Text
        NIV.Text = OIValue.Text
        NSFV.Text = OSFValue.Text
        NSRV.Text = OSRValue.Text

        Mycn.Close()

        Call InitialLoadAccountData()

    End Sub

    Public Sub InitialLoadNewAdjustment()
        Dim Mycn As OleDb.OleDbConnection
        Dim ds As New DataSet
        Dim dv As New DataView
        Mycn = New OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & My.Settings.dbloc & "';Jet OLEDB:Database Password='" & My.Settings.dbpass & "';")

        Mycn.Open()

        Dim ad As New OleDb.OleDbDataAdapter("SELECT * FROM mdb_BillingBase WHERE ID = " & ID.Text, Mycn)
        ad.Fill(ds, "Production")
        dv.Table = ds.Tables("Production")

        Mycn.Close()

        Dim dt As DataTable = ds.Tables("Production")

        Dim tmv As Double
        Dim dbill As Integer
        Dim ir As Double
        Dim sf As Double
        Dim sr As Double
        Dim iv As Double
        Dim sfv As Double
        Dim srv As Double

        Dim row As DataRow = dt.Rows(0)
        If IsDBNull(row("BillableMarketValue")) Then
            tmv = 0
        Else
            tmv = (row("BillableMarketValue"))
        End If

        If IsDBNull(row("DaysBillable")) Then
            dbill = 0
        Else
            dbill = (row("DaysBillable"))
        End If

        If IsDBNull(row("AAMRate")) Then
            ir = 0
        Else
            ir = (row("AAMRate"))
        End If

        If IsDBNull(row("SolicitorFirm")) Then
            sf = 0
        Else
            sf = (row("SolicitorFirm"))
        End If

        If IsDBNull(row("SolicitorRep")) Then
            sr = 0
        Else
            sr = (row("SolicitorRep"))
        End If

        If IsDBNull(row("AAMValue")) Then
            iv = 0
        Else
            iv = (row("AAMValue"))
        End If

        If IsDBNull(row("SolicitorFirmValue")) Then
            sfv = 0
        Else
            sfv = (row("SolicitorFirmValue"))
        End If

        If IsDBNull(row("SolicitorRepValue")) Then
            srv = 0
        Else
            srv = (row("SolicitorRepValue"))
        End If

        PortfolioID.Text = (row("PortfolioID"))

        OMarketValue.Text = Format(tmv, "#.00")
        ODays.Text = dbill
        OIRate.Text = Format(ir, ".##")
        OSFRate.Text = Format(sf, ".##")
        OSRRate.Text = Format(sr, ".##")
        OIValue.Text = Format(iv, "#.00")
        OSFValue.Text = Format(sfv, "#.00")
        OSRValue.Text = Format(srv, "#.00")

        ORMarketVal.Text = OMarketValue.Text
        ORDaysBill.Text = ODays.Text
        ORIRate.Text = OIRate.Text
        ORSFRate.Text = OSFRate.Text
        ORSRRate.Text = OSRRate.Text
        ORIValue.Text = OIValue.Text
        ORSFValue.Text = OSFValue.Text
        ORSRValue.Text = OSRValue.Text

        NMarketValue.Text = OMarketValue.Text
        NDays.Text = ODays.Text
        NIR.Text = OIRate.Text
        NSF.Text = OSFRate.Text
        NSR.Text = OSRRate.Text
        NIV.Text = OIValue.Text
        NSFV.Text = OSFValue.Text
        NSRV.Text = OSRValue.Text

        Mycn.Close()

        Call InitialLoadAccountData()

    End Sub

    Private Sub RevenueCenter_Adjustments_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Call LoadReasonCodes()
    End Sub

    Public Sub LoadReasonCodes()
        Try

            Dim conn As New OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & My.Settings.dbloc & "';Jet OLEDB:Database Password='" & My.Settings.dbpass & "';")

            Dim strSQL As String = "SELECT ID, ReasonText FROM mdb_BillingBaseAdjReasons" & _
            " GROUP BY ID, ReasonText" & _
            " ORDER BY ReasonText"

            Dim da As New OleDb.OleDbDataAdapter(strSQL, conn)
            Dim ds As New DataSet
            da.Fill(ds, "Users")

            With cboReason
                .DataSource = ds.Tables("Users")
                .DisplayMember = "ReasonText"
                .ValueMember = "ID"
                .SelectedIndex = 0
            End With

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "ERROR")
        End Try
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        If RadioButton1.Checked Then

            If IsDBNull(RIntFee) Or RIntFee.Text = "" Then
            Else
                If ckbIFPer.Checked Then
                    Dim nivalue As Double
                    Dim oival As Double = OIValue.Text
                    Dim rifee As Double = RIntFee.Text
                    nivalue = (oival - (oival * (rifee / 100)))
                    NIV.Text = Format(nivalue, "#.00")
                Else
                    Dim nivalue As Double
                    Dim oival As Double = OIValue.Text
                    Dim rifee As Double = RIntFee.Text
                    nivalue = (oival - rifee)
                    NIV.Text = Format(nivalue, "#.00")
                End If
            End If

            If IsDBNull(RSFFee) Or RSFFee.Text = "" Then
            Else
                If ckbSFPer.Checked Then
                    Dim nivalue As Double
                    Dim oival As Double = OSFValue.Text
                    Dim rifee As Double = RSFFee.Text
                    nivalue = (oival - (oival * (rifee / 100)))
                    NSFV.Text = Format(nivalue, "#.00")
                Else
                    Dim nivalue As Double
                    Dim oival As Double = OSFValue.Text
                    Dim rifee As Double = RSFFee.Text
                    nivalue = (oival - rifee)
                    NSFV.Text = Format(nivalue, "#.00")
                End If
            End If

            If IsDBNull(RSRFee) Or RSRFee.Text = "" Then
            Else
                If ckbSRPer.Checked Then
                    Dim nivalue As Double
                    Dim oival As Double = OSRValue.Text
                    Dim rifee As Double = RSRFee.Text
                    nivalue = (oival - (oival * (rifee / 100)))
                    NSRV.Text = Format(nivalue, "#.00")
                Else
                    Dim nivalue As Double
                    Dim oival As Double = OSRValue.Text
                    Dim rifee As Double = RSRFee.Text
                    nivalue = (oival - rifee)
                    NSRV.Text = Format(nivalue, "#.00")
                End If
            End If
        Else
            NMarketValue.Text = ORMarketVal.Text
            NDays.Text = ORDaysBill.Text
            NIR.Text = ORIRate.Text
            NSF.Text = ORSFRate.Text
            NSR.Text = ORSRRate.Text
            NIV.Text = ORIValue.Text
            NSFV.Text = ORSFValue.Text
            NSRV.Text = ORSRValue.Text

        End If

        Button12.Enabled = True

    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        Dim nifee As Double
        Dim nsffee As Double
        Dim nsrfee As Double
        Dim nmv As Double
        Dim db As Integer
        Dim nirate As Double
        Dim nsfrate As Double
        Dim nsrrate As Double

        nmv = ORMarketVal.Text
        db = ORDaysBill.Text
        nirate = ORIRate.Text
        If ORSFRate.Text = "" Then
            nsfrate = 0
        Else
            nsfrate = ORSFRate.Text
        End If

        If ORSRRate.Text = "" Then
            nsrrate = 0
        Else
            nsrrate = ORSRRate.Text
        End If


        nifee = ((nmv * db * nirate) / 36000)
        nsffee = ((nmv * db * nsfrate) / 36000)
        nsrfee = ((nmv * db * nsrrate) / 36000)

        NMarketValue.Text = Format(nmv, "#.00")
        NDays.Text = db
        NIR.Text = nirate
        NSF.Text = nsfrate
        NSR.Text = nsrrate
        NIV.Text = Format(nifee, "#.00")
        NSFV.Text = Format(nsffee, "#.00")
        NSRV.Text = Format(nsrfee, "#.00")
        ORIValue.Text = Format(nifee, "#.00")
        ORSFValue.Text = Format(nsffee, "#.00")
        ORSRValue.Text = Format(nsrfee, "#.00")

        Button12.Enabled = True

    End Sub

    Private Sub Button12_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button12.Click
        If AdjID.Text = "NEW" Then
            Call SaveNew()
            Call PullID()
        Else
            Call SaveOld()
        End If

    End Sub

    Public Sub SaveNew()
        '1. Save Values in Adj
        '2. Update BillingBase

        Dim Mycn As OleDb.OleDbConnection
        Dim Command As OleDb.OleDbCommand
        Dim SQLstr As String

        Try

            Mycn = New OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & My.Settings.dbloc & "';Jet OLEDB:Database Password='" & My.Settings.dbpass & "';")
            Mycn.Open()

            Dim adjtype As Integer
            Dim rifeetype As Integer
            Dim rsffeetype As Integer
            Dim rsrfeetype As Integer

            If RadioButton1.Checked Then
                adjtype = 1
            Else
                adjtype = 2
            End If

            If ckbIFPer.Checked Then
                rifeetype = 1
            Else
                rifeetype = 2
            End If

            If ckbSFPer.Checked Then
                rsffeetype = 1
            Else
                rsffeetype = 2
            End If

            If ckbSRPer.Checked Then
                rsrfeetype = 1
            Else
                rsrfeetype = 2
            End If

            Dim srintfee As Double
            If RIntFee.Text = "" Then
                srintfee = 0
            Else
                srintfee = RIntFee.Text
            End If

            Dim srsffee As Double
            If RSFFee.Text = "" Then
                srsffee = 0
            Else
                srsffee = RSFFee.Text
            End If

            Dim srsrfee As Double
            If RSRFee.Text = "" Then
                srsrfee = 0
            Else
                srsrfee = RSRFee.Text
            End If

            Dim nmv As Double
            Dim ndb As Integer
            Dim nir1 As Double
            Dim nsfr1 As Double
            Dim nsrr1 As Double
            Dim niv1 As Double
            Dim nsrv1 As Double
            Dim nsfv1 As Double

            If NMarketValue.Text = "" Then
                nmv = 0
            Else
                nmv = NMarketValue.Text
            End If

            If NDays.Text = "" Then
                ndb = 0
            Else
                ndb = NDays.Text
            End If

            If NIR.Text = "" Then
                nir1 = 0
            Else
                nir1 = NIR.Text
            End If

            If NSF.Text = "" Then
                nsfr1 = 0
            Else
                nsfr1 = NSF.Text
            End If

            If NSR.Text = "" Then
                nsrr1 = 0
            Else
                nsrr1 = NSR.Text
            End If

            If NIV.Text = "" Then
                niv1 = 0
            Else
                niv1 = NIV.Text
            End If

            If NSFV.Text = "" Then
                nsfv1 = 0
            Else
                nsfv1 = NSFV.Text
            End If

            If NSRV.Text = "" Then
                nsrv1 = 0
            Else
                nsrv1 = NSRV.Text
            End If

            Dim origmv As Double
            Dim origdb As Integer
            Dim origir As Double
            Dim origsfr As Double
            Dim origsrr As Double
            Dim origiv As Double
            Dim origsfv As Double
            Dim origsrv As Double

            If OMarketValue.Text = "" Then
                origmv = 0
            Else
                origmv = OMarketValue.Text
            End If

            If ODays.Text = "" Then
                origdb = 0
            Else
                origdb = ODays.Text
            End If

            If OIRate.Text = "" Then
                origir = 0
            Else
                origir = OIRate.Text
            End If

            If OSFRate.Text = "" Then
                origsfr = 0
            Else
                origsfr = OSFRate.Text
            End If

            If OSRRate.Text = "" Then
                origsrr = 0
            Else
                origsrr = OSRRate.Text
            End If

            If OIValue.Text = "" Then
                origiv = 0
            Else
                origiv = OIValue.Text
            End If

            If OSFValue.Text = "" Then
                origsfv = 0
            Else
                origsfv = OSFValue.Text
            End If

            If OSRValue.Text = "" Then
                origsrv = 0
            Else
                origsrv = OSRValue.Text
            End If

            SQLstr = "INSERT INTO mdb_BillingBaseAdj (BillingBaseID, OldAAMValue, OldSolicitorFirmValue, OldSolicitorRepValue, NewAAMValue, NewSolicitorFirmValue, NewSolicitorRepValue, OldTMV, NewTMV, OldDaysBillable, NewDaysBillable, OldIRate, NewIRate, OldSFRate, NewSFRate, OldSRRate, NewSRRate, ReasonCode, ReasonText, AdjustmentType, RIntFee, RIntFeeType, RSFFee, RSFFeeType, RSRFee, RSRFeeType, DateStamp, ChangedBy, Active)" & _
            "VALUES(" & ID.Text & "," & origiv & "," & origsfv & "," & origsrv & "," & niv1 & "," & nsfv1 & "," & nsrv1 & "," & origmv & "," & nmv & "," & origdb & "," & ndb & "," & origir & "," & nir1 & "," & origsfr & "," & nsfr1 & "," & origsrr & "," & nsrr1 & "," & cboReason.SelectedValue & ",'" & TextBox1.Text & "'," & adjtype & "," & srintfee & "," & rifeetype & "," & srsffee & "," & rsffeetype & "," & srsrfee & "," & rsrfeetype & ",#" & Format(Now()) & "#," & My.Settings.userid & ",-1)"
            Command = New OleDb.OleDbCommand(SQLstr, Mycn)
            Command.ExecuteNonQuery()

            Mycn.Close()

            Mycn.Open()


            SQLstr = "UPDATE mdb_BillingBase SET DaysBillable = " & NDays.Text & ", BillableMarketValue = " & nmv & ", AAMRate = " & nir1 & ", SolicitorFirm = " & nsfr1 & ", SolicitorRep = " & nsrr1 & ", AAMValue = " & niv1 & ", SolicitorFirmValue = " & nsfv1 & ", SolicitorRepValue = " & nsrv1 & ", Adjusted = True WHERE ID = " & ID.Text
            Command = New OleDb.OleDbCommand(SQLstr, Mycn)
            Command.ExecuteNonQuery()

            Mycn.Close()

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "ERROR")
        End Try

    End Sub

    Public Sub SaveOld()
        Dim adjtype As Integer
        Dim rifeetype As Integer
        Dim rsffeetype As Integer
        Dim rsrfeetype As Integer

        If RadioButton1.Checked Then
            adjtype = 1
        Else
            adjtype = 2
        End If

        If ckbIFPer.Checked Then
            rifeetype = 1
        Else
            rifeetype = 2
        End If

        If ckbSFPer.Checked Then
            rsffeetype = 1
        Else
            rsffeetype = 2
        End If

        If ckbSRPer.Checked Then
            rsrfeetype = 1
        Else
            rsrfeetype = 2
        End If

        Dim srintfee As Double
        If RIntFee.Text = "" Then
            srintfee = 0
        Else
            srintfee = RIntFee.Text
        End If

        Dim srsffee As Double
        If RSFFee.Text = "" Then
            srsffee = 0
        Else
            srsffee = RSFFee.Text
        End If

        Dim srsrfee As Double
        If RSRFee.Text = "" Then
            srsrfee = 0
        Else
            srsrfee = RSRFee.Text
        End If

        Dim nmv As Double
        Dim ndb As Integer
        Dim nir1 As Double
        Dim nsfr1 As Double
        Dim nsrr1 As Double
        Dim niv1 As Double
        Dim nsrv1 As Double
        Dim nsfv1 As Double

        If NMarketValue.Text = "" Then
            nmv = 0
        Else
            nmv = NMarketValue.Text
        End If

        If NDays.Text = "" Then
            ndb = 0
        Else
            ndb = NDays.Text
        End If

        If NIR.Text = "" Then
            nir1 = 0
        Else
            nir1 = NIR.Text
        End If

        If NSF.Text = "" Then
            nsfr1 = 0
        Else
            nsfr1 = NSF.Text
        End If

        If NSR.Text = "" Then
            nsrr1 = 0
        Else
            nsrr1 = NSR.Text
        End If

        If NIV.Text = "" Then
            niv1 = 0
        Else
            niv1 = NIV.Text
        End If

        If NSFV.Text = "" Then
            nsfv1 = 0
        Else
            nsfv1 = NSFV.Text
        End If

        If NSRV.Text = "" Then
            nsrv1 = 0
        Else
            nsrv1 = NSRV.Text
        End If

        Dim origmv As Double
        Dim origdb As Integer
        Dim origir As Double
        Dim origsfr As Double
        Dim origsrr As Double
        Dim origiv As Double
        Dim origsfv As Double
        Dim origsrv As Double

        If OMarketValue.Text = "" Then
            origmv = 0
        Else
            origmv = OMarketValue.Text
        End If

        If ODays.Text = "" Then
            origdb = 0
        Else
            origdb = ODays.Text
        End If

        If OIRate.Text = "" Then
            origir = 0
        Else
            origir = OIRate.Text
        End If

        If OSFRate.Text = "" Then
            origsfr = 0
        Else
            origsfr = OSFRate.Text
        End If

        If OSRRate.Text = "" Then
            origsrr = 0
        Else
            origsrr = OSRRate.Text
        End If

        If OIValue.Text = "" Then
            origiv = 0
        Else
            origiv = OIValue.Text
        End If

        If OSFValue.Text = "" Then
            origsfv = 0
        Else
            origsfv = OSFValue.Text
        End If

        If OSRValue.Text = "" Then
            origsrv = 0
        Else
            origsrv = OSRValue.Text
        End If

        Dim Mycn As OleDb.OleDbConnection
        Dim Command As OleDb.OleDbCommand
        Dim SQLstr As String

        Try

            Mycn = New OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & My.Settings.dbloc & "';Jet OLEDB:Database Password='" & My.Settings.dbpass & "';")

            Mycn.Open()

            SQLstr = "UPDATE mdb_BillingBaseAdj SET OldAAMValue = " & origiv & ", OldSolicitorFirmValue = " & origsfv & ", OldSolicitorRepValue = " & origsrv & ", NewAAMValue = " & niv1 & ", NewSolicitorFirmValue = " & nsfv1 & ", NewSolicitorRepValue = " & nsrv1 & ", OldTMV = " & origmv & ", NewTMV = " & nmv & ", OldDaysBillable = " & origdb & ", NewDaysBillable = " & ndb & ", OldIRate = " & origir & ", NewIRate = " & nir1 & ", OldSFRate = " & origsfr & ", NewSFRate = " & nsfr1 & ", OldSRRate = " & origsrr & ", NewSRRate = " & nsrr1 & ", ReasonCode = " & cboReason.SelectedValue & ", ReasonText = '" & TextBox1.Text & "', AdjustmentType = " & adjtype & ", RIntFee = " & srintfee & ", RIntFeeType = " & rifeetype & ", RSFFee = " & srsffee & ", RSFFeeType = " & rsffeetype & ", RSRFee = " & srsrfee & ", RSRFeeType = " & rsrfeetype & ", DateStamp = #" & Format(Now()) & "#, ChangedBy = " & My.Settings.userid & ", Active = -1 WHERE ID = " & AdjID.Text
            Command = New OleDb.OleDbCommand(SQLstr, Mycn)
            Command.ExecuteNonQuery()

            Mycn.Close()


            Mycn.Open()

            SQLstr = "UPDATE mdb_BillingBase SET DaysBillable = " & NDays.Text & ", BillableMarketValue = " & nmv & ", AAMRate = " & nir1 & ", SolicitorFirm = " & nsfr1 & ", SolicitorRep = " & nsrr1 & ", AAMValue = " & niv1 & ", SolicitorFirmValue = " & nsfv1 & ", SolicitorRepValue = " & nsrv1 & ", Adjusted = True WHERE ID = " & ID.Text
            Command = New OleDb.OleDbCommand(SQLstr, Mycn)
            Command.ExecuteNonQuery()

            Mycn.Close()

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "ERROR")
        End Try


    End Sub

    Public Sub PullID()
        Dim Mycn As OleDb.OleDbConnection
        Try

        
            Mycn = New OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & My.Settings.dbloc & "';Jet OLEDB:Database Password='" & My.Settings.dbpass & "';")
            Mycn.Open()

            Dim ds As New DataSet
            Dim dv As New DataView

            Dim ad As New OleDb.OleDbDataAdapter("SELECT Top 1 ID FROM mdb_BillingBaseAdj WHERE BillingBaseID = " & ID.Text & " ORDER BY ID Desc", Mycn)
            ad.Fill(ds, "Production")
            dv.Table = ds.Tables("Production")

            Mycn.Close()

            Dim dt As DataTable = ds.Tables("Production")
            Dim row As DataRow = dt.Rows(0)

            AdjID.Text = (row("ID"))

            Mycn.Close()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "ERROR")
        End Try
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        If AdjID.Text = "NEW" Then
            MsgBox("Record has not been saved.  You cannot delete an unsaved record.", MsgBoxStyle.Critical, "Cannot Process")
        Else
            Dim ir As MsgBoxResult
            ir = MsgBox("Are you sure you want to remove this adjustment?  This cannot be undone!", MsgBoxStyle.YesNo + MsgBoxStyle.Critical, "Confirm selection")
            If ir = MsgBoxResult.Yes Then

                Dim Mycn As OleDb.OleDbConnection
                Dim Command As OleDb.OleDbCommand
                Dim SQLstr As String

                Dim origmv As Double
                Dim origdb As Integer
                Dim origir As Double
                Dim origsfr As Double
                Dim origsrr As Double
                Dim origiv As Double
                Dim origsfv As Double
                Dim origsrv As Double

                If OMarketValue.Text = "" Then
                    origmv = 0
                Else
                    origmv = OMarketValue.Text
                End If

                If ODays.Text = "" Then
                    origdb = 0
                Else
                    origdb = ODays.Text
                End If

                If OIRate.Text = "" Then
                    origir = 0
                Else
                    origir = OIRate.Text
                End If

                If OSFRate.Text = "" Then
                    origsfr = 0
                Else
                    origsfr = OSFRate.Text
                End If

                If OSRRate.Text = "" Then
                    origsrr = 0
                Else
                    origsrr = OSRRate.Text
                End If

                If OIValue.Text = "" Then
                    origiv = 0
                Else
                    origiv = OIValue.Text
                End If

                If OSFValue.Text = "" Then
                    origsfv = 0
                Else
                    origsfv = OSFValue.Text
                End If

                If OSRValue.Text = "" Then
                    origsrv = 0
                Else
                    origsrv = OSRValue.Text
                End If

                Try

                    Mycn = New OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & My.Settings.dbloc & "';Jet OLEDB:Database Password='" & My.Settings.dbpass & "';")

                    Mycn.Open()

                    SQLstr = "UPDATE mdb_BillingBaseAdj SET Active = 0 WHERE ID = " & AdjID.Text
                    Command = New OleDb.OleDbCommand(SQLstr, Mycn)
                    Command.ExecuteNonQuery()

                    Mycn.Close()


                    Mycn.Open()

                    SQLstr = "UPDATE mdb_BillingBase SET DaysBillable = " & origdb & ", BillableMarketValue = " & origmv & ", AAMRate = " & origir & ", SolicitorFirm = " & origsfr & ", SolicitorRep = " & origsrr & ", AAMValue = " & origiv & ", SolicitorFirmValue = " & origsfv & ", SolicitorRepValue = " & origsrv & ", Adjusted = false WHERE ID = " & ID.Text
                    Command = New OleDb.OleDbCommand(SQLstr, Mycn)
                    Command.ExecuteNonQuery()

                    ckbActive.Checked = False

                    Mycn.Close()

                Catch ex As Exception
                    MsgBox(ex.Message, MsgBoxStyle.Critical, "ERROR")
                End Try

            Else
            End If
        End If
    End Sub

    Private Sub Timer1_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer1.Tick
        If ckbActive.Checked = False Then
            Label34.Visible = True
            Button12.Enabled = False
            Button1.Enabled = False
        Else
            Label34.Visible = False
            'If AdjID.Text = "NEW" Then
            Button12.Enabled = True
            'Else
            'End If
            Button1.Enabled = True
        End If
    End Sub
End Class