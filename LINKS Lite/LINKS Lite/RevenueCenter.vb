Imports System.IO
Imports System.Data.OleDb
'Imports Microsoft.Office.Interop.Excel


Public Class RevenueCenter
    Dim ProssBill As System.Threading.Thread
    Dim dtStudentGrade As DataTable
    Dim dtExcelData As DataTable

    Private Sub RadioButton2_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RadioButton2.CheckedChanged
        If RadioButton2.Checked Then
            ComboBox1.Enabled = False
            ComboBox2.Enabled = True
            Call LoadPortfolioGroups()
        Else
            ComboBox1.Enabled = True
            ComboBox2.Enabled = False
            Call LoadPortfolioCodes()
        End If
    End Sub

    Private Sub RadioButton1_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RadioButton1.CheckedChanged
        If RadioButton2.Checked Then
            ComboBox1.Enabled = False
            ComboBox2.Enabled = True
            Call LoadPortfolioGroups()
        Else
            ComboBox1.Enabled = True
            ComboBox2.Enabled = False
            Call LoadPortfolioCodes()
        End If
    End Sub

    Public Sub LoadPortfolioCodes()
        Try

            Dim conn As New OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & My.Settings.dbloc & "';Jet OLEDB:Database Password='" & My.Settings.dbpass & "';")

            Dim strSQL As String = "SELECT PortfolioID, PortfolioCode FROM dbo_vQbRowDefPortfolio"
            Dim da As New OleDb.OleDbDataAdapter(strSQL, conn)
            Dim ds As New DataSet
            da.Fill(ds, "Users")

            With ComboBox1
                .DataSource = ds.Tables("Users")
                .DisplayMember = "PortfolioCode"
                .ValueMember = "PortfolioID"
                .SelectedIndex = 0
            End With

            ComboBox1.Enabled = True

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "ERROR")
        End Try
    End Sub

    Public Sub LoadPortfolioGroups()
        Try

            Dim conn As New OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & My.Settings.dbloc & "';Jet OLEDB:Database Password='" & My.Settings.dbpass & "';")

            Dim strSQL As String = "SELECT PortfolioGroupID, PortfolioGroupCode FROM dbo_vQbRowDefPortfolioGroup"
            Dim da As New OleDb.OleDbDataAdapter(strSQL, conn)
            Dim ds As New DataSet
            da.Fill(ds, "Users")

            With ComboBox2
                .DataSource = ds.Tables("Users")
                .DisplayMember = "PortfolioGroupCode"
                .ValueMember = "PortfolioGroupID"
                .SelectedIndex = 0
            End With

            'Label14.Visible = False
            ComboBox2.Enabled = True
            'Call LoadDBData()

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "ERROR")
        End Try
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        '================================================================================================
        '*************************************** SECTION HEADER *****************************************
        'This section kicks off when user presses the "Process" button on the "Process Billing" tab
        'Created: 05/24/2014
        'Creator: Josh Colyer
        'Last Modified: 05/24/2014
        '================================================================================================

        '================================================================================================
        '*** BLOCK START ***
        'Check to ensure all values need for the billing process have been checked and are valid.
        '================================================================================================

        ProgressBar1.Value = 0

        Dim pstart As Integer
        Dim pend As Integer
        Dim datestring1 As Date
        datestring1 = DateTimePicker1.Text

        Dim datestring2 As Date
        datestring2 = DateTimePicker2.Text

        pstart = DatePart(DateInterval.Month, datestring1)
        pend = DatePart(DateInterval.Month, datestring2)

        Dim dteend As Date
        dteend = DateTimePicker2.Text

        If pstart = 12 Then
            pend = pend + 12
        End If

        If RadioButton1.Checked = False And RadioButton2.Checked = False Then
            Label4.Text = "Please select a Portfolio or Group"
            Label4.Visible = True
            GroupBox1.BackColor = Color.Red
            GroupBox1.ForeColor = Color.White
        Else
            GroupBox1.BackColor = Color.White
            GroupBox1.ForeColor = Color.Black
            If dteend > Format(Now()) Then
                Label4.Text = "Ending date cannot be in the future"
                GroupBox2.BackColor = Color.Red
                GroupBox2.ForeColor = Color.White
            Else
                'Dim datestring1 As Date
                datestring1 = DateTimePicker1.Text

                'Dim datestring2 As Date
                datestring2 = DateTimePicker2.Text
                If datestring1 >= datestring2 Then
                    Label4.Text = "Starting date cannot be greater than or equal to ending date."
                    GroupBox2.BackColor = Color.Red
                    GroupBox2.ForeColor = Color.White
                Else
                    GroupBox2.BackColor = Color.White
                    GroupBox2.ForeColor = Color.Black
                    If RadioButton3.Checked = False And RadioButton4.Checked = False Then
                        Label4.Text = "Please select the type of billing to run."
                        GroupBox3.BackColor = Color.Red
                        GroupBox3.ForeColor = Color.White
                    Else
                        GroupBox3.BackColor = Color.White
                        GroupBox3.ForeColor = Color.Black
                        Label4.Text = "All selections valid."
                        If (pend - pstart) > 3 Then
                            Dim msg As MsgBoxResult
                            msg = MsgBox("WARNING!" & vbNewLine & vbNewLine & "Period selected is greather than one quarter.  Are you sure you want to proceed?", MsgBoxStyle.Critical + MsgBoxStyle.YesNo, "WARNING!")
                            If msg = MsgBoxResult.Yes Then
                                Control.CheckForIllegalCrossThreadCalls = False
                                ProssBill = New System.Threading.Thread(AddressOf ProcessBilling)
                                ProssBill.Start()
                            Else
                                GoTo Line2
                            End If
                        Else
                            Control.CheckForIllegalCrossThreadCalls = False
                            ProssBill = New System.Threading.Thread(AddressOf ProcessBilling)
                            ProssBill.Start()
                        End If
                    End If
                End If
            End If
        End If
Line2:

        '================================================================================================
        '*** BLOCK END ***
        '================================================================================================


    End Sub

    Public Sub ProcessBilling()
        '================================================================================================
        '*************************************** SECTION HEADER *****************************************
        'This section handes the entire process of billing and is initiated on the ProssBill Thread
        'Created: 05/24/2014
        'Creator: Josh Colyer
        'Last Modified: 05/24/2014
        '================================================================================================
        Dim Mycn As OleDb.OleDbConnection
        Dim Command As OleDb.OleDbCommand
        Dim SQLstr As String

        RichTextBox1.Text = "Billing started." & vbNewLine & "Pulling a Period ID from database."

        'Try

        '================================================================================================
        '*** BLOCK START ***
        'Create a Period ID from database and return value to user.
        'Tables used in this function: mdb_BillingPeriods
        '================================================================================================

        Mycn = New OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & My.Settings.dbloc & "';Jet OLEDB:Database Password='" & My.Settings.dbpass & "';")
        Mycn.Open()

        'Passes Dates to find the start and end of qtr dates
        Dim QtrStart As Date = FirstDayOfQuarter(DateTimePicker2.Value)
        Dim QtrEnd As Date = LastDayOfQuarter(DateTimePicker2.Value)

        Dim rtype As Integer

        If RadioButton3.Checked = True Then
            rtype = 1
        Else
            rtype = 2
        End If

        ProgressBar1.Maximum = 51

        Label3.Visible = True
        Label3.Text = "Working..."
        Timer1.Enabled = True

        Label14.Visible = True
        Label14.Text = "Billing Started"

        'Insert Values in Database
        SQLstr = "INSERT INTO mdb_BillingPeriods(PeriodStart, PeriodEnd, RequestBy, RequestDateStamp, QtrStart, QtrEnd, TypeID)" & _
        "VALUES('" & DateTimePicker1.Text & "','" & DateTimePicker2.Text & "','" & Environ("USERNAME") & "', Format(Now()), #" & QtrStart & "#, #" & QtrEnd & "#," & rtype & ")"
        Command = New OleDb.OleDbCommand(SQLstr, Mycn)
        Command.ExecuteNonQuery()

        ProgressBar1.Value = ProgressBar1.Value + 1

        Mycn.Close()

        RichTextBox1.Text = RichTextBox1.Text & vbNewLine & "Period ID Created."
        Label14.Text = "Period ID Created"

        '*** Move onto pulling the ID ***

        Dim ds As New DataSet
        Dim dv As New DataView

        Mycn.Open()

        Dim ad As New OleDb.OleDbDataAdapter("SELECT Top 1 ID FROM mdb_BillingPeriods ORDER BY ID Desc", Mycn)
        ad.Fill(ds, "Production")
        dv.Table = ds.Tables("Production")

        Mycn.Close()

        Dim dt As DataTable = ds.Tables("Production")
        Dim pid As Integer

        Dim row As DataRow = dt.Rows(0)
        pid = (row("ID"))

        ProgressBar1.Value = ProgressBar1.Value + 1

        Mycn.Close()

        RichTextBox1.Text = RichTextBox1.Text & vbNewLine & "Period ID = " & pid

        '================================================================================================
        '*** BLOCK END ***
        '================================================================================================
        Label14.Text = "Calculating Days."
        '================================================================================================
        '*** BLOCK START ***
        'Find accounts to be billed based on selections, then insert into request table
        'Tables used in this function: mdb_BillingRequests, dbo_vQbRowDefPortfolio, dbo_vQbRowDefPortfolioGroupAssociation
        '================================================================================================
        If RadioButton1.Checked Then
            RichTextBox1.Text = RichTextBox1.Text & vbNewLine & "Billing ran for a single account.  Loading account into Request Table."
            Mycn.Open()

            SQLstr = "INSERT INTO mdb_BillingRequests(PortfolioID, PeriodID)" & _
            "VALUES(" & ComboBox1.SelectedValue & "," & pid & ")"

            Command = New OleDb.OleDbCommand(SQLstr, Mycn)
            Command.ExecuteNonQuery()
            RichTextBox1.Text = RichTextBox1.Text & vbNewLine & "Account Loaded."
            ProgressBar1.Value = ProgressBar1.Value + 1
            Mycn.Close()
        Else
            If RadioButton2.Checked Then
                RichTextBox1.Text = RichTextBox1.Text & vbNewLine & "Billing ran for a group.  Loading Accounts from group."
                Mycn.Open()

                SQLstr = "INSERT INTO mdb_BillingRequests(PortfolioID, PeriodID)" & _
                "SELECT MEMBERID, " & pid & " FROM dbo_vQbRowDefPortfolioGroupAssociation WHERE PortfolioGroupID = " & ComboBox2.SelectedValue

                Command = New OleDb.OleDbCommand(SQLstr, Mycn)
                Command.ExecuteNonQuery()
                RichTextBox1.Text = RichTextBox1.Text & vbNewLine & "Accounts Loaded."
                ProgressBar1.Value = ProgressBar1.Value + 1
                Mycn.Close()

            End If
        End If

        '================================================================================================
        '*** BLOCK END ***
        '================================================================================================

        '================================================================================================
        '*** BLOCK START ***
        'Calculates days billable for accounts opened for entire billing period
        'QID: D1
        'Tables used in this function: mdb_BillingDaysBillable, dbo_vQbRowDefPortfolio
        '================================================================================================

        RichTextBox1.Text = RichTextBox1.Text & vbNewLine & "Calculating days billable for accounts opened during entire period."
        Label14.Text = "Calculating days billable for accounts opened during entire period"
        Dim pstart As Integer
        Dim pend As Integer
        Dim datestring1 As Date
        datestring1 = DateTimePicker1.Text

        Dim datestring2 As Date
        datestring2 = DateTimePicker2.Text

        pstart = DatePart(DateInterval.Month, datestring1)
        pend = DatePart(DateInterval.Month, datestring2)

        If pstart = 12 Then
            pend = pend + 12
        End If

        Dim billdays As Integer
        billdays = ((pend - pstart) * 30)

        Mycn.Open()

        SQLstr = "INSERT INTO mdb_BillingDaysBillable (PortfolioID, PeriodID, DaysBillable, PortfolioStart, PortfolioEnd, QID)" & _
        "SELECT dbo_vQbRowDefPortfolio.PortfolioID, mdb_BillingRequests.PeriodID," & billdays & ", dbo_vQbRowDefPortfolio.StartDate, dbo_vQbRowDefPortfolio.CloseDate, 'D1'" & _
        " FROM (mdb_BillingRequests INNER JOIN dbo_vQbRowDefPortfolio ON mdb_BillingRequests.PortfolioID = dbo_vQbRowDefPortfolio.PortfolioID) INNER JOIN mdb_BillingPeriods ON mdb_BillingRequests.PeriodID = mdb_BillingPeriods.ID" & _
        " WHERE ((((CDate(dbo_vQbRowDefPortfolio.StartDate)) <= #" & datestring1 & "#) AND dbo_vQbRowDefPortfolio.StartDate Is Not Null) AND dbo_vQbRowDefPortfolio.CloseDate Is Null) AND mdb_BillingRequests.PeriodID = " & pid

        Command = New OleDb.OleDbCommand(SQLstr, Mycn)
        Command.ExecuteNonQuery()
        RichTextBox1.Text = RichTextBox1.Text & vbNewLine & "Days loaded."
        ProgressBar1.Value = ProgressBar1.Value + 1
        Mycn.Close()

        '================================================================================================
        '*** BLOCK END ***
        '================================================================================================

        '================================================================================================
        '*** BLOCK START ***
        'Calculates days billable for accounts opened for entire billing period - close date after end of period
        'QID: D1a
        'Tables used in this function: mdb_BillingDaysBillable, dbo_vQbRowDefPortfolio
        '================================================================================================

        RichTextBox1.Text = RichTextBox1.Text & vbNewLine & "Calculating days billable for accounts opened during entire period - Closed after period end."
        Label14.Text = "Calculating days billable for accounts opened during entire period - Closed after period end"

        datestring1 = DateTimePicker1.Text

        'Dim datestring2 As Date
        datestring2 = DateTimePicker2.Text

        pstart = DatePart(DateInterval.Month, datestring1)
        pend = DatePart(DateInterval.Month, datestring2)

        If pstart = 12 Then
            pend = pend + 12
        End If

        'Dim billdays As Integer
        billdays = ((pend - pstart) * 30)

        Mycn.Open()

        SQLstr = "INSERT INTO mdb_BillingDaysBillable (PortfolioID, PeriodID, DaysBillable, PortfolioStart, PortfolioEnd, QID)" & _
        "SELECT dbo_vQbRowDefPortfolio.PortfolioID, mdb_BillingRequests.PeriodID," & billdays & ", dbo_vQbRowDefPortfolio.StartDate, dbo_vQbRowDefPortfolio.CloseDate, 'D1a'" & _
        " FROM (mdb_BillingRequests INNER JOIN dbo_vQbRowDefPortfolio ON mdb_BillingRequests.PortfolioID = dbo_vQbRowDefPortfolio.PortfolioID) INNER JOIN mdb_BillingPeriods ON mdb_BillingRequests.PeriodID = mdb_BillingPeriods.ID" & _
        " WHERE ((((CDate(dbo_vQbRowDefPortfolio.StartDate)) <= #" & datestring1 & "#) AND dbo_vQbRowDefPortfolio.StartDate Is Not Null) AND ((CDate(dbo_vQbRowDefPortfolio.CloseDate)) >= #" & datestring2 & "#) AND dbo_vQbRowDefPortfolio.CloseDate Is Not Null) AND mdb_BillingRequests.PeriodID = " & pid

        Command = New OleDb.OleDbCommand(SQLstr, Mycn)
        Command.ExecuteNonQuery()
        RichTextBox1.Text = RichTextBox1.Text & vbNewLine & "Days loaded."
        ProgressBar1.Value = ProgressBar1.Value + 1
        Mycn.Close()

        '================================================================================================
        '*** BLOCK END ***
        '================================================================================================

        '================================================================================================
        '*** BLOCK START ***
        'Calculates days billable for accounts opened during billing period
        'Style: Arrears
        'QID: D2
        'Tables used in this function: mdb_BillingDaysBillable, dbo_vQbRowDefPortfolio
        '================================================================================================

        RichTextBox1.Text = RichTextBox1.Text & vbNewLine & "Calculating days billable for accounts opened during period - Arrears."
        Label14.Text = "Calculating days billable for accounts opened during period - Arrears"
        pstart = DatePart(DateInterval.Month, datestring1)
        pend = DatePart(DateInterval.Month, datestring2)

        'If pstart = 12 Then
        'pend = pend - 12
        'End If

        Mycn.Open()

        '"SELECT dbo_vQbRowDefPortfolio.PortfolioID, mdb_BillingRequests.PeriodID,(((-(Format (CDate(dbo_vQbRowDefPortfolio.StartDate), 'dd')) +(31))-(((Format(CDate(dbo_vQbRowDefPortfolio.StartDate), 'mm'))-" & pend & ")))*30), dbo_vQbRowDefPortfolio.StartDate, dbo_vQbRowDefPortfolio.CloseDate" & _
        SQLstr = "INSERT INTO mdb_BillingDaysBillable (PortfolioID, PeriodID, DaysBillable, PortfolioStart, PortfolioEnd, QID)" & _
        "SELECT dbo_vQbRowDefPortfolio.PortfolioID, mdb_BillingRequests.PeriodID,((-(Format(CDate(dbo_vQbRowDefPortfolio.StartDate), 'dd'))+(31)) + (( " & pend & "-(Format(CDate(dbo_vQbRowDefPortfolio.StartDate), 'mm')))*30)), dbo_vQbRowDefPortfolio.StartDate, dbo_vQbRowDefPortfolio.CloseDate, 'D2'" & _
        " FROM (mdb_BillingRequests INNER JOIN dbo_vQbRowDefPortfolio ON mdb_BillingRequests.PortfolioID = dbo_vQbRowDefPortfolio.PortfolioID) INNER JOIN mdb_BillingPeriods ON mdb_BillingRequests.PeriodID = mdb_BillingPeriods.ID" & _
        " WHERE ((((CDate(dbo_vQbRowDefPortfolio.StartDate)) > #" & datestring1 & "#) AND ((CDate(dbo_vQbRowDefPortfolio.StartDate)) <= #" & datestring2 & "#) AND dbo_vQbRowDefPortfolio.StartDate Is Not Null)) AND (dbo_vQbRowDefPortfolio.CloseDate Is Null) AND PeriodID = " & pid & " AND (((dbo_vQbRowDefPortfolio.BillInArrears)<>0) OR (dbo_vQbRowDefPortfolio.BillInArrears) is Null)"

        Command = New OleDb.OleDbCommand(SQLstr, Mycn)
        Command.ExecuteNonQuery()
        RichTextBox1.Text = RichTextBox1.Text & vbNewLine & "Days loaded."
        ProgressBar1.Value = ProgressBar1.Value + 1
        Mycn.Close()

        '================================================================================================
        '*** BLOCK END ***
        '================================================================================================

        '================================================================================================
        '*** BLOCK START ***
        'Calculates days billable for accounts opened during billing period - Closed after period
        'Style: Arrears
        'QID: D2a
        'Tables used in this function: mdb_BillingDaysBillable, dbo_vQbRowDefPortfolio
        '================================================================================================

        RichTextBox1.Text = RichTextBox1.Text & vbNewLine & "Calculating days billable for accounts opened during period - Arrears - Closed after period end."
        Label14.Text = "Calculating days billable for accounts opened during period - Arrears - Closed after period end"
        pstart = DatePart(DateInterval.Month, datestring1)
        pend = DatePart(DateInterval.Month, datestring2)

        'If pstart = 12 Then
        'pend = pend - 12
        'End If

        Mycn.Open()

        '"SELECT dbo_vQbRowDefPortfolio.PortfolioID, mdb_BillingRequests.PeriodID,(((-(Format (CDate(dbo_vQbRowDefPortfolio.StartDate), 'dd')) +(31))-(((Format(CDate(dbo_vQbRowDefPortfolio.StartDate), 'mm'))-" & pend & ")))*30), dbo_vQbRowDefPortfolio.StartDate, dbo_vQbRowDefPortfolio.CloseDate" & _
        SQLstr = "INSERT INTO mdb_BillingDaysBillable (PortfolioID, PeriodID, DaysBillable, PortfolioStart, PortfolioEnd, QID)" & _
        "SELECT dbo_vQbRowDefPortfolio.PortfolioID, mdb_BillingRequests.PeriodID,((-(Format(CDate(dbo_vQbRowDefPortfolio.StartDate), 'dd'))+(31)) + (( " & pend & "-(Format(CDate(dbo_vQbRowDefPortfolio.StartDate), 'mm')))*30)), dbo_vQbRowDefPortfolio.StartDate, dbo_vQbRowDefPortfolio.CloseDate, 'D2a'" & _
        " FROM (mdb_BillingRequests INNER JOIN dbo_vQbRowDefPortfolio ON mdb_BillingRequests.PortfolioID = dbo_vQbRowDefPortfolio.PortfolioID) INNER JOIN mdb_BillingPeriods ON mdb_BillingRequests.PeriodID = mdb_BillingPeriods.ID" & _
        " WHERE ((((CDate(dbo_vQbRowDefPortfolio.StartDate)) > #" & datestring1 & "#) AND ((CDate(dbo_vQbRowDefPortfolio.StartDate)) <= #" & datestring2 & "#) AND dbo_vQbRowDefPortfolio.StartDate Is Not Null) AND ((CDate(dbo_vQbRowDefPortfolio.CloseDate)) >= #" & datestring2 & "#) AND dbo_vQbRowDefPortfolio.CloseDate Is Not Null) AND PeriodID = " & pid & " AND (((dbo_vQbRowDefPortfolio.BillInArrears)<>0) OR (dbo_vQbRowDefPortfolio.BillInArrears) is Null)"

        Command = New OleDb.OleDbCommand(SQLstr, Mycn)
        Command.ExecuteNonQuery()
        RichTextBox1.Text = RichTextBox1.Text & vbNewLine & "Days loaded."
        ProgressBar1.Value = ProgressBar1.Value + 1
        Mycn.Close()

        '================================================================================================
        '*** BLOCK END ***
        '================================================================================================

        '================================================================================================
        '*** BLOCK START ***
        'Calculates days billable for accounts opened during billing period
        'Style: Advanced
        'QID: D3
        'Tables used in this function: mdb_BillingDaysBillable, dbo_vQbRowDefPortfolio
        '================================================================================================

        RichTextBox1.Text = RichTextBox1.Text & vbNewLine & "Calculating days billable for accounts opened during period - Advanced."
        Label14.Text = "Calculating days billable for accounts opened during period - Advanced"
        pstart = DatePart(DateInterval.Month, datestring1)
        pend = DatePart(DateInterval.Month, datestring2)

        Mycn.Open()

        SQLstr = "INSERT INTO mdb_BillingDaysBillable (PortfolioID, PeriodID, DaysBillable, PortfolioStart, PortfolioEnd, QID)" & _
        "SELECT dbo_vQbRowDefPortfolio.PortfolioID, mdb_BillingRequests.PeriodID,((-(Format(CDate(dbo_vQbRowDefPortfolio.StartDate), 'dd'))+(31)) + (( " & pend & "-(Format(CDate(dbo_vQbRowDefPortfolio.StartDate), 'mm')))*30)+90), dbo_vQbRowDefPortfolio.StartDate, dbo_vQbRowDefPortfolio.CloseDate, 'D3'" & _
        " FROM (mdb_BillingRequests INNER JOIN dbo_vQbRowDefPortfolio ON mdb_BillingRequests.PortfolioID = dbo_vQbRowDefPortfolio.PortfolioID) INNER JOIN mdb_BillingPeriods ON mdb_BillingRequests.PeriodID = mdb_BillingPeriods.ID" & _
        " WHERE ((((CDate(dbo_vQbRowDefPortfolio.StartDate)) > #" & datestring1 & "#) AND ((CDate(dbo_vQbRowDefPortfolio.StartDate)) <= #" & datestring2 & "#) AND dbo_vQbRowDefPortfolio.StartDate Is Not Null)) AND (dbo_vQbRowDefPortfolio.CloseDate Is Null) AND PeriodID = " & pid & " AND dbo_vQbRowDefPortfolio.BillInArrears = 0"

        Command = New OleDb.OleDbCommand(SQLstr, Mycn)
        Command.ExecuteNonQuery()
        RichTextBox1.Text = RichTextBox1.Text & vbNewLine & "Days loaded."
        ProgressBar1.Value = ProgressBar1.Value + 1
        Mycn.Close()

        '================================================================================================
        '*** BLOCK END ***
        '================================================================================================

        '================================================================================================
        '*** BLOCK START ***
        'Calculates days billable for accounts opened during billing period - Closed after period end
        'Style: Advanced
        'QID: D3a
        'Tables used in this function: mdb_BillingDaysBillable, dbo_vQbRowDefPortfolio
        '================================================================================================

        RichTextBox1.Text = RichTextBox1.Text & vbNewLine & "Calculating days billable for accounts opened during period - Advanced - Closed after period end."
        Label14.Text = "Calculating days billable for accounts opened during period - Advanced - Closed after period end"
        pstart = DatePart(DateInterval.Month, datestring1)
        pend = DatePart(DateInterval.Month, datestring2)

        Mycn.Open()

        SQLstr = "INSERT INTO mdb_BillingDaysBillable (PortfolioID, PeriodID, DaysBillable, PortfolioStart, PortfolioEnd, QID)" & _
        "SELECT dbo_vQbRowDefPortfolio.PortfolioID, mdb_BillingRequests.PeriodID,((-(Format(CDate(dbo_vQbRowDefPortfolio.StartDate), 'dd'))+(31)) + (( " & pend & "-(Format(CDate(dbo_vQbRowDefPortfolio.StartDate), 'mm')))*30)+90), dbo_vQbRowDefPortfolio.StartDate, dbo_vQbRowDefPortfolio.CloseDate, 'D3a'" & _
        " FROM (mdb_BillingRequests INNER JOIN dbo_vQbRowDefPortfolio ON mdb_BillingRequests.PortfolioID = dbo_vQbRowDefPortfolio.PortfolioID) INNER JOIN mdb_BillingPeriods ON mdb_BillingRequests.PeriodID = mdb_BillingPeriods.ID" & _
        " WHERE ((((CDate(dbo_vQbRowDefPortfolio.StartDate)) > #" & datestring1 & "#) AND ((CDate(dbo_vQbRowDefPortfolio.StartDate)) <= #" & datestring2 & "#) AND dbo_vQbRowDefPortfolio.StartDate Is Not Null) AND ((CDate(dbo_vQbRowDefPortfolio.CloseDate)) >= #" & datestring2 & "#) AND dbo_vQbRowDefPortfolio.CloseDate Is Not Null) AND PeriodID = " & pid & " AND dbo_vQbRowDefPortfolio.BillInArrears = 0"

        Command = New OleDb.OleDbCommand(SQLstr, Mycn)
        Command.ExecuteNonQuery()
        RichTextBox1.Text = RichTextBox1.Text & vbNewLine & "Days loaded."
        ProgressBar1.Value = ProgressBar1.Value + 1
        Mycn.Close()

        '================================================================================================
        '*** BLOCK END ***
        '================================================================================================

        '================================================================================================
        '*** BLOCK START ***
        'Calculates days billable for accounts closed during billing period
        'Style: Arrears
        'QID: D4a/D4b
        'Tables used in this function: mdb_BillingDaysBillable, dbo_vQbRowDefPortfolio
        '================================================================================================

        RichTextBox1.Text = RichTextBox1.Text & vbNewLine & "Calculating days billable for accounts closed during period - Arrears."
        Label14.Text = "Calculating days billable for accounts closed during period - Arrears"
        pstart = DatePart(DateInterval.Month, datestring1)
        pend = DatePart(DateInterval.Month, datestring2)

        Mycn.Open()

        If pstart = 12 Then
            SQLstr = "INSERT INTO mdb_BillingDaysBillable (PortfolioID, PeriodID, DaysBillable, PortfolioStart, PortfolioEnd, QID)" & _
            "SELECT dbo_vQbRowDefPortfolio.PortfolioID, mdb_BillingRequests.PeriodID,(((Format(CDate(dbo_vQbRowDefPortfolio.CloseDate), 'mm')+12)-(" & pstart & "))*30)-(- (Format(CDate(dbo_vQbRowDefPortfolio.CloseDate), 'dd'))+(31)), dbo_vQbRowDefPortfolio.StartDate, dbo_vQbRowDefPortfolio.CloseDate, 'D4a'" & _
            " FROM (mdb_BillingRequests INNER JOIN dbo_vQbRowDefPortfolio ON mdb_BillingRequests.PortfolioID = dbo_vQbRowDefPortfolio.PortfolioID) INNER JOIN mdb_BillingPeriods ON mdb_BillingRequests.PeriodID = mdb_BillingPeriods.ID" & _
            " WHERE ((((CDate(dbo_vQbRowDefPortfolio.CloseDate)) < #" & datestring2 & "#) AND ((CDate(dbo_vQbRowDefPortfolio.CloseDate)) > #" & datestring1 & "#) AND dbo_vQbRowDefPortfolio.CloseDate Is Not Null) AND ((CDate(dbo_vQbRowDefPortfolio.StartDate)) < #" & datestring1 & "#)) AND PeriodID = " & pid & " AND (((dbo_vQbRowDefPortfolio.BillInArrears)<>0) OR (dbo_vQbRowDefPortfolio.BillInArrears) is Null)"
        Else
            SQLstr = "INSERT INTO mdb_BillingDaysBillable (PortfolioID, PeriodID, DaysBillable, PortfolioStart, PortfolioEnd, QID)" & _
            "SELECT dbo_vQbRowDefPortfolio.PortfolioID, mdb_BillingRequests.PeriodID,(((Format(CDate(dbo_vQbRowDefPortfolio.CloseDate), 'mm'))-(" & pstart & "))*30)-(- (Format(CDate(dbo_vQbRowDefPortfolio.CloseDate), 'dd'))+(31)), dbo_vQbRowDefPortfolio.StartDate, dbo_vQbRowDefPortfolio.CloseDate, 'D4b'" & _
            " FROM (mdb_BillingRequests INNER JOIN dbo_vQbRowDefPortfolio ON mdb_BillingRequests.PortfolioID = dbo_vQbRowDefPortfolio.PortfolioID) INNER JOIN mdb_BillingPeriods ON mdb_BillingRequests.PeriodID = mdb_BillingPeriods.ID" & _
            " WHERE ((((CDate(dbo_vQbRowDefPortfolio.CloseDate)) < #" & datestring2 & "#) AND ((CDate(dbo_vQbRowDefPortfolio.CloseDate)) > #" & datestring1 & "#) AND dbo_vQbRowDefPortfolio.CloseDate Is Not Null) AND ((CDate(dbo_vQbRowDefPortfolio.StartDate)) < #" & datestring1 & "#)) AND PeriodID = " & pid & " AND (((dbo_vQbRowDefPortfolio.BillInArrears)<>0) OR (dbo_vQbRowDefPortfolio.BillInArrears) is Null)"
        End If

        Command = New OleDb.OleDbCommand(SQLstr, Mycn)
        Command.ExecuteNonQuery()
        RichTextBox1.Text = RichTextBox1.Text & vbNewLine & "Days loaded."
        ProgressBar1.Value = ProgressBar1.Value + 1
        Mycn.Close()

        '================================================================================================
        '*** BLOCK END ***
        '================================================================================================

        '================================================================================================
        '*** BLOCK START ***
        'Calculates days billable for accounts closed during billing period
        'Style: Advanced
        'QID: D5
        'Tables used in this function: mdb_BillingDaysBillable, dbo_vQbRowDefPortfolio
        '================================================================================================

        RichTextBox1.Text = RichTextBox1.Text & vbNewLine & "Calculating days billable for accounts closed during period - Advanced."
        Label14.Text = "Calculating days billable for accounts closed during period - Advanced"
        pstart = DatePart(DateInterval.Month, datestring1)
        pend = DatePart(DateInterval.Month, datestring2)

        Mycn.Open()

        SQLstr = "INSERT INTO mdb_BillingDaysBillable (PortfolioID, PeriodID, DaysBillable, PortfolioStart, PortfolioEnd, QID)" & _
            "SELECT dbo_vQbRowDefPortfolio.PortfolioID, mdb_BillingRequests.PeriodID,((-(" & pend & "-(Format(CDate(dbo_vQbRowDefPortfolio.CloseDate), 'mm')))*30)- (-(Format(CDate(dbo_vQbRowDefPortfolio.CloseDate), 'dd'))+(31))), dbo_vQbRowDefPortfolio.StartDate, dbo_vQbRowDefPortfolio.CloseDate, 'D5'" & _
            " FROM (mdb_BillingRequests INNER JOIN dbo_vQbRowDefPortfolio ON mdb_BillingRequests.PortfolioID = dbo_vQbRowDefPortfolio.PortfolioID) INNER JOIN mdb_BillingPeriods ON mdb_BillingRequests.PeriodID = mdb_BillingPeriods.ID" & _
            " WHERE ((((CDate(dbo_vQbRowDefPortfolio.CloseDate)) < #" & datestring2 & "#) AND ((CDate(dbo_vQbRowDefPortfolio.CloseDate)) > #" & datestring1 & "#) AND dbo_vQbRowDefPortfolio.CloseDate Is Not Null) AND ((CDate(dbo_vQbRowDefPortfolio.StartDate)) < #" & datestring1 & "#)) AND PeriodID = " & pid & " AND (((dbo_vQbRowDefPortfolio.BillInArrears)=0))"

        Command = New OleDb.OleDbCommand(SQLstr, Mycn)
        Command.ExecuteNonQuery()
        RichTextBox1.Text = RichTextBox1.Text & vbNewLine & "Days loaded."
        ProgressBar1.Value = ProgressBar1.Value + 1
        Mycn.Close()

        '================================================================================================
        '*** BLOCK END ***
        '================================================================================================

        '================================================================================================
        '*** BLOCK START ***
        'Calculates days billable for accounts open and closed during billing period
        'Style: N/A
        'QID: D6
        'Tables used in this function: mdb_BillingDaysBillable, dbo_vQbRowDefPortfolio
        '================================================================================================

        RichTextBox1.Text = RichTextBox1.Text & vbNewLine & "Calculating days billable for accounts opened and closed during period."
        Label14.Text = "Calculating days billable for accounts opened and closed during period"
        pstart = DatePart(DateInterval.Month, datestring1)
        pend = DatePart(DateInterval.Month, datestring2)

        Mycn.Open()

        SQLstr = "INSERT INTO mdb_BillingDaysBillable (PortfolioID, PeriodID, DaysBillable, PortfolioStart, PortfolioEnd, QID)" & _
            "SELECT dbo_vQbRowDefPortfolio.PortfolioID, mdb_BillingRequests.PeriodID,(((-(Format(CDate(dbo_vQbRowDefPortfolio.StartDate), 'dd'))+31)- (-(Format(CDate(dbo_vQbRowDefPortfolio.CloseDate), 'dd'))+31))+(((Format(CDate(dbo_vQbRowDefPortfolio.CloseDate), 'mm'))- (Format(CDate(dbo_vQbRowDefPortfolio.StartDate), 'mm')))*30)), dbo_vQbRowDefPortfolio.StartDate, dbo_vQbRowDefPortfolio.CloseDate, 'D6'" & _
            " FROM (mdb_BillingRequests INNER JOIN dbo_vQbRowDefPortfolio ON mdb_BillingRequests.PortfolioID = dbo_vQbRowDefPortfolio.PortfolioID) INNER JOIN mdb_BillingPeriods ON mdb_BillingRequests.PeriodID = mdb_BillingPeriods.ID" & _
            " WHERE ((((CDate(dbo_vQbRowDefPortfolio.CloseDate)) < #" & datestring2 & "#) AND ((CDate(dbo_vQbRowDefPortfolio.CloseDate)) > #" & datestring1 & "#) AND dbo_vQbRowDefPortfolio.CloseDate Is Not Null) AND ((CDate(dbo_vQbRowDefPortfolio.StartDate)) > #" & datestring1 & "#)) AND PeriodID = " & pid

        Command = New OleDb.OleDbCommand(SQLstr, Mycn)
        Command.ExecuteNonQuery()
        RichTextBox1.Text = RichTextBox1.Text & vbNewLine & "Days loaded."
        ProgressBar1.Value = ProgressBar1.Value + 1
        Mycn.Close()

        '================================================================================================
        '*** BLOCK END ***
        '================================================================================================

        '************************************************************************************************
        '************************************************************************************************
        '************************************* SECTION BREAK ********************************************
        '************************************************************************************************
        '************************************************************************************************

        '================================================================================================
        '*** BLOCK START ***
        'Determines the date needed for accounts opened all quarter
        'Style: EOQ
        'QID: V1
        'Tables used in this function: mdb_BillingDateQueue, dbo_vQbRowDefPortfolio
        '================================================================================================

        RichTextBox1.Text = RichTextBox1.Text & vbNewLine & "Loading valuation dates for accounts opened all quarter - EOQ."
        Label14.Text = "Loading valuation dates for accounts opened all quarter - EOQ"
        pstart = DatePart(DateInterval.Month, datestring1)
        pend = DatePart(DateInterval.Month, datestring2)

        Dim edte As String
        edte = Format(datestring2, "MMddyyyy")

        Mycn.Open()

        SQLstr = "INSERT INTO mdb_BillingDateQueue (PortfolioID, PortfolioCode, PeriodID, DateNeeded, DateNeededText, QID)" & _
            "SELECT dbo_vQbRowDefPortfolio.PortfolioID, dbo_vQbRowDefPortfolio.PortfolioCode, " & pid & ", #" & datestring2 & "#, '" & edte & "', 'V1'" & _
            " FROM (mdb_BillingRequests INNER JOIN dbo_vQbRowDefPortfolio ON mdb_BillingRequests.PortfolioID = dbo_vQbRowDefPortfolio.PortfolioID) INNER JOIN mdb_BillingPeriods ON mdb_BillingRequests.PeriodID = mdb_BillingPeriods.ID" & _
            " WHERE ((((CDate(dbo_vQbRowDefPortfolio.StartDate)) <= #" & datestring1 & "#) AND dbo_vQbRowDefPortfolio.StartDate Is Not Null) AND ((dbo_vQbRowDefPortfolio.CloseDate Is Null) OR ((CDate(dbo_vQbRowDefPortfolio.CloseDate)) >= #" & datestring2 & "#))) AND mdb_BillingRequests.PeriodID = " & pid & " AND ((dbo_vQbRowDefPortfolio.BillingValuationPeriod) = 'End_of_Qtr')"

        Command = New OleDb.OleDbCommand(SQLstr, Mycn)
        Command.ExecuteNonQuery()
        RichTextBox1.Text = RichTextBox1.Text & vbNewLine & "Dates Loaded."
        ProgressBar1.Value = ProgressBar1.Value + 1
        Mycn.Close()

        '================================================================================================
        '*** BLOCK END ***
        '================================================================================================

        '================================================================================================
        '*** BLOCK START ***
        'Determines the date needed for accounts opened all quarter
        'Style: SOQ
        'QID: V2
        'Tables used in this function: mdb_BillingDateQueue, dbo_vQbRowDefPortfolio
        '================================================================================================

        RichTextBox1.Text = RichTextBox1.Text & vbNewLine & "Loading valuation dates for accounts opened all quarter - SOQ."
        Label14.Text = "Loading valuation dates for accounts opened all quarter - SOQ"
        pstart = DatePart(DateInterval.Month, datestring1)
        pend = DatePart(DateInterval.Month, datestring2)

        Dim sdte As String
        sdte = Format(datestring1, "MMddyyyy")

        Mycn.Open()

        SQLstr = "INSERT INTO mdb_BillingDateQueue (PortfolioID, PortfolioCode, PeriodID, DateNeeded, DateNeededText, QID)" & _
            "SELECT dbo_vQbRowDefPortfolio.PortfolioID, dbo_vQbRowDefPortfolio.PortfolioCode, " & pid & ", #" & datestring1 & "#, '" & sdte & "', 'V2'" & _
            " FROM (mdb_BillingRequests INNER JOIN dbo_vQbRowDefPortfolio ON mdb_BillingRequests.PortfolioID = dbo_vQbRowDefPortfolio.PortfolioID) INNER JOIN mdb_BillingPeriods ON mdb_BillingRequests.PeriodID = mdb_BillingPeriods.ID" & _
            " WHERE ((((CDate(dbo_vQbRowDefPortfolio.StartDate)) <= #" & datestring1 & "#) AND dbo_vQbRowDefPortfolio.StartDate Is Not Null) AND ((dbo_vQbRowDefPortfolio.CloseDate Is Null) OR ((CDate(dbo_vQbRowDefPortfolio.CloseDate)) >= #" & datestring2 & "#))) AND mdb_BillingRequests.PeriodID = " & pid & " AND ((dbo_vQbRowDefPortfolio.BillingValuationPeriod) = 'Start_of_Qtr')"

        Command = New OleDb.OleDbCommand(SQLstr, Mycn)
        Command.ExecuteNonQuery()
        RichTextBox1.Text = RichTextBox1.Text & vbNewLine & "Dates Loaded."
        ProgressBar1.Value = ProgressBar1.Value + 1
        Mycn.Close()

        '================================================================================================
        '*** BLOCK END ***
        '================================================================================================

        '================================================================================================
        '*** BLOCK START ***
        'Determines the date needed for accounts closed intra quarter
        'Style: EOQ
        'QID: V3
        'Tables used in this function: mdb_BillingDateQueue, dbo_vQbRowDefPortfolio
        '================================================================================================

        RichTextBox1.Text = RichTextBox1.Text & vbNewLine & "Loading valuation dates for accounts closed intra quarter - EOQ."
        Label14.Text = "Loading valuation dates for accounts closed intra quarter - EOQ"
        pstart = DatePart(DateInterval.Month, datestring1)
        pend = DatePart(DateInterval.Month, datestring2)

        Mycn.Open()

        SQLstr = "INSERT INTO mdb_BillingDateQueue (PortfolioID, PortfolioCode, PeriodID, DateNeeded, DateNeededText, QID)" & _
            "SELECT dbo_vQbRowDefPortfolio.PortfolioID, dbo_vQbRowDefPortfolio.PortfolioCode, " & pid & ", DateAdd('d','-1',CDate(dbo_vQbRowDefPortfolio.CloseDate)) As SDate1, Format((SDate1), 'MMddyyyy'), 'V3'" & _
            " FROM (mdb_BillingRequests INNER JOIN dbo_vQbRowDefPortfolio ON mdb_BillingRequests.PortfolioID = dbo_vQbRowDefPortfolio.PortfolioID) INNER JOIN mdb_BillingPeriods ON mdb_BillingRequests.PeriodID = mdb_BillingPeriods.ID" & _
            " WHERE ((((CDate(dbo_vQbRowDefPortfolio.CloseDate)) < #" & datestring2 & "#) AND ((CDate(dbo_vQbRowDefPortfolio.CloseDate)) > #" & datestring1 & "#) AND dbo_vQbRowDefPortfolio.CloseDate Is Not Null) AND ((CDate(dbo_vQbRowDefPortfolio.StartDate)) < #" & datestring1 & "#)) AND PeriodID = " & pid & " AND ((dbo_vQbRowDefPortfolio.BillingValuationPeriod) = 'End_of_Qtr')"

        Command = New OleDb.OleDbCommand(SQLstr, Mycn)
        Command.ExecuteNonQuery()
        RichTextBox1.Text = RichTextBox1.Text & vbNewLine & "Dates Loaded."
        ProgressBar1.Value = ProgressBar1.Value + 1
        Mycn.Close()

        '================================================================================================
        '*** BLOCK END ***
        '================================================================================================

        '================================================================================================
        '*** BLOCK START ***
        'Determines the date needed for accounts closed intra quarter
        'Style: SOQ
        'QID: V4
        'Tables used in this function: mdb_BillingDateQueue, dbo_vQbRowDefPortfolio
        '================================================================================================

        RichTextBox1.Text = RichTextBox1.Text & vbNewLine & "Loading valuation dates for accounts closed intra quarter - SOQ."
        Label14.Text = "Loading valuation dates for accounts closed intra quarter - SOQ"
        pstart = DatePart(DateInterval.Month, datestring1)
        pend = DatePart(DateInterval.Month, datestring2)

        Mycn.Open()

        SQLstr = "INSERT INTO mdb_BillingDateQueue (PortfolioID, PortfolioCode, PeriodID, DateNeeded, DateNeededText, QID)" & _
            "SELECT dbo_vQbRowDefPortfolio.PortfolioID, dbo_vQbRowDefPortfolio.PortfolioCode, " & pid & ", #" & datestring1 & "#, '" & sdte & "', 'V4'" & _
            " FROM (mdb_BillingRequests INNER JOIN dbo_vQbRowDefPortfolio ON mdb_BillingRequests.PortfolioID = dbo_vQbRowDefPortfolio.PortfolioID) INNER JOIN mdb_BillingPeriods ON mdb_BillingRequests.PeriodID = mdb_BillingPeriods.ID" & _
            " WHERE ((((CDate(dbo_vQbRowDefPortfolio.CloseDate)) < #" & datestring2 & "#) AND ((CDate(dbo_vQbRowDefPortfolio.CloseDate)) > #" & datestring1 & "#) AND dbo_vQbRowDefPortfolio.CloseDate Is Not Null) AND ((CDate(dbo_vQbRowDefPortfolio.StartDate)) < #" & datestring1 & "#)) AND PeriodID = " & pid & " AND ((dbo_vQbRowDefPortfolio.BillingValuationPeriod) = 'Start_of_Qtr')"

        Command = New OleDb.OleDbCommand(SQLstr, Mycn)
        Command.ExecuteNonQuery()
        RichTextBox1.Text = RichTextBox1.Text & vbNewLine & "Dates Loaded."
        ProgressBar1.Value = ProgressBar1.Value + 1
        Mycn.Close()

        '================================================================================================
        '*** BLOCK END ***
        '================================================================================================

        '================================================================================================
        '*** BLOCK START ***
        'Determines the date needed for accounts opened intra quarter
        'Style: EOQ
        'QID: V5
        'Tables used in this function: mdb_BillingDateQueue, dbo_vQbRowDefPortfolio
        '================================================================================================

        RichTextBox1.Text = RichTextBox1.Text & vbNewLine & "Loading valuation dates for accounts opened intra quarter - EOQ."
        Label14.Text = "Loading valuation dates for accounts opened intra quarter - EOQ"
        pstart = DatePart(DateInterval.Month, datestring1)
        pend = DatePart(DateInterval.Month, datestring2)

        Mycn.Open()

        SQLstr = "INSERT INTO mdb_BillingDateQueue (PortfolioID, PortfolioCode, PeriodID, DateNeeded, DateNeededText, QID)" & _
            "SELECT dbo_vQbRowDefPortfolio.PortfolioID, dbo_vQbRowDefPortfolio.PortfolioCode, " & pid & ", #" & datestring2 & "#, '" & edte & "', 'V5'" & _
            " FROM (mdb_BillingRequests INNER JOIN dbo_vQbRowDefPortfolio ON mdb_BillingRequests.PortfolioID = dbo_vQbRowDefPortfolio.PortfolioID) INNER JOIN mdb_BillingPeriods ON mdb_BillingRequests.PeriodID = mdb_BillingPeriods.ID" & _
            " WHERE ((((CDate(dbo_vQbRowDefPortfolio.StartDate)) > #" & datestring1 & "#) AND ((CDate(dbo_vQbRowDefPortfolio.StartDate)) <= #" & datestring2 & "#) AND dbo_vQbRowDefPortfolio.StartDate Is Not Null) AND ((dbo_vQbRowDefPortfolio.CloseDate Is Null) OR ((CDate(dbo_vQbRowDefPortfolio.CloseDate)) >= #" & datestring2 & "#))) AND PeriodID = " & pid & " AND ((dbo_vQbRowDefPortfolio.BillingValuationPeriod) = 'End_of_Qtr')"

        Command = New OleDb.OleDbCommand(SQLstr, Mycn)
        Command.ExecuteNonQuery()
        RichTextBox1.Text = RichTextBox1.Text & vbNewLine & "Dates Loaded."
        ProgressBar1.Value = ProgressBar1.Value + 1
        Mycn.Close()

        '================================================================================================
        '*** BLOCK END ***
        '================================================================================================

        '================================================================================================
        '*** BLOCK START ***
        'Determines the date needed for accounts opened intra quarter
        'Style: SOQ
        'QID: V6
        'Tables used in this function: mdb_BillingDateQueue, dbo_vQbRowDefPortfolio
        '================================================================================================

        RichTextBox1.Text = RichTextBox1.Text & vbNewLine & "Loading valuation dates for accounts opened intra quarter - SOQ."
        Label14.Text = "Loading valuation dates for accounts opened intra quarter - SOQ"
        pstart = DatePart(DateInterval.Month, datestring1)
        pend = DatePart(DateInterval.Month, datestring2)

        Mycn.Open()

        SQLstr = "INSERT INTO mdb_BillingDateQueue (PortfolioID, PortfolioCode, PeriodID, DateNeeded, DateNeededText, QID)" & _
            "SELECT dbo_vQbRowDefPortfolio.PortfolioID, dbo_vQbRowDefPortfolio.PortfolioCode, " & pid & ", (CDate(dbo_vQbRowDefPortfolio.StartDate)) As SDate1, Format((SDate1), 'MMddyyyy'), 'V6'" & _
            " FROM (mdb_BillingRequests INNER JOIN dbo_vQbRowDefPortfolio ON mdb_BillingRequests.PortfolioID = dbo_vQbRowDefPortfolio.PortfolioID) INNER JOIN mdb_BillingPeriods ON mdb_BillingRequests.PeriodID = mdb_BillingPeriods.ID" & _
            " WHERE ((((CDate(dbo_vQbRowDefPortfolio.StartDate)) > #" & datestring1 & "#) AND ((CDate(dbo_vQbRowDefPortfolio.StartDate)) <= #" & datestring2 & "#) AND dbo_vQbRowDefPortfolio.StartDate Is Not Null) AND ((dbo_vQbRowDefPortfolio.CloseDate Is Null) OR ((CDate(dbo_vQbRowDefPortfolio.CloseDate)) >= #" & datestring2 & "#))) AND PeriodID = " & pid & " AND ((dbo_vQbRowDefPortfolio.BillingValuationPeriod) = 'Start_of_Qtr')"

        Command = New OleDb.OleDbCommand(SQLstr, Mycn)
        Command.ExecuteNonQuery()
        RichTextBox1.Text = RichTextBox1.Text & vbNewLine & "Dates Loaded."
        ProgressBar1.Value = ProgressBar1.Value + 1
        Mycn.Close()

        '================================================================================================
        '*** BLOCK END ***
        '================================================================================================

        '================================================================================================
        '*** BLOCK START ***
        'Determines the date needed for accounts opened and closed intra quarter
        'Style: EOQ
        'QID: V7
        'Tables used in this function: mdb_BillingDateQueue, dbo_vQbRowDefPortfolio
        '================================================================================================

        RichTextBox1.Text = RichTextBox1.Text & vbNewLine & "Loading valuation dates for accounts opened and closed intra quarter - EOQ."
        Label14.Text = "Loading valuation dates for accounts opened and closed intra quarter - EOQ"
        pstart = DatePart(DateInterval.Month, datestring1)
        pend = DatePart(DateInterval.Month, datestring2)

        Mycn.Open()

        SQLstr = "INSERT INTO mdb_BillingDateQueue (PortfolioID, PortfolioCode, PeriodID, DateNeeded, DateNeededText, QID)" & _
            "SELECT dbo_vQbRowDefPortfolio.PortfolioID, dbo_vQbRowDefPortfolio.PortfolioCode, " & pid & ", DateAdd('d','-1',CDate(dbo_vQbRowDefPortfolio.CloseDate)) As SDate1, Format((SDate1), 'MMddyyyy'), 'V3'" & _
            " FROM (mdb_BillingRequests INNER JOIN dbo_vQbRowDefPortfolio ON mdb_BillingRequests.PortfolioID = dbo_vQbRowDefPortfolio.PortfolioID) INNER JOIN mdb_BillingPeriods ON mdb_BillingRequests.PeriodID = mdb_BillingPeriods.ID" & _
            " WHERE ((((CDate(dbo_vQbRowDefPortfolio.CloseDate)) < #" & datestring2 & "#) AND ((CDate(dbo_vQbRowDefPortfolio.CloseDate)) > #" & datestring1 & "#) AND dbo_vQbRowDefPortfolio.CloseDate Is Not Null) AND ((CDate(dbo_vQbRowDefPortfolio.StartDate)) > #" & datestring1 & "#)) AND PeriodID = " & pid & " AND ((dbo_vQbRowDefPortfolio.BillingValuationPeriod) = 'End_of_Qtr')"

        Command = New OleDb.OleDbCommand(SQLstr, Mycn)
        Command.ExecuteNonQuery()
        RichTextBox1.Text = RichTextBox1.Text & vbNewLine & "Dates Loaded."
        ProgressBar1.Value = ProgressBar1.Value + 1
        Mycn.Close()

        '================================================================================================
        '*** BLOCK END ***
        '================================================================================================

        '================================================================================================
        '*** BLOCK START ***
        'Determines the date needed for accounts opened and closed intra quarter
        'Style: SOQ
        'QID: V8
        'Tables used in this function: mdb_BillingDateQueue, dbo_vQbRowDefPortfolio
        '================================================================================================

        RichTextBox1.Text = RichTextBox1.Text & vbNewLine & "Loading valuation dates for accounts opened and closed intra quarter - SOQ."
        Label14.Text = "Loading valuation dates for accounts opened and closed intra quarter - SOQ"
        pstart = DatePart(DateInterval.Month, datestring1)
        pend = DatePart(DateInterval.Month, datestring2)

        Mycn.Open()

        SQLstr = "INSERT INTO mdb_BillingDateQueue (PortfolioID, PortfolioCode, PeriodID, DateNeeded, DateNeededText, QID)" & _
            "SELECT dbo_vQbRowDefPortfolio.PortfolioID, dbo_vQbRowDefPortfolio.PortfolioCode, " & pid & ", (CDate(dbo_vQbRowDefPortfolio.StartDate)) As SDate1, Format((SDate1), 'MMddyyyy'), 'V8'" & _
            " FROM (mdb_BillingRequests INNER JOIN dbo_vQbRowDefPortfolio ON mdb_BillingRequests.PortfolioID = dbo_vQbRowDefPortfolio.PortfolioID) INNER JOIN mdb_BillingPeriods ON mdb_BillingRequests.PeriodID = mdb_BillingPeriods.ID" & _
            " WHERE ((((CDate(dbo_vQbRowDefPortfolio.CloseDate)) < #" & datestring2 & "#) AND ((CDate(dbo_vQbRowDefPortfolio.CloseDate)) > #" & datestring1 & "#) AND dbo_vQbRowDefPortfolio.CloseDate Is Not Null) AND ((CDate(dbo_vQbRowDefPortfolio.StartDate)) > #" & datestring1 & "#)) AND PeriodID = " & pid & " AND ((dbo_vQbRowDefPortfolio.BillingValuationPeriod) = 'Start_of_Qtr')"

        Command = New OleDb.OleDbCommand(SQLstr, Mycn)
        Command.ExecuteNonQuery()
        RichTextBox1.Text = RichTextBox1.Text & vbNewLine & "Dates Loaded."
        ProgressBar1.Value = ProgressBar1.Value + 1
        Mycn.Close()

        '================================================================================================
        '*** BLOCK END ***
        '================================================================================================

        '************************************************************************************************
        '************************************************************************************************
        '************************************* SECTION BREAK ********************************************
        '************************************************************************************************
        '************************************************************************************************
        Label14.Text = "Finding Valuation Dates"
        '================================================================================================
        '*** BLOCK START ***
        'Groups dates need and puts them in queue
        'Tables used in this function: mdb_BillingDateQueue, mdb_BillingValQueue
        '================================================================================================

        RichTextBox1.Text = RichTextBox1.Text & vbNewLine & "Grouping dates needed."

        pstart = DatePart(DateInterval.Month, datestring1)
        pend = DatePart(DateInterval.Month, datestring2)

        Mycn.Open()

        SQLstr = "INSERT INTO mdb_BillingValQueue (DateNeeded, DateNeededText, PortfolioCount, PID)" & _
            "SELECT DISTINCT mdb_BillingDateQueue.DateNeeded, mdb_BillingDateQueue.DateNeededText, Count(mdb_BillingDateQueue.PortfolioID) AS CountOfPortfolioID, mdb_BillingDateQueue.PeriodID" & _
        " FROM(mdb_BillingDateQueue)" & _
        " WHERE PeriodID = " & pid & _
        " GROUP BY mdb_BillingDateQueue.DateNeededText, mdb_BillingDateQueue.DateNeeded, mdb_BillingDateQueue.PeriodID"


        Command = New OleDb.OleDbCommand(SQLstr, Mycn)
        Command.ExecuteNonQuery()
        RichTextBox1.Text = RichTextBox1.Text & vbNewLine & "Dates Grouped."
        ProgressBar1.Value = ProgressBar1.Value + 1
        Mycn.Close()

        '================================================================================================
        '*** BLOCK END ***
        '================================================================================================

        '================================================================================================
        '*** BLOCK START ***
        'Create Portfolio Code txt File, BAT and send to APX.
        'Tables used in this function: mdb_BillingDateQueue, mdb_BillingValQueue, mdb_BillingFileQueue
        '================================================================================================

        RichTextBox1.Text = RichTextBox1.Text & vbNewLine & "Starting process to get valuations."
        Label14.Text = "Getting Valuations"
        Mycn.Open()

        Call LoadDateQueue()

        Do

            Dim strSQL As String = "SELECT Top 1 * FROM mdb_BillingValQueue WHERE Done = False AND PID = " & pid
            Dim queryString As String = String.Format(strSQL)
            Dim cmd As New OleDb.OleDbCommand(queryString, Mycn)
            Dim daq As New OleDb.OleDbDataAdapter(cmd)
            Dim dsq As New DataSet

            daq.Fill(dsq, "User")
            Dim dtq As DataTable = dsq.Tables("User")

            Dim row1 As DataRow = dtq.Rows(0)

            'Pull ID, PostDate, QueueDate
            Dim fileid As Integer = (row1("ID"))
            Dim dateneeded1 As Date = (row1("DateNeeded"))
            Dim dateneeded2 As String = (row1("DateNeededText"))
            Dim portcount As Integer = (row1("PortfolioCount"))

            Dim datefile As String = Format(dateneeded1, "MMddyy").ToString

            RichTextBox1.Text = RichTextBox1.Text & vbNewLine & "Finding Values for " & dateneeded1 & ". " & portcount & " portfolios affected."

            'Update Working = True where ID = ID
            SQLstr = "Update mdb_BillingValQueue SET Working = True WHERE ID = " & fileid

            Command = New OleDb.OleDbCommand(SQLstr, Mycn)
            Command.ExecuteNonQuery()

            Dim path As String
            path = "\\aamapxapps01\apx$\Automation\BillingDateRequests"

            If (System.IO.Directory.Exists(path)) Then
                System.IO.Directory.Delete(path, True)
            End If

            '***********************************Code works but doesnt make APX do what is needed*******************************************************
            'Dim AccessConn As New System.Data.OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & My.Settings.dbloc & "';Jet OLEDB:Database Password='" & My.Settings.dbpass & "';")

            'AccessConn.Open()
            'RichTextBox1.Text = RichTextBox1.Text & vbNewLine & "Creating APX File Nammed '" & datefile & ".txt'..."
            'Dim AccessCommand As New OleDb.OleDbCommand("SELECT PortfolioCode INTO [Text;HDR=NO;DATABASE=" & path & "].[" & datefile & ".txt] FROM mdb_BillingDateQueue WHERE PeriodID = " & pid & " AND DateNeededText = '" & dateneeded2 & "'", AccessConn)
            'System.IO.Directory.CreateDirectory(path)
            'AccessCommand.ExecuteNonQuery()

            'AccessConn.Close()
            '***********************************END*******************************************************

            '***********************************Queue Portfolios Needed***********************************
            Dim conn As New OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & My.Settings.dbloc & "';Jet OLEDB:Database Password='" & My.Settings.dbpass & "';")
            Dim PortSQL As String = "SELECT PortfolioCode" & _
                " FROM(mdb_BillingDateQueue)" & _
                " WHERE Done = False AND PeriodID = " & pid & " AND DateNeededText = '" & dateneeded2 & "'" & _
                " GROUP BY PortfolioCode"

            Dim pa As New OleDb.OleDbDataAdapter(PortSQL, conn)
            Dim ps As New DataSet
            pa.Fill(ps, "Users")

            With DataGridView4
                .DataSource = ps.Tables("Users")
                '.Columns(0).Visible = False
            End With

            Dim PortfolioIDQueue As String

            RichTextBox4.Text = ""

            For i As Integer = 0 To (Me.DataGridView4.Rows.Count - 1)
                PortfolioIDQueue = DataGridView4.Rows(i).Cells(0).Value

                If RichTextBox4.Text = "" Or IsDBNull(RichTextBox4) Then
                    RichTextBox4.Text = PortfolioIDQueue
                Else
                    RichTextBox4.Text = RichTextBox4.Text & " " & PortfolioIDQueue
                End If

            Next

            '***********************************End Dates Needed***********************************

            '***********************************CREATE 62 Bit BAT File***********************************
            If File.Exists("\\aamapxapps01\apx$\_BillingVals\" & dateneeded2 & ".xls") Then
                File.Delete("\\aamapxapps01\apx$\_BillingVals\" & dateneeded2 & ".xls")
            Else

            End If

            If ((RadioButton2.Checked) And (dateneeded1 = DateTimePicker2.Text)) Then
                RichTextBox2.Text = "c:" & vbNewLine & "cd C:\Program Files (x86)\Advent\ApxClient\4.0" & vbNewLine & "AdvScriptRunner REPRUN -mBillingValsFinal -p@" & ComboBox2.Text & " -x -b" & dateneeded2 & " -vs -t\\aamapxapps01\apx$\_BillingVals\" & dateneeded2 & ".xls"
            Else
                RichTextBox2.Text = "c:" & vbNewLine & "cd C:\Program Files (x86)\Advent\ApxClient\4.0" & vbNewLine & "AdvScriptRunner REPRUN -mBillingValsFinal " & Chr(34) & "-p" & RichTextBox4.Text & Chr(34) & " -x -b" & dateneeded2 & " -vs -t\\aamapxapps01\apx$\_BillingVals\" & dateneeded2 & ".xls"
            End If

            RichTextBox1.Text = RichTextBox1.Text & vbNewLine & "64 Bit Script Created..."

            RichTextBox1.Text = RichTextBox1.Text & vbNewLine & "Creating BAT File..."
            Label14.Text = "Creating BAT File"
            RichTextBox2.SaveFile("\\aamapxapps01\apx$\automation\_AutomatedBillingBATS\" & datefile & "64.bat", _
                RichTextBoxStreamType.PlainText)

            RichTextBox1.Text = RichTextBox1.Text & vbNewLine & "64 Bit BAT file created..."

            System.Diagnostics.Process.Start("\\aamapxapps01\apx$\Automation\_AutomatedBillingBATS\" & datefile & "64.BAT")
            RichTextBox1.Text = RichTextBox1.Text & vbNewLine & "Sent files to RepRunner for Processing..."

            SQLstr = "Update mdb_BillingValQueue SET Working = False, Done = True WHERE ID = " & fileid

            Command = New OleDb.OleDbCommand(SQLstr, Mycn)
            Command.ExecuteNonQuery()

            SQLstr = "INSERT INTO mdb_BillingFileQueue (PID, FileLocation, Hits, PathName, FileName)" & _
            " VALUES(" & pid & ", '\\aamapxapps01\apx$\_BillingVals\" & dateneeded2 & ".xls', 0, '\\aamapxapps01\apx$\_BillingVals\','" & dateneeded2 & ".xls')"

            Command = New OleDb.OleDbCommand(SQLstr, Mycn)
            Command.ExecuteNonQuery()

            Call LoadDateQueue()


        Loop Until DataGridView1.RowCount = 0

        '================================================================================================
        '*** BLOCK END ***
        '================================================================================================

        '================================================================================================
        '*** BLOCK START ***
        'Load values into database from APX Export
        'Tables used in this function: mdb_BillingValues, mdb_BillingFileQueue
        '================================================================================================

        RichTextBox1.Text = RichTextBox1.Text & vbNewLine & "Starting process to get valuations."
        Label14.Text = "Starting process to get valuations"
        Call LoadFileQueue()
        RichTextBox1.Text = RichTextBox1.Text & vbNewLine & "Looking for finished APX File."
        Label14.Text = "Looking for finished APX File"
        Do

            Dim strSQL As String = "SELECT Top 1 * FROM mdb_BillingFileQueue WHERE Done = False AND PID = " & pid
            Dim queryString As String = String.Format(strSQL)
            Dim cmd As New OleDb.OleDbCommand(queryString, Mycn)
            Dim daq As New OleDb.OleDbDataAdapter(cmd)
            Dim dsq As New DataSet

            daq.Fill(dsq, "User")
            Dim dtq As DataTable = dsq.Tables("User")


            If dtq.Rows.Count = 0 Then
                GoTo Line1
            Else

                Dim row1 As DataRow = dtq.Rows(0)

                'Pull ID, PostDate, QueueDate
                Dim fileid As Integer = (row1("ID"))
                Dim fileloc As String = (row1("FileLocation"))
                Dim fpath As String = (row1("PathName"))
                Dim fname As String = (row1("FileName"))
                Dim hit As Integer = (row1("Hits"))

                Label11.Text = fname
                Label12.Text = "Looking for file"

                'Update Working = True where ID = ID
                SQLstr = "Update mdb_BillingFileQueue SET Working = True WHERE ID = " & fileid

                Command = New OleDb.OleDbCommand(SQLstr, Mycn)
                Command.ExecuteNonQuery()

                If (System.IO.File.Exists(fileloc)) Then
                    RichTextBox1.Text = RichTextBox1.Text & vbNewLine & "File '" & fname & "' found.  Starting Import."
                    Label12.Text = "Found File"

                    '****************** IMPORT FILE TO DATAGRID ***************************
                    Dim _filename As String = fileloc
                    Dim _conn As String
                    Dim ds1 As New DataSet
                    Dim ds2 As New DataSet


                    _conn = "Provider=Microsoft.Jet.OLEDB.4.0;" & "Data Source=" & _filename & ";" & "Extended Properties=Excel 8.0;"
                    Dim _connection As OleDbConnection = New OleDbConnection(_conn)
                    Dim da As OleDbDataAdapter = New OleDbDataAdapter()
                    Dim _command As OleDbCommand = New OleDbCommand()

                    _command.Connection = _connection
                    _command.CommandText = "SELECT * FROM [Sheet1$] WHERE [PortfolioID] Is Not Null"
                    da.SelectCommand = _command
                    da.Fill(ds1, "sheet1")

                    DataGridView3.DataSource = ds1
                    DataGridView3.DataMember = "sheet1"
                    Label12.Text = "Found " & DataGridView3.RowCount.ToString & " rows of data."
                    '****************** END IMPORT ***************************
                    RichTextBox1.Text = RichTextBox1.Text & vbNewLine & "File '" & fname & "' Imported.  Sending data to database."
                    Label14.Text = "File '" & fname & "' Imported.  Sending data to database"
                    '****************** SEND DATA TO DATABASE ***************************
                    Dim PID1 As Integer
                    Dim Dte1 As Date
                    Dim IsCash As String
                    Dim Symbol As String
                    Dim Desc As String
                    Dim Qnty As Double
                    Dim Price As Double
                    Dim TMV As Double

                    ProgressBar2.Value = 0
                    ProgressBar2.Maximum = DataGridView3.RowCount


                    For i As Integer = 0 To (Me.DataGridView3.Rows.Count - 1)
                        If DataGridView3.Rows(i).Cells(0).ValueType Is GetType(String) Then

                        Else


                            PID1 = DataGridView3.Rows(i).Cells(0).Value
                            Dte1 = DataGridView3.Rows(i).Cells(1).Value
                            IsCash = DataGridView3.Rows(i).Cells(2).Value
                            If IsDBNull(DataGridView3.Rows(i).Cells(3).Value) Then
                                Symbol = ""
                            Else
                                Symbol = DataGridView3.Rows(i).Cells(3).Value
                            End If

                            Desc = DataGridView3.Rows(i).Cells(4).Value
                            If IsDBNull(DataGridView3.Rows(i).Cells(5).Value) Then
                                Qnty = DataGridView3.Rows(i).Cells(7).Value
                            Else
                                Qnty = DataGridView3.Rows(i).Cells(5).Value
                            End If

                            If IsDBNull(DataGridView3.Rows(i).Cells(6).Value) Then
                                Price = 1
                            Else
                                Price = DataGridView3.Rows(i).Cells(6).Value
                            End If

                            TMV = DataGridView3.Rows(i).Cells(7).Value

                            SQLstr = "Insert Into mdb_BillingValues ([PortfolioID], [Date1], [IsCash], [Symbol], [Desc], [Qnty], [Price], [MarketValue], [PID])" & _
                            " VALUES(" & PID1 & ",#" & Dte1 & "#,'" & IsCash & "','" & Symbol & "','" & Desc & "'," & Qnty & "," & Price & "," & TMV & "," & pid & ")"

                            Command = New OleDb.OleDbCommand(SQLstr, Mycn)
                            Command.ExecuteNonQuery()
                        End If
                        ProgressBar2.Value = ProgressBar2.Value + 1

                        Dim nleft As Integer
                        nleft = ProgressBar2.Maximum - ProgressBar2.Value

                        Label12.Text = nleft & " records left to import."

                    Next
                    RichTextBox1.Text = RichTextBox1.Text & vbNewLine & "File '" & fname & "' loaded and imported to database."
                    Label14.Text = "File '" & fname & "' loaded and imported to database"
                    '****************** END DATASEND ***************************

                    SQLstr = "Update mdb_BillingFileQueue SET Done = True, Hits = " & hit & " WHERE ID = " & fileid
                    Command = New OleDb.OleDbCommand(SQLstr, Mycn)
                    Command.ExecuteNonQuery()

                Else

                    hit = hit + 1
                    SQLstr = "Update mdb_BillingFileQueue SET Hits = " & hit & " WHERE ID = " & fileid

                    Command = New OleDb.OleDbCommand(SQLstr, Mycn)
                    Command.ExecuteNonQuery()

                    Label12.Text = "File not found.  # of hits: " & hit + 1

                    Pause(0.1)

                    If hit > 750 Then
                        Label12.Text = "ERROR: Maximum number of hits reached."
                        RichTextBox1.Text = RichTextBox1.Text & vbNewLine & "File '" & fname & "' cannot be found."
                        RichTextBox1.Text = RichTextBox1.Text & vbNewLine & "Maximum number of hits reached.  Process Ended in Error."
                        Label14.Text = "Process ended in error."
                        GoTo line2
                    Else

                    End If

                End If

                Call LoadFileQueue()

            End If
        Loop Until DataGridView2.RowCount = 0


Line1:
        Mycn.Close()

        '================================================================================================
        '*** BLOCK END ***
        '================================================================================================

        '************************************************************************************************
        '************************************************************************************************
        '************************************* SECTION BREAK ********************************************
        '************************************************************************************************
        '************************************************************************************************

        '================================================================================================
        '*** BLOCK START ***
        'Clean up valuations table - Delete end of period records to ensure no dupes.
        'Tables used in this function: mdb_BillingValues, mdb_BillingDatesQueue
        '================================================================================================

        RichTextBox1.Text = RichTextBox1.Text & vbNewLine & "Priming Temp Table."
        Label14.Text = "Priming Temp Table"
        Mycn.Open()

        SQLstr = "DELETE * FROM mdb_BillingTempIDs"

        Command = New OleDb.OleDbCommand(SQLstr, Mycn)
        Command.ExecuteNonQuery()
        RichTextBox1.Text = RichTextBox1.Text & vbNewLine & "Table Primed."
        Label14.Text = "Table Primed"
        ProgressBar1.Value = ProgressBar1.Value + 1
        Mycn.Close()


        Mycn.Open()

        SQLstr = "INSERT INTO mdb_BillingTempIDs (ValueID)" & _
        "SELECT mdb_BillingValues.ID" & _
        " FROM mdb_BillingDateQueue INNER JOIN mdb_BillingValues ON (mdb_BillingDateQueue.DateNeeded = mdb_BillingValues.Date1) AND (mdb_BillingDateQueue.PeriodID = mdb_BillingValues.PID) AND (mdb_BillingDateQueue.PortfolioID = mdb_BillingValues.PortfolioID)" & _
        " WHERE(((mdb_BillingDateQueue.PeriodID) = " & pid & "))" & _
        " GROUP BY mdb_BillingValues.ID"

        Command = New OleDb.OleDbCommand(SQLstr, Mycn)
        Command.ExecuteNonQuery()
        RichTextBox1.Text = RichTextBox1.Text & vbNewLine & "Clean Values Identified."
        Label14.Text = "Clean Values Identified"
        ProgressBar1.Value = ProgressBar1.Value + 1
        Mycn.Close()


        RichTextBox1.Text = RichTextBox1.Text & vbNewLine & "Cleaning Valuations Table."
        Label14.Text = "Cleaning Valuation Table"

        Mycn.Open()

        SQLstr = "DELETE * FROM mdb_BillingValues WHERE PID = " & pid & " AND ID Not In (SELECT ValueID FROM mdb_BillingTempIDs)" 

        Command = New OleDb.OleDbCommand(SQLstr, Mycn)
        Command.ExecuteNonQuery()
        RichTextBox1.Text = RichTextBox1.Text & vbNewLine & "Valuation Table Cleaned."
        Label14.Text = "Valuation Table Cleaned"
        ProgressBar1.Value = ProgressBar1.Value + 1
        Mycn.Close()

        '================================================================================================
        '*** BLOCK END ***
        '================================================================================================

        '************************************************************************************************
        '************************************************************************************************
        '************************************* SECTION BREAK ********************************************
        '************************************************************************************************
        '************************************************************************************************

        '================================================================================================
        '*** BLOCK START ***
        'Calculates billable values for accounts not on tiered billing - bill on cash
        'Tables used in this function: mdb_BillingValues, dbo_vQbRowDefPortfolio, mdb_BillingSumValues
        '================================================================================================

        'RichTextBox1.Text = RichTextBox1.Text & vbNewLine & "Calculating TMV for non Tiered Billing - Bill Cash."
        'Label14.Text = "Calculating TMV for non Tiered Billing - Bill Cash"
        'Mycn.Open()

        'SQLstr = "INSERT INTO mdb_BillingSumValues (PortfolioID, BillableValue, PID, AAMRate, FirmFee, RepFee)" & _
        '"SELECT mdb_BillingValues.PortfolioID, Sum(mdb_BillingValues.MarketValue) AS SumOfMarketValue, mdb_BillingValues.PID, dbo_vQbRowDefPortfolio.TieredBillingRate1, dbo_vQbRowDefPortfolio.SolicitorsFeeFirm, dbo_vQbRowDefPortfolio.SolicitorsFeeRep" & _
        '" FROM mdb_BillingValues INNER JOIN dbo_vQbRowDefPortfolio ON mdb_BillingValues.PortfolioID = dbo_vQbRowDefPortfolio.PortfolioID" & _
        '" WHERE(((mdb_BillingValues.PID) = " & pid & ") AND ((dbo_vQbRowDefPortfolio.ManagerCode) = 'AAM') AND ((dbo_vQbRowDefPortfolio.BillonCash) = 'Yes'))" & _
        '" GROUP BY mdb_BillingValues.PortfolioID, mdb_BillingValues.PID, dbo_vQbRowDefPortfolio.TieredBillingRate1, dbo_vQbRowDefPortfolio.SolicitorsFeeFirm, dbo_vQbRowDefPortfolio.SolicitorsFeeRep;"

        'Command = New OleDb.OleDbCommand(SQLstr, Mycn)
        'Command.ExecuteNonQuery()
        'RichTextBox1.Text = RichTextBox1.Text & vbNewLine & "TMV's loaded."
        'ProgressBar1.Value = ProgressBar1.Value + 1
        'Mycn.Close()

        '================================================================================================
        '*** BLOCK END ***
        '================================================================================================

        '================================================================================================
        '*** BLOCK START ***
        'Calculates billable values for accounts on tiered billing 1 - bill on cash
        'Tables used in this function: mdb_BillingValues, dbo_vQbRowDefPortfolio, mdb_BillingSumValues
        '================================================================================================
        
        RichTextBox1.Text = RichTextBox1.Text & vbNewLine & "Calculating TMV for Tiered Billing Bracket 1 - Bill Cash."
        Label14.Text = "Calculating TMV for Tiered Billing 1 - Bill Cash"
        Mycn.Open()

        SQLstr = "INSERT INTO mdb_BillingSumValues (PortfolioID, BillableValue, PID, AAMRate, FirmFee, RepFee)" & _
        " SELECT mdb_BillingValues.PortfolioID, IIF((dbo_vQbRowDefPortfolio.TieredBillingBracket1 > 0) AND (Sum(mdb_BillingValues.MarketValue) > dbo_vQbRowDefPortfolio.TieredBillingBracket1),dbo_vQbRowDefPortfolio.TieredBillingBracket1,Sum(mdb_BillingValues.MarketValue)) AS SumOfMarketValue, mdb_BillingValues.PID, dbo_vQbRowDefPortfolio.TieredBillingRate1, dbo_vQbRowDefPortfolio.SolicitorsFeeFirm, dbo_vQbRowDefPortfolio.SolicitorsFeeRep" & _
        " FROM mdb_BillingValues INNER JOIN dbo_vQbRowDefPortfolio ON mdb_BillingValues.PortfolioID = dbo_vQbRowDefPortfolio.PortfolioID" & _
        " WHERE (((mdb_BillingValues.PID)= " & pid & ") AND ((dbo_vQbRowDefPortfolio.ManagerCode)='AAM') AND ((dbo_vQbRowDefPortfolio.BillonCash)='Yes') AND ((dbo_vQbRowDefPortfolio.BillingMethodCode)='t'))" & _
        " GROUP BY mdb_BillingValues.PortfolioID, mdb_BillingValues.PID, dbo_vQbRowDefPortfolio.TieredBillingRate1, dbo_vQbRowDefPortfolio.SolicitorsFeeFirm, dbo_vQbRowDefPortfolio.SolicitorsFeeRep, dbo_vQbRowDefPortfolio.TieredBillingBracket1, dbo_vQbRowDefPortfolio.TieredBillingBracket1From;"
        Command = New OleDb.OleDbCommand(SQLstr, Mycn)
        Command.ExecuteNonQuery()
        RichTextBox1.Text = RichTextBox1.Text & vbNewLine & "TMV's loaded."
        ProgressBar1.Value = ProgressBar1.Value + 1
        Mycn.Close()

        '================================================================================================
        '*** BLOCK END ***
        '================================================================================================

        '================================================================================================
        '*** BLOCK START ***
        'Calculates billable values for accounts on tiered billing 2 - bill on cash
        'Tables used in this function: mdb_BillingValues, dbo_vQbRowDefPortfolio, mdb_BillingSumValues
        '================================================================================================

        RichTextBox1.Text = RichTextBox1.Text & vbNewLine & "Calculating TMV for Tiered Billing Bracket 2 - Bill Cash."
        Label14.Text = "Calculating TMV for Tiered Billing 2 - Bill Cash"
        Mycn.Open()

        SQLstr = "INSERT INTO mdb_BillingSumValues (PortfolioID, BillableValue, PID, AAMRate, FirmFee, RepFee)" & _
        " SELECT mdb_BillingValues.PortfolioID, IIf(dbo_vQbRowDefPortfolio.TieredBillingBracket2 > 0,dbo_vQbRowDefPortfolio.TieredBillingBracket2 - dbo_vQbRowDefPortfolio.TieredBillingBracket1From,Sum(mdb_BillingValues.MarketValue) - dbo_vQbRowDefPortfolio.TieredBillingBracket1From) AS SumOfMarketValue, mdb_BillingValues.PID, IIF(dbo_vQbRowDefPortfolio.TieredBillingRate2 Is Null,dbo_vQbRowDefPortfolio.TieredBillingRate6,dbo_vQbRowDefPortfolio.TieredBillingRate2), dbo_vQbRowDefPortfolio.SolicitorsFeeFirm, dbo_vQbRowDefPortfolio.SolicitorsFeeRep" & _
        " FROM mdb_BillingValues INNER JOIN dbo_vQbRowDefPortfolio ON mdb_BillingValues.PortfolioID = dbo_vQbRowDefPortfolio.PortfolioID" & _
        " WHERE (((mdb_BillingValues.PID)= " & pid & ") AND ((Select Sum(mdb_BillingValues.MarketValue) FROM mdb_BillingValues WHERE mdb_BillingValues.PortfolioID = dbo_vQbRowDefPortfolio.PortfolioID AND mdb_BillingValues.PID = " & pid & ")>dbo_vQbRowDefPortfolio.TieredBillingBracket1From) AND ((dbo_vQbRowDefPortfolio.ManagerCode)='AAM') AND ((dbo_vQbRowDefPortfolio.BillonCash)='Yes') AND ((dbo_vQbRowDefPortfolio.BillingMethodCode)='t'))" & _
        " GROUP BY mdb_BillingValues.PortfolioID, mdb_BillingValues.PID, dbo_vQbRowDefPortfolio.TieredBillingRate2, dbo_vQbRowDefPortfolio.TieredBillingRate6, dbo_vQbRowDefPortfolio.SolicitorsFeeFirm, dbo_vQbRowDefPortfolio.SolicitorsFeeRep, dbo_vQbRowDefPortfolio.TieredBillingBracket2, dbo_vQbRowDefPortfolio.TieredBillingBracket1From;"
        Command = New OleDb.OleDbCommand(SQLstr, Mycn)
        Command.ExecuteNonQuery()
        RichTextBox1.Text = RichTextBox1.Text & vbNewLine & "TMV's loaded."
        ProgressBar1.Value = ProgressBar1.Value + 1
        Mycn.Close()

        '================================================================================================
        '*** BLOCK END ***
        '================================================================================================

        '================================================================================================
        '*** BLOCK START ***
        'Calculates billable values for accounts on tiered billing 3 - bill on cash
        'Tables used in this function: mdb_BillingValues, dbo_vQbRowDefPortfolio, mdb_BillingSumValues
        '================================================================================================

        RichTextBox1.Text = RichTextBox1.Text & vbNewLine & "Calculating TMV for Tiered Billing Bracket 3 - Bill Cash."
        Label14.Text = "Calculating TMV for Tiered Billing 3 - Bill Cash"
        Mycn.Open()

        SQLstr = "INSERT INTO mdb_BillingSumValues (PortfolioID, BillableValue, PID, AAMRate, FirmFee, RepFee)" & _
        " SELECT mdb_BillingValues.PortfolioID, IIf(dbo_vQbRowDefPortfolio.TieredBillingBracket3 > 0,dbo_vQbRowDefPortfolio.TieredBillingBracket3 - dbo_vQbRowDefPortfolio.TieredBillingBracket2From,Sum(mdb_BillingValues.MarketValue) - dbo_vQbRowDefPortfolio.TieredBillingBracket2From) AS SumOfMarketValue, mdb_BillingValues.PID, IIF(dbo_vQbRowDefPortfolio.TieredBillingRate3 Is Null,dbo_vQbRowDefPortfolio.TieredBillingRate6,dbo_vQbRowDefPortfolio.TieredBillingRate3), dbo_vQbRowDefPortfolio.SolicitorsFeeFirm, dbo_vQbRowDefPortfolio.SolicitorsFeeRep" & _
        " FROM mdb_BillingValues INNER JOIN dbo_vQbRowDefPortfolio ON mdb_BillingValues.PortfolioID = dbo_vQbRowDefPortfolio.PortfolioID" & _
        " WHERE (((mdb_BillingValues.PID)= " & pid & ") AND ((Select Sum(mdb_BillingValues.MarketValue) FROM mdb_BillingValues WHERE mdb_BillingValues.PortfolioID = dbo_vQbRowDefPortfolio.PortfolioID AND mdb_BillingValues.PID = " & pid & ")>dbo_vQbRowDefPortfolio.TieredBillingBracket2From) AND ((dbo_vQbRowDefPortfolio.ManagerCode)='AAM') AND ((dbo_vQbRowDefPortfolio.BillonCash)='Yes') AND ((dbo_vQbRowDefPortfolio.BillingMethodCode)='t'))" & _
        " GROUP BY mdb_BillingValues.PortfolioID, mdb_BillingValues.PID, dbo_vQbRowDefPortfolio.TieredBillingRate3, dbo_vQbRowDefPortfolio.TieredBillingRate6, dbo_vQbRowDefPortfolio.SolicitorsFeeFirm, dbo_vQbRowDefPortfolio.SolicitorsFeeRep, dbo_vQbRowDefPortfolio.TieredBillingBracket3, dbo_vQbRowDefPortfolio.TieredBillingBracket2From;"
        Command = New OleDb.OleDbCommand(SQLstr, Mycn)
        Command.ExecuteNonQuery()
        RichTextBox1.Text = RichTextBox1.Text & vbNewLine & "TMV's loaded."
        ProgressBar1.Value = ProgressBar1.Value + 1
        Mycn.Close()

        '================================================================================================
        '*** BLOCK END ***
        '================================================================================================

        '================================================================================================
        '*** BLOCK START ***
        'Calculates billable values for accounts on tiered billing 4 - bill on cash
        'Tables used in this function: mdb_BillingValues, dbo_vQbRowDefPortfolio, mdb_BillingSumValues
        '================================================================================================

        RichTextBox1.Text = RichTextBox1.Text & vbNewLine & "Calculating TMV for Tiered Billing Bracket 4 - Bill Cash."
        Label14.Text = "Calculating TMV for Tiered Billing 4 - Bill Cash"
        Mycn.Open()

        SQLstr = "INSERT INTO mdb_BillingSumValues (PortfolioID, BillableValue, PID, AAMRate, FirmFee, RepFee)" & _
        " SELECT mdb_BillingValues.PortfolioID, IIf(dbo_vQbRowDefPortfolio.TieredBillingBracket4 > 0,dbo_vQbRowDefPortfolio.TieredBillingBracket4 - dbo_vQbRowDefPortfolio.TieredBillingBracket3From,Sum(mdb_BillingValues.MarketValue) - dbo_vQbRowDefPortfolio.TieredBillingBracket3From) AS SumOfMarketValue, mdb_BillingValues.PID, IIF(dbo_vQbRowDefPortfolio.TieredBillingRate4 Is Null,dbo_vQbRowDefPortfolio.TieredBillingRate6,dbo_vQbRowDefPortfolio.TieredBillingRate4), dbo_vQbRowDefPortfolio.SolicitorsFeeFirm, dbo_vQbRowDefPortfolio.SolicitorsFeeRep" & _
        " FROM mdb_BillingValues INNER JOIN dbo_vQbRowDefPortfolio ON mdb_BillingValues.PortfolioID = dbo_vQbRowDefPortfolio.PortfolioID" & _
        " WHERE (((mdb_BillingValues.PID)= " & pid & ") AND ((Select Sum(mdb_BillingValues.MarketValue) FROM mdb_BillingValues WHERE mdb_BillingValues.PortfolioID = dbo_vQbRowDefPortfolio.PortfolioID AND mdb_BillingValues.PID = " & pid & ")>dbo_vQbRowDefPortfolio.TieredBillingBracket3From) AND ((dbo_vQbRowDefPortfolio.ManagerCode)='AAM') AND ((dbo_vQbRowDefPortfolio.BillonCash)='Yes') AND ((dbo_vQbRowDefPortfolio.BillingMethodCode)='t'))" & _
        " GROUP BY mdb_BillingValues.PortfolioID, mdb_BillingValues.PID, dbo_vQbRowDefPortfolio.TieredBillingRate4, dbo_vQbRowDefPortfolio.TieredBillingRate6,dbo_vQbRowDefPortfolio.SolicitorsFeeFirm, dbo_vQbRowDefPortfolio.SolicitorsFeeRep, dbo_vQbRowDefPortfolio.TieredBillingBracket4, dbo_vQbRowDefPortfolio.TieredBillingBracket3From;"
        Command = New OleDb.OleDbCommand(SQLstr, Mycn)
        Command.ExecuteNonQuery()
        RichTextBox1.Text = RichTextBox1.Text & vbNewLine & "TMV's loaded."
        ProgressBar1.Value = ProgressBar1.Value + 1
        Mycn.Close()

        '================================================================================================
        '*** BLOCK END ***
        '================================================================================================

        '================================================================================================
        '*** BLOCK START ***
        'Calculates billable values for accounts on tiered billing 5 - bill on cash
        'Tables used in this function: mdb_BillingValues, dbo_vQbRowDefPortfolio, mdb_BillingSumValues
        '================================================================================================

        RichTextBox1.Text = RichTextBox1.Text & vbNewLine & "Calculating TMV for Tiered Billing Bracket 5 - Bill Cash."
        Label14.Text = "Calculating TMV for Tiered Billing 5 - Bill Cash"
        Mycn.Open()

        SQLstr = "INSERT INTO mdb_BillingSumValues (PortfolioID, BillableValue, PID, AAMRate, FirmFee, RepFee)" & _
        " SELECT mdb_BillingValues.PortfolioID, IIf(dbo_vQbRowDefPortfolio.TieredBillingBracket5 > 0,dbo_vQbRowDefPortfolio.TieredBillingBracket5 - dbo_vQbRowDefPortfolio.TieredBillingBracket4From,Sum(mdb_BillingValues.MarketValue) - dbo_vQbRowDefPortfolio.TieredBillingBracket4From) AS SumOfMarketValue, mdb_BillingValues.PID, IIF(dbo_vQbRowDefPortfolio.TieredBillingRate5 Is Null,dbo_vQbRowDefPortfolio.TieredBillingRate6,dbo_vQbRowDefPortfolio.TieredBillingRate5), dbo_vQbRowDefPortfolio.SolicitorsFeeFirm, dbo_vQbRowDefPortfolio.SolicitorsFeeRep" & _
        " FROM mdb_BillingValues INNER JOIN dbo_vQbRowDefPortfolio ON mdb_BillingValues.PortfolioID = dbo_vQbRowDefPortfolio.PortfolioID" & _
        " WHERE (((mdb_BillingValues.PID)= " & pid & ") AND ((Select Sum(mdb_BillingValues.MarketValue) FROM mdb_BillingValues WHERE mdb_BillingValues.PortfolioID = dbo_vQbRowDefPortfolio.PortfolioID AND mdb_BillingValues.PID = " & pid & ")>dbo_vQbRowDefPortfolio.TieredBillingBracket4From) AND ((dbo_vQbRowDefPortfolio.ManagerCode)='AAM') AND ((dbo_vQbRowDefPortfolio.BillonCash)='Yes') AND ((dbo_vQbRowDefPortfolio.BillingMethodCode)='t'))" & _
        " GROUP BY mdb_BillingValues.PortfolioID, mdb_BillingValues.PID, dbo_vQbRowDefPortfolio.TieredBillingRate5, dbo_vQbRowDefPortfolio.TieredBillingRate6, dbo_vQbRowDefPortfolio.SolicitorsFeeFirm, dbo_vQbRowDefPortfolio.SolicitorsFeeRep, dbo_vQbRowDefPortfolio.TieredBillingBracket5, dbo_vQbRowDefPortfolio.TieredBillingBracket4From;"
        Command = New OleDb.OleDbCommand(SQLstr, Mycn)
        Command.ExecuteNonQuery()
        RichTextBox1.Text = RichTextBox1.Text & vbNewLine & "TMV's loaded."
        ProgressBar1.Value = ProgressBar1.Value + 1
        Mycn.Close()

        '================================================================================================
        '*** BLOCK END ***
        '================================================================================================

        '================================================================================================
        '*** BLOCK START ***
        'Calculates billable values for accounts on tiered billing 6 - bill on cash
        'Tables used in this function: mdb_BillingValues, dbo_vQbRowDefPortfolio, mdb_BillingSumValues
        '================================================================================================

        RichTextBox1.Text = RichTextBox1.Text & vbNewLine & "Calculating TMV for Tiered Billing Bracket 6 - Bill Cash."
        Label14.Text = "Calculating TMV for Tiered Billing 6 - Bill Cash"
        Mycn.Open()

        SQLstr = "INSERT INTO mdb_BillingSumValues (PortfolioID, BillableValue, PID, AAMRate, FirmFee, RepFee)" & _
        " SELECT mdb_BillingValues.PortfolioID, (Sum(mdb_BillingValues.MarketValue) - dbo_vQbRowDefPortfolio.TieredBillingBracket5From) AS SumOfMarketValue, mdb_BillingValues.PID, dbo_vQbRowDefPortfolio.TieredBillingRate6, dbo_vQbRowDefPortfolio.SolicitorsFeeFirm, dbo_vQbRowDefPortfolio.SolicitorsFeeRep" & _
        " FROM mdb_BillingValues INNER JOIN dbo_vQbRowDefPortfolio ON mdb_BillingValues.PortfolioID = dbo_vQbRowDefPortfolio.PortfolioID" & _
        " WHERE (((mdb_BillingValues.PID)= " & pid & ") AND ((Select Sum(mdb_BillingValues.MarketValue) FROM mdb_BillingValues WHERE mdb_BillingValues.PortfolioID = dbo_vQbRowDefPortfolio.PortfolioID AND mdb_BillingValues.PID = " & pid & ")>dbo_vQbRowDefPortfolio.TieredBillingBracket5From) AND ((dbo_vQbRowDefPortfolio.ManagerCode)='AAM') AND ((dbo_vQbRowDefPortfolio.BillonCash)='Yes') AND ((dbo_vQbRowDefPortfolio.BillingMethodCode)='t'))" & _
        " GROUP BY mdb_BillingValues.PortfolioID, mdb_BillingValues.PID, dbo_vQbRowDefPortfolio.TieredBillingRate6, dbo_vQbRowDefPortfolio.SolicitorsFeeFirm, dbo_vQbRowDefPortfolio.SolicitorsFeeRep, dbo_vQbRowDefPortfolio.TieredBillingBracket5From;"
        Command = New OleDb.OleDbCommand(SQLstr, Mycn)
        Command.ExecuteNonQuery()
        RichTextBox1.Text = RichTextBox1.Text & vbNewLine & "TMV's loaded."
        ProgressBar1.Value = ProgressBar1.Value + 1
        Mycn.Close()

        ProgressBar1.Maximum = ProgressBar1.Maximum + 5
        '================================================================================================
        '*** BLOCK END ***
        '================================================================================================

        '================================================================================================
        '*** BLOCK START ***
        'Calculates billable values for accounts not on tiered billing - bill without cash
        'Tables used in this function: mdb_BillingValues, dbo_vQbRowDefPortfolio, mdb_BillingSumValues
        '================================================================================================

        RichTextBox1.Text = RichTextBox1.Text & vbNewLine & "Calculating TMV for non Tiered Billing - Bill Without Cash."
        Label14.Text = "Calculating TMV for non Tiered Billing - Bill Without Cash"
        Mycn.Open()

        SQLstr = "INSERT INTO mdb_BillingSumValues (PortfolioID, BillableValue, PID, AAMRate, FirmFee, RepFee)" & _
        "SELECT mdb_BillingValues.PortfolioID, Sum(mdb_BillingValues.MarketValue) AS SumOfMarketValue, mdb_BillingValues.PID, dbo_vQbRowDefPortfolio.TieredBillingRate1, dbo_vQbRowDefPortfolio.SolicitorsFeeFirm, dbo_vQbRowDefPortfolio.SolicitorsFeeRep" & _
        " FROM mdb_BillingValues INNER JOIN dbo_vQbRowDefPortfolio ON mdb_BillingValues.PortfolioID = dbo_vQbRowDefPortfolio.PortfolioID" & _
        " WHERE(((mdb_BillingValues.PID) = " & pid & ") AND ((dbo_vQbRowDefPortfolio.ManagerCode) = 'AAM') AND ((dbo_vQbRowDefPortfolio.BillonCash) = 'No') AND ((mdb_BillingValues.IsCash) = 'n'))" & _
        " GROUP BY mdb_BillingValues.PortfolioID, mdb_BillingValues.PID, dbo_vQbRowDefPortfolio.TieredBillingRate1, dbo_vQbRowDefPortfolio.SolicitorsFeeFirm, dbo_vQbRowDefPortfolio.SolicitorsFeeRep;"

        Command = New OleDb.OleDbCommand(SQLstr, Mycn)
        Command.ExecuteNonQuery()
        RichTextBox1.Text = RichTextBox1.Text & vbNewLine & "TMV's loaded."
        ProgressBar1.Value = ProgressBar1.Value + 1
        Mycn.Close()

        '================================================================================================
        '*** BLOCK END ***
        '================================================================================================

        '================================================================================================
        '*** BLOCK START ***
        'Calculates billable values for third party accounts
        'Tables used in this function: mdb_BillingValues, dbo_vQbRowDefPortfolio, mdb_BillingSumValues, env_ExternalRate
        '================================================================================================

        RichTextBox1.Text = RichTextBox1.Text & vbNewLine & "Calculating TMV for third party accounts."
        Label14.Text = "Calculating TMV for third party accounts"
        Mycn.Open()

        SQLstr = "INSERT INTO mdb_BillingSumValues (PortfolioID, BillableValue, PID, AAMRate)" & _
        "SELECT mdb_BillingValues.PortfolioID, Sum(mdb_BillingValues.MarketValue) AS SumOfMarketValue, mdb_BillingValues.PID, env_ExternalRate.Rate" & _
        " FROM (mdb_BillingValues INNER JOIN dbo_vQbRowDefPortfolio ON mdb_BillingValues.PortfolioID = dbo_vQbRowDefPortfolio.PortfolioID) INNER JOIN env_ExternalRate ON dbo_vQbRowDefPortfolio.ManagerCode = env_ExternalRate.MCode" & _
        " WHERE (((mdb_BillingValues.PID)=" & pid & ") AND ((dbo_vQbRowDefPortfolio.ManagerCode)<>'AAM'))" & _
        " GROUP BY mdb_BillingValues.PortfolioID, mdb_BillingValues.PID, env_ExternalRate.Rate, dbo_vQbRowDefPortfolio.SolicitorsFeeFirm, dbo_vQbRowDefPortfolio.SolicitorsFeeRep;"

        Command = New OleDb.OleDbCommand(SQLstr, Mycn)
        Command.ExecuteNonQuery()
        RichTextBox1.Text = RichTextBox1.Text & vbNewLine & "TMV's loaded."
        ProgressBar1.Value = ProgressBar1.Value + 1
        Mycn.Close()

        '================================================================================================
        '*** BLOCK END ***
        '================================================================================================

        '************************************************************************************************
        '************************************************************************************************
        '************************************* SECTION BREAK ********************************************
        '************************************************************************************************
        '************************************************************************************************

        '================================================================================================
        '*** BLOCK START ***
        'Calculates final billing
        'Tables used in this function: mdb_BillingDaysBillable, mdb_BillingDateQueue, mdb_BillingSumValues
        '================================================================================================

        RichTextBox1.Text = RichTextBox1.Text & vbNewLine & "Calculating Billing."
        Label14.Text = "Calculating Billing"
        Mycn.Open()

        SQLstr = "INSERT INTO mdb_BillingBase (PortfolioID, DaysBillable, DateOfValue, BillableMarketValue, AAMRate, SolicitorFirm, SolicitorRep, AAMValue, SolicitorFirmValue, SolicitorRepValue, PID)" & _
        "SELECT mdb_BillingSumValues.PortfolioID, mdb_BillingDaysBillable.DaysBillable, mdb_BillingDateQueue.DateNeeded, mdb_BillingSumValues.BillableValue, mdb_BillingSumValues.AAMRate, mdb_BillingSumValues.FirmFee, mdb_BillingSumValues.RepFee, ((mdb_BillingSumValues.BillableValue) * (mdb_BillingDaysBillable.DaysBillable) * (mdb_BillingSumValues.AAMRate)/36000) As AAMValue, ((mdb_BillingSumValues.BillableValue) * (mdb_BillingDaysBillable.DaysBillable) * (mdb_BillingSumValues.FirmFee)/36000) As FirmValue, ((mdb_BillingSumValues.BillableValue) * (mdb_BillingDaysBillable.DaysBillable) * (mdb_BillingSumValues.RepFee)/36000) As RepValue, mdb_BillingSumValues.PID" & _
        " FROM (mdb_BillingSumValues INNER JOIN mdb_BillingDateQueue ON (mdb_BillingSumValues.PID = mdb_BillingDateQueue.PeriodID) AND (mdb_BillingSumValues.PortfolioID = mdb_BillingDateQueue.PortfolioID)) INNER JOIN mdb_BillingDaysBillable ON (mdb_BillingSumValues.PID = mdb_BillingDaysBillable.PeriodID) AND (mdb_BillingSumValues.PortfolioID = mdb_BillingDaysBillable.PortfolioID)" & _
        " WHERE(((mdb_BillingSumValues.PID) = " & pid & "))" & _
        " GROUP BY mdb_BillingSumValues.PortfolioID, mdb_BillingDaysBillable.DaysBillable, mdb_BillingDateQueue.DateNeeded, mdb_BillingSumValues.BillableValue, mdb_BillingSumValues.AAMRate, mdb_BillingSumValues.FirmFee, mdb_BillingSumValues.RepFee, mdb_BillingSumValues.PID;"

        Command = New OleDb.OleDbCommand(SQLstr, Mycn)
        Command.ExecuteNonQuery()
        RichTextBox1.Text = RichTextBox1.Text & vbNewLine & "Fees Calculated."
        ProgressBar1.Value = ProgressBar1.Value + 1
        Mycn.Close()

        '================================================================================================
        '*** BLOCK END ***
        '================================================================================================

        '************************************************************************************************
        '************************************************************************************************
        '************************************* SECTION BREAK ********************************************
        '************************************************************************************************
        '************************************************************************************************

        RichTextBox1.Text = RichTextBox1.Text & vbNewLine & "Validating Data."
        Label14.Text = "Validating Data"
        '================================================================================================
        '*** BLOCK START ***
        'Filtering accounts that were successfully billed
        'Tables used in this function: mdb_BillingRequests, mdb_BillingExceptions, dbo_vQbRowDefPortfolio
        '================================================================================================

        RichTextBox1.Text = RichTextBox1.Text & vbNewLine & "Filtering accounts that were billed."

        Mycn.Open()

        SQLstr = "DELETE * FROM mdb_BillingRequests WHERE ((mdb_BillingRequests.PeriodID = " & pid & ") And PortfolioID In (SELECT PortfolioID from mdb_BillingBase WHERE PID = " & pid & "))"

        Command = New OleDb.OleDbCommand(SQLstr, Mycn)
        Command.ExecuteNonQuery()
        RichTextBox1.Text = RichTextBox1.Text & vbNewLine & "Done."
        ProgressBar1.Value = ProgressBar1.Value + 1
        Mycn.Close()

        '================================================================================================
        '*** BLOCK END ***
        '================================================================================================

        '================================================================================================
        '*** BLOCK START ***
        'Finds accounts opened after end of period
        'Tables used in this function: mdb_BillingRequests, mdb_BillingExceptions, dbo_vQbRowDefPortfolio
        '================================================================================================

        RichTextBox1.Text = RichTextBox1.Text & vbNewLine & "Filtering accounts opened after end of period."

        Mycn.Open()

        SQLstr = "INSERT INTO mdb_BillingExceptions (PortfolioID, PID, ReasonCode)" & _
        "SELECT mdb_BillingRequests.PortfolioID, " & pid & ",1" & _
        " FROM mdb_BillingRequests INNER JOIN dbo_vQbRowDefPortfolio ON mdb_BillingRequests.PortfolioID = dbo_vQbRowDefPortfolio.PortfolioID" & _
        " WHERE(((mdb_BillingRequests.PeriodID) = " & pid & ") And ((CDate([dbo_vQbRowDefPortfolio].[StartDate])) > #" & datestring2 & "#) And ((dbo_vQbRowDefPortfolio.StartDate) Is Not Null))" & _
        " GROUP BY mdb_BillingRequests.PortfolioID"

        Command = New OleDb.OleDbCommand(SQLstr, Mycn)
        Command.ExecuteNonQuery()
        RichTextBox1.Text = RichTextBox1.Text & vbNewLine & "Done."
        ProgressBar1.Value = ProgressBar1.Value + 1
        Mycn.Close()

        '================================================================================================
        '*** BLOCK END ***
        '================================================================================================

        '================================================================================================
        '*** BLOCK START ***
        'Filtering accounts already flagged for a known exception
        'Tables used in this function: mdb_BillingRequests, mdb_BillingExceptions, dbo_vQbRowDefPortfolio
        '================================================================================================

        RichTextBox1.Text = RichTextBox1.Text & vbNewLine & "Filtering accounts aready flagged for an exception."

        Mycn.Open()

        SQLstr = "DELETE * FROM mdb_BillingRequests WHERE ((mdb_BillingRequests.PeriodID = " & pid & ") And PortfolioID In (SELECT PortfolioID from mdb_BillingExceptions WHERE PID = " & pid & "))"

        Command = New OleDb.OleDbCommand(SQLstr, Mycn)
        Command.ExecuteNonQuery()
        RichTextBox1.Text = RichTextBox1.Text & vbNewLine & "Done."
        ProgressBar1.Value = ProgressBar1.Value + 1
        Mycn.Close()

        '================================================================================================
        '*** BLOCK END ***
        '================================================================================================

        '================================================================================================
        '*** BLOCK START ***
        'Finds accounts closed before start of period
        'Tables used in this function: mdb_BillingRequests, mdb_BillingExceptions, dbo_vQbRowDefPortfolio
        '================================================================================================

        RichTextBox1.Text = RichTextBox1.Text & vbNewLine & "Filtering accounts closed before start of period."

        Mycn.Open()

        SQLstr = "INSERT INTO mdb_BillingExceptions (PortfolioID, PID, ReasonCode)" & _
        "SELECT mdb_BillingRequests.PortfolioID, " & pid & ",2" & _
        " FROM mdb_BillingRequests INNER JOIN dbo_vQbRowDefPortfolio ON mdb_BillingRequests.PortfolioID = dbo_vQbRowDefPortfolio.PortfolioID" & _
        " WHERE(((mdb_BillingRequests.PeriodID) = " & pid & ") And ((CDate([dbo_vQbRowDefPortfolio].[CloseDate])) < #" & datestring1 & "#) And ((dbo_vQbRowDefPortfolio.CloseDate) Is Not Null))" & _
        " GROUP BY mdb_BillingRequests.PortfolioID"

        Command = New OleDb.OleDbCommand(SQLstr, Mycn)
        Command.ExecuteNonQuery()
        RichTextBox1.Text = RichTextBox1.Text & vbNewLine & "Done."
        ProgressBar1.Value = ProgressBar1.Value + 1
        Mycn.Close()

        '================================================================================================
        '*** BLOCK END ***
        '================================================================================================

        '================================================================================================
        '*** BLOCK START ***
        'Filtering accounts already flagged for a known exception
        'Tables used in this function: mdb_BillingRequests, mdb_BillingExceptions, dbo_vQbRowDefPortfolio
        '================================================================================================

        RichTextBox1.Text = RichTextBox1.Text & vbNewLine & "Filtering accounts aready flagged for an exception."

        Mycn.Open()

        SQLstr = "DELETE * FROM mdb_BillingRequests WHERE ((mdb_BillingRequests.PeriodID = " & pid & ") And PortfolioID In (SELECT PortfolioID from mdb_BillingExceptions WHERE PID = " & pid & "))"

        Command = New OleDb.OleDbCommand(SQLstr, Mycn)
        Command.ExecuteNonQuery()
        RichTextBox1.Text = RichTextBox1.Text & vbNewLine & "Done."
        ProgressBar1.Value = ProgressBar1.Value + 1
        Mycn.Close()

        '================================================================================================
        '*** BLOCK END ***
        '================================================================================================


        '================================================================================================
        '*** BLOCK START ***
        'Finds accounts not set to open or closed
        'Tables used in this function: mdb_BillingRequests, mdb_BillingExceptions, dbo_vQbRowDefPortfolio
        '================================================================================================

        RichTextBox1.Text = RichTextBox1.Text & vbNewLine & "Filtering accounts not set to Open or Closed."

        Mycn.Open()

        SQLstr = "INSERT INTO mdb_BillingExceptions (PortfolioID, PID, ReasonCode)" & _
        "SELECT mdb_BillingRequests.PortfolioID, " & pid & ",3" & _
        " FROM mdb_BillingRequests INNER JOIN dbo_vQbRowDefPortfolio ON mdb_BillingRequests.PortfolioID = dbo_vQbRowDefPortfolio.PortfolioID" & _
        " WHERE (((mdb_BillingRequests.PeriodID)=" & pid & ") AND ((dbo_vQbRowDefPortfolio.PortfolioStatus)<>'Open') AND ((dbo_vQbRowDefPortfolio.PortfolioStatus)<>'Closed'))" & _
        " GROUP BY mdb_BillingRequests.PortfolioID"

        Command = New OleDb.OleDbCommand(SQLstr, Mycn)
        Command.ExecuteNonQuery()
        RichTextBox1.Text = RichTextBox1.Text & vbNewLine & "Done."
        ProgressBar1.Value = ProgressBar1.Value + 1
        Mycn.Close()

        '================================================================================================
        '*** BLOCK END ***
        '================================================================================================

        '================================================================================================
        '*** BLOCK START ***
        'Filtering accounts already flagged for a known exception
        'Tables used in this function: mdb_BillingRequests, mdb_BillingExceptions, dbo_vQbRowDefPortfolio
        '================================================================================================

        RichTextBox1.Text = RichTextBox1.Text & vbNewLine & "Filtering accounts aready flagged for an exception."

        Mycn.Open()

        SQLstr = "DELETE * FROM mdb_BillingRequests WHERE ((mdb_BillingRequests.PeriodID = " & pid & ") And PortfolioID In (SELECT PortfolioID from mdb_BillingExceptions WHERE PID = " & pid & "))"

        Command = New OleDb.OleDbCommand(SQLstr, Mycn)
        Command.ExecuteNonQuery()
        RichTextBox1.Text = RichTextBox1.Text & vbNewLine & "Done."
        ProgressBar1.Value = ProgressBar1.Value + 1
        Mycn.Close()

        '================================================================================================
        '*** BLOCK END ***
        '================================================================================================

        
        '================================================================================================
        '*** BLOCK START ***
        'Finds accounts where active billing is null
        'Tables used in this function: mdb_BillingRequests, mdb_BillingExceptions, dbo_vQbRowDefPortfolio
        '================================================================================================

        RichTextBox1.Text = RichTextBox1.Text & vbNewLine & "Filtering accounts where active billing is undefined."

        Mycn.Open()

        SQLstr = "INSERT INTO mdb_BillingExceptions (PortfolioID, PID, ReasonCode)" & _
        "SELECT mdb_BillingRequests.PortfolioID, " & pid & ",5" & _
        " FROM mdb_BillingRequests INNER JOIN dbo_vQbRowDefPortfolio ON mdb_BillingRequests.PortfolioID = dbo_vQbRowDefPortfolio.PortfolioID" & _
        " WHERE (((mdb_BillingRequests.PeriodID)=" & pid & ") AND ((dbo_vQbRowDefPortfolio.ActiveBilling) Is Null))" & _
        " GROUP BY mdb_BillingRequests.PortfolioID"

        Command = New OleDb.OleDbCommand(SQLstr, Mycn)
        Command.ExecuteNonQuery()
        RichTextBox1.Text = RichTextBox1.Text & vbNewLine & "Done."
        ProgressBar1.Value = ProgressBar1.Value + 1
        Mycn.Close()

        '================================================================================================
        '*** BLOCK END ***
        '================================================================================================

        '================================================================================================
        '*** BLOCK START ***
        'Filtering accounts already flagged for a known exception
        'Tables used in this function: mdb_BillingRequests, mdb_BillingExceptions, dbo_vQbRowDefPortfolio
        '================================================================================================

        RichTextBox1.Text = RichTextBox1.Text & vbNewLine & "Filtering accounts aready flagged for an exception."

        Mycn.Open()

        SQLstr = "DELETE * FROM mdb_BillingRequests WHERE ((mdb_BillingRequests.PeriodID = " & pid & ") And PortfolioID In (SELECT PortfolioID from mdb_BillingExceptions WHERE PID = " & pid & "))"

        Command = New OleDb.OleDbCommand(SQLstr, Mycn)
        Command.ExecuteNonQuery()
        RichTextBox1.Text = RichTextBox1.Text & vbNewLine & "Done."
        ProgressBar1.Value = ProgressBar1.Value + 1
        Mycn.Close()

        '================================================================================================
        '*** BLOCK END ***
        '================================================================================================


        '================================================================================================
        '*** BLOCK START ***
        'Finds accounts where bill on cash is null
        'Tables used in this function: mdb_BillingRequests, mdb_BillingExceptions, dbo_vQbRowDefPortfolio
        '================================================================================================

        RichTextBox1.Text = RichTextBox1.Text & vbNewLine & "Filtering accounts where bill on cash is undefined."

        Mycn.Open()

        SQLstr = "INSERT INTO mdb_BillingExceptions (PortfolioID, PID, ReasonCode)" & _
        "SELECT mdb_BillingRequests.PortfolioID, " & pid & ",6" & _
        " FROM mdb_BillingRequests INNER JOIN dbo_vQbRowDefPortfolio ON mdb_BillingRequests.PortfolioID = dbo_vQbRowDefPortfolio.PortfolioID" & _
        " WHERE (((mdb_BillingRequests.PeriodID)=" & pid & ") AND ((dbo_vQbRowDefPortfolio.BillonCash) Is Null))" & _
        " GROUP BY mdb_BillingRequests.PortfolioID"

        Command = New OleDb.OleDbCommand(SQLstr, Mycn)
        Command.ExecuteNonQuery()
        RichTextBox1.Text = RichTextBox1.Text & vbNewLine & "Done."
        ProgressBar1.Value = ProgressBar1.Value + 1
        Mycn.Close()

        '================================================================================================
        '*** BLOCK END ***
        '================================================================================================

        '================================================================================================
        '*** BLOCK START ***
        'Filtering accounts already flagged for a known exception
        'Tables used in this function: mdb_BillingRequests, mdb_BillingExceptions, dbo_vQbRowDefPortfolio
        '================================================================================================

        RichTextBox1.Text = RichTextBox1.Text & vbNewLine & "Filtering accounts aready flagged for an exception."

        Mycn.Open()

        SQLstr = "DELETE * FROM mdb_BillingRequests WHERE ((mdb_BillingRequests.PeriodID = " & pid & ") And PortfolioID In (SELECT PortfolioID from mdb_BillingExceptions WHERE PID = " & pid & "))"

        Command = New OleDb.OleDbCommand(SQLstr, Mycn)
        Command.ExecuteNonQuery()
        RichTextBox1.Text = RichTextBox1.Text & vbNewLine & "Done."
        ProgressBar1.Value = ProgressBar1.Value + 1
        Mycn.Close()

        '================================================================================================
        '*** BLOCK END ***
        '================================================================================================


        '================================================================================================
        '*** BLOCK START ***
        'Finds accounts where billing valuation period is null
        'Tables used in this function: mdb_BillingRequests, mdb_BillingExceptions, dbo_vQbRowDefPortfolio
        '================================================================================================

        RichTextBox1.Text = RichTextBox1.Text & vbNewLine & "Filtering accounts where billing valuation period is undefined."

        Mycn.Open()

        SQLstr = "INSERT INTO mdb_BillingExceptions (PortfolioID, PID, ReasonCode)" & _
        "SELECT mdb_BillingRequests.PortfolioID, " & pid & ",7" & _
        " FROM mdb_BillingRequests INNER JOIN dbo_vQbRowDefPortfolio ON mdb_BillingRequests.PortfolioID = dbo_vQbRowDefPortfolio.PortfolioID" & _
        " WHERE (((mdb_BillingRequests.PeriodID)=" & pid & ") AND ((dbo_vQbRowDefPortfolio.BillingValuationPeriod) Is Null))" & _
        " GROUP BY mdb_BillingRequests.PortfolioID"

        Command = New OleDb.OleDbCommand(SQLstr, Mycn)
        Command.ExecuteNonQuery()
        RichTextBox1.Text = RichTextBox1.Text & vbNewLine & "Done."
        ProgressBar1.Value = ProgressBar1.Value + 1
        Mycn.Close()

        '================================================================================================
        '*** BLOCK END ***
        '================================================================================================

        '================================================================================================
        '*** BLOCK START ***
        'Filtering accounts already flagged for a known exception
        'Tables used in this function: mdb_BillingRequests, mdb_BillingExceptions, dbo_vQbRowDefPortfolio
        '================================================================================================

        RichTextBox1.Text = RichTextBox1.Text & vbNewLine & "Filtering accounts aready flagged for an exception."

        Mycn.Open()

        SQLstr = "DELETE * FROM mdb_BillingRequests WHERE ((mdb_BillingRequests.PeriodID = " & pid & ") And PortfolioID In (SELECT PortfolioID from mdb_BillingExceptions WHERE PID = " & pid & "))"

        Command = New OleDb.OleDbCommand(SQLstr, Mycn)
        Command.ExecuteNonQuery()
        RichTextBox1.Text = RichTextBox1.Text & vbNewLine & "Done."
        ProgressBar1.Value = ProgressBar1.Value + 1
        Mycn.Close()

        '================================================================================================
        '*** BLOCK END ***
        '================================================================================================


        '================================================================================================
        '*** BLOCK START ***
        'Finds accounts where Bill in Arrears is null
        'Tables used in this function: mdb_BillingRequests, mdb_BillingExceptions, dbo_vQbRowDefPortfolio
        '================================================================================================

        RichTextBox1.Text = RichTextBox1.Text & vbNewLine & "Filtering accounts where billing style is undefined."

        Mycn.Open()

        SQLstr = "INSERT INTO mdb_BillingExceptions (PortfolioID, PID, ReasonCode)" & _
        "SELECT mdb_BillingRequests.PortfolioID, " & pid & ",8" & _
        " FROM mdb_BillingRequests INNER JOIN dbo_vQbRowDefPortfolio ON mdb_BillingRequests.PortfolioID = dbo_vQbRowDefPortfolio.PortfolioID" & _
        " WHERE (((mdb_BillingRequests.PeriodID)=" & pid & ") AND ((dbo_vQbRowDefPortfolio.BillInArrears) Is Null))" & _
        " GROUP BY mdb_BillingRequests.PortfolioID"

        Command = New OleDb.OleDbCommand(SQLstr, Mycn)
        Command.ExecuteNonQuery()
        RichTextBox1.Text = RichTextBox1.Text & vbNewLine & "Done."
        ProgressBar1.Value = ProgressBar1.Value + 1
        Mycn.Close()

        '================================================================================================
        '*** BLOCK END ***
        '================================================================================================

        '================================================================================================
        '*** BLOCK START ***
        'Filtering accounts already flagged for a known exception
        'Tables used in this function: mdb_BillingRequests, mdb_BillingExceptions, dbo_vQbRowDefPortfolio
        '================================================================================================

        RichTextBox1.Text = RichTextBox1.Text & vbNewLine & "Filtering accounts aready flagged for an exception."

        Mycn.Open()

        SQLstr = "DELETE * FROM mdb_BillingRequests WHERE ((mdb_BillingRequests.PeriodID = " & pid & ") And PortfolioID In (SELECT PortfolioID from mdb_BillingExceptions WHERE PID = " & pid & "))"

        Command = New OleDb.OleDbCommand(SQLstr, Mycn)
        Command.ExecuteNonQuery()
        RichTextBox1.Text = RichTextBox1.Text & vbNewLine & "Done."
        ProgressBar1.Value = ProgressBar1.Value + 1
        Mycn.Close()

        '================================================================================================
        '*** BLOCK END ***
        '================================================================================================


        '================================================================================================
        '*** BLOCK START ***
        'Finds accounts where Rate is null - Prop SMA
        'Tables used in this function: mdb_BillingRequests, mdb_BillingExceptions, dbo_vQbRowDefPortfolio
        '================================================================================================

        RichTextBox1.Text = RichTextBox1.Text & vbNewLine & "Filtering accounts where billing rate is undefined - AAM Managed."

        Mycn.Open()

        SQLstr = "INSERT INTO mdb_BillingExceptions (PortfolioID, PID, ReasonCode)" & _
        "SELECT mdb_BillingRequests.PortfolioID, " & pid & ",9" & _
        " FROM mdb_BillingRequests INNER JOIN dbo_vQbRowDefPortfolio ON mdb_BillingRequests.PortfolioID = dbo_vQbRowDefPortfolio.PortfolioID" & _
        " WHERE (((mdb_BillingRequests.PeriodID)=" & pid & ") AND ((dbo_vQbRowDefPortfolio.TieredBillingRate1) Is Null) AND (dbo_vQbRowDefPortfolio.ManagerCode = 'AAM'))" & _
        " GROUP BY mdb_BillingRequests.PortfolioID"

        Command = New OleDb.OleDbCommand(SQLstr, Mycn)
        Command.ExecuteNonQuery()
        RichTextBox1.Text = RichTextBox1.Text & vbNewLine & "Done."
        ProgressBar1.Value = ProgressBar1.Value + 1
        Mycn.Close()

        '================================================================================================
        '*** BLOCK END ***
        '================================================================================================

        '================================================================================================
        '*** BLOCK START ***
        'Filtering accounts already flagged for a known exception
        'Tables used in this function: mdb_BillingRequests, mdb_BillingExceptions, dbo_vQbRowDefPortfolio
        '================================================================================================

        RichTextBox1.Text = RichTextBox1.Text & vbNewLine & "Filtering accounts aready flagged for an exception."

        Mycn.Open()

        SQLstr = "DELETE * FROM mdb_BillingRequests WHERE ((mdb_BillingRequests.PeriodID = " & pid & ") And PortfolioID In (SELECT PortfolioID from mdb_BillingExceptions WHERE PID = " & pid & "))"

        Command = New OleDb.OleDbCommand(SQLstr, Mycn)
        Command.ExecuteNonQuery()
        RichTextBox1.Text = RichTextBox1.Text & vbNewLine & "Done."
        ProgressBar1.Value = ProgressBar1.Value + 1
        Mycn.Close()

        '================================================================================================
        '*** BLOCK END ***
        '================================================================================================


        '================================================================================================
        '*** BLOCK START ***
        'Finds accounts where Rate is null - Non Prop SMA
        'Tables used in this function: mdb_BillingRequests, mdb_BillingExceptions, dbo_vQbRowDefPortfolio
        '================================================================================================

        RichTextBox1.Text = RichTextBox1.Text & vbNewLine & "Filtering accounts where billing rate is undefined - External."

        Mycn.Open()

        SQLstr = "INSERT INTO mdb_BillingExceptions (PortfolioID, PID, ReasonCode)" & _
        "SELECT mdb_BillingRequests.PortfolioID, " & pid & ",10" & _
        " FROM (mdb_BillingRequests INNER JOIN dbo_vQbRowDefPortfolio ON mdb_BillingRequests.PortfolioID = dbo_vQbRowDefPortfolio.PortfolioID) INNER JOIN env_ExternalRate ON dbo_vQbRowDefPortfolio.ManagerCode = env_ExternalRate.MCode" & _
        " WHERE (((mdb_BillingRequests.PeriodID)=" & pid & ") AND ((dbo_vQbRowDefPortfolio.ManagerCode)<>'AAM') AND ((env_ExternalRate.Rate) Is Null) AND ((env_ExternalRate.Active)=True))" & _
        " GROUP BY mdb_BillingRequests.PortfolioID;"

        Command = New OleDb.OleDbCommand(SQLstr, Mycn)
        Command.ExecuteNonQuery()
        RichTextBox1.Text = RichTextBox1.Text & vbNewLine & "Done."
        ProgressBar1.Value = ProgressBar1.Value + 1
        Mycn.Close()

        '================================================================================================
        '*** BLOCK END ***
        '================================================================================================

        '================================================================================================
        '*** BLOCK START ***
        'Filtering accounts already flagged for a known exception
        'Tables used in this function: mdb_BillingRequests, mdb_BillingExceptions, dbo_vQbRowDefPortfolio
        '================================================================================================

        RichTextBox1.Text = RichTextBox1.Text & vbNewLine & "Filtering accounts aready flagged for an exception."

        Mycn.Open()

        SQLstr = "DELETE * FROM mdb_BillingRequests WHERE ((mdb_BillingRequests.PeriodID = " & pid & ") And PortfolioID In (SELECT PortfolioID from mdb_BillingExceptions WHERE PID = " & pid & "))"

        Command = New OleDb.OleDbCommand(SQLstr, Mycn)
        Command.ExecuteNonQuery()
        RichTextBox1.Text = RichTextBox1.Text & vbNewLine & "Done."
        ProgressBar1.Value = ProgressBar1.Value + 1
        Mycn.Close()

        '================================================================================================
        '*** BLOCK END ***
        '================================================================================================


        '================================================================================================
        '*** BLOCK START ***
        'Finds accounts with no calculated value
        'Tables used in this function: mdb_BillingRequests, mdb_BillingExceptions, mdb_BillingSumValues
        '================================================================================================

        RichTextBox1.Text = RichTextBox1.Text & vbNewLine & "Filtering accounts with no value found."

        Mycn.Open()

        SQLstr = "INSERT INTO mdb_BillingExceptions (PortfolioID, PID, ReasonCode)" & _
        "SELECT mdb_BillingRequests.PortfolioID, " & pid & ",4" & _
        " FROM mdb_BillingRequests" & _
        " WHERE mdb_BillingRequests.PeriodID = " & pid & " AND PortfolioID Not In (SELECT mdb_BillingSumValues.PortfolioID FROM mdb_BillingSumValues WHERE mdb_BillingSumValues.PID = " & pid & ")"

        Command = New OleDb.OleDbCommand(SQLstr, Mycn)
        Command.ExecuteNonQuery()
        RichTextBox1.Text = RichTextBox1.Text & vbNewLine & "Done."
        ProgressBar1.Value = ProgressBar1.Value + 1
        Mycn.Close()

        '================================================================================================
        '*** BLOCK END ***
        '================================================================================================


        '================================================================================================
        '*** BLOCK START ***
        'Filtering accounts already flagged for a known exception
        'Tables used in this function: mdb_BillingRequests, mdb_BillingExceptions, dbo_vQbRowDefPortfolio
        '================================================================================================

        RichTextBox1.Text = RichTextBox1.Text & vbNewLine & "Filtering accounts aready flagged for an exception."

        Mycn.Open()

        SQLstr = "DELETE * FROM mdb_BillingRequests WHERE ((mdb_BillingRequests.PeriodID = " & pid & ") And PortfolioID In (SELECT PortfolioID from mdb_BillingExceptions WHERE PID = " & pid & "))"

        Command = New OleDb.OleDbCommand(SQLstr, Mycn)
        Command.ExecuteNonQuery()
        RichTextBox1.Text = RichTextBox1.Text & vbNewLine & "Done."
        ProgressBar1.Value = ProgressBar1.Value + 1
        Mycn.Close()

        '================================================================================================
        '*** BLOCK END ***
        '================================================================================================

        '================================================================================================
        '*** BLOCK START ***
        'Finds accounts with an unknown exception
        'Tables used in this function: mdb_BillingRequests, mdb_BillingExceptions
        '================================================================================================

        RichTextBox1.Text = RichTextBox1.Text & vbNewLine & "Filtering accounts with an unknown exception."

        Mycn.Open()

        SQLstr = "INSERT INTO mdb_BillingExceptions (PortfolioID, PID, ReasonCode)" & _
        "SELECT mdb_BillingRequests.PortfolioID, " & pid & ",11" & _
        " FROM mdb_BillingRequests" & _
        " WHERE (((mdb_BillingRequests.PeriodID)=" & pid & "))" & _
        " GROUP BY mdb_BillingRequests.PortfolioID"

        Command = New OleDb.OleDbCommand(SQLstr, Mycn)
        Command.ExecuteNonQuery()
        RichTextBox1.Text = RichTextBox1.Text & vbNewLine & "Done."
        ProgressBar1.Value = ProgressBar1.Value + 1
        Mycn.Close()

        '================================================================================================
        '*** BLOCK END ***
        '================================================================================================

        '================================================================================================
        '*** BLOCK START ***
        'Clear Request Table
        'Tables used in this function: mdb_BillingRequests
        '================================================================================================

        RichTextBox1.Text = RichTextBox1.Text & vbNewLine & "Clearing Request Table."

        Mycn.Open()

        SQLstr = "DELETE * FROM mdb_BillingRequests WHERE (PeriodID = " & pid & ")"

        Command = New OleDb.OleDbCommand(SQLstr, Mycn)
        Command.ExecuteNonQuery()
        RichTextBox1.Text = RichTextBox1.Text & vbNewLine & "Done."
        ProgressBar1.Value = ProgressBar1.Value + 1
        Mycn.Close()

        '================================================================================================
        '*** BLOCK END ***
        '================================================================================================


        '************************************************************************************************
        '************************************************************************************************
        '************************************* SECTION BREAK ********************************************
        '************************************************************************************************
        '************************************************************************************************
        'progressbar 42

        If rtype = 1 Then
            Dim pbar As Integer
            pbar = ProgressBar1.Maximum + 3
            ProgressBar1.Maximum = pbar

            RichTextBox1.Text = RichTextBox1.Text & vbNewLine & "Cleaning Billed Records."

            '================================================================================================
            '*** BLOCK START ***
            'Filtering accounts that were rebilled
            'Tables used in this function: mdb_BillingBase, mdb_BillingExceptions, dbo_vQbRowDefPortfolio
            '================================================================================================

            RichTextBox1.Text = RichTextBox1.Text & vbNewLine & "Filtering accounts that were re-billed."

            Mycn.Open()

            SQLstr = "DELETE * FROM mdb_BillingBase WHERE ((Locked = False) AND PID In (Select PID from mdb_BillingPeriods WHERE TypeID = 1 AND PeriodStart = #" & datestring1 & "# AND PeriodEnd = #" & datestring2 & "# AND PID <> " & pid & "))"

            Command = New OleDb.OleDbCommand(SQLstr, Mycn)
            Command.ExecuteNonQuery()
            RichTextBox1.Text = RichTextBox1.Text & vbNewLine & "Done."
            ProgressBar1.Value = ProgressBar1.Value + 1
            Mycn.Close()

            '================================================================================================
            '*** BLOCK END ***
            '================================================================================================
            GoTo line3 'TEMP FIX AROUND LOCKED RECORDS BUG - DELETES LOTS OF GOOD RECORDS.
            '================================================================================================
            '*** BLOCK START ***
            'Insert portfolios into temp table that have already been locked.
            'Tables used in this function: mdb_BillingRequests, mdb_BillingExceptions, dbo_vQbRowDefPortfolio
            '================================================================================================

            RichTextBox1.Text = RichTextBox1.Text & vbNewLine & "Finding accounts with a locked value."

            Mycn.Open()

            SQLstr = "DELETE * FROM mdb_BillingTempLocked"

            Command = New OleDb.OleDbCommand(SQLstr, Mycn)
            Command.ExecuteNonQuery()
            Mycn.Close()

            Mycn.Open()

            Dim sdate As Date
            Dim edate As Date
            sdate = DateTimePicker1.Text
            edate = DateTimePicker2.Text

            SQLstr = "INSERT INTO mdb_BillingTempLocked (PortfolioID)" & _
            "Select PortfolioID FROM mdb_BillingBase WHERE Locked = True AND PID In (Select PID from mdb_BillingPeriods WHERE TypeID = 1 AND PeriodStart = #" & sdate & "# AND PeriodEnd = #" & edate & "# AND PID <> " & pid & ")"

            Command = New OleDb.OleDbCommand(SQLstr, Mycn)
            Command.ExecuteNonQuery()
            RichTextBox1.Text = RichTextBox1.Text & vbNewLine & "Done."
            ProgressBar1.Value = ProgressBar1.Value + 1
            Mycn.Close()

            '================================================================================================
            '*** BLOCK END ***
            '================================================================================================

            '================================================================================================
            '*** BLOCK START ***
            'Filtering accounts that were already locked
            'Tables used in this function: mdb_BillingRequests, mdb_BillingExceptions, dbo_vQbRowDefPortfolio
            '================================================================================================

            RichTextBox1.Text = RichTextBox1.Text & vbNewLine & "Filtering accounts that have been locked for period."

            Mycn.Open()

            SQLstr = "DELETE * FROM mdb_BillingBase WHERE ((Locked = False) AND PID = " & pid & " AND PortfolioID In (Select PortfolioID FROM mdb_BillingTempLocked))"

            Command = New OleDb.OleDbCommand(SQLstr, Mycn)
            Command.ExecuteNonQuery()
            RichTextBox1.Text = RichTextBox1.Text & vbNewLine & "Done."
            ProgressBar1.Value = ProgressBar1.Value + 1
            Mycn.Close()

            '================================================================================================
            '*** BLOCK END ***
            '================================================================================================


        Else
            RichTextBox1.Text = RichTextBox1.Text & vbNewLine & "On the fly billing ran - removing old records."
            Dim pbar As Integer
            pbar = ProgressBar1.Maximum + 1
            ProgressBar1.Maximum = pbar

            '================================================================================================
            '*** BLOCK START ***
            'Filtering accounts that were rebilled
            'Tables used in this function: mdb_BillingBase, mdb_BillingExceptions, dbo_vQbRowDefPortfolio
            '================================================================================================

            'RichTextBox1.Text = RichTextBox1.Text & vbNewLine & "Filtering accounts that were re-billed."

            Mycn.Open()

            SQLstr = "DELETE * FROM mdb_BillingBase WHERE ((Locked = False) AND PID In (Select PID from mdb_BillingPeriods WHERE TypeID = 2 AND PeriodStart = #" & datestring1 & "# AND PeriodEnd = #" & datestring2 & "# AND PID <> " & pid & "))"

            Command = New OleDb.OleDbCommand(SQLstr, Mycn)
            Command.ExecuteNonQuery()
            RichTextBox1.Text = RichTextBox1.Text & vbNewLine & "Done."
            ProgressBar1.Value = ProgressBar1.Value + 1
            Mycn.Close()

            '================================================================================================
            '*** BLOCK END ***
            '================================================================================================


        End If
line3:
        Timer1.Enabled = False
        Label3.Text = "DONE!"
        Label14.Text = "Process Finished.  Values Stored under Period ID: " & pid

line2:
        'Catch ex As Exception
        'MsgBox(ex.Message, MsgBoxStyle.Critical, "ERROR")
        'End Try
    End Sub

    Public Sub Pause(ByVal dblSecs As Double)
        Const OneSec As Double = 1.0# / (1440.0# * 60.0#)
        Dim dblWaitTil As Date
        Now.AddSeconds(OneSec)
        dblWaitTil = Now.AddSeconds(OneSec).AddSeconds(dblSecs)
        Do Until Now > dblWaitTil
            Application.DoEvents() ' Allow windows messages to be processed
        Loop
    End Sub

    Public Function FirstDayOfQuarter(ByVal DateIn As DateTime) _
          As DateTime
        Dim intQuarterNum As Integer = (Month(DateIn) - 1) \ 3
        Return DateSerial(Year(DateIn), 3 * intQuarterNum + 1, 0)
    End Function


    Public Function LastDayOfQuarter(ByVal DateIn As Date) As Date
        Dim intQuarterNum As Integer = (Month(DateIn) - 1) \ 3 + 1
        Return DateSerial(Year(DateIn), 3 * intQuarterNum + 1, 0)
    End Function

    Public Sub LoadDateQueue()
        Try
            Dim conn As New OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & My.Settings.dbloc & "';Jet OLEDB:Database Password='" & My.Settings.dbpass & "';")

            Dim strSQL As String

            strSQL = "SELECT ID, DateNeeded As [Date], DateNeededText As [Text], PortfolioCount As [Portfolios], Done" & _
            " FROM(mdb_BillingValQueue)" & _
            " WHERE(((Done) = False))" & _
            " GROUP BY ID, DateNeeded, DateNeededText, PortfolioCount, Done"

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
    End Sub

    Public Sub LoadFileQueue()
        Try
            Dim conn As New OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & My.Settings.dbloc & "';Jet OLEDB:Database Password='" & My.Settings.dbpass & "';")

            Dim strSQL As String

            strSQL = "SELECT ID, FileLocation, Done" & _
            " FROM(mdb_BillingFileQueue)" & _
            " WHERE(((Done) = False))" & _
            " GROUP BY ID, FileLocation, Done"

            Dim da As New OleDb.OleDbDataAdapter(strSQL, conn)
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

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        MsgBox("You selected @" & ComboBox2.Text & ".")
    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)

        ''create connection 
        Dim mycn As New OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & My.Settings.dbloc & "';Jet OLEDB:Database Password='" & My.Settings.dbpass & "';")
        mycn.Open()

        'Command = New OleDb.OleDbCommand(SQLstr, Mycn)
        'Command.ExecuteNonQuery()
        'Dim comm As SqlCommand = New SqlCommand()
        'comm.Connection = conn

        ''insert data to sql database row by row
        Dim PID As Integer
        Dim Dte1 As Date
        Dim IsCash As String
        Dim Symbol As String
        Dim Desc As String
        Dim Qnty As Double
        Dim Price As Double
        Dim TMV As Double
        'Dim PID As Integer

        Dim Command As OleDb.OleDbCommand
        Dim SQLstr As String

        For i As Integer = 0 To (Me.DataGridView3.Rows.Count - 1)
            PID = DataGridView3.Rows(i).Cells(0).Value
            Dte1 = DataGridView3.Rows(i).Cells(1).Value
            IsCash = DataGridView3.Rows(i).Cells(2).Value
            Symbol = DataGridView3.Rows(i).Cells(3).Value
            Desc = DataGridView3.Rows(i).Cells(4).Value
            If IsDBNull(DataGridView3.Rows(i).Cells(5).Value) Then
                Qnty = DataGridView3.Rows(i).Cells(7).Value
            Else
                Qnty = DataGridView3.Rows(i).Cells(5).Value
            End If

            If IsDBNull(DataGridView3.Rows(i).Cells(6).Value) Then
                Price = 1
            Else
                Price = DataGridView3.Rows(i).Cells(6).Value
            End If

            TMV = DataGridView3.Rows(i).Cells(7).Value


            SQLstr = "Insert Into mdb_BillingValues ([PortfolioID], [Date1], [IsCash], [Symbol], [Desc], [Qnty], [Price], [MarketValue])" & _
            " VALUES(" & PID & ",#" & Dte1 & "#,'" & IsCash & "','" & Symbol & "','" & Desc & "'," & Qnty & "," & Price & "," & TMV & ")"

            Command = New OleDb.OleDbCommand(SQLstr, mycn)
            Command.ExecuteNonQuery()
        Next

        mycn.Close()



    End Sub

    Private Sub convertExcelToCSV(ByVal sourceFile As String, ByVal worksheetName As String, ByVal targetFile As String)

        Try
            Dim conn As New OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & My.Settings.dbloc & "';Jet OLEDB:Database Password='" & My.Settings.dbpass & "';")

            Dim strSQL As String

            strSQL = "SELECT PortfolioCode" & _
            " FROM(mdb_BillingDateQueue)" & _
            " WHERE(((Done) = False))" & _
            " GROUP BY ID, FileLocation, Done"

            Dim da As New OleDb.OleDbDataAdapter(strSQL, conn)
            Dim ds As New DataSet
            da.Fill(ds, "Users")

            With DataGridView4
                .DataSource = ds.Tables("Users")
                .Columns(0).Visible = False
            End With

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "ERROR")
        End Try

    End Sub

    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Dim _filename As String = "C:\03312014.xls"

        Dim _conn As String

        Dim ds1 As New DataSet

        Dim ds2 As New DataSet


        _conn = "Provider=Microsoft.Jet.OLEDB.4.0;" & "Data Source=" & _filename & ";" & "Extended Properties=Excel 8.0;"

        Dim _connection As OleDbConnection = New OleDbConnection(_conn)

        Dim da As OleDbDataAdapter = New OleDbDataAdapter()

        Dim _command As OleDbCommand = New OleDbCommand()

        _command.Connection = _connection

        _command.CommandText = "SELECT * FROM [Sheet1$] WHERE [PortfolioID] Is Not Null"

        da.SelectCommand = _command

        Try

            da.Fill(ds1, "sheet1")

            MessageBox.Show("The import is complete!")

            Me.DataGridView3.DataSource = ds1

            Me.DataGridView3.DataMember = "sheet1"

        Catch e1 As Exception

            MessageBox.Show("Import Failed, correct Column name in the sheet!")

        End Try

    End Sub


    Private Sub Button5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Call ImportToDB()

    End Sub

    Public Sub ImportToDB()
        'Create a new instance of dtExcelData datatable
        dtExcelData = New DataTable

        'Set the cursor to wait cursor
        Windows.Forms.Cursor.Current = Cursors.WaitCursor

        'Call ReadExcelFile function to get the data from excel
        dtExcelData = ReadExcelFile()

        'Reset the progressbar
        Me.ProgressBar2.Value = 0

        'Get the DataTable row count to set the progressbar maximum value
        Me.ProgressBar2.Maximum = dtExcelData.Rows.Count
        Me.ProgressBar2.Visible = True

        'Use looping to read the value of field in each row in DataTable
        For i = 0 To dtExcelData.Rows.Count - 1
            'Check if the student number has a value
            If Not IsDBNull(dtExcelData.Rows(i)(0)) Then

                'call save procedure and pass the row varialble i(row index) as parameter to save each the value of each field
                SaveToDB(i)

                'Increase the value of progressbar to inform the user.
                Me.ProgressBar2.Value = Me.ProgressBar2.Value + 1

            End If

        Next

        dtExcelData = Nothing

        'Call the Load_Data procedure that will load the data and display it to DataGrid
        'Load_Data()

        'Inform the user that the importing of data has been finished
        MsgBox("Data has successfully imported.", MsgBoxStyle.OkOnly, "Import Export Demo")

    End Sub

    Public Sub SaveToDB(ByVal iRowIndex As Long)
        Dim conn As New OleDbConnection
        'Dim sConnString As String
        Dim cmd As New OleDbCommand
        Dim sSQL As String = String.Empty

        'Try
        'Check if the path has a backslash in the end of string
        'If Microsoft.VisualBasic.Right(Application.StartupPath, 1) = "\" Then
        'sConnString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & Application.StartupPath & "dbexport.accdb;Persist Security Info=False;"
        'Else
        'sConnString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & Application.StartupPath & "\dbexport.accdb;Persist Security Info=False;"
        'End If

        'Dim mycn As New OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & My.Settings.dbloc & "';Jet OLEDB:Database Password='" & My.Settings.dbpass & "';")

        'create a new instance of connection
        conn = New OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & My.Settings.dbloc & "';Jet OLEDB:Database Password='" & My.Settings.dbpass & "';")

        'open the connection to be used by command object
        conn.Open()

        'Set the command's connection to our opened connection
        cmd.Connection = conn

        'Set the command type to CommandType.Text in order to use SQL statment constructed here 
        'in code editor
        cmd.CommandType = CommandType.Text

        'Set the comment text to insert the data to database
        cmd.CommandText = "INSERT INTO mdb_BillingValues ([PortfolioID], [Date1], [IsCash], [Symbol], [Desc], [Qnty], [Price], [MarketValue]) VALUES(@PortfolioID, @Dte1, @IsCash, @Symbol, @Desc, @Qnty, @Price, @TMV)"

        'SQLstr = "Insert Into mdb_BillingValues ([PortfolioID], [Date1], [IsCash], [Symbol], [Desc], [Qnty], [Price], [MarketValue])" & _
        '" VALUES(" & PID & ",#" & Dte1 & "#,'" & IsCash & "','" & Symbol & "','" & Desc & "'," & Qnty & "," & Price & "," & TMV & ")"

        'Add parameters in order to set the values in the query
        cmd.Parameters.Add("@PortfolioID", OleDbType.VarChar).Value = dtExcelData.Rows(iRowIndex)(0)
        cmd.Parameters.Add("@Dte1", OleDbType.Date).Value = dtExcelData.Rows(iRowIndex)(1)
        cmd.Parameters.Add("@IsCash", OleDbType.VarChar).Value = dtExcelData.Rows(iRowIndex)(2)
        cmd.Parameters.Add("@Symbol", OleDbType.VarChar).Value = dtExcelData.Rows(iRowIndex)(3)
        cmd.Parameters.Add("@Desc", OleDbType.VarChar).Value = dtExcelData.Rows(iRowIndex)(4)
        'cmd.Parameters.Add("@Qnty", OleDbType.Double).Value = dtExcelData.Rows(iRowIndex)(1)
        cmd.Parameters.Add("@Qnty", OleDbType.Double).Value = IIf(Not IsDBNull(dtExcelData.Rows(iRowIndex)(5)), dtExcelData.Rows(iRowIndex)(7), Nothing)
        'cmd.Parameters.Add("@Price", OleDbType.Double).Value = dtExcelData.Rows(iRowIndex)(1)
        cmd.Parameters.Add("@Price", OleDbType.Double).Value = IIf(Not IsDBNull(dtExcelData.Rows(iRowIndex)(6)), dtExcelData.Rows(iRowIndex)(7), Nothing)
        cmd.Parameters.Add("@TMV", OleDbType.Double).Value = dtExcelData.Rows(iRowIndex)(7)
        'This is just a sample of how to check if the field is null.
        'cmd.Parameters.Add("@grade", OleDbType.Numeric).Value = IIf(Not IsDBNull(dtExcelData.Rows(iRowIndex)(2)), dtExcelData.Rows(iRowIndex)(2), Nothing)
        cmd.ExecuteNonQuery()


        'Catch ex As Exception
        'MsgBox(ErrorToString)
        'Finally
        conn.Close()
        'End Try
    End Sub

    Function ReadExcelFile()
        'Use OleDbDataAdapter  to provide communication between the DataTable and the OleDb Data Sources
        Dim da As New OleDbDataAdapter

        'Use DataTable as storage of data from excel
        Dim dt As New DataTable

        'Use OleDbCommand to execute our SQL statement
        Dim cmd As New OleDbCommand

        'Use OleDbConnection that will be used by OleDbCommand to connect to excel file
        Dim xlsConn As OleDbConnection

        Dim sPath As String = String.Empty

        sPath = "\\aamapxapps01\apx$\_BillingVals\03312014.xls"

        'Create a new instance of connection and set the datasource value to excel's path
        xlsConn = New OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & sPath & ";Extended Properties=Excel 12.0")

        'Use try catch block to handle some or all possible errors that may occur in a 
        'given block of code, while still running code.

        Try
            'Open the connection
            xlsConn.Open()

            'Set the command connection to opened connection
            cmd.Connection = xlsConn

            'Set the command type to CommandType.Text in order to use SQL statment constructed here 
            'in code editor
            cmd.CommandType = CommandType.Text

            'Assigned the command text to query the excel as shown below
            cmd.CommandText = ("select * from [Sheet1$]")

            'Assign the cmd to dataadapter
            da.SelectCommand = cmd

            'Fill the datatable with data from excel file using DataAdapter
            da.Fill(dt)

        Catch
            'This block Handle the exception.
            MsgBox(ErrorToString)
        Finally
            'We need to close the connection and set to nothing. This code will still execute even the code raised an error
            xlsConn.Close()
            xlsConn = Nothing
        End Try
        Return dt
    End Function

    Private Sub Timer1_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer1.Tick
        If Label3.Text = "Working..." Then
            Label3.Text = "Working"
        Else
            If Label3.Text = "Working" Then
                Label3.Text = "Working."
            Else
                If Label3.Text = "Working." Then
                    Label3.Text = "Working.."
                Else
                    If Label3.Text = "Working.." Then
                        Label3.Text = "Working..."
                    End If
                End If
            End If
        End If
    End Sub

    Private Sub RevenueCenter_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        'Passes Dates to find the start and end of qtr dates
        Dim dte As Date
        Dim tdy As Date = Format(Now())
        dte = DateAdd(DateInterval.Month, -1, tdy)

        Dim QtrStart As Date = FirstDayOfQuarter(dte)
        Dim QtrEnd As Date = LastDayOfQuarter(dte)

        DateTimePicker1.Value = QtrStart
        DateTimePicker2.Value = QtrEnd

        DateTimePicker3.Value = QtrStart
        DateTimePicker4.Value = QtrEnd

        Call LoadExceptionReasons()
        Call LoadExceptionPeriods()
    End Sub

    Private Sub GroupBox7_Enter(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles GroupBox7.Enter

    End Sub

    Public Sub LoadExceptionPeriods()
        Try

            Dim conn As New OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & My.Settings.dbloc & "';Jet OLEDB:Database Password='" & My.Settings.dbpass & "';")

            Dim strSQL As String = "SELECT mdb_BillingPeriods.ID" & _
            " FROM(mdb_BillingPeriods)" & _
            " ORDER BY mdb_BillingPeriods.ID DESC;"

            Dim da As New OleDb.OleDbDataAdapter(strSQL, conn)
            Dim ds As New DataSet
            da.Fill(ds, "Users")

            With ComboBox3
                .DataSource = ds.Tables("Users")
                .DisplayMember = "ID"
                .ValueMember = "ID"
                .SelectedIndex = 0
            End With

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "ERROR")
        End Try
    End Sub

    Public Sub LoadExceptionReasons()
        Try

            Dim conn As New OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & My.Settings.dbloc & "';Jet OLEDB:Database Password='" & My.Settings.dbpass & "';")

            Dim strSQL As String = "SELECT mdb_BillingExceptionsReasons.ID, mdb_BillingExceptionsReasons.ReasonText" & _
            " FROM mdb_BillingExceptionsReasons;"

            Dim da As New OleDb.OleDbDataAdapter(strSQL, conn)
            Dim ds As New DataSet
            da.Fill(ds, "Users")

            With ComboBox4
                .DataSource = ds.Tables("Users")
                .DisplayMember = "ReasonText"
                .ValueMember = "ID"
                .SelectedIndex = 0
            End With

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "ERROR")
        End Try
    End Sub

    Private Sub Button3_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        Call LoadExceptionPeriods()
    End Sub

    Private Sub CheckBox1_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckBox1.CheckedChanged
        If CheckBox1.Checked Then
            ComboBox4.Enabled = True
        Else
            ComboBox4.Enabled = False
        End If
    End Sub

    Private Sub CheckBox2_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckBox2.CheckedChanged
        If CheckBox2.Checked Then
            TextBox1.Enabled = True
        Else
            TextBox1.Enabled = False
        End If
    End Sub

    Private Sub Button2_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        Try
            Dim conn As New OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & My.Settings.dbloc & "';Jet OLEDB:Database Password='" & My.Settings.dbpass & "';")

            Dim strSQL As String

            strSQL = "SELECT mdb_BillingExceptions.ID, dbo_vQbRowDefPortfolio.PortfolioCode As [Portfolio Code], dbo_vQbRowDefPortfolio.ReportHeading1 As [Account Name], dbo_vQbRowDefPortfolio.StartDate As [Start Date], dbo_vQbRowDefPortfolio.CloseDate As [Close Date], dbo_vQbRowDefPortfolio.PortfolioStatus As [Portfolio Status], mdb_BillingExceptionsReasons.ReasonText As [Exception Reason], mdb_BillingExceptionsReasons.ResolutionText As [Resolution], mdb_BillingPeriods.PeriodStart As [Period Start], mdb_BillingPeriods.PeriodEnd As [Period End]" & _
            " FROM ((mdb_BillingExceptions INNER JOIN dbo_vQbRowDefPortfolio ON mdb_BillingExceptions.PortfolioID = dbo_vQbRowDefPortfolio.PortfolioID) INNER JOIN mdb_BillingExceptionsReasons ON mdb_BillingExceptions.ReasonCode = mdb_BillingExceptionsReasons.ID) INNER JOIN mdb_BillingPeriods ON mdb_BillingExceptions.PID = mdb_BillingPeriods.ID" & _
            " WHERE mdb_BillingExceptions.PID = " & ComboBox3.SelectedValue

            If CheckBox1.Checked Then
                strSQL = strSQL & " AND mdb_BillingExceptions.ReasonCode = " & ComboBox4.SelectedValue
            Else

            End If

            If CheckBox2.Checked Then
                strSQL = strSQL & " AND dbo_vQbRowDefPortfolio.PortfolioCode Like '%" & TextBox1.Text & "%'"
            Else

            End If

            Dim da As New OleDb.OleDbDataAdapter(strSQL, conn)
            Dim ds As New DataSet
            da.Fill(ds, "Users")

            With DataGridView5
                .DataSource = ds.Tables("Users")
                .Columns(0).Visible = False
            End With

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "ERROR")
        End Try
    End Sub

    Private Sub ckbCustodian_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ckbCustodian.CheckedChanged
        If ckbCustodian.Checked Then
            cboCustodian.Enabled = True
            Call LoadCustodianList()
        Else
            cboCustodian.Enabled = False
        End If
    End Sub

    Public Sub LoadCustodianList()
        Try

            Dim conn As New OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & My.Settings.dbloc & "';Jet OLEDB:Database Password='" & My.Settings.dbpass & "';")

            Dim strSQL As String = "SELECT DISTINCT dbo_vQbRowDefPortfolio.Custodian7" & _
            " FROM(dbo_vQbRowDefPortfolio)" & _
            " WHERE(((dbo_vQbRowDefPortfolio.Custodian7) Is Not Null))" & _
            " GROUP BY dbo_vQbRowDefPortfolio.Custodian7" & _
            " ORDER BY dbo_vQbRowDefPortfolio.Custodian7"

            Dim da As New OleDb.OleDbDataAdapter(strSQL, conn)
            Dim ds As New DataSet
            da.Fill(ds, "Users")

            With cboCustodian
                .DataSource = ds.Tables("Users")
                .DisplayMember = "Custodian7"
                .ValueMember = "Custodian7"
                .SelectedIndex = 0
            End With

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "ERROR")
        End Try
    End Sub

    Private Sub ckbExtFirm_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ckbExtFirm.CheckedChanged
        If ckbExtFirm.Checked Then
            cboExtFirm.Enabled = True
            Call LoadExtFirmList()
        Else
            cboExtFirm.Enabled = False
        End If
    End Sub

    Public Sub LoadExtFirmList()
        Try

            Dim conn As New OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & My.Settings.dbloc & "';Jet OLEDB:Database Password='" & My.Settings.dbpass & "';")

            Dim strSQL As String = "SELECT DISTINCT dbo_vQbRowDefPortfolio.ExternalFirm" & _
            " FROM(dbo_vQbRowDefPortfolio)" & _
            " WHERE(((dbo_vQbRowDefPortfolio.ExternalFirm) Is Not Null))" & _
            " GROUP BY dbo_vQbRowDefPortfolio.ExternalFirm" & _
            " ORDER BY dbo_vQbRowDefPortfolio.ExternalFirm"

            Dim da As New OleDb.OleDbDataAdapter(strSQL, conn)
            Dim ds As New DataSet
            da.Fill(ds, "Users")

            With cboExtFirm
                .DataSource = ds.Tables("Users")
                .DisplayMember = "ExternalFirm"
                .ValueMember = "ExternalFirm"
                .SelectedIndex = 0
            End With

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "ERROR")
        End Try
    End Sub

    Private Sub ckbExtRep_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ckbExtRep.CheckedChanged
        If ckbExtRep.Checked Then
            cboExtRep.Enabled = True
            Call LoadExtRepList()
        Else
            cboExtRep.Enabled = False
        End If
    End Sub

    Public Sub LoadExtRepList()
        Try

            Dim conn As New OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & My.Settings.dbloc & "';Jet OLEDB:Database Password='" & My.Settings.dbpass & "';")

            Dim strSQL As String = "SELECT DISTINCT dbo_vQbRowDefPortfolio.ExternalAdvisor" & _
            " FROM(dbo_vQbRowDefPortfolio)" & _
            " WHERE(((dbo_vQbRowDefPortfolio.ExternalAdvisor) Is Not Null))" & _
            " GROUP BY dbo_vQbRowDefPortfolio.ExternalAdvisor" & _
            " ORDER BY dbo_vQbRowDefPortfolio.ExternalAdvisor"


            Dim da As New OleDb.OleDbDataAdapter(strSQL, conn)
            Dim ds As New DataSet
            da.Fill(ds, "Users")

            With cboExtRep
                .DataSource = ds.Tables("Users")
                .DisplayMember = "ExternalAdvisor"
                .ValueMember = "ExternalAdvisor"
                .SelectedIndex = 0
            End With

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "ERROR")
        End Try
    End Sub

    Private Sub ckbRegion_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ckbRegion.CheckedChanged
        If ckbRegion.Checked Then
            cboRegion.Enabled = True
            Call LoadRegionList()
        Else
            cboRegion.Enabled = False
        End If
    End Sub

    Public Sub LoadRegionList()
        Try

            Dim conn As New OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & My.Settings.dbloc & "';Jet OLEDB:Database Password='" & My.Settings.dbpass & "';")

            Dim strSQL As String = "SELECT DISTINCT dbo_vQbRowDefPortfolio.SalesRegion" & _
            " FROM(dbo_vQbRowDefPortfolio)" & _
            " WHERE(((dbo_vQbRowDefPortfolio.SalesRegion) Is Not Null))" & _
            " GROUP BY dbo_vQbRowDefPortfolio.SalesRegion" & _
            " ORDER BY dbo_vQbRowDefPortfolio.SalesRegion"

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

    Private Sub ckbTerritory_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ckbTerritory.CheckedChanged
        If ckbTerritory.Checked Then
            cboTerritory.Enabled = True
            Call LoadTerritoryList()
        Else
            cboTerritory.Enabled = False
        End If
    End Sub

    Public Sub LoadTerritoryList()
        Try

            Dim conn As New OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & My.Settings.dbloc & "';Jet OLEDB:Database Password='" & My.Settings.dbpass & "';")

            Dim strSQL As String = "SELECT DISTINCT dbo_vQbRowDefPortfolio.AAMTeamName" & _
            " FROM(dbo_vQbRowDefPortfolio)" & _
            " WHERE(((dbo_vQbRowDefPortfolio.AAMTeamName) Is Not Null))" & _
            " GROUP BY dbo_vQbRowDefPortfolio.AAMTeamName" & _
            " ORDER BY dbo_vQbRowDefPortfolio.AAMTeamName"

            Dim da As New OleDb.OleDbDataAdapter(strSQL, conn)
            Dim ds As New DataSet
            da.Fill(ds, "Users")

            With cboTerritory
                .DataSource = ds.Tables("Users")
                .DisplayMember = "AAMTeamName"
                .ValueMember = "AAMTeamName"
                .SelectedIndex = 0
            End With

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "ERROR")
        End Try
    End Sub

    Private Sub ckbIntRep_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ckbIntRep.CheckedChanged
        If ckbIntRep.Checked Then
            cboIntRep.Enabled = True
            Call LoadIntRepList()
        Else
            cboIntRep.Enabled = False
        End If
    End Sub

    Public Sub LoadIntRepList()
        Try

            Dim conn As New OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & My.Settings.dbloc & "';Jet OLEDB:Database Password='" & My.Settings.dbpass & "';")

            Dim strSQL As String = "SELECT DISTINCT dbo_vQbRowDefPortfolio.AAMRepName" & _
            " FROM(dbo_vQbRowDefPortfolio)" & _
            " WHERE(((dbo_vQbRowDefPortfolio.AAMRepName) Is Not Null))" & _
            " GROUP BY dbo_vQbRowDefPortfolio.AAMRepName" & _
            " ORDER BY dbo_vQbRowDefPortfolio.AAMRepName;"


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

    Private Sub ckbAdvRole_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ckbAdvRole.CheckedChanged
        If ckbAdvRole.Checked Then
            cboAdvRole.Enabled = True
            Call LoadAdvRoleList()
        Else
            cboAdvRole.Enabled = False
        End If
    End Sub

    Public Sub LoadAdvRoleList()
        Try

            Dim conn As New OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & My.Settings.dbloc & "';Jet OLEDB:Database Password='" & My.Settings.dbpass & "';")

            Dim strSQL As String = "SELECT DISTINCT dbo_vQbRowDefPortfolio.AdvisorRoleType5" & _
            " FROM(dbo_vQbRowDefPortfolio)" & _
            " WHERE(((dbo_vQbRowDefPortfolio.AdvisorRoleType5) Is Not Null))" & _
            " GROUP BY dbo_vQbRowDefPortfolio.AdvisorRoleType5" & _
            " ORDER BY dbo_vQbRowDefPortfolio.AdvisorRoleType5"


            Dim da As New OleDb.OleDbDataAdapter(strSQL, conn)
            Dim ds As New DataSet
            da.Fill(ds, "Users")

            With cboAdvRole
                .DataSource = ds.Tables("Users")
                .DisplayMember = "AdvisorRoleType5"
                .ValueMember = "AdvisorRoleType5"
                .SelectedIndex = 0
            End With

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "ERROR")
        End Try
    End Sub

    Private Sub ckbStatus_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ckbStatus.CheckedChanged
        If ckbStatus.Checked Then
            cboStatus.Enabled = True
            Call LoadStatusList()
        Else
            cboStatus.Enabled = False
        End If
    End Sub

    Public Sub LoadStatusList()
        Try

            Dim conn As New OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & My.Settings.dbloc & "';Jet OLEDB:Database Password='" & My.Settings.dbpass & "';")

            Dim strSQL As String = "SELECT DISTINCT dbo_vQbRowDefPortfolio.PortfolioStatus" & _
            " FROM(dbo_vQbRowDefPortfolio)" & _
            " WHERE(((dbo_vQbRowDefPortfolio.PortfolioStatus) Is Not Null))" & _
            " GROUP BY dbo_vQbRowDefPortfolio.PortfolioStatus" & _
            " ORDER BY dbo_vQbRowDefPortfolio.PortfolioStatus"

            Dim da As New OleDb.OleDbDataAdapter(strSQL, conn)
            Dim ds As New DataSet
            da.Fill(ds, "Users")

            With cboStatus
                .DataSource = ds.Tables("Users")
                .DisplayMember = "PortfolioStatus"
                .ValueMember = "PortfolioStatus"
                .SelectedIndex = 0
            End With

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "ERROR")
        End Try
    End Sub

    Private Sub ckbPlatform_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ckbPlatform.CheckedChanged
        If ckbPlatform.Checked Then
            cboPlatform.Enabled = True
            Call LoadPlatformList()
        Else
            cboPlatform.Enabled = False
        End If
    End Sub

    Public Sub LoadPlatformList()
        Try

            Dim conn As New OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & My.Settings.dbloc & "';Jet OLEDB:Database Password='" & My.Settings.dbpass & "';")

            Dim strSQL As String = "SELECT DISTINCT dbo_vQbRowDefPortfolio.Platform" & _
            " FROM(dbo_vQbRowDefPortfolio)" & _
            " WHERE(((dbo_vQbRowDefPortfolio.Platform) Is Not Null))" & _
            " GROUP BY dbo_vQbRowDefPortfolio.Platform" & _
            " ORDER BY dbo_vQbRowDefPortfolio.Platform"

            Dim da As New OleDb.OleDbDataAdapter(strSQL, conn)
            Dim ds As New DataSet
            da.Fill(ds, "Users")

            With cboPlatform
                .DataSource = ds.Tables("Users")
                .DisplayMember = "Platform"
                .ValueMember = "Platform"
                .SelectedIndex = 0
            End With

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "ERROR")
        End Try
    End Sub

    Private Sub ckbBillType_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ckbBillType.CheckedChanged
        If ckbBillType.Checked Then
            cboBillType.Enabled = True
            'Items in this combo box are pre-populated in the form.
        Else
            cboBillType.Enabled = False
        End If
    End Sub

    Private Sub ckbManager_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ckbManager.CheckedChanged
        If ckbManager.Checked Then
            cboManager.Enabled = True
            Call LoadManagerList()
        Else
            cboManager.Enabled = False
        End If
    End Sub

    Public Sub LoadManagerList()
        Try

            Dim conn As New OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & My.Settings.dbloc & "';Jet OLEDB:Database Password='" & My.Settings.dbpass & "';")

            Dim strSQL As String = "SELECT DISTINCT dbo_vQbRowDefPortfolio.ManagerCode" & _
            " FROM(dbo_vQbRowDefPortfolio)" & _
            " WHERE(((dbo_vQbRowDefPortfolio.ManagerCode) Is Not Null))" & _
            " GROUP BY dbo_vQbRowDefPortfolio.ManagerCode" & _
            " ORDER BY dbo_vQbRowDefPortfolio.ManagerCode"


            Dim da As New OleDb.OleDbDataAdapter(strSQL, conn)
            Dim ds As New DataSet
            da.Fill(ds, "Users")

            With cboManager
                .DataSource = ds.Tables("Users")
                .DisplayMember = "ManagerCode"
                .ValueMember = "ManagerCode"
                .SelectedIndex = 0
            End With

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "ERROR")
        End Try
    End Sub

    Private Sub ckbDiscipline_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ckbDiscipline.CheckedChanged
        If ckbDiscipline.Checked Then
            cboDiscipline.Enabled = True
            Call LoadDisciplineList()
        Else
            cboDiscipline.Enabled = False
        End If
    End Sub

    Public Sub LoadDisciplineList()
        Try

            Dim conn As New OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & My.Settings.dbloc & "';Jet OLEDB:Database Password='" & My.Settings.dbpass & "';")

            Dim strSQL As String = "SELECT DISTINCT dbo_vQbRowDefPortfolio.Discipline6" & _
            " FROM(dbo_vQbRowDefPortfolio)" & _
            " WHERE(((dbo_vQbRowDefPortfolio.Discipline6) Is Not Null))" & _
            " GROUP BY dbo_vQbRowDefPortfolio.Discipline6" & _
            " ORDER BY dbo_vQbRowDefPortfolio.Discipline6"



            Dim da As New OleDb.OleDbDataAdapter(strSQL, conn)
            Dim ds As New DataSet
            da.Fill(ds, "Users")

            With cboDiscipline
                .DataSource = ds.Tables("Users")
                .DisplayMember = "Discipline6"
                .ValueMember = "Discipline6"
                .SelectedIndex = 0
            End With

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "ERROR")
        End Try
    End Sub

    Private Sub Button4_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        Call LoadVerifyBilling()
    End Sub

    Public Sub LoadVerifyBilling()
        Try
            Dim conn As New OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & My.Settings.dbloc & "';Jet OLEDB:Database Password='" & My.Settings.dbpass & "';")

            Dim strSQL As String

            strSQL = "SELECT mdb_BillingBase.ID, mdb_BillingBase.PortfolioID, dbo_vQbRowDefPortfolio.PortfolioCode As [Portfolio Code], dbo_vQbRowDefPortfolio.ReportHeading1 As [Account Name], mdb_BillingBase.AAMRate As [Internal Rate], mdb_BillingBase.SolicitorFirm As [Solicitor Firm Rate], mdb_BillingBase.SolicitorRep As [Solicitor Rep Fee], mdb_BillingBase.DaysBillable As [Days Billable], mdb_BillingBase.DateOfValue As [Date of Valuation], mdb_BillingBase.BillableMarketValue As [Billable Market Value], mdb_BillingBase.AAMValue As [Internal Billed Amount], mdb_BillingBase.SolicitorFirmValue As [Solicitor Firm Billed Amount], mdb_BillingBase.SolicitorRepValue As [Solicitor Rep Billed Amount], dbo_vQbRowDefPortfolio.StartDate As [Portfolio Start Date], dbo_vQbRowDefPortfolio.CloseDate As [Portfolio Close Date], dbo_vQbRowDefPortfolio.Discipline6 As [Discipline], dbo_vQbRowDefPortfolio.Custodian7 As [Custodian], dbo_vQbRowDefPortfolio.SalesDepartment As [Sales Department], dbo_vQbRowDefPortfolio.SalesRegion As [Sales Region], dbo_vQbRowDefPortfolio.AAMTeamName As [Territory], dbo_vQbRowDefPortfolio.AAMRepName As [Internal Rep], dbo_vQbRowDefPortfolio.Platform, dbo_vQbRowDefPortfolio.ExternalAdvisor As [External Advisor], dbo_vQbRowDefPortfolio.ExternalFirm As [External Firm], dbo_vQbRowDefPortfolio.BillingValuationPeriod As [Valuation Period], dbo_vQbRowDefPortfolio.CollectionSource As [Collection Source], mdb_BillingPeriods.PeriodStart As [Period Start], mdb_BillingPeriods.PeriodEnd As [Period End], IIF(dbo_vQbRowDefPortfolio.BillInArrears=1,'Arrears','Advanced') As [Type of Billing], mdb_BillingBase.Locked, mdb_BillingBase.Verified, mdb_BillingBase.Adjusted" & _
            " FROM (mdb_BillingBase INNER JOIN dbo_vQbRowDefPortfolio ON mdb_BillingBase.PortfolioID = dbo_vQbRowDefPortfolio.PortfolioID) INNER JOIN mdb_BillingPeriods ON mdb_BillingBase.PID = mdb_BillingPeriods.ID" & _
            " WHERE mdb_BillingPeriods.PeriodStart = #" & DateTimePicker3.Value & "# AND mdb_BillingPeriods.PeriodEnd = #" & DateTimePicker4.Value & "#"

            If ckbCustodian.Checked Then
                strSQL = strSQL & " AND dbo_vQbRowDefPortfolio.Custodian7 = '" & cboCustodian.Text & "'"
            Else
            End If

            If ckbExtFirm.Checked Then
                strSQL = strSQL & " AND dbo_vQbRowDefPortfolio.ExternalFirm = '" & cboExtFirm.Text & "'"
            Else
            End If

            If ckbExtRep.Checked Then
                strSQL = strSQL & " AND dbo_vQbRowDefPortfolio.ExternalAdvisor = '" & cboExtRep.Text & "'"
            Else
            End If

            If ckbRegion.Checked Then
                strSQL = strSQL & " AND dbo_vQbRowDefPortfolio.SalesRegion = '" & cboRegion.Text & "'"
            Else
            End If

            If ckbTerritory.Checked Then
                strSQL = strSQL & " AND dbo_vQbRowDefPortfolio.AAMTeamName = '" & cboTerritory.Text & "'"
            Else
            End If

            If ckbIntRep.Checked Then
                strSQL = strSQL & " AND dbo_vQbRowDefPortfolio.AAMRepName = '" & cboIntRep.Text & "'"
            Else
            End If

            If ckbAdvRole.Checked Then
                strSQL = strSQL & " AND dbo_vQbRowDefPortfolio.AdvisorRoleType5 = '" & cboAdvRole.Text & "'"
            Else
            End If

            If ckbStatus.Checked Then
                strSQL = strSQL & " AND dbo_vQbRowDefPortfolio.PortfolioStatus = '" & cboStatus.Text & "'"
            Else
            End If

            If ckbPlatform.Checked Then
                strSQL = strSQL & " AND dbo_vQbRowDefPortfolio.Platform = '" & cboPlatform.Text & "'"
            Else
            End If

            If ckbManager.Checked Then
                strSQL = strSQL & " AND dbo_vQbRowDefPortfolio.ManagerCode = '" & cboManager.Text & "'"
            Else
            End If

            If ckbDiscipline.Checked Then
                strSQL = strSQL & " AND dbo_vQbRowDefPortfolio.Discipline6 = '" & cboDiscipline.Text & "'"
            Else
            End If

            If ckbBillType.Checked Then
                If cboBillType.Text = "Advanced" Then
                    strSQL = strSQL & " AND dbo_vQbRowDefPortfolio.BillInArrears = 0"
                Else
                    strSQL = strSQL & " AND dbo_vQbRowDefPortfolio.BillInArrears = 1"
                End If
            Else
            End If

            If ckbLocked.Checked Then
                strSQL = strSQL & " AND mdb_BillingBase.Locked = True"
            Else
            End If

            If ckbValidated.Checked Then
                strSQL = strSQL & " AND mdb_BillingBase.Verified = True"
            Else
            End If

            If ckbPortfolio.Checked Then
                strSQL = strSQL & " AND dbo_vQbRowDefPortfolio.PortfolioCode Like '%" & txtPortfolio.Text & "%'"
            Else
            End If

            If RadioButton5.Checked Then
                strSQL = strSQL & " AND mdb_BillingPeriods.TypeID = 1"
                Button6.Enabled = True
                Button7.Enabled = True
                Button10.Enabled = True
                'Button11.Enabled = True
            Else
                strSQL = strSQL & " AND mdb_BillingPeriods.TypeID = 2"
                Button6.Enabled = False
                Button7.Enabled = False
                Button10.Enabled = False
                'Button11.Enabled = False
            End If

            If ckbAdjusted.Checked Then
                strSQL = strSQL & " AND mdb_BillingBase.Adjusted = True"
            Else
            End If

            Dim da As New OleDb.OleDbDataAdapter(strSQL, conn)
            Dim ds As New DataSet
            da.Fill(ds, "Users")

            With DataGridView6
                .DataSource = ds.Tables("Users")
                .Columns(0).Visible = False
                .Columns(1).Visible = False
                .Columns(9).DefaultCellStyle.Format = "c"
                .Columns(10).DefaultCellStyle.Format = "c"
                .Columns(11).DefaultCellStyle.Format = "c"
                .Columns(12).DefaultCellStyle.Format = "c"
                .Columns(9).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                .Columns(10).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                .Columns(11).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                .Columns(12).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
            End With

            TextBox2.Text = DataGridView6.RowCount

            If DataGridView6.RowCount > 0 Then
                Dim taum As Integer = 0
                Dim ival As Integer = 0

                For index As Integer = 0 To DataGridView6.RowCount - 1
                    taum += Convert.ToInt32(DataGridView6.Rows(index).Cells(9).Value)
                    ival += Convert.ToInt32(DataGridView6.Rows(index).Cells(10).Value)
                Next

                TextBox3.Text = Format(taum, "$#,###.##")
                TextBox4.Text = Format(ival, "$#,###.##")
                TextBox5.Text = Format((((ival * 4) / taum) * 100), ".##")
            Else
                TextBox3.Text = "0.00"
                TextBox4.Text = "0.00"
                TextBox5.Text = ".00"
            End If

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "ERROR")
        End Try
    End Sub

    Private Sub Button5_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button5.Click
        ckbLocked.Checked = False
        ckbValidated.Checked = False
        ckbCustodian.Checked = False
        ckbExtFirm.Checked = False
        ckbExtRep.Checked = False
        ckbRegion.Checked = False
        ckbTerritory.Checked = False
        ckbIntRep.Checked = False
        ckbAdvRole.Checked = False
        ckbStatus.Checked = False
        ckbPlatform.Checked = False
        ckbBillType.Checked = False
        ckbManager.Checked = False
        ckbDiscipline.Checked = False
        ckbPortfolio.Checked = False
        txtPortfolio.Text = ""
    End Sub

    Private Sub Button6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button6.Click
        Dim ir As MsgBoxResult
        ir = MsgBox("Are you sure you want to Lock all records?", MsgBoxStyle.YesNo + MsgBoxStyle.Information, "Are you sure?")
        If ir = MsgBoxResult.Yes Then
            Try
                Dim Mycn As OleDb.OleDbConnection
                Dim Command As OleDb.OleDbCommand
                Dim SQLstr As String

                Mycn = New OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & My.Settings.dbloc & "';Jet OLEDB:Database Password='" & My.Settings.dbpass & "';")

                ProgressBar3.Value = 0
                ProgressBar3.Maximum = DataGridView6.RowCount

                If DataGridView6.RowCount > 1 Then
                    For index As Integer = 0 To DataGridView6.RowCount - 1
                        Mycn.Open()

                        Label36.Text = "Locking Portfolio: " & DataGridView6.Rows(index).Cells(2).Value
                        SQLstr = "UPDATE mdb_BillingBase SET Locked = TRUE, LockBy = " & My.Settings.userid & ", LockDateStamp = Format(Now())" & _
                        " WHERE Locked = False AND ID = " & DataGridView6.Rows(index).Cells(0).Value
                        Command = New OleDb.OleDbCommand(SQLstr, Mycn)
                        Command.ExecuteNonQuery()
                        Mycn.Close()
                        Pause(0.1)
                        ProgressBar3.Value = ProgressBar3.Value + 1
                    Next

                Else
                    
                End If

                Label36.Text = "Refreshing Grid"
                Call LoadVerifyBilling()
                Label36.Text = "Done!"

            Catch ex As Exception
                MsgBox(ex.Message, MsgBoxStyle.Critical, "ERROR")
            End Try
        Else

        End If
    End Sub

    Private Sub Button7_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button7.Click
        Dim ir As MsgBoxResult
        ir = MsgBox("Are you sure you want to Validate all records?" & vbNewLine & "This means you have verified the values and the accounts are ready to be billed." & vbNewLine & "NOTE: Only accounts that have been Locked will be validated!", MsgBoxStyle.YesNo + MsgBoxStyle.Information, "Are you sure?")
        If ir = MsgBoxResult.Yes Then
            Try
                Dim Mycn As OleDb.OleDbConnection
                Dim Command As OleDb.OleDbCommand
                Dim SQLstr As String

                Mycn = New OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & My.Settings.dbloc & "';Jet OLEDB:Database Password='" & My.Settings.dbpass & "';")

                If DataGridView6.RowCount > 1 Then
                    ProgressBar3.Value = 0
                    ProgressBar3.Maximum = DataGridView6.RowCount

                    For index As Integer = 0 To DataGridView6.RowCount - 1
                        Mycn.Open()
                        Label36.Text = "Validating Portfolio: " & DataGridView6.Rows(index).Cells(2).Value
                        SQLstr = "UPDATE mdb_BillingBase SET Verified = TRUE, VerifyBy = " & My.Settings.userid & ", VerifyDateStamp = Format(Now())" & _
                        " WHERE Locked = True AND Verified = False AND ID = " & DataGridView6.Rows(index).Cells(0).Value
                        Command = New OleDb.OleDbCommand(SQLstr, Mycn)
                        Command.ExecuteNonQuery()
                        Mycn.Close()
                        Pause(0.1)
                        ProgressBar3.Value = ProgressBar3.Value + 1
                    Next

                Else

                End If

                Label36.Text = "Refreshing Grid"
                Call LoadVerifyBilling()
                Label36.Text = "Done!"

            Catch ex As Exception
                MsgBox(ex.Message, MsgBoxStyle.Critical, "ERROR")
            End Try
        Else

        End If
    End Sub

    Private Sub ckbPortfolio_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ckbPortfolio.CheckedChanged
        If ckbPortfolio.Checked Then
            txtPortfolio.Enabled = True
        Else
            txtPortfolio.Enabled = False
        End If
    End Sub

    Private Sub Button8_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        If DataGridView6.RowCount = 0 Then
            GoTo line1
        Else
            GoTo line2
        End If

line2:
        Dim wapp As Microsoft.Office.Interop.Excel.Application
        Dim wsheet As Microsoft.Office.Interop.Excel.Worksheet
        Dim wbook As Microsoft.Office.Interop.Excel.Workbook
        wapp = New Microsoft.Office.Interop.Excel.Application
        wapp.Visible = False
        wbook = wapp.Workbooks.Add()
        wsheet = wbook.ActiveSheet

        Try
            Dim iX As Integer
            Dim iY As Integer
            Dim iC As Integer
            Dim CC As Integer = DataGridView6.Columns.Count

            ProgressBar3.Value = 0
            ProgressBar3.Maximum = DataGridView6.Columns.Count

            Label36.Text = "Setting up Excel Columns"

            For iC = 0 To DataGridView6.Columns.Count - 1
                wsheet.Cells(1, iC + 1).Value = DataGridView6.Columns(iC).HeaderText
                wsheet.Cells(1, iC + 1).font.bold = True
                wsheet.Cells(1, iC + 1).horizontalalignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
                wsheet.Rows(1).autofit()
                Pause(0.1)
                ProgressBar3.Value = ProgressBar3.Value + 1
            Next

            ProgressBar3.Value = 0

            Dim clmn As Integer
            clmn = DataGridView6.Columns.Count

            Dim rws As Integer
            rws = DataGridView6.Rows.Count

            ProgressBar3.Maximum = (clmn * rws)

            Label36.Text = "Sending data to Excel"

            For iX = 0 To DataGridView6.Rows.Count - 1
                For iY = 0 To DataGridView6.Columns.Count - 1
                    wsheet.Cells(iX + 2, iY + 1).value = DataGridView6(iY, iX).Value.ToString
                    wsheet.Cells(iX + 2, iY + 1).horizontalalignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignLeft
                    Pause(0.1)
                    ProgressBar3.Value = ProgressBar3.Value + 1
                    Label36.Text = "Left to Import: " & ProgressBar3.Maximum - ProgressBar3.Value
                Next
            Next

            'MsgBox("cnt Value: " & cnt & " - Rows: " & DataGridView6.Rows.Count & " - Columns: " & DataGridView6.Columns.Count)

            Label36.Text = "Starting Excel"
            wapp.Visible = True
            'wapp.SaveWorkspace("C:\test.txt")
            Label36.Text = "DONE!"

        Catch ex As Exception
            MsgBox(ex.Message)
        End Try

        wapp.UserControl = True
        GoTo line3

line1:
        MsgBox("Cannot process request.  Please check data.", MsgBoxStyle.Information, "Invalid Results.")

line3:
    End Sub

    Private Sub Button12_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button12.Click
        Dim path As String
        path = "C:\_BillingValues"

        If (System.IO.Directory.Exists(path)) Then
            'System.IO.Directory.Delete(path, True)
        Else
            System.IO.Directory.CreateDirectory(path)
        End If

        Dim dte As String
        dte = Format(Now(), "yyyyMMdd")

        Dim tme As String
        tme = Format(Now(), "HHmmss")

        Dim fname As String
        fname = "billing_" & dte & "_" & tme & ".csv"

        Dim AccessConn As New System.Data.OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & My.Settings.dbloc & "';Jet OLEDB:Database Password='" & My.Settings.dbpass & "';")

        AccessConn.Open()

        Dim strSQL As String

        'strSQL = "SELECT mdb_BillingBase.ID, mdb_BillingBase.PortfolioID, dbo_vQbRowDefPortfolio.PortfolioCode As [Portfolio Code], dbo_vQbRowDefPortfolio.ReportHeading1 As [Account Name], mdb_BillingBase.AAMRate As [Internal Rate], mdb_BillingBase.SolicitorFirm As [Solicitor Firm Rate], mdb_BillingBase.SolicitorRep As [Solicitor Rep Fee], mdb_BillingBase.DaysBillable As [Days Billable], mdb_BillingBase.DateOfValue As [Date of Valuation], mdb_BillingBase.BillableMarketValue As [Billable Market Value], mdb_BillingBase.AAMValue As [Internal Billed Amount], mdb_BillingBase.SolicitorFirmValue As [Solicitor Firm Billed Amount], mdb_BillingBase.SolicitorRepValue As [Solicitor Rep Billed Amount], dbo_vQbRowDefPortfolio.StartDate As [Portfolio Start Date], dbo_vQbRowDefPortfolio.CloseDate As [Portfolio Close Date], dbo_vQbRowDefPortfolio.Discipline6 As [Discipline], dbo_vQbRowDefPortfolio.Custodian7 As [Custodian], dbo_vQbRowDefPortfolio.SalesDepartment As [Sales Department], dbo_vQbRowDefPortfolio.SalesRegion As [Sales Region], dbo_vQbRowDefPortfolio.AAMTeamName As [Territory], dbo_vQbRowDefPortfolio.AAMRepName As [Internal Rep], dbo_vQbRowDefPortfolio.Platform, dbo_vQbRowDefPortfolio.ExternalAdvisor As [External Advisor], dbo_vQbRowDefPortfolio.ExternalFirm As [External Firm], dbo_vQbRowDefPortfolio.BillingValuationPeriod As [Valuation Period], dbo_vQbRowDefPortfolio.CollectionSource As [Collection Source], mdb_BillingPeriods.PeriodStart As [Period Start], mdb_BillingPeriods.PeriodEnd As [Period End], IIF(dbo_vQbRowDefPortfolio.BillInArrears=1,'Arrears','Advanced') As [Type of Billing], mdb_BillingBase.Locked, mdb_BillingBase.Verified, mdb_BillingBase.Adjusted" & _
        '" FROM (mdb_BillingBase INNER JOIN dbo_vQbRowDefPortfolio ON mdb_BillingBase.PortfolioID = dbo_vQbRowDefPortfolio.PortfolioID) INNER JOIN mdb_BillingPeriods ON mdb_BillingBase.PID = mdb_BillingPeriods.ID" & _
        strSQL = " WHERE mdb_BillingPeriods.PeriodStart = #" & DateTimePicker3.Value & "# AND mdb_BillingPeriods.PeriodEnd = #" & DateTimePicker4.Value & "#"

        If ckbCustodian.Checked Then
            strSQL = strSQL & " AND dbo_vQbRowDefPortfolio.Custodian7 = '" & cboCustodian.Text & "'"
        Else
        End If

        If ckbExtFirm.Checked Then
            strSQL = strSQL & " AND dbo_vQbRowDefPortfolio.ExternalFirm = '" & cboExtFirm.Text & "'"
        Else
        End If

        If ckbExtRep.Checked Then
            strSQL = strSQL & " AND dbo_vQbRowDefPortfolio.ExternalAdvisor = '" & cboExtRep.Text & "'"
        Else
        End If

        If ckbRegion.Checked Then
            strSQL = strSQL & " AND dbo_vQbRowDefPortfolio.SalesRegion = '" & cboRegion.Text & "'"
        Else
        End If

        If ckbTerritory.Checked Then
            strSQL = strSQL & " AND dbo_vQbRowDefPortfolio.AAMTeamName = '" & cboTerritory.Text & "'"
        Else
        End If

        If ckbIntRep.Checked Then
            strSQL = strSQL & " AND dbo_vQbRowDefPortfolio.AAMRepName = '" & cboIntRep.Text & "'"
        Else
        End If

        If ckbAdvRole.Checked Then
            strSQL = strSQL & " AND dbo_vQbRowDefPortfolio.AdvisorRoleType5 = '" & cboAdvRole.Text & "'"
        Else
        End If

        If ckbStatus.Checked Then
            strSQL = strSQL & " AND dbo_vQbRowDefPortfolio.PortfolioStatus = '" & cboStatus.Text & "'"
        Else
        End If

        If ckbPlatform.Checked Then
            strSQL = strSQL & " AND dbo_vQbRowDefPortfolio.Platform = '" & cboPlatform.Text & "'"
        Else
        End If

        If ckbManager.Checked Then
            strSQL = strSQL & " AND dbo_vQbRowDefPortfolio.ManagerCode = '" & cboManager.Text & "'"
        Else
        End If

        If ckbDiscipline.Checked Then
            strSQL = strSQL & " AND dbo_vQbRowDefPortfolio.Discipline6 = '" & cboDiscipline.Text & "'"
        Else
        End If

        If ckbBillType.Checked Then
            If cboBillType.Text = "Advanced" Then
                strSQL = strSQL & " AND dbo_vQbRowDefPortfolio.BillInArrears = 0"
            Else
                strSQL = strSQL & " AND dbo_vQbRowDefPortfolio.BillInArrears = 1"
            End If
        Else
        End If

        If ckbLocked.Checked Then
            strSQL = strSQL & " AND mdb_BillingBase.Locked = True"
        Else
        End If

        If ckbValidated.Checked Then
            strSQL = strSQL & " AND mdb_BillingBase.Verified = True"
        Else
        End If

        If ckbPortfolio.Checked Then
            strSQL = strSQL & " AND dbo_vQbRowDefPortfolio.PortfolioCode Like '%" & txtPortfolio.Text & "%'"
        Else
        End If

        If RadioButton5.Checked Then
            strSQL = strSQL & " AND mdb_BillingPeriods.TypeID = 1"
        Else
            strSQL = strSQL & " AND mdb_BillingPeriods.TypeID = 2"
        End If

        If ckbAdjusted.Checked Then
            strSQL = strSQL & " AND mdb_BillingBase.Adjusted = True"
        Else
        End If

        Dim AccessCommand As New OleDb.OleDbCommand("SELECT dbo_vQbRowDefPortfolio.PortfolioCode As [Portfolio Code], dbo_vQbRowDefPortfolio.ReportHeading1 As [Account Name], mdb_BillingBase.AAMRate As [Internal Rate], mdb_BillingBase.SolicitorFirm As [Solicitor Firm Rate], mdb_BillingBase.SolicitorRep As [Solicitor Rep Fee], mdb_BillingBase.DaysBillable As [Days Billable], mdb_BillingBase.DateOfValue As [Date of Valuation], mdb_BillingBase.BillableMarketValue As [Billable Market Value], mdb_BillingBase.AAMValue As [Internal Billed Amount], mdb_BillingBase.SolicitorFirmValue As [Solicitor Firm Billed Amount], mdb_BillingBase.SolicitorRepValue As [Solicitor Rep Billed Amount], dbo_vQbRowDefPortfolio.StartDate As [Portfolio Start Date], dbo_vQbRowDefPortfolio.CloseDate As [Portfolio Close Date], dbo_vQbRowDefPortfolio.Discipline6 As [Discipline], dbo_vQbRowDefPortfolio.Custodian7 As [Custodian], dbo_vQbRowDefPortfolio.SalesDepartment As [Sales Department], dbo_vQbRowDefPortfolio.SalesRegion As [Sales Region], dbo_vQbRowDefPortfolio.AAMTeamName As [Territory], dbo_vQbRowDefPortfolio.AAMRepName As [Internal Rep], dbo_vQbRowDefPortfolio.Platform, dbo_vQbRowDefPortfolio.ExternalAdvisor As [External Advisor], dbo_vQbRowDefPortfolio.ExternalFirm As [External Firm], dbo_vQbRowDefPortfolio.BillingValuationPeriod As [Valuation Period], dbo_vQbRowDefPortfolio.CollectionSource As [Collection Source], mdb_BillingPeriods.PeriodStart As [Period Start], mdb_BillingPeriods.PeriodEnd As [Period End], IIF(dbo_vQbRowDefPortfolio.BillInArrears=1,'Arrears','Advanced') As [Type of Billing], mdb_BillingBase.Locked, mdb_BillingBase.Verified, mdb_BillingBase.Adjusted INTO [Text;HDR=YES;DATABASE=" & path & "].[" & fname & "] FROM (mdb_BillingBase INNER JOIN dbo_vQbRowDefPortfolio ON mdb_BillingBase.PortfolioID = dbo_vQbRowDefPortfolio.PortfolioID) INNER JOIN mdb_BillingPeriods ON mdb_BillingBase.PID = mdb_BillingPeriods.ID" & strSQL, AccessConn)

        AccessCommand.ExecuteNonQuery()
        AccessConn.Close()

        Dim ir As MsgBoxResult
        Dim fullnme As String
        fullnme = path & "\" & fname

        ir = MsgBox("Export finished.  File saved as '" & fullnme & "." & vbNewLine & vbNewLine & "Would you like to open the file now?", MsgBoxStyle.YesNo + MsgBoxStyle.Information, "Done")

        If ir = MsgBoxResult.Yes Then

            Dim xlsApp As Microsoft.Office.Interop.Excel.Application
            Dim xlsWorkBook As Microsoft.Office.Interop.Excel.Workbook

            xlsApp = New Microsoft.Office.Interop.Excel.Application

            xlsWorkBook = xlsApp.Workbooks.Open(fullnme)
            xlsApp.Visible = True

        Else
        End If


    End Sub

    Private Sub Button9_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub

    Private Sub ProcessAdjustmentToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ProcessAdjustmentToolStripMenuItem.Click
        If DataGridView6.RowCount = 0 Then

        Else
            If DataGridView6.SelectedCells(29).Value = 0 Then
                MsgBox("Record must be locked before values can be adjusted.", MsgBoxStyle.Critical, "Record not locked.")
            Else
                If DataGridView6.SelectedCells(30).Value = -1 Then
                    MsgBox("Record has already been validated.  Validated records cannot be changed.", MsgBoxStyle.Critical, "Record already validated.")
                Else
                    If DataGridView6.SelectedCells(31).Value = -1 Then
                        Dim child As New RevenueCenter_Adjustments
                        child.MdiParent = Home
                        child.Show()
                        child.ID.Text = DataGridView6.SelectedCells(0).Value
                        child.PortfolioID.Text = DataGridView6.SelectedCells(1).Value
                        'child.AdjID.Text = "NEW"
                        Call child.PullID()
                        Call child.InitialLoadOldAdjustment()
                    Else
                        Dim child As New RevenueCenter_Adjustments
                        child.MdiParent = Home
                        child.Show()
                        child.ID.Text = DataGridView6.SelectedCells(0).Value
                        child.AdjID.Text = "NEW"
                        child.ckbActive.Checked = True
                        Call child.InitialLoadNewAdjustment()
                    End If
                End If
            End If

        End If
    End Sub

    Private Sub Button10_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button10.Click
        Dim path As String
        path = "C:\_BillingValues"

        If (System.IO.Directory.Exists(path)) Then
            'System.IO.Directory.Delete(path, True)
        Else
            System.IO.Directory.CreateDirectory(path)
        End If

        Dim dte As String
        dte = Format(Now(), "yyyyMMdd")

        Dim tme As String
        tme = Format(Now(), "HHmm")

        Dim fname As String
        fname = "billing_" & dte & "_" & tme & ".csv"

        Dim AccessConn As New System.Data.OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & My.Settings.dbloc & "';Jet OLEDB:Database Password='" & My.Settings.dbpass & "';")


        'Process Schwab Billing
        fname = "SchwabBilling_" & dte & "_" & tme & ".csv"
        AccessConn.Open()

        Dim AccessCommand As New OleDb.OleDbCommand("SELECT dbo_vQbRowDefPortfolio.PortfolioCode, Format(mdb_BillingBase.AAMValue, '0.00') INTO [Text;HDR=NO;DATABASE=" & path & "].[" & fname & "]" & _
        " FROM (mdb_BillingBase INNER JOIN dbo_vQbRowDefPortfolio ON mdb_BillingBase.PortfolioID = dbo_vQbRowDefPortfolio.PortfolioID) INNER JOIN mdb_BillingPeriods ON mdb_BillingBase.PID = mdb_BillingPeriods.ID" & _
        " WHERE(((dbo_vQbRowDefPortfolio.CollectionSource) = 'Custodian') And ((dbo_vQbRowDefPortfolio.Custodian7) = 'Schwab') And ((mdb_BillingBase.Verified) = True) And ((mdb_BillingPeriods.PeriodStart) = #" & DateTimePicker3.Value & "#) And ((mdb_BillingPeriods.PeriodEnd) = #" & DateTimePicker4.Value & "#))" & _
        " GROUP BY dbo_vQbRowDefPortfolio.PortfolioCode, mdb_BillingBase.AAMValue", AccessConn)

        AccessCommand.ExecuteNonQuery()
        AccessConn.Close()

        Dim schwabfullnme As String
        schwabfullnme = path & "\" & fname

        'Process TDA Billing
        fname = "TDABilling_" & dte & "_" & tme & ".csv"
        AccessConn.Open()

        Dim AccessCommand1 As New OleDb.OleDbCommand("SELECT dbo_vQbRowDefPortfolio.PortfolioCode, 'q', Format(mdb_BillingBase.AAMValue, '0.00') INTO [Text;HDR=NO;DATABASE=" & path & "].[" & fname & "]" & _
        " FROM (mdb_BillingBase INNER JOIN dbo_vQbRowDefPortfolio ON mdb_BillingBase.PortfolioID = dbo_vQbRowDefPortfolio.PortfolioID) INNER JOIN mdb_BillingPeriods ON mdb_BillingBase.PID = mdb_BillingPeriods.ID" & _
        " WHERE(((dbo_vQbRowDefPortfolio.CollectionSource) = 'Custodian') And ((dbo_vQbRowDefPortfolio.Custodian7) = 'TD_Ameritrade') And ((mdb_BillingBase.Verified) = True) And ((mdb_BillingPeriods.PeriodStart) = #" & DateTimePicker3.Value & "#) And ((mdb_BillingPeriods.PeriodEnd) = #" & DateTimePicker4.Value & "#))" & _
        " GROUP BY dbo_vQbRowDefPortfolio.PortfolioCode, mdb_BillingBase.AAMValue", AccessConn)

        AccessCommand1.ExecuteNonQuery()
        AccessConn.Close()

        Dim TDAfullnme As String
        TDAfullnme = path & "\" & fname

        'Process NFS Billing
        fname = "NFSBilling_" & dte & "_" & tme & ".txt"
        AccessConn.Open()

        Dim AccessCommand2 As New OleDb.OleDbCommand("SELECT dbo_vQbRowDefPortfolio.PortfolioCode, Format(mdb_BillingBase.AAMValue, '0.00') INTO [Text;HDR=NO;FMT=Delimited;DATABASE=" & path & "].[" & fname & "]" & _
        " FROM (mdb_BillingBase INNER JOIN dbo_vQbRowDefPortfolio ON mdb_BillingBase.PortfolioID = dbo_vQbRowDefPortfolio.PortfolioID) INNER JOIN mdb_BillingPeriods ON mdb_BillingBase.PID = mdb_BillingPeriods.ID" & _
        " WHERE(((dbo_vQbRowDefPortfolio.CollectionSource) = 'Custodian') And ((dbo_vQbRowDefPortfolio.Custodian7) = 'NFS') And ((mdb_BillingBase.Verified) = True) And ((mdb_BillingPeriods.PeriodStart) = #" & DateTimePicker3.Value & "#) And ((mdb_BillingPeriods.PeriodEnd) = #" & DateTimePicker4.Value & "#))" & _
        " GROUP BY dbo_vQbRowDefPortfolio.PortfolioCode, mdb_BillingBase.AAMValue", AccessConn)

        AccessCommand2.ExecuteNonQuery()
        AccessConn.Close()

        Dim NFSfullnme As String
        NFSfullnme = path & "\" & fname


        Dim ir As MsgBoxResult
        

        ir = MsgBox("Export finished." & vbNewLine & vbNewLine & "Would you like to open the files now?", MsgBoxStyle.YesNo + MsgBoxStyle.Information, "Done")

        If ir = MsgBoxResult.Yes Then

            Dim xlsApp As Microsoft.Office.Interop.Excel.Application
            Dim xlsWorkBook As Microsoft.Office.Interop.Excel.Workbook

            xlsApp = New Microsoft.Office.Interop.Excel.Application

            xlsWorkBook = xlsApp.Workbooks.Open(schwabfullnme)
            xlsApp.Visible = True

            xlsApp = New Microsoft.Office.Interop.Excel.Application

            xlsWorkBook = xlsApp.Workbooks.Open(TDAfullnme)
            xlsApp.Visible = True

            System.Diagnostics.Process.Start("Notepad.Exe", NFSfullnme)

        Else
        End If
    End Sub

    Private Sub ChangeValidationStatusToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ChangeValidationStatusToolStripMenuItem.Click
        If DataGridView6.RowCount = 0 Then
        Else
            If DataGridView6.SelectedCells(29).Value = 0 Then
                MsgBox("Record must be locked before the record can be validated.", MsgBoxStyle.Critical, "Record not locked.")
            Else
                Try
                    Dim Mycn As OleDb.OleDbConnection
                    Dim Command As OleDb.OleDbCommand
                    Dim SQLstr As String

                    Mycn = New OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & My.Settings.dbloc & "';Jet OLEDB:Database Password='" & My.Settings.dbpass & "';")

                    If DataGridView6.SelectedCells(30).Value = -1 Then
                        Mycn.Open()
                        Label36.Text = "Un-Validating Portfolio: " & DataGridView6.SelectedCells(2).Value
                        SQLstr = "UPDATE mdb_BillingBase SET Verified = FALSE, VerifyBy = Null, VerifyDateStamp = Null" & _
                        " WHERE ID = " & DataGridView6.SelectedCells(0).Value
                        Command = New OleDb.OleDbCommand(SQLstr, Mycn)
                        Command.ExecuteNonQuery()
                        Mycn.Close()
                        ProgressBar3.Value = ProgressBar3.Maximum
                        DataGridView6.SelectedCells(30).Value = 0
                    Else
                        Mycn.Open()
                        Label36.Text = "Validating Portfolio: " & DataGridView6.SelectedCells(2).Value
                        SQLstr = "UPDATE mdb_BillingBase SET Verified = TRUE, VerifyBy = " & My.Settings.userid & ", VerifyDateStamp = Format(Now())" & _
                        " WHERE ID = " & DataGridView6.SelectedCells(0).Value
                        Command = New OleDb.OleDbCommand(SQLstr, Mycn)
                        Command.ExecuteNonQuery()
                        Mycn.Close()
                        ProgressBar3.Value = ProgressBar3.Maximum
                        DataGridView6.SelectedCells(30).Value = -1
                    End If

                Catch ex As Exception
                    MsgBox(ex.Message)
                End Try
            End If
        End If
    End Sub

    Private Sub ChangeLockStatusToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ChangeLockStatusToolStripMenuItem.Click
        If DataGridView6.RowCount = 0 Then
        Else
            If DataGridView6.SelectedCells(30).Value = -1 Then
                MsgBox("Record has been validated.  You must unvalidate the record before removing the lock.", MsgBoxStyle.Critical, "Record Validated.")
            Else
                Try
                    Dim Mycn As OleDb.OleDbConnection
                    Dim Command As OleDb.OleDbCommand
                    Dim SQLstr As String

                    Mycn = New OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & My.Settings.dbloc & "';Jet OLEDB:Database Password='" & My.Settings.dbpass & "';")

                    If DataGridView6.SelectedCells(29).Value = -1 Then
                        Mycn.Open()
                        Label36.Text = "Un-Validating Portfolio: " & DataGridView6.SelectedCells(2).Value
                        SQLstr = "UPDATE mdb_BillingBase SET Locked = FALSE, LockBy = Null, LockDateStamp = Null" & _
                        " WHERE ID = " & DataGridView6.SelectedCells(0).Value
                        Command = New OleDb.OleDbCommand(SQLstr, Mycn)
                        Command.ExecuteNonQuery()
                        Mycn.Close()
                        ProgressBar3.Value = ProgressBar3.Maximum
                        DataGridView6.SelectedCells(29).Value = 0
                    Else
                        Mycn.Open()
                        Label36.Text = "Validating Portfolio: " & DataGridView6.SelectedCells(2).Value
                        SQLstr = "UPDATE mdb_BillingBase SET Locked = TRUE, LockBy = " & My.Settings.userid & ", LockDateStamp = Format(Now())" & _
                        " WHERE ID = " & DataGridView6.SelectedCells(0).Value
                        Command = New OleDb.OleDbCommand(SQLstr, Mycn)
                        Command.ExecuteNonQuery()
                        Mycn.Close()
                        ProgressBar3.Value = ProgressBar3.Maximum
                        DataGridView6.SelectedCells(29).Value = -1
                    End If

                Catch ex As Exception
                    MsgBox(ex.Message)
                End Try
            End If
        End If
    End Sub

    Private Sub Button9_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button9.Click
        Dim child As New RevenueCenter_Address
        child.MdiParent = Home
        child.Show()
        child.ID.Text = "NEW"
    End Sub

    Private Sub Button11_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button11.Click
        Dim child As New RevenueCenterViewFirms
        child.MdiParent = Home
        child.Show()
    End Sub
End Class