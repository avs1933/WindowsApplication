Imports System.IO
Imports Microsoft.Office.Interop

Public Class rfp_Report
    Dim thContacts As System.Threading.Thread
    Dim thEmployees As System.Threading.Thread
    Dim thStage As System.Threading.Thread
    Dim thReason As System.Threading.Thread
    Dim thInsert As System.Threading.Thread
    Dim thUpdate As System.Threading.Thread

    Private Sub rfp_Report_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        If Button1.Enabled = False Then
            MsgBox("Please wait for the form to save before closing.", MsgBoxStyle.Information, "Saving.")
            e.Cancel = True
        Else
            e.Cancel = False
        End If
    End Sub

    Private Sub rfp_Report_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Control.CheckForIllegalCrossThreadCalls = False
        thContacts = New System.Threading.Thread(AddressOf LoadContacts)
        thContacts.Start()

        thEmployees = New System.Threading.Thread(AddressOf LoadEmployees)
        thEmployees.Start()

        thStage = New System.Threading.Thread(AddressOf LoadStage)
        thStage.Start()

        thReason = New System.Threading.Thread(AddressOf LoadReason)
        thReason.Start()

        Call LoadDeliveryMethod()
        Call LoadDisciplines()
        Call pulldisclosure()

    End Sub

    Public Sub pulldisclosure()
        If ckbDisclosure.Checked Then
            Try
                Dim conn As New OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & My.Settings.dbloc & "';Jet OLEDB:Database Password='" & My.Settings.dbpass & "';")

                Dim strSQL As String = "SELECT * FROM rfp_Disclosures WHERE Default = -1"
                Dim queryString As String = String.Format(strSQL)
                Dim cmd As New OleDb.OleDbCommand(queryString, conn)
                Dim da As New OleDb.OleDbDataAdapter(cmd)
                Dim ds As New DataSet

                da.Fill(ds, "User")
                Dim dt As DataTable = ds.Tables("User")

                Dim row1 As DataRow = dt.Rows(0)


                If IsDBNull(row1("Disclosure")) Then
                    rtbDisclosure.Text = "UNKNOWN"
                    rtbDisclosureCopy.Text = "UNKNOWN"
                Else
                    rtbDisclosure.Text = (row1("Disclosure"))
                    rtbDisclosureCopy.Text = (row1("Disclosure"))
                End If

                txtDisclosureID.Text = (row1("ID"))

            Catch ex As Exception
                MsgBox(ex.Message, MsgBoxStyle.Critical, "ERROR")
            End Try
        Else
            If txtID.Text = "NEW" Then
            Else
                Try
                    Dim conn As New OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & My.Settings.dbloc & "';Jet OLEDB:Database Password='" & My.Settings.dbpass & "';")

                    Dim strSQL As String = "SELECT * FROM rfp_Disclosures WHERE Default = 0 and RFPID = " & txtID.Text
                    Dim queryString As String = String.Format(strSQL)
                    Dim cmd As New OleDb.OleDbCommand(queryString, conn)
                    Dim da As New OleDb.OleDbDataAdapter(cmd)
                    Dim ds As New DataSet

                    da.Fill(ds, "User")
                    Dim dt As DataTable = ds.Tables("User")
                    If dt.Rows.Count > 0 Then
                        Dim row1 As DataRow = dt.Rows(0)

                        If IsDBNull(row1("Disclosure")) Then
                            rtbDisclosure.Text = "UNKNOWN"
                            rtbDisclosureCopy.Text = "UNKNOWN"
                        Else
                            rtbDisclosure.Text = (row1("Disclosure"))
                            rtbDisclosureCopy.Text = (row1("Disclosure"))
                        End If

                        txtDisclosureID.Text = (row1("ID"))
                    Else
                        txtDisclosureID.Text = "NEW"
                    End If

                Catch ex As Exception
                    MsgBox(ex.Message, MsgBoxStyle.Critical, "ERROR")
                End Try
            End If
        End If
    End Sub

    Public Sub estrev()
        Dim tmv As Double
        If IsDBNull(txtAUM) Or txtAUM.Text = "" Then
            tmv = 0
        Else
            tmv = txtAUM.Text
        End If

        If ckbFlatRate.Checked = True Then
            Dim frate As Double
            If IsDBNull(txtFlatRate) Or txtFlatRate.Text = "" Then
                frate = 0
            Else
                Dim numericcheck As Boolean
                numericcheck = IsNumeric(txtFlatRate.Text)
                If numericcheck = True Then
                    frate = txtFlatRate.Text
                    Label42.Visible = False
                Else
                    Label42.Text = "Error - Rates must be entered as '0.00'"
                    Label42.Visible = True
                End If
                'frate = txtFlatRate.Text / 100
            End If

            Dim fvalue As Double
            fvalue = tmv * frate / 100
            Dim minfee As Double
            If IsDBNull(txtMinFee) Or txtMinFee.Text = "" Then
                minfee = 0
            Else
                minfee = txtMinFee.Text
            End If
            If fvalue < minfee Then
                lblMinFee.ForeColor = Color.Red
                lblMinFee.Text = "Calculated Fee of " & Format(fvalue, "$#,###.00") & " does not meet the minimum fee of " & Format(minfee, "$#,###.00") & "."
                lblMinFee.Visible = True
                fvalue = minfee
            Else
                lblMinFee.ForeColor = Color.Black
                lblMinFee.Text = "Calculated Fee of " & Format(fvalue, "$#,###.00") & " meets the minimum fee of " & Format(minfee, "$#,###.00") & "."
                lblMinFee.Visible = True
            End If
            txtEstRev.Text = Format(fvalue, "$#,###.00")
            txtVelocity.Text = Format(frate, "#.00")
        Else
            Dim f6rate As Double
            If IsDBNull(txtF6Rate) Or txtF6Rate.Text = "" Then
                f6rate = 0
            Else
                Dim numericcheck As Boolean
                numericcheck = IsNumeric(txtF6Rate.Text)
                If numericcheck = True Then
                    f6rate = txtF6Rate.Text
                    Label42.Visible = False
                Else
                    Label42.Text = "Error - Rates must be entered as '0.00'"
                    Label42.Visible = True
                End If
            End If
            '*******Change color of expected rate field************
            If txtF1To.Text = "" And txtF6Rate.Text = "" Then
                txtF6Rate.BackColor = Color.Red
                txtF6Rate.ForeColor = Color.White
                txtF1Rate.BackColor = Color.White
                txtF1Rate.ForeColor = Color.Black
            Else
                txtF6Rate.BackColor = Color.White
                txtF6Rate.ForeColor = Color.Black
                If txtF1To.Text <> "" And txtF1Rate.Text = "" Then
                    txtF1Rate.BackColor = Color.Red
                    txtF1Rate.ForeColor = Color.Black
                Else
                    txtF1Rate.BackColor = Color.White
                    txtF1Rate.ForeColor = Color.Black
                End If
            End If

            If txtF2To.Text = "" And txtF6Rate.Text = "" And txtF1To.Text <> "" Then
                txtF6Rate.BackColor = Color.Red
                txtF6Rate.ForeColor = Color.White
                txtF2Rate.BackColor = Color.White
                txtF2Rate.ForeColor = Color.Black
            Else
                txtF6Rate.BackColor = Color.White
                txtF6Rate.ForeColor = Color.Black
                If txtF2To.Text <> "" And txtF2Rate.Text = "" Then
                    txtF2Rate.BackColor = Color.Red
                    txtF2Rate.ForeColor = Color.Black
                Else
                    txtF2Rate.BackColor = Color.White
                    txtF2Rate.ForeColor = Color.Black
                End If
            End If

            If txtF3To.Text = "" And txtF6Rate.Text = "" And txtF1To.Text <> "" And txtF2To.Text <> "" Then
                txtF6Rate.BackColor = Color.Red
                txtF6Rate.ForeColor = Color.White
                txtF3Rate.BackColor = Color.White
                txtF3Rate.ForeColor = Color.Black
            Else
                txtF6Rate.BackColor = Color.White
                txtF6Rate.ForeColor = Color.Black
                If txtF3To.Text <> "" And txtF3Rate.Text = "" Then
                    txtF3Rate.BackColor = Color.Red
                    txtF3Rate.ForeColor = Color.Black
                Else
                    txtF3Rate.BackColor = Color.White
                    txtF3Rate.ForeColor = Color.Black
                End If
            End If

            If txtF4To.Text = "" And txtF6Rate.Text = "" And txtF1To.Text <> "" And txtF2To.Text <> "" And txtF3To.Text <> "" Then
                txtF6Rate.BackColor = Color.Red
                txtF6Rate.ForeColor = Color.White
                txtF4Rate.BackColor = Color.White
                txtF4Rate.ForeColor = Color.Black
            Else
                txtF6Rate.BackColor = Color.White
                txtF6Rate.ForeColor = Color.Black
                If txtF4To.Text <> "" And txtF4Rate.Text = "" Then
                    txtF4Rate.BackColor = Color.Red
                    txtF4Rate.ForeColor = Color.Black
                Else
                    txtF4Rate.BackColor = Color.White
                    txtF4Rate.ForeColor = Color.Black
                End If
            End If

            If txtF5To.Text = "" And txtF6Rate.Text = "" And txtF1To.Text <> "" And txtF2To.Text <> "" And txtF3To.Text <> "" And txtF4To.Text <> "" Then
                txtF6Rate.BackColor = Color.Red
                txtF6Rate.ForeColor = Color.White
                txtF5Rate.BackColor = Color.White
                txtF5Rate.ForeColor = Color.Black
            Else
                txtF6Rate.BackColor = Color.White
                txtF6Rate.ForeColor = Color.Black
                If txtF5To.Text <> "" And txtF5Rate.Text = "" Then
                    txtF5Rate.BackColor = Color.Red
                    txtF5Rate.ForeColor = Color.Black
                Else
                    txtF5Rate.BackColor = Color.White
                    txtF5Rate.ForeColor = Color.Black
                End If
            End If

            If txtF1To.Text = "" Then
                txtF6Rate.BackColor = Color.Red
                txtF6Rate.ForeColor = Color.White
            Else
                If txtF2To.Text = "" And txtF1To.Text <> "" Then
                    txtF6Rate.BackColor = Color.Red
                    txtF6Rate.ForeColor = Color.White
                Else
                    If txtF3To.Text = "" And txtF1To.Text <> "" And txtF2To.Text <> "" Then
                        txtF6Rate.BackColor = Color.Red
                        txtF6Rate.ForeColor = Color.White
                    Else
                        If txtF4To.Text = "" And txtF1To.Text <> "" And txtF2To.Text <> "" And txtF3To.Text <> "" Then
                            txtF6Rate.BackColor = Color.Red
                            txtF6Rate.ForeColor = Color.White
                        Else
                            If txtF5To.Text = "" And txtF1To.Text <> "" And txtF2To.Text <> "" And txtF3To.Text <> "" And txtF4To.Text <> "" Then
                                txtF6Rate.BackColor = Color.Red
                                txtF6Rate.ForeColor = Color.White
                            Else
                                If txtF6From.Text <> "" And txtF6Rate.Text = "" Then
                                    txtF6Rate.BackColor = Color.Red
                                    txtF6Rate.ForeColor = Color.White
                                Else
                                    txtF6Rate.BackColor = Color.White
                                    txtF6Rate.ForeColor = Color.Black
                                End If
                            End If
                        End If
                    End If
                End If
            End If

            If txtF6Rate.Text <> "" Then
                txtF6Rate.BackColor = Color.White
                txtF6Rate.ForeColor = Color.Black
            End If

            '*****Bracket 1*****
            Dim f1from As Double
            Dim f1to As Double
            Dim f1rate As Double
            Dim f1level As Double
            Dim f1above As Boolean
            Dim f1diff As Double
            Dim f1value As Double
            If IsDBNull(txtF1From) Or txtF1From.Text = "" Then
                f1from = 0
            Else
                f1from = txtF1From.Text
            End If
            If IsDBNull(txtF1To) Or txtF1To.Text = "" Then
                f1to = 0
            Else
                f1to = txtF1To.Text
            End If
            If IsDBNull(txtF1Rate) Or txtF1Rate.Text = "" Then
                f1rate = 0
            Else
                Dim numericcheck As Boolean
                numericcheck = IsNumeric(txtF1Rate.Text)
                If numericcheck = True Then
                    f1rate = txtF1Rate.Text
                    Label42.Visible = False
                Else
                    Label42.Text = "Error - Rates must be entered as '0.00'"
                    Label42.Visible = True
                End If

            End If
            f1level = f1to - f1from
            f1above = tmv > f1to
            f1diff = tmv - f1from
            If IsDBNull(txtF1To) Or txtF1To.Text = "" Then
                If txtF6Rate.Text = "" Then
                Else
                    f1value = tmv * f6rate / 100
                End If
            Else
                If f1diff > 0 And f1above = True Then
                    f1value = f1level * f1rate / 100
                    txtF1Value.Text = Format(f1value, "$#,###.00")
                Else
                    If f1diff > 0 And f1above = False Then
                        f1value = ((f1diff * f1rate) / 100)
                        txtF1Value.Text = Format(f1diff, "$#,###.00")
                    Else
                        f1value = 0
                    End If
                End If
            End If
            '*****Bracket 2*****
            Dim f2from As Double
            Dim f2to As Double
            Dim f2rate As Double
            Dim f2level As Double
            Dim f2above As Boolean
            Dim f2diff As Double
            Dim f2value As Double
            If IsDBNull(txtF2From) Or txtF2From.Text = "" Then
                f2from = 0
            Else
                f2from = txtF2From.Text
            End If
            If IsDBNull(txtF2To) Or txtF2To.Text = "" Then
                f2to = 0
            Else
                f2to = txtF2To.Text
            End If
            If IsDBNull(txtF2Rate) Or txtF2Rate.Text = "" Then
                f2rate = 0
            Else
                Dim numericcheck As Boolean
                numericcheck = IsNumeric(txtF2Rate.Text)
                If numericcheck = True Then
                    f2rate = txtF2Rate.Text
                    Label42.Visible = False
                Else
                    Label42.Text = "Error - Rates must be entered as '0.00'"
                    Label42.Visible = True
                End If

            End If
            f2level = f2to - f2from
            f2above = tmv > f2to
            f2diff = tmv - f2from
            If ((IsDBNull(txtF2To) Or txtF2To.Text = "") And txtF1To.Text <> "") Then
                If txtF6Rate.Text = "" Then
                Else
                    f2value = f2diff * f6rate / 100
                End If
            Else
                If f2diff > 0 And f2above = True Then
                    f2value = f2level * f2rate / 100
                    txtF2Value.Text = Format(f2value, "$#,###.00")
                Else
                    If f2diff > 0 And f2above = False Then
                        f2value = ((f2diff * f2rate) / 100)
                        txtF2Value.Text = Format(f2value, "$#,###.00")
                    Else
                        f2value = 0
                        txtF2Value.Text = Format(f2value, "$#,###.00")
                    End If
                End If
            End If
            '*****Bracket 3*****
            Dim f3from As Double
            Dim f3to As Double
            Dim f3rate As Double
            Dim f3level As Double
            Dim f3above As Boolean
            Dim f3diff As Double
            Dim f3value As Double
            If IsDBNull(txtF3From) Or txtF3From.Text = "" Then
                f3from = 0
            Else
                f3from = txtF3From.Text
            End If
            If IsDBNull(txtF3To) Or txtF3To.Text = "" Then
                f3to = 0
            Else
                f3to = txtF3To.Text
            End If
            If IsDBNull(txtF3Rate) Or txtF3Rate.Text = "" Then
                f3rate = 0
            Else
                Dim numericcheck As Boolean
                numericcheck = IsNumeric(txtF3Rate.Text)
                If numericcheck = True Then
                    f3rate = txtF3Rate.Text
                    Label42.Visible = False
                Else
                    Label42.Text = "Error - Rates must be entered as '0.00'"
                    Label42.Visible = True
                End If
            End If
            f3level = f3to - f3from
            f3above = tmv > f3to
            f3diff = tmv - f3from
            If ((IsDBNull(txtF3To) Or txtF3To.Text = "") And (txtF1To.Text <> "" And txtF2To.Text <> "")) Then
                If txtF6Rate.Text = "" Then
                Else
                    f3value = f3diff * f6rate / 100
                End If
            Else
                If f3diff > 0 And f3above = True Then
                    f3value = f3level * f3rate / 100
                    txtF3Value.Text = Format(f3value, "$#,###.00")
                Else
                    If f3diff > 0 And f3above = False Then
                        f3value = ((f3diff * f3rate) / 100)
                        txtF3Value.Text = Format(f3value, "$#,###.00")
                    Else
                        f3value = 0
                        txtF3Value.Text = Format(f3value, "$#,###.00")
                    End If
                End If
            End If
            '*****Bracket 4*****
            Dim f4from As Double
            Dim f4to As Double
            Dim f4rate As Double
            Dim f4level As Double
            Dim f4above As Boolean
            Dim f4diff As Double
            Dim f4value As Double
            If IsDBNull(txtF4From) Or txtF4From.Text = "" Then
                f4from = 0
            Else
                f4from = txtF4From.Text
            End If
            If IsDBNull(txtF4To) Or txtF4To.Text = "" Then
                f4to = 0
            Else
                f4to = txtF4To.Text
            End If
            If IsDBNull(txtF4Rate) Or txtF4Rate.Text = "" Then
                f4rate = 0
            Else
                Dim numericcheck As Boolean
                numericcheck = IsNumeric(txtF4Rate.Text)
                If numericcheck = True Then
                    f4rate = txtF4Rate.Text
                    Label42.Visible = False
                Else
                    Label42.Text = "Error - Rates must be entered as '0.00'"
                    Label42.Visible = True
                End If
            End If
            f4level = f4to - f4from
            f4above = tmv > f4to
            f4diff = tmv - f4from
            If ((IsDBNull(txtF4To) Or txtF4To.Text = "") And (txtF1To.Text <> "" And txtF2To.Text <> "" And txtF3To.Text <> "")) Then
                If txtF6Rate.Text = "" Then
                Else
                    f4value = f4diff * f6rate / 100
                End If
            Else
                If f4diff > 0 And f4above = True Then
                    f4value = f4level * f4rate / 100
                    txtF4Value.Text = Format(f4value, "$#,###.00")
                Else
                    If f4diff > 0 And f4above = False Then
                        f4value = ((f4diff * f4rate) / 100)
                        txtF4Value.Text = Format(f4value, "$#,###.00")
                    Else
                        f4value = 0
                        txtF4Value.Text = Format(f4value, "$#,###.00")
                    End If
                End If
            End If
            '*****Bracket 5*****
            Dim f5from As Double
            Dim f5to As Double
            Dim f5rate As Double
            Dim f5level As Double
            Dim f5above As Boolean
            Dim f5diff As Double
            Dim f5value As Double
            If IsDBNull(txtF5From) Or txtF5From.Text = "" Then
                f5from = 0
            Else
                f5from = txtF5From.Text
            End If
            If IsDBNull(txtF5To) Or txtF5To.Text = "" Then
                f5to = 0
            Else
                f5to = txtF5To.Text
            End If
            If IsDBNull(txtF5Rate) Or txtF5Rate.Text = "" Then
                f5rate = 0
            Else
                Dim numericcheck As Boolean
                numericcheck = IsNumeric(txtF5Rate.Text)
                If numericcheck = True Then
                    f5rate = txtF5Rate.Text
                    Label42.Visible = False
                Else
                    Label42.Text = "Error - Rates must be entered as '0.00'"
                    Label42.Visible = True
                End If
            End If
            f5level = f5to - f5from
            f5above = tmv > f5to
            f5diff = tmv - f5from
            If ((IsDBNull(txtF5To) Or txtF5To.Text = "") And (txtF1To.Text <> "" And txtF2To.Text <> "" And txtF3To.Text <> "" And txtF4To.Text <> "")) Then
                If txtF6Rate.Text = "" Then
                Else
                    f5value = f5diff * f6rate / 100
                End If
            Else
                If f5diff > 0 And f5above = True Then
                    f5value = f5level * f5rate / 100
                    txtF5Value.Text = Format(f5value, "$#,###.00")
                Else
                    If f5diff > 0 And f5above = False Then
                        f5value = ((f5diff * f5rate) / 100)
                        txtF5Value.Text = Format(f5value, "$#,###.00")
                    Else
                        f5value = 0
                        txtF5Value.Text = Format(f5value, "$#,###.00")
                    End If
                End If
            End If
            '*****Bracket 6*****
            Dim f6from As Double

            Dim f6level As Double
            Dim f6above As Boolean
            Dim f6diff As Double
            Dim f6value As Double
            If IsDBNull(txtF6From) Or txtF6From.Text = "" Then
                f6from = 0
            Else
                f6from = txtF6From.Text
            End If


            f6level = tmv - f6from
            f6above = tmv > f6from
            f6diff = tmv - f6from
            If txtF6From.Text = "" Then
                'do nothing
            Else
                If f6diff > 0 And f6above = True Then
                    f6value = f6level * f6rate / 100
                    txtF6Value.Text = Format(f6value, "$#,###.00")
                Else
                    If f6diff > 0 And f6above = False Then
                        f6value = ((f6diff * f6rate) / 100)
                        txtF6Value.Text = Format(f6value, "$#,###.00")
                    Else
                        f6value = 0
                        txtF6Value.Text = Format(f6value, "$#,###.00")
                    End If
                End If
            End If
            Dim ttlrev As Double
            ttlrev = f1value + f2value + f3value + f4value + f5value + f6value
            Dim minfee As Double
            If IsDBNull(txtMinFee) Or txtMinFee.Text = "" Then
                minfee = 0
            Else
                minfee = txtMinFee.Text
            End If

            If ttlrev < minfee Then
                lblMinFee.ForeColor = Color.Red
                lblMinFee.Text = "Calculated Fee of " & Format(ttlrev, "$#,###.00") & " does not meet the minimum fee of " & Format(minfee, "$#,###.00") & "."
                lblMinFee.Visible = True
                ttlrev = minfee
            Else
                lblMinFee.ForeColor = Color.Black
                lblMinFee.Text = "Calculated Fee of " & Format(ttlrev, "$#,###.00") & " meets the minimum fee of " & Format(minfee, "$#,###.00") & "."
                lblMinFee.Visible = True
            End If
            Dim vel As Double
            vel = ((ttlrev / tmv) * 100)
            txtEstRev.Text = Format(ttlrev, "$#,###.00")
            txtVelocity.Text = Format(vel, "#.00")
        End If

    End Sub

    Public Sub LoadFirmData()
        If ckbLockData.Checked Then
        Else

            If IsDBNull(cboContact.SelectedValue) Then

            Else
                Try
                    Dim conn As New OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & My.Settings.dbloc & "';Jet OLEDB:Database Password='" & My.Settings.dbpass & "';")

                    Dim strSQL As String = "SELECT * FROM map_APXFirms WHERE ContactID = " & cboContact.SelectedValue
                    Dim queryString As String = String.Format(strSQL)
                    Dim cmd As New OleDb.OleDbCommand(queryString, conn)
                    Dim da As New OleDb.OleDbDataAdapter(cmd)
                    Dim ds As New DataSet

                    da.Fill(ds, "User")
                    Dim dt As DataTable = ds.Tables("User")

                    Dim row1 As DataRow = dt.Rows(0)


                    If IsDBNull(row1("AddressLine1")) Then
                        txtAddress.Text = "UNKNOWN"
                    Else
                        txtAddress.Text = (row1("AddressLine1"))
                    End If

                    If IsDBNull(row1("AddressCity")) Then
                        txtCity.Text = "UNKNOWN"
                    Else
                        txtCity.Text = (row1("AddressCity"))
                    End If

                    If IsDBNull(row1("AddressStateCode")) Then
                        txtState.Text = "UNKNOWN"
                    Else
                        txtState.Text = (row1("AddressStateCode"))
                    End If

                    If IsDBNull(row1("AddressPostalCode")) Then
                        txtZip.Text = "UNKNOWN"
                    Else
                        txtZip.Text = (row1("AddressPostalCode"))
                    End If

                    If IsDBNull(row1("BusinessPhone")) Then
                        txtPhone.Text = "UNKNOWN"
                    Else
                        txtPhone.Text = (row1("BusinessPhone"))
                    End If

                    If IsDBNull(row1("Email")) Then
                        txtEmail.Text = "UNKNOWN"
                    Else
                        txtEmail.Text = (row1("Email"))
                    End If

                    If IsDBNull(row1("URL")) Then
                        txtWebsite.Text = "UNKNOWN"
                    Else
                        txtWebsite.Text = (row1("URL"))
                    End If

                    'Dim msgbox1 As MsgBoxResult

                    If (row1("DeliveryName")) = txtCompanyName.Text Then

                    Else
                        Dim ir As MsgBoxResult

                        ir = MsgBox("The company name you entered does not match the Advent contact record.  Would you like to update the company name to match the advent record?", MsgBoxStyle.YesNo + MsgBoxStyle.Information, "Match APX Contact")

                        If ir = MsgBoxResult.Yes Then
                            If IsDBNull(row1("DeliveryName")) Then
                                txtCompanyName.Text = "UNKNOWN"
                            Else
                                txtCompanyName.Text = (row1("DeliveryName"))
                            End If
                        Else
                        End If
                    End If
                    ckbLockData.Checked = True

                Catch ex As Exception
                    MsgBox(ex.Message, MsgBoxStyle.Critical, "ERROR")
                End Try
            End If
        End If
    End Sub

    Public Sub LoadContacts()
        Try

            Dim conn As New OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & My.Settings.dbloc & "';Jet OLEDB:Database Password='" & My.Settings.dbpass & "';")

            Dim strSQL As String = "SELECT ContactID, ContactName FROM dbo_vQbRowDefContact WHERE ContactType = 4" & _
            " GROUP BY ContactID, ContactName" & _
            " ORDER BY ContactName"
            Dim da As New OleDb.OleDbDataAdapter(strSQL, conn)
            Dim ds As New DataSet
            da.Fill(ds, "Users")

            With cboContact
                .DataSource = ds.Tables("Users")
                .DisplayMember = "ContactName"
                .ValueMember = "ContactID"
                .SelectedIndex = 0
            End With

            lblContactLoad.Visible = False
            cboContact.Enabled = True

            If txtID.Text = "NEW" Then

            Else
                cboContact.SelectedValue = txtlContactID.Text
            End If


        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "ERROR")
        End Try
    End Sub

    Public Sub LoadDeliveryMethod()
        Try

            Dim conn As New OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & My.Settings.dbloc & "';Jet OLEDB:Database Password='" & My.Settings.dbpass & "';")

            Dim strSQL As String = "SELECT ID, Method FROM rfp_DeliveryMethod WHERE Active = -1"
            Dim da As New OleDb.OleDbDataAdapter(strSQL, conn)
            Dim ds As New DataSet
            da.Fill(ds, "Users")

            With cboDeliveryMethod
                .DataSource = ds.Tables("Users")
                .DisplayMember = "Method"
                .ValueMember = "ID"
                .SelectedIndex = 0
            End With

            If txtID.Text = "NEW" Then

            Else
                'cboContact.SelectedValue = txtlContactID.Text
            End If


        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "ERROR")
        End Try
    End Sub

    Public Sub LoadDisciplines()
        Try

            Dim conn As New OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & My.Settings.dbloc & "';Jet OLEDB:Database Password='" & My.Settings.dbpass & "';")

            Dim strSQL As String = "SELECT DisplayValue FROM AAM_DisciplineQuery"
            Dim da As New OleDb.OleDbDataAdapter(strSQL, conn)
            Dim ds As New DataSet
            da.Fill(ds, "Users")

            With cboDiscipline
                .DataSource = ds.Tables("Users")
                .DisplayMember = "DisplayValue"
                .ValueMember = "DisplayValue"
                .SelectedIndex = 0
            End With

            If txtID.Text = "NEW" Then

            Else
                'cboContact.SelectedValue = txtlContactID.Text
            End If


        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "ERROR")
        End Try
    End Sub

    Public Sub LoadEmployees()
        'Try

        Dim conn As New OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & My.Settings.dbloc & "';Jet OLEDB:Database Password='" & My.Settings.dbpass & "';")

        Dim strSQL As String = "SELECT ContactID, ContactName FROM dbo_vQbRowDefContact WHERE Status = 'Employee'" & _
        " GROUP BY ContactID, ContactName" & _
        " ORDER BY ContactName"
        Dim da As New OleDb.OleDbDataAdapter(strSQL, conn)
        Dim ds As New DataSet
        da.Fill(ds, "Users")

        With cboEmployee
            .DataSource = ds.Tables("Users")
            .DisplayMember = "ContactName"
            .ValueMember = "ContactID"
            .SelectedIndex = 0
        End With

        lblEmployee.Visible = False
        cboEmployee.Enabled = True

        If txtID.Text = "NEW" Then

        Else
            cboEmployee.SelectedValue = txtlRqstBy.Text
        End If


        'Catch ex As Exception
        'MsgBox(ex.Message, MsgBoxStyle.Critical, "ERROR")
        'End Try
    End Sub

    Public Sub LoadStage()
        'Try

        Dim conn As New OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & My.Settings.dbloc & "';Jet OLEDB:Database Password='" & My.Settings.dbpass & "';")

        Dim strSQL As String = "SELECT ID, StageName FROM rfp_ReportStage WHERE IsActive = -1"
        '" GROUP BY ID, StageName" & _
        '" ORDER BY StageName"
        Dim da As New OleDb.OleDbDataAdapter(strSQL, conn)
        Dim ds As New DataSet
        da.Fill(ds, "Users")

        With cboStage
            .DataSource = ds.Tables("Users")
            .DisplayMember = "StageName"
            .ValueMember = "ID"
            .SelectedIndex = 0
        End With

        If CheckBox1.Checked Then
            cboStage.Enabled = True
        Else
            cboStage.Enabled = False
        End If
        'Call LoadDBData()

        If txtID.Text = "NEW" Then
            Dim nw As Integer
            nw = 1
            cboStage.SelectedValue = nw
        Else
            cboStage.SelectedValue = txtlStage.Text
        End If

        'Catch ex As Exception
        'MsgBox(ex.Message, MsgBoxStyle.Critical, "ERROR")
        'End Try
    End Sub

    Public Sub LoadReason()
        'Try

        Dim conn As New OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & My.Settings.dbloc & "';Jet OLEDB:Database Password='" & My.Settings.dbpass & "';")

        Dim strSQL As String = "SELECT ID, ReasonName FROM rfp_ReportReasons WHERE IsActive = -1" & _
        " GROUP BY ID, ReasonName" & _
        " ORDER BY ReasonName"
        Dim da As New OleDb.OleDbDataAdapter(strSQL, conn)
        Dim ds As New DataSet
        da.Fill(ds, "Users")

        With cboReason
            .DataSource = ds.Tables("Users")
            .DisplayMember = "ReasonName"
            .ValueMember = "ID"
            .SelectedIndex = 0
        End With

        If CheckBox1.Checked Then
            cboReason.Enabled = True
        Else
            cboReason.Enabled = False
        End If

        If txtID.Text = "NEW" Then
            Dim nw As Integer
            nw = 1
            cboReason.SelectedValue = nw
        Else
            If IsDBNull(txtlReason) Or txtlReason.Text = "" Then

            Else
                cboReason.SelectedValue = txtlReason.Text
            End If

        End If

        'Catch ex As Exception
        'MsgBox(ex.Message, MsgBoxStyle.Critical, "ERROR")
        'End Try
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        Me.Close()
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        If txtID.Text = "NEW" Then
            Button1.Enabled = False
            Control.CheckForIllegalCrossThreadCalls = False
            thInsert = New System.Threading.Thread(AddressOf InsertRec)
            thInsert.Start()
        Else
            Button1.Enabled = False
            Control.CheckForIllegalCrossThreadCalls = False
            thUpdate = New System.Threading.Thread(AddressOf UpdateRec)
            thUpdate.Start()
        End If
    End Sub

    Public Sub InsertRec()
        Dim Mycn As OleDb.OleDbConnection
        Dim Command As OleDb.OleDbCommand
        Dim SQLstr As String

        'Try
        Mycn = New OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & My.Settings.roledbloc & "';Jet OLEDB:Database Password='" & My.Settings.dbpass & "';")

        Dim ds1 As New DataSet
        Dim eds1 As New DataGridView
        Dim dv1 As New DataView

        Mycn.Open()

        Dim ds As Date
        If IsDBNull(txtDateStarted) Or txtDateStarted.Text = "" Then
            ds = "12/31/2099"
        Else
            ds = txtDateStarted.Text
        End If

        Dim df As Date
        If IsDBNull(txtDateFinished) Or txtDateFinished.Text = "" Then
            df = "12/31/2099"
        Else
            df = txtDateFinished.Text
        End If

        Dim title As String
        title = txtRptName.Text
        title = Replace(title, "'", "''")

        Dim CName As String
        CName = txtCompanyName.Text
        CName = Replace(CName, "'", "''")

        Dim PCName As String
        PCName = txtPrimaryContact.Text
        PCName = Replace(PCName, "'", "''")

        Dim Addy As String
        Addy = txtAddress.Text
        Addy = Replace(Addy, "'", "''")

        Dim Cty As String
        Cty = txtCity.Text
        Cty = Replace(Cty, "'", "''")

        Dim CNotes As String
        CNotes = rtbCustomization.Text
        CNotes = Replace(CNotes, "'", "''")

        Dim Narr As String
        Narr = rtbNarrative.Text
        Narr = Replace(Narr, "'", "''")

        Dim b1to As Double
        Dim b1rate As Double
        Dim b2to As Double
        Dim b2rate As Double
        Dim b3to As Double
        Dim b3rate As Double
        Dim b4to As Double
        Dim b4rate As Double
        Dim b5to As Double
        Dim b5rate As Double
        Dim b6rate As Double
        Dim flatrate As Double
        Dim minfee As Double
        Dim aum As Double

        If IsDBNull(txtFlatRate) Or txtFlatRate.Text = "" Then
            flatrate = 0
        Else
            flatrate = txtFlatRate.Text
        End If

        If IsDBNull(txtF1Rate) Or txtF1Rate.Text = "" Then
            b1rate = 0
        Else
            b1rate = txtF1Rate.Text
        End If

        If IsDBNull(txtF2Rate) Or txtF2Rate.Text = "" Then
            b2rate = 0
        Else
            b2rate = txtF2Rate.Text
        End If

        If IsDBNull(txtF3Rate) Or txtF3Rate.Text = "" Then
            b3rate = 0
        Else
            b3rate = txtF3Rate.Text
        End If

        If IsDBNull(txtF4Rate) Or txtF4Rate.Text = "" Then
            b4rate = 0
        Else
            b4rate = txtF4Rate.Text
        End If

        If IsDBNull(txtF5Rate) Or txtF5Rate.Text = "" Then
            b5rate = 0
        Else
            b5rate = txtF5Rate.Text
        End If

        If IsDBNull(txtF6Rate) Or txtF6Rate.Text = "" Then
            b6rate = 0
        Else
            b6rate = txtF6Rate.Text
        End If

        If IsDBNull(txtF1To) Or txtF1To.Text = "" Then
            b1to = 0
        Else
            b1to = txtF1To.Text
            b1to = Format(b1to, "#.00")
        End If

        If IsDBNull(txtF2To) Or txtF2To.Text = "" Then
            b2to = 0
        Else
            b2to = txtF2To.Text
            b2to = Format(b2to, "#.00")
        End If

        If IsDBNull(txtF3To) Or txtF3To.Text = "" Then
            b3to = 0
        Else
            b3to = txtF3To.Text
            b3to = Format(b3to, "#.00")
        End If

        If IsDBNull(txtF4To) Or txtF4To.Text = "" Then
            b4to = 0
        Else
            b4to = txtF4To.Text
            b4to = Format(b4to, "#.00")
        End If

        If IsDBNull(txtF5To) Or txtF5To.Text = "" Then
            b5to = 0
        Else
            b5to = txtF5To.Text
            b5to = Format(b5to, "#.00")
        End If

        If IsDBNull(txtMinFee) Or txtMinFee.Text = "" Then
            minfee = 0
        Else
            minfee = txtMinFee.Text
            minfee = Format(minfee, "#.00")
        End If

        If IsDBNull(txtAUM) Or txtAUM.Text = "" Then
            aum = 0
        Else
            aum = txtAUM.Text
            aum = Format(aum, "#.00")
        End If


        SQLstr = "INSERT INTO rfp_Reports([Title], [ContactID], [DueDate], [RequestedBy], [Stage], [Reason], [Working], [DateStarted], [DateFinished], [WorkedBy], [CompanyName],[PrimaryContact],[PriContactShow],[Address],[City],[State],[Zip],[Phone],[Fax],[Website],[Email],[ProposedStrategy],[CustomizationNotes],[DeliveryMethod],[TrackingNumber], [DataLock], [AUM], [FlatRate],[FRate],[MinFee], [B1To],[B1Rate],[B2To],[B2Rate],[B3To],[B3Rate],[B4To],[B4Rate],[B5To],[B5Rate],[B6Rate],[HasImage],[ImageLocation],[DefaultDisclosure], [Narrative], [NarrativeFont], [PBFirm],[PBPerson],[PBAddress],[PBCity],[PBState],[PBZip],[PBPhone])" & _
        "VALUES('" & title & "'," & cboContact.SelectedValue & ",#" & DateTimePicker1.Text & "#," & cboEmployee.SelectedValue & "," & cboStage.SelectedValue & "," & cboReason.SelectedValue & "," & CheckBox1.CheckState & ",#" & ds & "#,#" & df & "#,'" & txtWorkedBy.Text & "','" & CName & "','" & PCName & "'," & ckbPriContact.Checked & ",'" & Addy & "','" & Cty & "','" & txtState.Text & "','" & txtZip.Text & "','" & txtPhone.Text & "','" & txtFax.Text & "','" & txtWebsite.Text & "','" & txtEmail.Text & "','" & cboDiscipline.Text & "','" & CNotes & "'," & cboDeliveryMethod.SelectedValue & ",'" & txtTrackingNumber.Text & "'," & ckbLockData.Checked & "," & aum & "," & ckbFlatRate.CheckState & "," & flatrate & "," & minfee & "," & b1to & "," & b1rate & "," & b2to & "," & b2rate & "," & b3to & "," & b3rate & "," & b4to & "," & b4rate & "," & b5to & "," & b5rate & "," & b6rate & "," & ckbHasImage.Checked & ",'" & txtImageLocation.Text & "'," & ckbDisclosure.Checked & ",'" & Narr & "','" & rtbNarrative.Font.ToString & "','" & txtPBFirm.Text & "','" & txtPBPerson.Text & "','" & txtPBAddress.Text & "','" & txtPBCity.Text & "','" & txtPBState.Text & "','" & txtPBZip.Text & "','" & txtPBPhone.Text & "')"

        Command = New OleDb.OleDbCommand(SQLstr, Mycn)
        Command.ExecuteNonQuery()

        Mycn.Close()

        'Me.Close()

        Call LoadID()

        Button1.Enabled = True

        'Catch ex As Exception
        'MsgBox(ex.Message, MsgBoxStyle.Critical, "ERROR")

        'End Try
    End Sub

    Public Sub UpdateRec()
        Dim Mycn As OleDb.OleDbConnection
        Dim Command As OleDb.OleDbCommand
        Dim SQLstr As String

        'Try
        Mycn = New OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & My.Settings.roledbloc & "';Jet OLEDB:Database Password='" & My.Settings.dbpass & "';")

        Dim ds1 As New DataSet
        Dim eds1 As New DataGridView
        Dim dv1 As New DataView

        Mycn.Open()

        Dim ds As Date
        If IsDBNull(txtDateStarted) Or txtDateStarted.Text = "" Then
            ds = "12/31/2099"
        Else
            ds = txtDateStarted.Text
        End If

        Dim df As Date
        If IsDBNull(txtDateFinished) Or txtDateFinished.Text = "" Then
            df = "12/31/2099"
        Else
            df = txtDateFinished.Text
        End If

        Dim title As String
        title = txtRptName.Text
        title = Replace(title, "'", "''")

        Dim CName As String
        CName = txtCompanyName.Text
        CName = Replace(CName, "'", "''")

        Dim PCName As String
        PCName = txtPrimaryContact.Text
        PCName = Replace(PCName, "'", "''")

        Dim Addy As String
        Addy = txtAddress.Text
        Addy = Replace(Addy, "'", "''")

        Dim Cty As String
        Cty = txtCity.Text
        Cty = Replace(Cty, "'", "''")

        Dim CNotes As String
        CNotes = rtbCustomization.Text
        CNotes = Replace(CNotes, "'", "''")

        Dim Narr As String
        Narr = rtbNarrative.Text
        Narr = Replace(Narr, "'", "''")

        Dim b1to As Double
        Dim b1rate As Double
        Dim b2to As Double
        Dim b2rate As Double
        Dim b3to As Double
        Dim b3rate As Double
        Dim b4to As Double
        Dim b4rate As Double
        Dim b5to As Double
        Dim b5rate As Double
        Dim b6rate As Double
        Dim flatrate As Double
        Dim minfee As Double
        Dim aum As Double

        If IsDBNull(txtFlatRate) Or txtFlatRate.Text = "" Then
            flatrate = 0
        Else
            flatrate = txtFlatRate.Text
        End If

        If IsDBNull(txtF1Rate) Or txtF1Rate.Text = "" Then
            b1rate = 0
        Else
            b1rate = txtF1Rate.Text
        End If

        If IsDBNull(txtF2Rate) Or txtF2Rate.Text = "" Then
            b2rate = 0
        Else
            b2rate = txtF2Rate.Text
        End If

        If IsDBNull(txtF3Rate) Or txtF3Rate.Text = "" Then
            b3rate = 0
        Else
            b3rate = txtF3Rate.Text
        End If

        If IsDBNull(txtF4Rate) Or txtF4Rate.Text = "" Then
            b4rate = 0
        Else
            b4rate = txtF4Rate.Text
        End If

        If IsDBNull(txtF5Rate) Or txtF5Rate.Text = "" Then
            b5rate = 0
        Else
            b5rate = txtF5Rate.Text
        End If

        If IsDBNull(txtF6Rate) Or txtF6Rate.Text = "" Then
            b6rate = 0
        Else
            b6rate = txtF6Rate.Text
        End If

        If IsDBNull(txtF1To) Or txtF1To.Text = "" Then
            b1to = 0
        Else
            b1to = txtF1To.Text
            b1to = Format(b1to, "#.00")
        End If

        If IsDBNull(txtF2To) Or txtF2To.Text = "" Then
            b2to = 0
        Else
            b2to = txtF2To.Text
            b2to = Format(b2to, "#.00")
        End If

        If IsDBNull(txtF3To) Or txtF3To.Text = "" Then
            b3to = 0
        Else
            b3to = txtF3To.Text
            b3to = Format(b3to, "#.00")
        End If

        If IsDBNull(txtF4To) Or txtF4To.Text = "" Then
            b4to = 0
        Else
            b4to = txtF4To.Text
            b4to = Format(b4to, "#.00")
        End If

        If IsDBNull(txtF5To) Or txtF5To.Text = "" Then
            b5to = 0
        Else
            b5to = txtF5To.Text
            b5to = Format(b5to, "#.00")
        End If

        If IsDBNull(txtMinFee) Or txtMinFee.Text = "" Then
            minfee = 0
        Else
            minfee = txtMinFee.Text
            minfee = Format(minfee, "#.00")
        End If

        If IsDBNull(txtAUM) Or txtAUM.Text = "" Then
            aum = 0
        Else
            aum = txtAUM.Text
            aum = Format(aum, "#.00")
        End If


        SQLstr = "Update rfp_Reports SET [Title] = '" & title & "', [ContactID] = " & cboContact.SelectedValue & ", [DueDate] = #" & DateTimePicker1.Text & "#, [RequestedBy] = " & cboEmployee.SelectedValue & ", [Stage] = " & cboStage.SelectedValue & ", [Reason] = " & cboReason.SelectedValue & ", [Working] = " & CheckBox1.CheckState & ", [DateStarted] = #" & ds & "#, [DateFinished] = #" & df & "#, [WorkedBy] = '" & txtWorkedBy.Text & "', [CompanyName] = '" & CName & "', [PrimaryContact] = '" & PCName & "',[PriContactShow] = " & ckbPriContact.Checked & ",[Address] = '" & Addy & "',[City] = '" & Cty & "',[State] = '" & txtState.Text & "',[Zip] = '" & txtZip.Text & "',[Phone]='" & txtPhone.Text & "',[Fax]='" & txtFax.Text & "',[Website] = '" & txtWebsite.Text & "',[Email]='" & txtEmail.Text & "',[ProposedStrategy]='" & cboDiscipline.Text & "',[CustomizationNotes] = '" & CNotes & "', [DeliveryMethod]=" & cboDeliveryMethod.SelectedValue & ",[TrackingNumber]='" & txtTrackingNumber.Text & "', [DataLock] = " & ckbLockData.Checked & ",[AUM] = " & aum & ",[FlatRate] = " & ckbFlatRate.CheckState & ",[FRate] = " & flatrate & ",[MinFee] = " & minfee & ",[B1To] = " & b1to & ",[B1Rate]=" & b1rate & ",[B2To]=" & b2to & ",[B2Rate]=" & b2rate & ",[B3To]=" & b3to & ",[B3Rate]=" & b3rate & ",[B4To]=" & b4to & ",[B4Rate]=" & b4rate & ",[B5To]=" & b5to & ",[B5Rate]=" & b5rate & ",[B6Rate]=" & b6rate & ",[HasImage]=" & ckbHasImage.Checked & ",[ImageLocation]='" & txtImageLocation.Text & "',[DefaultDisclosure] = " & ckbDisclosure.Checked & ",[Narrative] = '" & Narr & "',[NarrativeFont]='" & rtbNarrative.Font.ToString & "',[PBFirm]='" & txtPBFirm.Text & "',[PBPerson] = '" & txtPBPerson.Text & "',[PBAddress]='" & txtPBAddress.Text & "',[PBCity]='" & txtPBCity.Text & "',[PBState]='" & txtPBState.Text & "',[PBZip]='" & txtPBZip.Text & "',[PBPhone]='" & txtPBPhone.Text & "' WHERE ID = " & txtID.Text

        Command = New OleDb.OleDbCommand(SQLstr, Mycn)
        Command.ExecuteNonQuery()

        Mycn.Close()

        Call TrackChanges()

        Button1.Enabled = True
        'Me.Close()

        'Catch ex As Exception
        'MsgBox(ex.Message, MsgBoxStyle.Critical, "ERROR")

        'End Try
    End Sub

    Public Sub TrackChanges()
        Dim Mycn As OleDb.OleDbConnection
        Dim Command As OleDb.OleDbCommand
        Dim SQLstr As String
        'Try
        Mycn = New OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & My.Settings.dbloc & "';Jet OLEDB:Database Password='" & My.Settings.dbpass & "';")

        Dim title As String
        title = txtRptName.Text
        title = Replace(title, "'", "''")

        Dim title2 As String
        title2 = txtLrtpnme.Text
        title2 = Replace(title2, "'", "''")

        If txtID.Text = txtlID.Text Then

        Else
            Mycn.Open()
            SQLstr = "INSERT INTO rfp_ReportsTracking([ReportID], [Field], [NewValue], [ChangedBy], [DateStamp])" & _
            "VALUES(" & txtID.Text & ",'NEW RECORD','" & txtID.Text & "','" & Environ("USERNAME") & "',#" & Format(Now()) & "#)"

            Command = New OleDb.OleDbCommand(SQLstr, Mycn)
            Command.ExecuteNonQuery()

            Mycn.Close()

            txtlID.Text = txtID.Text
        End If

        'Mycn.Open()
        If title = title2 Then
            'nothing changed
        Else
            Mycn.Open()
            SQLstr = "INSERT INTO rfp_ReportsTracking([ReportID], [Field], [OldValue], [NewValue], [ChangedBy], [DateStamp])" & _
            "VALUES(" & txtID.Text & ",'Report Name','" & title2 & "','" & title & "','" & Environ("USERNAME") & "',#" & Format(Now()) & "#)"

            Command = New OleDb.OleDbCommand(SQLstr, Mycn)
            Command.ExecuteNonQuery()

            Mycn.Close()

            txtLrtpnme.Text = title
        End If

        If IsDBNull(txtlContactID) Or txtlContactID.Text = "" Then
            Mycn.Open()
            SQLstr = "INSERT INTO rfp_ReportsTracking([ReportID], [Field], [NewValue], [ChangedBy], [DateStamp])" & _
            "VALUES(" & txtID.Text & ",'Contact Name','" & cboContact.SelectedValue & "','" & Environ("USERNAME") & "',#" & Format(Now()) & "#)"

            Command = New OleDb.OleDbCommand(SQLstr, Mycn)
            Command.ExecuteNonQuery()

            Mycn.Close()

            txtlContactID.Text = cboContact.SelectedValue
        Else
            If cboContact.SelectedValue = txtlContactID.Text Then
                'nothing changed
            Else

                Mycn.Open()
                SQLstr = "INSERT INTO rfp_ReportsTracking([ReportID], [Field], [OldValue], [NewValue], [ChangedBy], [DateStamp])" & _
                "VALUES(" & txtID.Text & ",'Contact Name','" & txtlContactID.Text & "','" & cboContact.SelectedValue & "','" & Environ("USERNAME") & "',#" & Format(Now()) & "#)"

                Command = New OleDb.OleDbCommand(SQLstr, Mycn)
                Command.ExecuteNonQuery()

                Mycn.Close()

                txtlContactID.Text = cboContact.SelectedValue
            End If

        End If

        If IsDBNull(txtlDueDate) Or txtlDueDate.Text = "" Then
            Mycn.Open()
            SQLstr = "INSERT INTO rfp_ReportsTracking([ReportID], [Field], [NewValue], [ChangedBy], [DateStamp])" & _
            "VALUES(" & txtID.Text & ",'Due Date','" & DateTimePicker1.Text & "','" & Environ("USERNAME") & "',#" & Format(Now()) & "#)"

            Command = New OleDb.OleDbCommand(SQLstr, Mycn)
            Command.ExecuteNonQuery()

            Mycn.Close()

            txtlDueDate.Text = DateTimePicker1.Text
        Else
            If DateTimePicker1.Text = txtlDueDate.Text Then
                'nothing changed
            Else

                Mycn.Open()
                SQLstr = "INSERT INTO rfp_ReportsTracking([ReportID], [Field], [OldValue], [NewValue], [ChangedBy], [DateStamp])" & _
                "VALUES(" & txtID.Text & ",'Due Date','" & txtlDueDate.Text & "','" & DateTimePicker1.Text & "','" & Environ("USERNAME") & "',#" & Format(Now()) & "#)"

                Command = New OleDb.OleDbCommand(SQLstr, Mycn)
                Command.ExecuteNonQuery()

                Mycn.Close()

                txtlDueDate.Text = DateTimePicker1.Text
            End If

        End If

        If IsDBNull(txtlRqstBy) Or txtlRqstBy.Text = "" Then
            Mycn.Open()
            SQLstr = "INSERT INTO rfp_ReportsTracking([ReportID], [Field],[NewValue], [ChangedBy], [DateStamp])" & _
            "VALUES(" & txtID.Text & ",'RequestedBy','" & cboEmployee.SelectedValue & "','" & Environ("USERNAME") & "',#" & Format(Now()) & "#)"

            Command = New OleDb.OleDbCommand(SQLstr, Mycn)
            Command.ExecuteNonQuery()

            Mycn.Close()

            txtlRqstBy.Text = cboEmployee.SelectedValue
        Else
            If cboEmployee.SelectedValue = txtlRqstBy.Text Then
                'nothing changed
            Else

                Mycn.Open()
                SQLstr = "INSERT INTO rfp_ReportsTracking([ReportID], [Field], [OldValue], [NewValue], [ChangedBy], [DateStamp])" & _
                "VALUES(" & txtID.Text & ",'RequestedBy','" & txtlRqstBy.Text & "','" & cboEmployee.SelectedValue & "','" & Environ("USERNAME") & "',#" & Format(Now()) & "#)"

                Command = New OleDb.OleDbCommand(SQLstr, Mycn)
                Command.ExecuteNonQuery()

                Mycn.Close()

                txtlRqstBy.Text = cboEmployee.SelectedValue
            End If

        End If

        If CheckBox1.Checked = CheckBox2.Checked Then
            'nothing changed
        Else
            Mycn.Open()
            SQLstr = "INSERT INTO rfp_ReportsTracking([ReportID], [Field], [OldValue], [NewValue], [ChangedBy], [DateStamp])" & _
            "VALUES(" & txtID.Text & ",'Working','" & CheckBox2.Checked & "','" & CheckBox1.Checked & "','" & Environ("USERNAME") & "',#" & Format(Now()) & "#)"

            Command = New OleDb.OleDbCommand(SQLstr, Mycn)
            Command.ExecuteNonQuery()

            Mycn.Close()

            CheckBox2.Checked = CheckBox1.Checked
        End If

        If IsDBNull(txtlStage) Or txtlStage.Text = "" Then
            Mycn.Open()
            SQLstr = "INSERT INTO rfp_ReportsTracking([ReportID], [Field], [NewValue], [ChangedBy], [DateStamp])" & _
            "VALUES(" & txtID.Text & ",'Stage','" & cboStage.SelectedValue & "','" & Environ("USERNAME") & "',#" & Format(Now()) & "#)"

            Command = New OleDb.OleDbCommand(SQLstr, Mycn)
            Command.ExecuteNonQuery()

            Mycn.Close()

            txtlStage.Text = cboStage.SelectedValue
        Else
            If cboStage.SelectedValue = txtlStage.Text Then
                'nothing changed
            Else

                Mycn.Open()
                SQLstr = "INSERT INTO rfp_ReportsTracking([ReportID], [Field], [OldValue], [NewValue], [ChangedBy], [DateStamp])" & _
                "VALUES(" & txtID.Text & ",'Stage','" & txtlStage.Text & "','" & cboStage.SelectedValue & "','" & Environ("USERNAME") & "',#" & Format(Now()) & "#)"

                Command = New OleDb.OleDbCommand(SQLstr, Mycn)
                Command.ExecuteNonQuery()

                Mycn.Close()

                txtlStage.Text = cboStage.SelectedValue
            End If

        End If

        If IsDBNull(txtlReason) Or txtlReason.Text = "" Then
            Mycn.Open()
            SQLstr = "INSERT INTO rfp_ReportsTracking([ReportID], [Field], [NewValue], [ChangedBy], [DateStamp])" & _
            "VALUES(" & txtID.Text & ",'Reason','" & cboReason.SelectedValue & "','" & Environ("USERNAME") & "',#" & Format(Now()) & "#)"

            Command = New OleDb.OleDbCommand(SQLstr, Mycn)
            Command.ExecuteNonQuery()

            Mycn.Close()

            txtlReason.Text = cboReason.SelectedValue
        Else
            If cboReason.SelectedValue = txtlReason.Text Then
                'nothing changed
            Else

                Mycn.Open()
                SQLstr = "INSERT INTO rfp_ReportsTracking([ReportID], [Field], [OldValue], [NewValue], [ChangedBy], [DateStamp])" & _
                "VALUES(" & txtID.Text & ",'Reason','" & txtlReason.Text & "','" & cboReason.SelectedValue & "','" & Environ("USERNAME") & "',#" & Format(Now()) & "#)"

                Command = New OleDb.OleDbCommand(SQLstr, Mycn)
                Command.ExecuteNonQuery()

                Mycn.Close()

                txtlReason.Text = cboReason.SelectedValue
            End If

        End If

        Dim CName As String
        CName = txtCompanyName.Text
        CName = Replace(CName, "'", "''")

        Dim CName1 As String
        CName1 = txtCompanyNameCopy.Text
        CName1 = Replace(CName1, "'", "''")

        If IsDBNull(txtCompanyNameCopy) Or txtCompanyNameCopy.Text = "" Then
            Mycn.Open()
            SQLstr = "INSERT INTO rfp_ReportsTracking([ReportID], [Field], [NewValue], [ChangedBy], [DateStamp])" & _
            "VALUES(" & txtID.Text & ",'Company Name','" & CName & "','" & Environ("USERNAME") & "',#" & Format(Now()) & "#)"

            Command = New OleDb.OleDbCommand(SQLstr, Mycn)
            Command.ExecuteNonQuery()

            Mycn.Close()

            txtCompanyNameCopy.Text = txtCompanyName.Text
        Else
            If txtCompanyName.Text = txtCompanyNameCopy.Text Then
                'nothing changed
            Else

                Mycn.Open()
                SQLstr = "INSERT INTO rfp_ReportsTracking([ReportID], [Field], [OldValue], [NewValue], [ChangedBy], [DateStamp])" & _
                "VALUES(" & txtID.Text & ",'Company Name','" & CName1 & "','" & CName & "','" & Environ("USERNAME") & "',#" & Format(Now()) & "#)"

                Command = New OleDb.OleDbCommand(SQLstr, Mycn)
                Command.ExecuteNonQuery()

                Mycn.Close()

                txtCompanyNameCopy.Text = txtCompanyName.Text
            End If

        End If

        Dim PCName As String
        PCName = txtPrimaryContact.Text
        PCName = Replace(PCName, "'", "''")

        Dim PCName1 As String
        PCName1 = txtPrimaryContactCopy.Text
        PCName1 = Replace(PCName1, "'", "''")

        If IsDBNull(txtPrimaryContactCopy) Or txtPrimaryContactCopy.Text = "" Then
            Mycn.Open()
            SQLstr = "INSERT INTO rfp_ReportsTracking([ReportID], [Field], [NewValue], [ChangedBy], [DateStamp])" & _
            "VALUES(" & txtID.Text & ",'Primary Contact','" & PCName & "','" & Environ("USERNAME") & "',#" & Format(Now()) & "#)"

            Command = New OleDb.OleDbCommand(SQLstr, Mycn)
            Command.ExecuteNonQuery()

            Mycn.Close()

            txtPrimaryContactCopy.Text = txtPrimaryContact.Text
        Else
            If txtPrimaryContact.Text = txtPrimaryContactCopy.Text Then
                'nothing changed
            Else

                Mycn.Open()
                SQLstr = "INSERT INTO rfp_ReportsTracking([ReportID], [Field], [OldValue], [NewValue], [ChangedBy], [DateStamp])" & _
                "VALUES(" & txtID.Text & ",'Primary Contact','" & PCName1 & "','" & PCName & "','" & Environ("USERNAME") & "',#" & Format(Now()) & "#)"

                Command = New OleDb.OleDbCommand(SQLstr, Mycn)
                Command.ExecuteNonQuery()

                Mycn.Close()

                txtPrimaryContactCopy.Text = txtPrimaryContact.Text
            End If

        End If

        If ckbPriContact.Checked = ckbPriContactCopy.Checked Then
            'nothing changed
        Else

            Mycn.Open()
            SQLstr = "INSERT INTO rfp_ReportsTracking([ReportID], [Field], [OldValue], [NewValue], [ChangedBy], [DateStamp])" & _
            "VALUES(" & txtID.Text & ",'Show Primary Contact','" & ckbPriContactCopy.Checked & "','" & ckbPriContact.Checked & "','" & Environ("USERNAME") & "',#" & Format(Now()) & "#)"

            Command = New OleDb.OleDbCommand(SQLstr, Mycn)
            Command.ExecuteNonQuery()

            Mycn.Close()

            ckbPriContactCopy.Checked = ckbPriContact.Checked
        End If

        Dim addy As String
        addy = txtAddress.Text
        addy = Replace(addy, "'", "''")

        Dim addy1 As String
        addy1 = txtAddressCopy.Text
        addy1 = Replace(addy1, "'", "''")

        If IsDBNull(txtAddressCopy) Or txtAddressCopy.Text = "" Then
            Mycn.Open()
            SQLstr = "INSERT INTO rfp_ReportsTracking([ReportID], [Field], [NewValue], [ChangedBy], [DateStamp])" & _
            "VALUES(" & txtID.Text & ",'Address','" & addy & "','" & Environ("USERNAME") & "',#" & Format(Now()) & "#)"

            Command = New OleDb.OleDbCommand(SQLstr, Mycn)
            Command.ExecuteNonQuery()

            Mycn.Close()

            txtAddressCopy.Text = txtAddress.Text
        Else
            If txtAddress.Text = txtAddressCopy.Text Then
                'nothing changed
            Else

                Mycn.Open()
                SQLstr = "INSERT INTO rfp_ReportsTracking([ReportID], [Field], [OldValue], [NewValue], [ChangedBy], [DateStamp])" & _
                "VALUES(" & txtID.Text & ",'Address','" & addy1 & "','" & addy & "','" & Environ("USERNAME") & "',#" & Format(Now()) & "#)"

                Command = New OleDb.OleDbCommand(SQLstr, Mycn)
                Command.ExecuteNonQuery()
                Mycn.Close()

                txtAddressCopy.Text = txtAddress.Text
            End If
        End If

        Dim cty As String
        cty = txtCity.Text
        cty = Replace(cty, "'", "''")

        Dim cty1 As String
        cty1 = txtCityCopy.Text
        cty1 = Replace(cty1, "'", "''")

        If IsDBNull(txtAddressCopy) Or txtAddressCopy.Text = "" Then
            Mycn.Open()
            SQLstr = "INSERT INTO rfp_ReportsTracking([ReportID], [Field], [NewValue], [ChangedBy], [DateStamp])" & _
            "VALUES(" & txtID.Text & ",'City','" & cty & "','" & Environ("USERNAME") & "',#" & Format(Now()) & "#)"

            Command = New OleDb.OleDbCommand(SQLstr, Mycn)
            Command.ExecuteNonQuery()

            Mycn.Close()

            txtCityCopy.Text = txtCity.Text
        Else
            If txtCity.Text = txtCityCopy.Text Then
                'nothing changed
            Else

                Mycn.Open()
                SQLstr = "INSERT INTO rfp_ReportsTracking([ReportID], [Field], [OldValue], [NewValue], [ChangedBy], [DateStamp])" & _
                "VALUES(" & txtID.Text & ",'City','" & cty1 & "','" & cty & "','" & Environ("USERNAME") & "',#" & Format(Now()) & "#)"

                Command = New OleDb.OleDbCommand(SQLstr, Mycn)
                Command.ExecuteNonQuery()
                Mycn.Close()

                txtCityCopy.Text = txtCity.Text
            End If
        End If

        If IsDBNull(txtStateCopy) Or txtStateCopy.Text = "" Then
            Mycn.Open()
            SQLstr = "INSERT INTO rfp_ReportsTracking([ReportID], [Field], [NewValue], [ChangedBy], [DateStamp])" & _
            "VALUES(" & txtID.Text & ",'State','" & txtState.Text & "','" & Environ("USERNAME") & "',#" & Format(Now()) & "#)"

            Command = New OleDb.OleDbCommand(SQLstr, Mycn)
            Command.ExecuteNonQuery()

            Mycn.Close()

            txtStateCopy.Text = txtState.Text
        Else
            If txtState.Text = txtStateCopy.Text Then
                'nothing changed
            Else

                Mycn.Open()
                SQLstr = "INSERT INTO rfp_ReportsTracking([ReportID], [Field], [OldValue], [NewValue], [ChangedBy], [DateStamp])" & _
                "VALUES(" & txtID.Text & ",'State','" & txtStateCopy.Text & "','" & txtState.Text & "','" & Environ("USERNAME") & "',#" & Format(Now()) & "#)"

                Command = New OleDb.OleDbCommand(SQLstr, Mycn)
                Command.ExecuteNonQuery()
                Mycn.Close()

                txtStateCopy.Text = txtState.Text
            End If
        End If

        If IsDBNull(txtZipCopy) Or txtZipCopy.Text = "" Then
            Mycn.Open()
            SQLstr = "INSERT INTO rfp_ReportsTracking([ReportID], [Field], [NewValue], [ChangedBy], [DateStamp])" & _
            "VALUES(" & txtID.Text & ",'Zip','" & txtZip.Text & "','" & Environ("USERNAME") & "',#" & Format(Now()) & "#)"

            Command = New OleDb.OleDbCommand(SQLstr, Mycn)
            Command.ExecuteNonQuery()

            Mycn.Close()

            txtZipCopy.Text = txtZip.Text
        Else
            If txtZip.Text = txtZipCopy.Text Then
                'nothing changed
            Else

                Mycn.Open()
                SQLstr = "INSERT INTO rfp_ReportsTracking([ReportID], [Field], [OldValue], [NewValue], [ChangedBy], [DateStamp])" & _
                "VALUES(" & txtID.Text & ",'Zip','" & txtZipCopy.Text & "','" & txtZip.Text & "','" & Environ("USERNAME") & "',#" & Format(Now()) & "#)"

                Command = New OleDb.OleDbCommand(SQLstr, Mycn)
                Command.ExecuteNonQuery()
                Mycn.Close()

                txtZipCopy.Text = txtZip.Text
            End If
        End If

        If IsDBNull(txtPhoneCopy) Or txtPhoneCopy.Text = "" Then
            Mycn.Open()
            SQLstr = "INSERT INTO rfp_ReportsTracking([ReportID], [Field], [NewValue], [ChangedBy], [DateStamp])" & _
            "VALUES(" & txtID.Text & ",'Phone','" & txtPhone.Text & "','" & Environ("USERNAME") & "',#" & Format(Now()) & "#)"

            Command = New OleDb.OleDbCommand(SQLstr, Mycn)
            Command.ExecuteNonQuery()

            Mycn.Close()

            txtPhoneCopy.Text = txtPhone.Text
        Else
            If txtPhone.Text = txtPhoneCopy.Text Then
                'nothing changed
            Else

                Mycn.Open()
                SQLstr = "INSERT INTO rfp_ReportsTracking([ReportID], [Field], [OldValue], [NewValue], [ChangedBy], [DateStamp])" & _
                "VALUES(" & txtID.Text & ",'Phone','" & txtPhoneCopy.Text & "','" & txtPhone.Text & "','" & Environ("USERNAME") & "',#" & Format(Now()) & "#)"

                Command = New OleDb.OleDbCommand(SQLstr, Mycn)
                Command.ExecuteNonQuery()
                Mycn.Close()

                txtPhoneCopy.Text = txtPhone.Text
            End If
        End If

        If IsDBNull(txtFaxCopy) Or txtFaxCopy.Text = "" Then
            Mycn.Open()
            SQLstr = "INSERT INTO rfp_ReportsTracking([ReportID], [Field], [NewValue], [ChangedBy], [DateStamp])" & _
            "VALUES(" & txtID.Text & ",'Fax','" & txtFax.Text & "','" & Environ("USERNAME") & "',#" & Format(Now()) & "#)"

            Command = New OleDb.OleDbCommand(SQLstr, Mycn)
            Command.ExecuteNonQuery()

            Mycn.Close()

            txtFaxCopy.Text = txtFax.Text
        Else
            If txtFax.Text = txtFaxCopy.Text Then
                'nothing changed
            Else

                Mycn.Open()
                SQLstr = "INSERT INTO rfp_ReportsTracking([ReportID], [Field], [OldValue], [NewValue], [ChangedBy], [DateStamp])" & _
                "VALUES(" & txtID.Text & ",'Fax','" & txtFaxCopy.Text & "','" & txtFax.Text & "','" & Environ("USERNAME") & "',#" & Format(Now()) & "#)"

                Command = New OleDb.OleDbCommand(SQLstr, Mycn)
                Command.ExecuteNonQuery()
                Mycn.Close()

                txtFaxCopy.Text = txtFax.Text
            End If
        End If

        If IsDBNull(txtWebsiteCopy) Or txtWebsiteCopy.Text = "" Then
            Mycn.Open()
            SQLstr = "INSERT INTO rfp_ReportsTracking([ReportID], [Field], [NewValue], [ChangedBy], [DateStamp])" & _
            "VALUES(" & txtID.Text & ",'Website','" & txtWebsite.Text & "','" & Environ("USERNAME") & "',#" & Format(Now()) & "#)"

            Command = New OleDb.OleDbCommand(SQLstr, Mycn)
            Command.ExecuteNonQuery()

            Mycn.Close()

            txtWebsiteCopy.Text = txtWebsite.Text
        Else
            If txtWebsite.Text = txtWebsiteCopy.Text Then
                'nothing changed
            Else

                Mycn.Open()
                SQLstr = "INSERT INTO rfp_ReportsTracking([ReportID], [Field], [OldValue], [NewValue], [ChangedBy], [DateStamp])" & _
                "VALUES(" & txtID.Text & ",'Website','" & txtWebsiteCopy.Text & "','" & txtWebsite.Text & "','" & Environ("USERNAME") & "',#" & Format(Now()) & "#)"

                Command = New OleDb.OleDbCommand(SQLstr, Mycn)
                Command.ExecuteNonQuery()
                Mycn.Close()

                txtWebsiteCopy.Text = txtWebsite.Text
            End If
        End If

        If IsDBNull(txtEmailCopy) Or txtEmailCopy.Text = "" Then
            Mycn.Open()
            SQLstr = "INSERT INTO rfp_ReportsTracking([ReportID], [Field], [NewValue], [ChangedBy], [DateStamp])" & _
            "VALUES(" & txtID.Text & ",'Email','" & txtEmail.Text & "','" & Environ("USERNAME") & "',#" & Format(Now()) & "#)"

            Command = New OleDb.OleDbCommand(SQLstr, Mycn)
            Command.ExecuteNonQuery()

            Mycn.Close()

            txtEmailCopy.Text = txtEmail.Text
        Else
            If txtEmail.Text = txtEmailCopy.Text Then
                'nothing changed
            Else

                Mycn.Open()
                SQLstr = "INSERT INTO rfp_ReportsTracking([ReportID], [Field], [OldValue], [NewValue], [ChangedBy], [DateStamp])" & _
                "VALUES(" & txtID.Text & ",'Email','" & txtEmailCopy.Text & "','" & txtEmail.Text & "','" & Environ("USERNAME") & "',#" & Format(Now()) & "#)"

                Command = New OleDb.OleDbCommand(SQLstr, Mycn)
                Command.ExecuteNonQuery()
                Mycn.Close()

                txtEmailCopy.Text = txtEmail.Text
            End If
        End If

        If IsDBNull(txtDisciplineCopy) Or txtDisciplineCopy.Text = "" Then
            Mycn.Open()
            SQLstr = "INSERT INTO rfp_ReportsTracking([ReportID], [Field], [NewValue], [ChangedBy], [DateStamp])" & _
            "VALUES(" & txtID.Text & ",'Discipline','" & cboDiscipline.Text & "','" & Environ("USERNAME") & "',#" & Format(Now()) & "#)"

            Command = New OleDb.OleDbCommand(SQLstr, Mycn)
            Command.ExecuteNonQuery()

            Mycn.Close()

            txtDisciplineCopy.Text = cboDiscipline.Text
        Else
            If cboDiscipline.Text = txtDisciplineCopy.Text Then
                'nothing changed
            Else

                Mycn.Open()
                SQLstr = "INSERT INTO rfp_ReportsTracking([ReportID], [Field], [OldValue], [NewValue], [ChangedBy], [DateStamp])" & _
                "VALUES(" & txtID.Text & ",'Discipline','" & txtDisciplineCopy.Text & "','" & cboDiscipline.Text & "','" & Environ("USERNAME") & "',#" & Format(Now()) & "#)"

                Command = New OleDb.OleDbCommand(SQLstr, Mycn)
                Command.ExecuteNonQuery()
                Mycn.Close()

                txtDisciplineCopy.Text = cboDiscipline.Text
            End If
        End If

        If IsDBNull(txtDeliveryMethodCopy) Or txtDeliveryMethodCopy.Text = "" Then
            Mycn.Open()
            SQLstr = "INSERT INTO rfp_ReportsTracking([ReportID], [Field], [NewValue], [ChangedBy], [DateStamp])" & _
            "VALUES(" & txtID.Text & ",'Delivery Method','" & cboDeliveryMethod.SelectedValue & "','" & Environ("USERNAME") & "',#" & Format(Now()) & "#)"

            Command = New OleDb.OleDbCommand(SQLstr, Mycn)
            Command.ExecuteNonQuery()

            Mycn.Close()

            txtDeliveryMethodCopy.Text = cboDeliveryMethod.SelectedValue
        Else
            If cboDeliveryMethod.SelectedValue = txtDeliveryMethodCopy.Text Then
                'nothing changed
            Else

                Mycn.Open()
                SQLstr = "INSERT INTO rfp_ReportsTracking([ReportID], [Field], [OldValue], [NewValue], [ChangedBy], [DateStamp])" & _
                "VALUES(" & txtID.Text & ",'Delivery Method','" & txtDeliveryMethodCopy.Text & "','" & cboDeliveryMethod.SelectedValue & "','" & Environ("USERNAME") & "',#" & Format(Now()) & "#)"

                Command = New OleDb.OleDbCommand(SQLstr, Mycn)
                Command.ExecuteNonQuery()
                Mycn.Close()

                txtDeliveryMethodCopy.Text = cboDeliveryMethod.SelectedValue
            End If
        End If

        If IsDBNull(txtTrackingNumberCopy) Or txtTrackingNumberCopy.Text = "" Then
            Mycn.Open()
            SQLstr = "INSERT INTO rfp_ReportsTracking([ReportID], [Field], [NewValue], [ChangedBy], [DateStamp])" & _
            "VALUES(" & txtID.Text & ",'Tracking Number','" & txtTrackingNumber.Text & "','" & Environ("USERNAME") & "',#" & Format(Now()) & "#)"

            Command = New OleDb.OleDbCommand(SQLstr, Mycn)
            Command.ExecuteNonQuery()

            Mycn.Close()

            txtTrackingNumberCopy.Text = txtTrackingNumber.Text
        Else
            If txtTrackingNumber.Text = txtTrackingNumberCopy.Text Then
                'nothing changed
            Else

                Mycn.Open()
                SQLstr = "INSERT INTO rfp_ReportsTracking([ReportID], [Field], [OldValue], [NewValue], [ChangedBy], [DateStamp])" & _
                "VALUES(" & txtID.Text & ",'Tracking Number','" & txtTrackingNumberCopy.Text & "','" & txtTrackingNumber.Text & "','" & Environ("USERNAME") & "',#" & Format(Now()) & "#)"

                Command = New OleDb.OleDbCommand(SQLstr, Mycn)
                Command.ExecuteNonQuery()
                Mycn.Close()

                txtTrackingNumberCopy.Text = txtTrackingNumber.Text
            End If
        End If

        Dim CNote As String
        CNote = rtbCustomization.Text
        CNote = Replace(CNote, "'", "''")

        Dim CNote1 As String
        CNote1 = rtbCustomizationCopy.Text
        CNote1 = Replace(CNote1, "'", "''")

        If IsDBNull(rtbCustomizationCopy) Or rtbCustomizationCopy.Text = "" Then
            Mycn.Open()
            SQLstr = "INSERT INTO rfp_ReportsTracking([ReportID], [Field], [NewValue], [ChangedBy], [DateStamp])" & _
            "VALUES(" & txtID.Text & ",'Customization Notes','" & CNote & "','" & Environ("USERNAME") & "',#" & Format(Now()) & "#)"

            Command = New OleDb.OleDbCommand(SQLstr, Mycn)
            Command.ExecuteNonQuery()

            Mycn.Close()

            rtbCustomizationCopy.Text = rtbCustomization.Text
        Else
            If rtbCustomization.Text = rtbCustomizationCopy.Text Then
                'nothing changed
            Else

                Mycn.Open()
                SQLstr = "INSERT INTO rfp_ReportsTracking([ReportID], [Field], [OldValue], [NewValue], [ChangedBy], [DateStamp])" & _
                "VALUES(" & txtID.Text & ",'Customization Notes','" & CNote1 & "','" & CNote & "','" & Environ("USERNAME") & "',#" & Format(Now()) & "#)"

                Command = New OleDb.OleDbCommand(SQLstr, Mycn)
                Command.ExecuteNonQuery()
                Mycn.Close()

                rtbCustomizationCopy.Text = rtbCustomization.Text
            End If
        End If

        Dim b1to As Double
        If txtF1To.Text = "" Then
            b1to = 0
        Else
            b1to = txtF1To.Text
            b1to = Format(b1to, "#.00")
        End If


        If IsDBNull(txtF1ToCopy) Or txtF1ToCopy.Text = "" Then
            Mycn.Open()
            SQLstr = "INSERT INTO rfp_ReportsTracking([ReportID], [Field], [NewValue], [ChangedBy], [DateStamp])" & _
            "VALUES(" & txtID.Text & ",'Fee 1 To','" & b1to & "','" & Environ("USERNAME") & "',#" & Format(Now()) & "#)"

            Command = New OleDb.OleDbCommand(SQLstr, Mycn)
            Command.ExecuteNonQuery()

            Mycn.Close()

            txtF1ToCopy.Text = b1to
        Else
            If b1to = txtF1ToCopy.Text Then
                'nothing changed
            Else

                Mycn.Open()
                SQLstr = "INSERT INTO rfp_ReportsTracking([ReportID], [Field], [OldValue], [NewValue], [ChangedBy], [DateStamp])" & _
                "VALUES(" & txtID.Text & ",'Fee 1 To','" & txtF1ToCopy.Text & "','" & b1to & "','" & Environ("USERNAME") & "',#" & Format(Now()) & "#)"

                Command = New OleDb.OleDbCommand(SQLstr, Mycn)
                Command.ExecuteNonQuery()
                Mycn.Close()

                txtF1ToCopy.Text = b1to
            End If
        End If

        If IsDBNull(txtFlatRateCopy) Or txtFlatRateCopy.Text = "" Then
            Mycn.Open()
            SQLstr = "INSERT INTO rfp_ReportsTracking([ReportID], [Field], [NewValue], [ChangedBy], [DateStamp])" & _
            "VALUES(" & txtID.Text & ",'Flat Rate','" & txtFlatRate.Text & "','" & Environ("USERNAME") & "',#" & Format(Now()) & "#)"

            Command = New OleDb.OleDbCommand(SQLstr, Mycn)
            Command.ExecuteNonQuery()

            Mycn.Close()

            txtFlatRateCopy.Text = txtFlatRate.Text
        Else
            If txtFlatRate.Text = txtFlatRateCopy.Text Then
                'nothing changed
            Else

                Mycn.Open()
                SQLstr = "INSERT INTO rfp_ReportsTracking([ReportID], [Field], [OldValue], [NewValue], [ChangedBy], [DateStamp])" & _
                "VALUES(" & txtID.Text & ",'Flat Rate','" & txtFlatRateCopy.Text & "','" & txtFlatRate.Text & "','" & Environ("USERNAME") & "',#" & Format(Now()) & "#)"

                Command = New OleDb.OleDbCommand(SQLstr, Mycn)
                Command.ExecuteNonQuery()
                Mycn.Close()

                txtFlatRateCopy.Text = txtFlatRate.Text
            End If
        End If

        If IsDBNull(txtF1RateCopy) Or txtF1RateCopy.Text = "" Then
            Mycn.Open()
            SQLstr = "INSERT INTO rfp_ReportsTracking([ReportID], [Field], [NewValue], [ChangedBy], [DateStamp])" & _
            "VALUES(" & txtID.Text & ",'Fee 1 Rate','" & txtF1Rate.Text & "','" & Environ("USERNAME") & "',#" & Format(Now()) & "#)"

            Command = New OleDb.OleDbCommand(SQLstr, Mycn)
            Command.ExecuteNonQuery()

            Mycn.Close()

            txtF1RateCopy.Text = txtF1Rate.Text
        Else
            If txtF1Rate.Text = txtF1RateCopy.Text Then
                'nothing changed
            Else

                Mycn.Open()
                SQLstr = "INSERT INTO rfp_ReportsTracking([ReportID], [Field], [OldValue], [NewValue], [ChangedBy], [DateStamp])" & _
                "VALUES(" & txtID.Text & ",'Fee 1 Rate','" & txtF1RateCopy.Text & "','" & txtF1Rate.Text & "','" & Environ("USERNAME") & "',#" & Format(Now()) & "#)"

                Command = New OleDb.OleDbCommand(SQLstr, Mycn)
                Command.ExecuteNonQuery()
                Mycn.Close()

                txtF1RateCopy.Text = txtF1Rate.Text
            End If
        End If

        Dim b2to As Double
        If txtF2To.Text = "" Then
            b2to = 0
        Else
            b2to = txtF2To.Text
            b2to = Format(b2to, "#.00")
        End If


        If IsDBNull(txtF2ToCopy) Or txtF2ToCopy.Text = "" Then
            Mycn.Open()
            SQLstr = "INSERT INTO rfp_ReportsTracking([ReportID], [Field], [NewValue], [ChangedBy], [DateStamp])" & _
            "VALUES(" & txtID.Text & ",'Fee 2 To','" & b2to & "','" & Environ("USERNAME") & "',#" & Format(Now()) & "#)"

            Command = New OleDb.OleDbCommand(SQLstr, Mycn)
            Command.ExecuteNonQuery()

            Mycn.Close()

            txtF2ToCopy.Text = b2to
        Else
            If b2to = txtF2ToCopy.Text Then
                'nothing changed
            Else

                Mycn.Open()
                SQLstr = "INSERT INTO rfp_ReportsTracking([ReportID], [Field], [OldValue], [NewValue], [ChangedBy], [DateStamp])" & _
                "VALUES(" & txtID.Text & ",'Fee 2 To','" & txtF2ToCopy.Text & "','" & b2to & "','" & Environ("USERNAME") & "',#" & Format(Now()) & "#)"

                Command = New OleDb.OleDbCommand(SQLstr, Mycn)
                Command.ExecuteNonQuery()
                Mycn.Close()

                txtF2ToCopy.Text = b2to
            End If
        End If

        If IsDBNull(txtF2RateCopy) Or txtF2RateCopy.Text = "" Then
            Mycn.Open()
            SQLstr = "INSERT INTO rfp_ReportsTracking([ReportID], [Field], [NewValue], [ChangedBy], [DateStamp])" & _
            "VALUES(" & txtID.Text & ",'Fee 2 Rate','" & txtF2Rate.Text & "','" & Environ("USERNAME") & "',#" & Format(Now()) & "#)"

            Command = New OleDb.OleDbCommand(SQLstr, Mycn)
            Command.ExecuteNonQuery()

            Mycn.Close()

            txtF2RateCopy.Text = txtF2Rate.Text
        Else
            If txtF2Rate.Text = txtF2RateCopy.Text Then
                'nothing changed
            Else

                Mycn.Open()
                SQLstr = "INSERT INTO rfp_ReportsTracking([ReportID], [Field], [OldValue], [NewValue], [ChangedBy], [DateStamp])" & _
                "VALUES(" & txtID.Text & ",'Fee 2 Rate','" & txtF2RateCopy.Text & "','" & txtF2Rate.Text & "','" & Environ("USERNAME") & "',#" & Format(Now()) & "#)"

                Command = New OleDb.OleDbCommand(SQLstr, Mycn)
                Command.ExecuteNonQuery()
                Mycn.Close()

                txtF2RateCopy.Text = txtF2Rate.Text
            End If
        End If

        Dim b3to As Double
        If txtF3To.Text = "" Then
            b3to = 0
        Else
            b3to = txtF3To.Text
            b3to = Format(b3to, "#.00")
        End If


        If IsDBNull(txtF3ToCopy) Or txtF3ToCopy.Text = "" Then
            Mycn.Open()
            SQLstr = "INSERT INTO rfp_ReportsTracking([ReportID], [Field], [NewValue], [ChangedBy], [DateStamp])" & _
            "VALUES(" & txtID.Text & ",'Fee 3 To','" & b3to & "','" & Environ("USERNAME") & "',#" & Format(Now()) & "#)"

            Command = New OleDb.OleDbCommand(SQLstr, Mycn)
            Command.ExecuteNonQuery()

            Mycn.Close()

            txtF3ToCopy.Text = b3to
        Else
            If b3to = txtF3ToCopy.Text Then
                'nothing changed
            Else

                Mycn.Open()
                SQLstr = "INSERT INTO rfp_ReportsTracking([ReportID], [Field], [OldValue], [NewValue], [ChangedBy], [DateStamp])" & _
                "VALUES(" & txtID.Text & ",'Fee 3 To','" & txtF3ToCopy.Text & "','" & b3to & "','" & Environ("USERNAME") & "',#" & Format(Now()) & "#)"

                Command = New OleDb.OleDbCommand(SQLstr, Mycn)
                Command.ExecuteNonQuery()
                Mycn.Close()

                txtF3ToCopy.Text = b3to
            End If
        End If

        If IsDBNull(txtF3RateCopy) Or txtF3RateCopy.Text = "" Then
            Mycn.Open()
            SQLstr = "INSERT INTO rfp_ReportsTracking([ReportID], [Field], [NewValue], [ChangedBy], [DateStamp])" & _
            "VALUES(" & txtID.Text & ",'Fee 3 Rate','" & txtF3Rate.Text & "','" & Environ("USERNAME") & "',#" & Format(Now()) & "#)"

            Command = New OleDb.OleDbCommand(SQLstr, Mycn)
            Command.ExecuteNonQuery()

            Mycn.Close()

            txtF3RateCopy.Text = txtF3Rate.Text
        Else
            If txtF3Rate.Text = txtF3RateCopy.Text Then
                'nothing changed
            Else

                Mycn.Open()
                SQLstr = "INSERT INTO rfp_ReportsTracking([ReportID], [Field], [OldValue], [NewValue], [ChangedBy], [DateStamp])" & _
                "VALUES(" & txtID.Text & ",'Fee 3 Rate','" & txtF3RateCopy.Text & "','" & txtF3Rate.Text & "','" & Environ("USERNAME") & "',#" & Format(Now()) & "#)"

                Command = New OleDb.OleDbCommand(SQLstr, Mycn)
                Command.ExecuteNonQuery()
                Mycn.Close()

                txtF3RateCopy.Text = txtF3Rate.Text
            End If
        End If

        Dim b4to As Double
        If txtF4To.Text = "" Then
            b4to = 0
        Else
            b4to = txtF4To.Text
            b4to = Format(b4to, "#.00")
        End If


        If IsDBNull(txtF4ToCopy) Or txtF4ToCopy.Text = "" Then
            Mycn.Open()
            SQLstr = "INSERT INTO rfp_ReportsTracking([ReportID], [Field], [NewValue], [ChangedBy], [DateStamp])" & _
            "VALUES(" & txtID.Text & ",'Fee 4 To','" & b4to & "','" & Environ("USERNAME") & "',#" & Format(Now()) & "#)"

            Command = New OleDb.OleDbCommand(SQLstr, Mycn)
            Command.ExecuteNonQuery()

            Mycn.Close()

            txtF4ToCopy.Text = b4to
        Else
            If b4to = txtF4ToCopy.Text Then
                'nothing changed
            Else

                Mycn.Open()
                SQLstr = "INSERT INTO rfp_ReportsTracking([ReportID], [Field], [OldValue], [NewValue], [ChangedBy], [DateStamp])" & _
                "VALUES(" & txtID.Text & ",'Fee 4 To','" & txtF4ToCopy.Text & "','" & b4to & "','" & Environ("USERNAME") & "',#" & Format(Now()) & "#)"

                Command = New OleDb.OleDbCommand(SQLstr, Mycn)
                Command.ExecuteNonQuery()
                Mycn.Close()

                txtF4ToCopy.Text = b4to
            End If
        End If

        If IsDBNull(txtF4RateCopy) Or txtF4RateCopy.Text = "" Then
            Mycn.Open()
            SQLstr = "INSERT INTO rfp_ReportsTracking([ReportID], [Field], [NewValue], [ChangedBy], [DateStamp])" & _
            "VALUES(" & txtID.Text & ",'Fee 4 Rate','" & txtF4Rate.Text & "','" & Environ("USERNAME") & "',#" & Format(Now()) & "#)"

            Command = New OleDb.OleDbCommand(SQLstr, Mycn)
            Command.ExecuteNonQuery()

            Mycn.Close()

            txtF4RateCopy.Text = txtF4Rate.Text
        Else
            If txtF4Rate.Text = txtF4RateCopy.Text Then
                'nothing changed
            Else

                Mycn.Open()
                SQLstr = "INSERT INTO rfp_ReportsTracking([ReportID], [Field], [OldValue], [NewValue], [ChangedBy], [DateStamp])" & _
                "VALUES(" & txtID.Text & ",'Fee 4 Rate','" & txtF4RateCopy.Text & "','" & txtF4Rate.Text & "','" & Environ("USERNAME") & "',#" & Format(Now()) & "#)"

                Command = New OleDb.OleDbCommand(SQLstr, Mycn)
                Command.ExecuteNonQuery()
                Mycn.Close()

                txtF4RateCopy.Text = txtF4Rate.Text
            End If
        End If

        Dim b5to As Double
        If txtF5To.Text = "" Then
            b5to = 0
        Else
            b5to = txtF5To.Text
            b5to = Format(b5to, "#.00")
        End If


        If IsDBNull(txtF5ToCopy) Or txtF5ToCopy.Text = "" Then
            Mycn.Open()
            SQLstr = "INSERT INTO rfp_ReportsTracking([ReportID], [Field], [NewValue], [ChangedBy], [DateStamp])" & _
            "VALUES(" & txtID.Text & ",'Fee 5 To','" & b5to & "','" & Environ("USERNAME") & "',#" & Format(Now()) & "#)"

            Command = New OleDb.OleDbCommand(SQLstr, Mycn)
            Command.ExecuteNonQuery()

            Mycn.Close()

            txtF5ToCopy.Text = b5to
        Else
            If b5to = txtF5ToCopy.Text Then
                'nothing changed
            Else

                Mycn.Open()
                SQLstr = "INSERT INTO rfp_ReportsTracking([ReportID], [Field], [OldValue], [NewValue], [ChangedBy], [DateStamp])" & _
                "VALUES(" & txtID.Text & ",'Fee 5 To','" & txtF5ToCopy.Text & "','" & b5to & "','" & Environ("USERNAME") & "',#" & Format(Now()) & "#)"

                Command = New OleDb.OleDbCommand(SQLstr, Mycn)
                Command.ExecuteNonQuery()
                Mycn.Close()

                txtF5ToCopy.Text = b5to
            End If
        End If

        If IsDBNull(txtF5RateCopy) Or txtF5RateCopy.Text = "" Then
            Mycn.Open()
            SQLstr = "INSERT INTO rfp_ReportsTracking([ReportID], [Field], [NewValue], [ChangedBy], [DateStamp])" & _
            "VALUES(" & txtID.Text & ",'Fee 5 Rate','" & txtF5Rate.Text & "','" & Environ("USERNAME") & "',#" & Format(Now()) & "#)"

            Command = New OleDb.OleDbCommand(SQLstr, Mycn)
            Command.ExecuteNonQuery()

            Mycn.Close()

            txtF5RateCopy.Text = txtF5Rate.Text
        Else
            If txtF5Rate.Text = txtF5RateCopy.Text Then
                'nothing changed
            Else

                Mycn.Open()
                SQLstr = "INSERT INTO rfp_ReportsTracking([ReportID], [Field], [OldValue], [NewValue], [ChangedBy], [DateStamp])" & _
                "VALUES(" & txtID.Text & ",'Fee 5 Rate','" & txtF5RateCopy.Text & "','" & txtF5Rate.Text & "','" & Environ("USERNAME") & "',#" & Format(Now()) & "#)"

                Command = New OleDb.OleDbCommand(SQLstr, Mycn)
                Command.ExecuteNonQuery()
                Mycn.Close()

                txtF5RateCopy.Text = txtF5Rate.Text
            End If
        End If

        If IsDBNull(txtF6RateCopy) Or txtF6RateCopy.Text = "" Then
            Mycn.Open()
            SQLstr = "INSERT INTO rfp_ReportsTracking([ReportID], [Field], [NewValue], [ChangedBy], [DateStamp])" & _
            "VALUES(" & txtID.Text & ",'Fee 6 Rate','" & txtF6Rate.Text & "','" & Environ("USERNAME") & "',#" & Format(Now()) & "#)"

            Command = New OleDb.OleDbCommand(SQLstr, Mycn)
            Command.ExecuteNonQuery()

            Mycn.Close()

            txtF6RateCopy.Text = txtF6Rate.Text
        Else
            If txtF6Rate.Text = txtF6RateCopy.Text Then
                'nothing changed
            Else

                Mycn.Open()
                SQLstr = "INSERT INTO rfp_ReportsTracking([ReportID], [Field], [OldValue], [NewValue], [ChangedBy], [DateStamp])" & _
                "VALUES(" & txtID.Text & ",'Fee 6 Rate','" & txtF6RateCopy.Text & "','" & txtF6Rate.Text & "','" & Environ("USERNAME") & "',#" & Format(Now()) & "#)"

                Command = New OleDb.OleDbCommand(SQLstr, Mycn)
                Command.ExecuteNonQuery()
                Mycn.Close()

                txtF6RateCopy.Text = txtF6Rate.Text
            End If
        End If

        Dim aum As Double
        If txtAUM.Text = "" Then
            aum = 0
        Else
            aum = txtAUM.Text
            aum = Format(aum, "#.00")
        End If


        If IsDBNull(txtAUMCopy) Or txtAUMCopy.Text = "" Then
            Mycn.Open()
            SQLstr = "INSERT INTO rfp_ReportsTracking([ReportID], [Field], [NewValue], [ChangedBy], [DateStamp])" & _
            "VALUES(" & txtID.Text & ",'AUM','" & aum & "','" & Environ("USERNAME") & "',#" & Format(Now()) & "#)"

            Command = New OleDb.OleDbCommand(SQLstr, Mycn)
            Command.ExecuteNonQuery()

            Mycn.Close()

            txtAUMCopy.Text = aum
        Else
            If aum = txtAUMCopy.Text Then
                'nothing changed
            Else

                Mycn.Open()
                SQLstr = "INSERT INTO rfp_ReportsTracking([ReportID], [Field], [OldValue], [NewValue], [ChangedBy], [DateStamp])" & _
                "VALUES(" & txtID.Text & ",'AUM','" & txtAUMCopy.Text & "','" & aum & "','" & Environ("USERNAME") & "',#" & Format(Now()) & "#)"

                Command = New OleDb.OleDbCommand(SQLstr, Mycn)
                Command.ExecuteNonQuery()
                Mycn.Close()

                txtAUMCopy.Text = aum
            End If
        End If

        Dim minfee As Double
        If txtMinFee.Text = "" Then
            minfee = 0
        Else
            minfee = txtMinFee.Text
            minfee = Format(minfee, "#.00")
        End If


        If IsDBNull(txtMinFeeCopy) Or txtMinFeeCopy.Text = "" Then
            Mycn.Open()
            SQLstr = "INSERT INTO rfp_ReportsTracking([ReportID], [Field], [NewValue], [ChangedBy], [DateStamp])" & _
            "VALUES(" & txtID.Text & ",'Minimum Fee','" & minfee & "','" & Environ("USERNAME") & "',#" & Format(Now()) & "#)"

            Command = New OleDb.OleDbCommand(SQLstr, Mycn)
            Command.ExecuteNonQuery()

            Mycn.Close()

            txtMinFeeCopy.Text = minfee
        Else
            If minfee = txtMinFeeCopy.Text Then
                'nothing changed
            Else

                Mycn.Open()
                SQLstr = "INSERT INTO rfp_ReportsTracking([ReportID], [Field], [OldValue], [NewValue], [ChangedBy], [DateStamp])" & _
                "VALUES(" & txtID.Text & ",'Minimum Fee','" & txtMinFeeCopy.Text & "','" & minfee & "','" & Environ("USERNAME") & "',#" & Format(Now()) & "#)"

                Command = New OleDb.OleDbCommand(SQLstr, Mycn)
                Command.ExecuteNonQuery()
                Mycn.Close()

                txtMinFeeCopy.Text = minfee
            End If
        End If

        If IsDBNull(txtImageLocationCopy) Or txtImageLocationCopy.Text = "" Then
            Mycn.Open()
            SQLstr = "INSERT INTO rfp_ReportsTracking([ReportID], [Field], [NewValue], [ChangedBy], [DateStamp])" & _
            "VALUES(" & txtID.Text & ",'Image Location','" & txtImageLocation.Text & "','" & Environ("USERNAME") & "',#" & Format(Now()) & "#)"

            Command = New OleDb.OleDbCommand(SQLstr, Mycn)
            Command.ExecuteNonQuery()

            Mycn.Close()

            txtImageLocationCopy.Text = txtImageLocation.Text
        Else
            If txtImageLocation.Text = txtImageLocationCopy.Text Then
                'nothing changed
            Else

                Mycn.Open()
                SQLstr = "INSERT INTO rfp_ReportsTracking([ReportID], [Field], [OldValue], [NewValue], [ChangedBy], [DateStamp])" & _
                "VALUES(" & txtID.Text & ",'Image Location','" & txtImageLocationCopy.Text & "','" & txtImageLocation.Text & "','" & Environ("USERNAME") & "',#" & Format(Now()) & "#)"

                Command = New OleDb.OleDbCommand(SQLstr, Mycn)
                Command.ExecuteNonQuery()
                Mycn.Close()

                txtImageLocationCopy.Text = txtImageLocation.Text
            End If
        End If

        If ckbHasImageCopy.Checked = False Then
            Mycn.Open()
            SQLstr = "INSERT INTO rfp_ReportsTracking([ReportID], [Field], [NewValue], [ChangedBy], [DateStamp])" & _
            "VALUES(" & txtID.Text & ",'Has Image','" & ckbHasImage.CheckState & "','" & Environ("USERNAME") & "',#" & Format(Now()) & "#)"

            Command = New OleDb.OleDbCommand(SQLstr, Mycn)
            Command.ExecuteNonQuery()

            Mycn.Close()

            ckbHasImageCopy.Checked = ckbHasImage.Checked
        Else
            If ckbHasImage.Checked = ckbHasImageCopy.Checked Then
                'nothing changed
            Else

                Mycn.Open()
                SQLstr = "INSERT INTO rfp_ReportsTracking([ReportID], [Field], [OldValue], [NewValue], [ChangedBy], [DateStamp])" & _
                "VALUES(" & txtID.Text & ",'Has Image','" & ckbHasImageCopy.CheckState & "','" & ckbHasImage.CheckState & "','" & Environ("USERNAME") & "',#" & Format(Now()) & "#)"

                Command = New OleDb.OleDbCommand(SQLstr, Mycn)
                Command.ExecuteNonQuery()
                Mycn.Close()

                ckbHasImageCopy.Checked = ckbHasImage.Checked
            End If
        End If

        If ckbDisclosureCopy.Checked = False Then
            Mycn.Open()
            SQLstr = "INSERT INTO rfp_ReportsTracking([ReportID], [Field], [NewValue], [ChangedBy], [DateStamp])" & _
            "VALUES(" & txtID.Text & ",'Default Disclosure','" & ckbDisclosure.CheckState & "','" & Environ("USERNAME") & "',#" & Format(Now()) & "#)"

            Command = New OleDb.OleDbCommand(SQLstr, Mycn)
            Command.ExecuteNonQuery()

            Mycn.Close()

            ckbDisclosureCopy.Checked = ckbDisclosure.Checked
        Else
            If ckbDisclosure.Checked = ckbDisclosureCopy.Checked Then
                'nothing changed
            Else

                Mycn.Open()
                SQLstr = "INSERT INTO rfp_ReportsTracking([ReportID], [Field], [OldValue], [NewValue], [ChangedBy], [DateStamp])" & _
                "VALUES(" & txtID.Text & ",'Default Disclosure','" & ckbDisclosureCopy.CheckState & "','" & ckbDisclosure.CheckState & "','" & Environ("USERNAME") & "',#" & Format(Now()) & "#)"

                Command = New OleDb.OleDbCommand(SQLstr, Mycn)
                Command.ExecuteNonQuery()
                Mycn.Close()

                ckbDisclosureCopy.Checked = ckbDisclosure.Checked
            End If
        End If

        Dim disc1 As String
        disc1 = rtbDisclosure.Text
        disc1 = Replace(disc1, "'", "''")

        Dim disc2 As String
        disc2 = rtbDisclosureCopy.Text
        disc2 = Replace(disc2, "'", "''")

        If IsDBNull(rtbDisclosureCopy) Or rtbDisclosureCopy.Text = "" Then
            Mycn.Open()
            SQLstr = "INSERT INTO rfp_ReportsTracking([ReportID], [Field], [NewValue], [ChangedBy], [DateStamp])" & _
            "VALUES(" & txtID.Text & ",'Disclosure','" & disc1 & "','" & Environ("USERNAME") & "',#" & Format(Now()) & "#)"

            Command = New OleDb.OleDbCommand(SQLstr, Mycn)
            Command.ExecuteNonQuery()

            Mycn.Close()

            rtbDisclosureCopy.Text = rtbDisclosure.Text
        Else
            If rtbDisclosure.Text = rtbDisclosureCopy.Text Then
                'nothing changed
            Else

                Mycn.Open()
                SQLstr = "INSERT INTO rfp_ReportsTracking([ReportID], [Field], [OldValue], [NewValue], [ChangedBy], [DateStamp])" & _
                "VALUES(" & txtID.Text & ",'Disclosure','" & disc2 & "','" & disc1 & "','" & Environ("USERNAME") & "',#" & Format(Now()) & "#)"

                Command = New OleDb.OleDbCommand(SQLstr, Mycn)
                Command.ExecuteNonQuery()
                Mycn.Close()

                rtbDisclosureCopy.Text = rtbDisclosure.Text
            End If
        End If

        'Catch ex As Exception
        'MsgBox(ex.Message, MsgBoxStyle.Critical, "ERROR")

        'End Try

line1:

    End Sub

    Public Sub LoadID()
        Dim Mycn As OleDb.OleDbConnection

        'Try

        Mycn = New OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & My.Settings.roledbloc & "';Jet OLEDB:Database Password='" & My.Settings.dbpass & "';")

        Mycn.Open()

        Dim sqlstring As String

        Dim title As String
        title = txtRptName.Text
        title = Replace(title, "'", "''")

        sqlstring = "SELECT Top 1 ID FROM rfp_Reports WHERE Title = '" & title & "'" & _
        " GROUP BY ID" & _
        " ORDER BY ID DESC"

        Dim queryString As String = String.Format(sqlstring)
        Dim cmd As New OleDb.OleDbCommand(queryString, Mycn)
        Dim da As New OleDb.OleDbDataAdapter(cmd)
        Dim ds As New DataSet

        da.Fill(ds, "User")
        Dim dt As DataTable = ds.Tables("User")
        If dt.Rows.Count > 0 Then

            Dim row As DataRow = dt.Rows(0)

            txtID.Text = (row("ID"))

            Call LoadNewAnsw()
        Else

        End If
        Mycn.Close()
        Mycn.Dispose()

        Call TrackChanges()

        'Catch ex As Exception
        'MsgBox(ex.Message, MsgBoxStyle.Critical, "ERROR")
        'End Try
    End Sub

    Private Sub cmdAdd_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdAdd.Click
        If txtID.Text = "NEW" Then
            Dim ir As MsgBoxResult
            ir = MsgBox("You Must save this record before adding questions.  Would you like to save now?", MsgBoxStyle.YesNo + MsgBoxStyle.Information, "Save Record")

            If ir = MsgBoxResult.Yes Then
                Call InsertRec()
                Dim child As New rfp_ReportQuestion
                child.MdiParent = Home
                child.txtID.Text = "NEW"
                child.txtRFPID.Text = Me.txtID.Text
                child.Show()
                Call child.LoadSortOrders()
            Else
                GoTo line1
            End If
        Else
            Dim child As New rfp_ReportQuestion
            child.MdiParent = Home
            child.txtID.Text = "NEW"
            child.txtRFPID.Text = Me.txtID.Text
            child.Show()
            Call child.LoadSortOrders()
        End If

line1:
    End Sub

    Private Sub CheckBox1_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckBox1.CheckedChanged
        If IsDBNull(txtDateStarted) Or txtDateStarted.Text = "" Then
            If CheckBox1.Checked Then
                txtDateStarted.Text = Format(Now())
                txtWorkedBy.Text = Environ("USERNAME")
            Else
                txtDateStarted.Text = ""
            End If
        Else

        End If

        If CheckBox1.Checked Then
            cboStage.Enabled = True
            cboReason.Enabled = True
        Else
            cboStage.Enabled = False
            cboReason.Enabled = False
        End If

    End Sub

    Public Sub LoadNewAnsw()
        If txtID.Text = "NEW" Then
            'do nothing
        Else
            Try

                Dim conn As New OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & My.Settings.dbloc & "';Jet OLEDB:Database Password='" & My.Settings.dbpass & "';")

                Dim strSQL As String = "SELECT ID, SortOrder, Question FROM rfp_ReportQuestions" & _
                " WHERE IsActive = -1 AND ReportID = " & txtID.Text & " AND StageID = 1" & _
                " ORDER BY SortOrder"
                Dim da As New OleDb.OleDbDataAdapter(strSQL, conn)
                Dim ds As New DataSet
                da.Fill(ds, "Users")

                With dgvNew
                    .DataSource = ds.Tables("Users")
                    .Columns(0).Visible = False
                End With
                dgvNew.Visible = True
                'Label3.Visible = False

            Catch ex As Exception
                MsgBox(ex.Message, MsgBoxStyle.Critical, "ERROR")
            End Try
        End If
    End Sub

    Public Sub LoadWorkingAnsw()
        If txtID.Text = "NEW" Then
        Else


            Try

                Dim conn As New OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & My.Settings.dbloc & "';Jet OLEDB:Database Password='" & My.Settings.dbpass & "';")

                Dim strSQL As String = "SELECT ID, SortOrder, Question FROM rfp_ReportQuestions" & _
                " WHERE IsActive = -1 AND ReportID = " & txtID.Text & " AND StageID <> 5 AND StageID <> 1" & _
                " ORDER BY SortOrder"

                Dim da As New OleDb.OleDbDataAdapter(strSQL, conn)
                Dim ds As New DataSet
                da.Fill(ds, "Users")

                With DataGridView1
                    .DataSource = ds.Tables("Users")
                    .Columns(0).Visible = False
                End With
                'dgvNew.Visible = True
                'Label3.Visible = False

            Catch ex As Exception
                MsgBox(ex.Message, MsgBoxStyle.Critical, "ERROR")
            End Try
        End If
    End Sub

    Public Sub LoadFinishedAnsw()
        If txtID.Text = "NEW" Then
            'do nothing
        Else


            Try
                Dim conn As New OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & My.Settings.dbloc & "';Jet OLEDB:Database Password='" & My.Settings.dbpass & "';")

                Dim strSQL As String = "SELECT ID, SortOrder, Question FROM rfp_ReportQuestions" & _
                " WHERE IsActive = -1 AND ReportID = " & txtID.Text & " AND StageID = 5" & _
                " ORDER BY SortOrder"
                Dim da As New OleDb.OleDbDataAdapter(strSQL, conn)
                Dim ds As New DataSet
                da.Fill(ds, "Users")

                With DataGridView2
                    .DataSource = ds.Tables("Users")
                    .Columns(0).Visible = False
                End With
                'dgvNew.Visible = True
                'Label3.Visible = False

            Catch ex As Exception
                MsgBox(ex.Message, MsgBoxStyle.Critical, "ERROR")
            End Try
        End If
    End Sub

    Private Sub cmdRefresh_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdRefresh.Click
        Call LoadNewAnsw()
        Call LoadWorkingAnsw()
        Call LoadFinishedAnsw()
    End Sub

    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        If txtID.Text = "NEW" Then

        Else
            Dim child As New rfp_ReportsChangeLog
            child.MdiParent = Home
            child.Show()
            child.txtID.Text = txtID.Text
            Try

                Dim conn As New OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & My.Settings.dbloc & "';Jet OLEDB:Database Password='" & My.Settings.dbpass & "';")

                Dim strSQL As String = "SELECT ID, Field, OldValue As [Old Value], NewValue As [New Value], ChangedBy As [User], DateStamp As [Date Stamp] FROM rfp_ReportsTracking WHERE ReportID = " & txtID.Text
                Dim da As New OleDb.OleDbDataAdapter(strSQL, conn)
                Dim ds As New DataSet
                da.Fill(ds, "Users")

                With child.DataGridView1
                    .DataSource = ds.Tables("Users")
                    '.Columns(0).Visible = False
                End With
                child.DataGridView1.Visible = True

            Catch ex As Exception
                MsgBox(ex.Message, MsgBoxStyle.Critical, "ERROR")
                Label3.Text = "ERROR"
            End Try
        End If


    End Sub

    Private Sub WorkToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles WorkToolStripMenuItem.Click
        If Permissions.RFPQuestionView.Checked Then

        Else
            MsgBox("You do not have permission to perform this function.", MsgBoxStyle.Exclamation, "Permission Issue")
            GoTo line1
        End If

        If dgvNew.RowCount < 1 Then

        Else
            Dim child As New rfp_Answers
            child.MdiParent = Home
            child.Show()
            child.txtID.Text = dgvNew.SelectedCells(0).Value
            child.txtRptID.Text = txtID.Text
            child.rtbQuestion.Text = dgvNew.SelectedCells(2).Value
            Call child.LoadForm()
        End If

line1:
    End Sub

    Private Sub EditToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub

    Private Sub dgvNew_CellContentClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvNew.CellContentClick

    End Sub

    Private Sub dgvNew_CellContentDoubleClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvNew.CellContentDoubleClick
        Dim child As New rfp_Answers
        child.MdiParent = Home
        child.Show()
        child.txtID.Text = dgvNew.SelectedCells(0).Value
        child.txtRptID.Text = txtID.Text
        child.rtbQuestion.Text = dgvNew.SelectedCells(2).Value
        Call child.LoadForm()
    End Sub

    Private Sub DataGridView1_CellContentClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick

    End Sub

    Private Sub DataGridView1_CellContentDoubleClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView1.CellContentDoubleClick
        Dim child As New rfp_Answers
        child.MdiParent = Home
        child.Show()
        child.txtID.Text = DataGridView1.SelectedCells(0).Value
        child.txtRptID.Text = txtID.Text
        child.rtbQuestion.Text = DataGridView1.SelectedCells(2).Value
        Call child.LoadForm()
    End Sub

    Private Sub DataGridView2_CellContentClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView2.CellContentClick

    End Sub

    Private Sub DataGridView2_CellContentDoubleClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView2.CellContentDoubleClick
        Dim child As New rfp_Answers
        child.MdiParent = Home
        child.Show()
        child.txtID.Text = DataGridView2.SelectedCells(0).Value
        child.txtRptID.Text = txtID.Text
        child.rtbQuestion.Text = DataGridView2.SelectedCells(2).Value
        Call child.LoadForm()
    End Sub

    Private Sub DeleteToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles DeleteToolStripMenuItem.Click
        If Permissions.RFPQuestionDelete.Checked Then

        Else
            MsgBox("You do not have permission to perform this function.", MsgBoxStyle.Exclamation, "Permission Issue")
            GoTo line1
        End If

        If dgvNew.RowCount > 0 Then

            Dim ir As MsgBoxResult
            ir = MsgBox("Are you sure you want to delete this question?", MsgBoxStyle.YesNo + MsgBoxStyle.Critical, "Confirm Delete")
            If ir = MsgBoxResult.Yes Then

                Dim Mycn As OleDb.OleDbConnection
                Dim Command As OleDb.OleDbCommand
                Dim SQLstr As String

                Try
                    Mycn = New OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & My.Settings.roledbloc & "';Jet OLEDB:Database Password='" & My.Settings.dbpass & "';")

                    Mycn.Open()
                    SQLstr = "Update rfp_ReportQuestions SET [IsActive] = False WHERE ID = " & dgvNew.SelectedCells(0).Value
                    Command = New OleDb.OleDbCommand(SQLstr, Mycn)
                    Command.ExecuteNonQuery()
                    Mycn.Close()

                    Mycn.Open()
                    SQLstr = "INSERT INTO rfp_ResponseTracking([QuestionID], [Field], [NewValue], [ChangedBy], [DateStamp])" & _
                    "VALUES(" & dgvNew.SelectedCells(0).Value & ",'Question Deleted','0','" & Environ("USERNAME") & "',#" & Format(Now()) & "#)"
                    Command = New OleDb.OleDbCommand(SQLstr, Mycn)
                    Command.ExecuteNonQuery()
                    Mycn.Close()

                    Call LoadNewAnsw()

                Catch ex As Exception
                    MsgBox(ex.Message, MsgBoxStyle.Critical, "ERROR")
                End Try
            Else
            End If
        Else
        End If

line1:

    End Sub

    Private Sub ChangeSortToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ChangeSortToolStripMenuItem.Click
        If Permissions.RFPQuestionDelete.Checked Then

        Else
            MsgBox("You do not have permission to perform this function.", MsgBoxStyle.Exclamation, "Permission Issue")
            GoTo line1
        End If

        If dgvNew.RowCount > 0 Then
            Dim child As New rfp_QuestionSortOrder
            child.MdiParent = Home
            child.Show()
            child.txtID.Text = dgvNew.SelectedCells(0).Value
            child.txtOldSort.Text = dgvNew.SelectedCells(1).Value
            child.txtNewSort.Text = dgvNew.SelectedCells(1).Value
        Else
        End If

line1:

    End Sub

    Private Sub DateTimePicker1_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles DateTimePicker1.ValueChanged
        Dim dte As Date
        dte = DateTimePicker1.Text
        Dim dte1 As Date
        dte1 = Format(Now())
        Dim dayslft As Integer
        dayslft = DateDiff(DateInterval.Day, dte1, dte)
        If dayslft < 0 Then
            lblDaysLeft.ForeColor = Color.Red
            lblDaysLeft.Text = "This RFP is " & -dayslft & " Day(s) Past Due."
        Else
            If dayslft < 5 Then
                lblDaysLeft.ForeColor = Color.Red
                lblDaysLeft.Text = dayslft & " Day(s) Left before Due Date."
            Else
                lblDaysLeft.ForeColor = Color.Black
                lblDaysLeft.Text = dayslft & " Day(s) Left before Due Date."
            End If
        End If
    End Sub

    Private Sub TabPage1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TabPage1.Click

    End Sub

    Private Sub Button6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button6.Click
        If IsDBNull(txtEmail) Or txtEmail.Text = "" Then
            MsgBox("Please enter an email address.", MsgBoxStyle.Information, "No Email Address")
        Else
            Dim oApp As Microsoft.Office.Interop.Outlook._Application
            oApp = New Microsoft.Office.Interop.Outlook.Application

            Dim oMsg As Microsoft.Office.Interop.Outlook._MailItem
            oMsg = oApp.CreateItem(Microsoft.Office.Interop.Outlook.OlItemType.olMailItem)

            Dim msgto As String
            msgto = txtEmail.Text

            Dim sbjt As String
            sbjt = "Regarding RFP Titled: " & txtRptName.Text

            oMsg.Subject = sbjt

            oMsg.To = msgto
            oMsg.Display(True)
        End If

    End Sub

    Private Sub Button5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button5.Click
        If IsDBNull(txtWebsite) Or txtWebsite.Text = "" Then
            MsgBox("Please enter a website to open.", MsgBoxStyle.Information, "No Website")
        Else
            Dim webaddress As String = txtWebsite.Text
            Process.Start(webaddress)
        End If
    End Sub

    Private Sub cboContact_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles cboContact.LostFocus
        Call LoadFirmData()
    End Sub

    Private Sub Button7_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button7.Click
        Dim ir As New MsgBoxResult
        ir = MsgBox("Are you sure you want to update all contact fields with APX Data?  This will overwrite any manually updated fields.", MsgBoxStyle.YesNo + MsgBoxStyle.Information, "Pull APX Data?")
        If ir = MsgBoxResult.Yes Then
            ckbLockData.Checked = False
            Call LoadFirmData()
        Else

        End If
    End Sub

    Private Sub txtF1To_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtF1To.LostFocus
        Dim val As Double
        val = txtF1To.Text
        txtF1To.Text = Format(val, "$#,###.00")
    End Sub

    Private Sub txtF1To_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtF1To.TextChanged
        Dim f2frm As Double
        If txtF1To.Text = "" Then
            txtF2From.Text = ""
        Else
            f2frm = txtF1To.Text + 0.01
            txtF2From.Text = f2frm
        End If

        Call estrev()
    End Sub

    Private Sub txtF2To_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtF2To.LostFocus
        Dim val As Double
        val = txtF2To.Text
        txtF2To.Text = Format(val, "$#,###.00")
    End Sub

    Private Sub txtF2To_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtF2To.TextChanged
        Dim ffrm As Double
        If txtF2To.Text = "" Then
            txtF3From.Text = ""
        Else
            ffrm = txtF2To.Text + 0.01
            txtF3From.Text = ffrm
        End If

        Call estrev()
    End Sub

    Private Sub txtF3To_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtF3To.LostFocus
        Dim val As Double
        val = txtF3To.Text
        txtF3To.Text = Format(val, "$#,###.00")
    End Sub

    Private Sub txtF3To_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtF3To.TextChanged
        Dim ffrm As Double
        If txtF3To.Text = "" Then
            txtF4From.Text = ""
        Else
            ffrm = txtF3To.Text + 0.01
            txtF4From.Text = ffrm
        End If

        Call estrev()
    End Sub

    Private Sub txtF4To_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtF4To.LostFocus
        Dim val As Double
        val = txtF4To.Text
        txtF4To.Text = Format(val, "$#,###.00")
    End Sub

    Private Sub txtF4To_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtF4To.TextChanged
        Dim ffrm As Double
        If txtF4To.Text = "" Then
            txtF5From.Text = ""
        Else
            ffrm = txtF4To.Text + 0.01
            txtF5From.Text = ffrm
        End If

        Call estrev()
    End Sub

    Private Sub txtF5To_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtF5To.LostFocus
        Dim val As Double
        val = txtF5To.Text
        txtF5To.Text = Format(val, "$#,###.00")
    End Sub

    Private Sub txtF5To_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtF5To.TextChanged
        Dim ffrm As Double
        If txtF5To.Text = "" Then
            txtF6From.Text = ""
        Else
            ffrm = txtF5To.Text + 0.01
            txtF6From.Text = ffrm
        End If

        Call estrev()
    End Sub

    Private Sub txtAUM_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtAUM.LostFocus
        Dim val As Double
        val = txtAUM.Text
        txtAUM.Text = Format(val, "$#,###.00")
    End Sub

    Private Sub txtAUM_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtAUM.TextChanged
        Call estrev()
    End Sub

    Private Sub txtF1Rate_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtF1Rate.TextChanged
        Call estrev()
    End Sub

    Private Sub txtF2Rate_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtF2Rate.TextChanged
        Call estrev()
    End Sub

    Private Sub txtF3Rate_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtF3Rate.TextChanged
        Call estrev()
    End Sub

    Private Sub txtF4Rate_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtF4Rate.TextChanged
        Call estrev()
    End Sub

    Private Sub txtF5Rate_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtF5Rate.TextChanged
        Call estrev()
    End Sub

    Private Sub txtF6Rate_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtF6Rate.TextChanged
        Call estrev()
    End Sub

    Private Sub txtMinFee_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtMinFee.LostFocus
        Dim val As Double
        val = txtMinFee.Text
        txtMinFee.Text = Format(val, "$#,###.00")
    End Sub

    Private Sub txtMinFee_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtMinFee.TextChanged
        Call estrev()
    End Sub

    Private Sub ckbFlatRate_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ckbFlatRate.CheckedChanged
        If ckbFlatRate.Checked = True Then
            txtFlatRate.Enabled = True
            txtF1To.Enabled = False
            txtF1Rate.Enabled = False
            txtF2To.Enabled = False
            txtF2Rate.Enabled = False
            txtF3To.Enabled = False
            txtF3Rate.Enabled = False
            txtF4To.Enabled = False
            txtF4Rate.Enabled = False
            txtF5To.Enabled = False
            txtF5Rate.Enabled = False
            txtF6Rate.Enabled = False
            Call estrev()
        Else
            txtFlatRate.Enabled = False
            txtF1To.Enabled = True
            txtF1Rate.Enabled = True
            txtF2To.Enabled = True
            txtF2Rate.Enabled = True
            txtF3To.Enabled = True
            txtF3Rate.Enabled = True
            txtF4To.Enabled = True
            txtF4Rate.Enabled = True
            txtF5To.Enabled = True
            txtF5Rate.Enabled = True
            txtF6Rate.Enabled = True
            Call estrev()
        End If
    End Sub

    Private Sub txtFlatRate_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Call estrev()
    End Sub

    Private Sub txtFlatRate_MaskInputRejected(ByVal sender As System.Object, ByVal e As System.Windows.Forms.MaskInputRejectedEventArgs)

    End Sub

    Private Sub txtFlatRate_TextChanged1(ByVal sender As Object, ByVal e As System.EventArgs)
        Call estrev()
    End Sub

    Private Sub TabPage3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TabPage3.Click

    End Sub

    Private Sub Button9_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button9.Click
        Call loadimage()

    End Sub

    Public Sub loadimage()
        If txtImageLocation.Text = "" Then
        Else
            PictureBox1.Image = Image.FromFile(txtImageLocation.Text)
        End If
    End Sub

    Private Sub Button8_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button8.Click
        'OpenFileDialog1.Title = "Please Select an Image"
        'OpenFileDialog1.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.Desktop) '"C:"
        'OpenFileDialog1.DefaultExt = "jpg"
        'OpenFileDialog1.Filter = "Image files (*.bmp, *jpg, *png)|*.bmp;*jpg;*png|All files (*.*)|*.*"
        'Dim fdate As Date
        'fdate = DateTimePicker1.Text
        'OpenFileDialog1.FileName = "account_positions_" & Format(fdate, "yyyy") & "_" & Format(fdate, "MM") & "_" & Format(fdate, "dd") & ".csv"
        'OpenFileDialog1.ShowDialog()
        If txtID.Text = "NEW" Then
            MsgBox("You must save this record before adding an image.", MsgBoxStyle.Information, "Record not saved")
        Else


            Dim SAVE_PATH As String = "\\monumentco1\data\RFPImages"

            Dim FileDialog As New OpenFileDialog

            Dim fso = My.Computer.FileSystem



            With FileDialog

                If Not fso.DirectoryExists(SAVE_PATH) Then

                    Try

                        fso.CreateDirectory(SAVE_PATH)

                    Catch ex As Exception

                        MessageBox.Show("Unable to create folder '" & SAVE_PATH.ToLower & _
            "'.  Images will be saved in '" & Application.StartupPath.ToLower & _
            "'.", Text, MessageBoxButtons.OK, MessageBoxIcon.Warning)

                        SAVE_PATH = Application.StartupPath

                    End Try

                End If



                '.InitialDirectory = SAVE_PATH
                .InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.Desktop) '"C:"
                .Filter = "All Graphic Files|*.bmp;*.gif;*.jpg;*.jpeg;*.png;|" & _
                "Graphic Interchange Format (*.gif)|*.gif|" & _
                "Portable Network Graphics (*.png)|*.png|" & _
                "JPEG File Interchange Format (*.jpg;*.jpeg)|*.jpg;*.jpeg|" & _
                "Windows Bitmap (*.bmp)|*.bmp"

                .FilterIndex = 1

                .FileName = ""



                If .ShowDialog = Windows.Forms.DialogResult.OK Then
                    Dim extention As String
                    extention = Path.GetExtension(.FileName)
                    Dim nfilename As String
                    nfilename = txtID.Text & extention
                    If .FileName.ToUpper = (SAVE_PATH & "\" & nfilename.ToUpper) Then

                        MessageBox.Show("The file '" & nfilename.ToLower & "' cannot be copied onto itself.", _
                                        Text, MessageBoxButtons.OK, MessageBoxIcon.Error)

                        Exit Sub

                    ElseIf fso.FileExists(SAVE_PATH & "\" & txtID.Text & "." & extention) Then

                        Dim Response As DialogResult

                        Response = MessageBox.Show("The file '" & nfilename.ToLower & "' already exist in the destination folder '" & _
            SAVE_PATH & "'" & vbCrLf & vbCrLf & "Do you want to overwite it?", Text, MessageBoxButtons.YesNo, _
            MessageBoxIcon.Question, MessageBoxDefaultButton.Button2)


                        If Response = Windows.Forms.DialogResult.No Then Exit Sub

                    End If

                    Try

                        fso.CopyFile(.FileName, SAVE_PATH & "\" & nfilename, True)
                        txtImageLocation.Text = SAVE_PATH & "\" & nfilename
                        Call loadimage()
                    Catch ex As Exception

                        MessageBox.Show(ex.Message, Text, MessageBoxButtons.OK, MessageBoxIcon.Error)

                    End Try

                End If

            End With

        End If
    End Sub

    Private Sub OpenFileDialog1_FileOk(ByVal sender As System.Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles OpenFileDialog1.FileOk
        Dim strm As System.IO.Stream
        strm = OpenFileDialog1.OpenFile()

        txtImageLocation.Text = OpenFileDialog1.FileName.ToString

        If Not (strm Is Nothing) Then
            'insert code to read the file data
            strm.Close()
        End If
    End Sub

    Private Sub ckbHasImage_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ckbHasImage.CheckedChanged
        If ckbHasImage.Checked = True Then
            Button8.Enabled = True
            Button9.Enabled = True
        Else
            Button9.Enabled = False
            Button8.Enabled = False
        End If
    End Sub

    Private Sub Button10_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button10.Click
        If rtbDisclosure.Text.Length > 0 Then
            Dim wordapp As New Word.Application
            wordapp.Visible = False
            Dim doc As Word.Document = wordapp.Documents.Add
            Dim range As Word.Range
            range = doc.Range
            range.Text = rtbDisclosure.Text
            doc.Activate()
            doc.CheckSpelling()
            Dim chars() As Char = {CType(vbCr, Char), CType(vbLf, Char)}
            rtbDisclosure.Text = doc.Range().Text.Trim(chars)
            doc.Close(SaveChanges:=False)
            wordapp.Quit()
            MsgBox("Spell Check Finished.", MsgBoxStyle.Information, "Done")
        End If
    End Sub

    Private Sub ckbDisclosure_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ckbDisclosure.CheckedChanged
        If ckbDisclosure.Checked = False Then
            rtbDisclosure.Enabled = True
            Button18.Enabled = True
            Button10.Enabled = True
            Button19.Enabled = True
            Call pulldisclosure()
        Else
            rtbDisclosure.Enabled = False
            Button18.Enabled = False
            Button10.Enabled = False
            Button19.Enabled = False
            Call pulldisclosure()
        End If
    End Sub

    Private Sub Button12_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button12.Click
        Call loadmktassoc()
        Call loadmktaval()
    End Sub

    Public Sub loadmktaval()
        Try

            Dim conn As New OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & My.Settings.dbloc & "';Jet OLEDB:Database Password='" & My.Settings.dbpass & "';")

            Dim strSQL As String = "SELECT rfp_MarketingMaterials.ID, rfp_MarketingMaterials.ItemName, rfp_MarketingMaterials.ItemDescription, rfp_MarketingMaterials.LastUpdate" & _
            " FROM(rfp_MarketingMaterials)" & _
            " WHERE (((rfp_MarketingMaterials.ID) Not In (Select MarketingID FROM rfp_MarketingAssociation WHERE RFPID = " & txtID.Text & ")) AND ((rfp_MarketingMaterials.IsActive)=True))" & _
            " GROUP BY rfp_MarketingMaterials.ID, rfp_MarketingMaterials.ItemName, rfp_MarketingMaterials.ItemDescription, rfp_MarketingMaterials.LastUpdate;"


            Dim da As New OleDb.OleDbDataAdapter(strSQL, conn)
            Dim ds As New DataSet
            da.Fill(ds, "Users")

            With dgvMarketingAvailable
                .DataSource = ds.Tables("Users")
                .Columns(0).Visible = False
            End With

            Label55.Text = "Available (" & dgvMarketingAvailable.RowCount.ToString & " records):"

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "ERROR")
        End Try
    End Sub

    Public Sub loadmktassoc()
        Try

            Dim conn As New OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & My.Settings.dbloc & "';Jet OLEDB:Database Password='" & My.Settings.dbpass & "';")

            Dim strSQL As String = "SELECT rfp_MarketingAssociation.ID, rfp_MarketingAssociation.SortOrder, rfp_MarketingMaterials.ItemName, rfp_MarketingMaterials.ItemDescription" & _
            " FROM rfp_MarketingMaterials INNER JOIN rfp_MarketingAssociation ON rfp_MarketingMaterials.ID = rfp_MarketingAssociation.MarketingID" & _
            " WHERE(((rfp_MarketingAssociation.RFPID) = " & txtID.Text & "))" & _
            " GROUP BY rfp_MarketingAssociation.ID, rfp_MarketingAssociation.SortOrder, rfp_MarketingMaterials.ItemName, rfp_MarketingMaterials.ItemDescription" & _
            " ORDER BY rfp_MarketingAssociation.SortOrder;"


            Dim da As New OleDb.OleDbDataAdapter(strSQL, conn)
            Dim ds As New DataSet
            da.Fill(ds, "Users")

            With dgvMarketingAssociated
                .DataSource = ds.Tables("Users")
                .Columns(0).Visible = False
            End With

            Label54.Text = "Associated (" & dgvMarketingAssociated.RowCount.ToString & " records):"

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "ERROR")
        End Try
    End Sub

    Private Sub Button18_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button18.Click
        If txtDisclosureID.Text = "NEW" Then
            Call InsertDisclosure()

        Else
            Call UpdateDisclosure()
        End If
        MsgBox("Disclosure Saved.", MsgBoxStyle.Information, "Saved")
    End Sub

    Public Sub InsertDisclosure()
        Dim Mycn As OleDb.OleDbConnection
        Dim Command As OleDb.OleDbCommand
        Dim SQLstr As String

        Try
            Mycn = New OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & My.Settings.roledbloc & "';Jet OLEDB:Database Password='" & My.Settings.dbpass & "';")

            Dim ds1 As New DataSet
            Dim eds1 As New DataGridView
            Dim dv1 As New DataView

            Mycn.Open()

            Dim disc As String
            disc = rtbDisclosure.Text
            disc = Replace(disc, "'", "''")

            SQLstr = "INSERT INTO rfp_Disclosures ([Disclosure],[SetBy],[LastUpdate],[RFPID])" & _
            "VALUES('" & disc & "'," & My.Settings.userid & ",#" & Format(Now()) & "#," & txtID.Text & ")"

            Command = New OleDb.OleDbCommand(SQLstr, Mycn)
            Command.ExecuteNonQuery()

            Mycn.Close()

            Call pulldisclosure()

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "ERROR")
        End Try
    End Sub

    Public Sub UpdateDisclosure()
        Dim Mycn As OleDb.OleDbConnection
        Dim Command As OleDb.OleDbCommand
        Dim SQLstr As String

        Try
            Mycn = New OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & My.Settings.roledbloc & "';Jet OLEDB:Database Password='" & My.Settings.dbpass & "';")

            Dim ds1 As New DataSet
            Dim eds1 As New DataGridView
            Dim dv1 As New DataView

            Mycn.Open()

            Dim disc As String
            disc = rtbDisclosure.Text
            disc = Replace(disc, "'", "''")

            SQLstr = "UPDATE rfp_Disclosures SET [Disclosure]='" & disc & "',[SetBy] = " & My.Settings.userid & ",[LastUpdate] = #" & Format(Now()) & "#,[RFPID] = " & txtID.Text & " WHERE ID = " & txtDisclosureID.Text

            Command = New OleDb.OleDbCommand(SQLstr, Mycn)
            Command.ExecuteNonQuery()

            Mycn.Close()

            Call pulldisclosure()

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "ERROR")
        End Try
    End Sub

    Private Sub Button19_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button19.Click
        If txtDisclosureID.Text = "NEW" Then
        Else
            Dim ir As MsgBoxResult
            ir = MsgBox("Are you sure you want to delete this disclosure?", MsgBoxStyle.YesNo + MsgBoxStyle.Information, "Confirm")
            If ir = MsgBoxResult.Yes Then
                Dim Mycn As OleDb.OleDbConnection
                Dim Command As OleDb.OleDbCommand
                Dim SQLstr As String

                Try
                    Mycn = New OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & My.Settings.roledbloc & "';Jet OLEDB:Database Password='" & My.Settings.dbpass & "';")
                    Dim ds1 As New DataSet
                    Dim eds1 As New DataGridView
                    Dim dv1 As New DataView
                    Mycn.Open()
                    SQLstr = "DELETE * FROM rfp_Disclosures WHERE ID = " & txtDisclosureID.Text
                    Command = New OleDb.OleDbCommand(SQLstr, Mycn)
                    Command.ExecuteNonQuery()
                    Mycn.Close()
                    Call pulldisclosure()
                Catch ex As Exception
                    MsgBox(ex.Message, MsgBoxStyle.Critical, "ERROR")
                End Try
            Else

            End If
        End If
    End Sub

    Private Sub txtFlatRate_TextChanged_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtFlatRate.TextChanged
        Call estrev()
    End Sub

    Private Sub CopyToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CopyToolStripMenuItem.Click
        Clipboard.SetDataObject(rtbDisclosure.SelectedText)
    End Sub

    Private Sub PasteToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PasteToolStripMenuItem.Click
        ' Declares an IDataObject to hold the data returned from the clipboard.
        ' Retrieves the data from the clipboard.
        Dim iData As IDataObject = Clipboard.GetDataObject()

        ' Determines whether the data is in a format you can use.
        If iData.GetDataPresent(DataFormats.Text) Then
            ' Yes it is, so display it in a text box.
            rtbDisclosure.SelectedText = CType(iData.GetData(DataFormats.Text), String)
        Else
            ' No it is not.
            rtbDisclosure.SelectedText = "Could not retrieve data off the clipboard."
        End If
    End Sub

    Private Sub Button20_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button20.Click
        FontDialog1.Font = rtbNarrative.Font
        FontDialog1.ShowDialog()
        rtbNarrative.Font = FontDialog1.Font
        'MsgBox(rtbNarrative.Font.ToString)
    End Sub

    Private Sub Button21_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button21.Click
        If rtbNarrative.Text.Length > 0 Then
            Dim wordapp As New Word.Application
            wordapp.Visible = False
            Dim doc As Word.Document = wordapp.Documents.Add
            Dim range As Word.Range
            range = doc.Range
            range.Text = rtbNarrative.Text
            doc.Activate()
            doc.CheckSpelling()
            Dim chars() As Char = {CType(vbCr, Char), CType(vbLf, Char)}
            rtbNarrative.Text = doc.Range().Text.Trim(chars)
            doc.Close(SaveChanges:=False)
            wordapp.Quit()
            MsgBox("Spell Check Finished.", MsgBoxStyle.Information, "Done")
        End If
    End Sub

    
    Private Sub MenuStrip1_ItemClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.ToolStripItemClickedEventArgs)

    End Sub

    Private Sub CopyToolStripMenuItem1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Clipboard.SetDataObject(rtbDisclosure.SelectedText)
    End Sub

    Private Sub PasteToolStripMenuItem1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        ' Declares an IDataObject to hold the data returned from the clipboard.
        ' Retrieves the data from the clipboard.
        Dim iData As IDataObject = Clipboard.GetDataObject()

        ' Determines whether the data is in a format you can use.
        If iData.GetDataPresent(DataFormats.Text) Then
            ' Yes it is, so display it in a text box.
            rtbDisclosure.SelectedText = CType(iData.GetData(DataFormats.Text), String)
        Else
            ' No it is not.
            rtbDisclosure.SelectedText = "Could not retrieve data off the clipboard."
        End If
    End Sub

    Private Sub CopyToolStripMenuItem1_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CopyToolStripMenuItem1.Click
        Clipboard.SetDataObject(rtbNarrative.SelectedText)
    End Sub

    Private Sub PasteToolStripMenuItem1_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PasteToolStripMenuItem1.Click
        ' Declares an IDataObject to hold the data returned from the clipboard.
        ' Retrieves the data from the clipboard.
        Dim iData As IDataObject = Clipboard.GetDataObject()

        ' Determines whether the data is in a format you can use.
        If iData.GetDataPresent(DataFormats.Text) Then
            ' Yes it is, so display it in a text box.
            rtbNarrative.SelectedText = CType(iData.GetData(DataFormats.Text), String)
        Else
            ' No it is not.
            rtbNarrative.SelectedText = "Could not retrieve data off the clipboard."
        End If
    End Sub

    Private Sub TextBox2_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtPBPerson.TextChanged

    End Sub

    Private Sub cboEmployee_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cboEmployee.SelectedIndexChanged
        'txtPBPerson.Text = cboEmployee.SelectedText.ToString
    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        If txtID.Text = "NEW" Then
            'do nothing
        Else
            Dim Mycn As OleDb.OleDbConnection
            Dim Command As OleDb.OleDbCommand
            Dim SQLstr As String
            Try
                Mycn = New OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & My.Settings.roledbloc & "';Jet OLEDB:Database Password='" & My.Settings.dbpass & "';")
                Dim ds1 As New DataSet
                Dim eds1 As New DataGridView
                Dim dv1 As New DataView
                Mycn.Open()
                SQLstr = "Update rfp_ReportQueue SET Finished = True WHERE Finished = false"
                Command = New OleDb.OleDbCommand(SQLstr, Mycn)
                Command.ExecuteNonQuery()
                Mycn.Close()

                Mycn.Open()
                SQLstr = "INSERT INTO rfp_ReportQueue (RFPID, RequestBy, RequestDate)" & _
                " Values(" & txtID.Text & "," & My.Settings.userid & ",#" & Format(Now()) & "#)"
                Command = New OleDb.OleDbCommand(SQLstr, Mycn)
                Command.ExecuteNonQuery()
                Mycn.Close()

                Dim child As New RFP_ReportViewer
                child.MdiParent = Home
                child.Show()

            Catch ex As Exception
                MsgBox(ex.Message, MsgBoxStyle.Critical, "ERROR")
            End Try
        End If
    End Sub

    Private Sub Button22_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button22.Click
        Call CreateWordDocTemplate()
    End Sub

    Public Sub CreateWordDocTemplate()
        Dim oWord As Word.Application
        Dim oDoc As Word.Document

        oWord = CreateObject("Word.Application")
        oWord.Visible = True
        oDoc = oWord.Documents.Add("\\monumentco1\data\toolboxfiles\RFPReport.dotx")

        oDoc.Bookmarks.Item("RFPName").Range.Text = txtRptName.Text
        oDoc.Bookmarks.Item("RFPDate").Range.Text = Format(Now, "MMMM dd, yyyy")
        oDoc.Bookmarks.Item("RFPSubName").Range.Text = txtCompanyName.Text
        If ckbPriContact.Checked Then
            oDoc.Bookmarks.Item("RFPSubContact").Range.Text = txtPrimaryContact.Text
        Else
            oDoc.Bookmarks.Item("RFPSubContact").Range.Paragraphs(1).Range.Delete()
        End If
        oDoc.Bookmarks.Item("RFPSubAddress").Range.Text = txtAddress.Text
        oDoc.Bookmarks.Item("RFPSubCity").Range.Text = txtCity.Text
        oDoc.Bookmarks.Item("RFPSubState").Range.Text = txtState.Text
        oDoc.Bookmarks.Item("RFPSubZip").Range.Text = txtZip.Text
        oDoc.Bookmarks.Item("RFPSubPhone").Range.Text = txtPhone.Text
        'oDoc.Bookmarks.Item("test").Range.Text = txtAddress.Text
        oDoc.Bookmarks.Item("SBFirm").Range.Text = txtPBFirm.Text
        oDoc.Bookmarks.Item("SBContact").Range.Text = txtPBPerson.Text
        oDoc.Bookmarks.Item("SBAddress").Range.Text = txtPBAddress.Text
        oDoc.Bookmarks.Item("SBCity").Range.Text = txtPBCity.Text
        oDoc.Bookmarks.Item("SBState").Range.Text = txtPBState.Text
        oDoc.Bookmarks.Item("SBZip").Range.Text = txtPBZip.Text
        oDoc.Bookmarks.Item("SBPhone").Range.Text = txtPBPhone.Text

        If ckbCoverLetter.Checked Then
            oDoc.Bookmarks.Item("P1Header").Range.Text = "INTRODUCTION LETTER"
            oDoc.Bookmarks.Item("Letter").Range.Text = rtbNarrative.Text
        Else

        End If
    End Sub

    Public Sub CreateWordDocReport()
        Dim oWord As Word.Application
        Dim oDoc As Word.Document
        Dim oTable As Word.Table
        Dim oPara1 As Word.Paragraph, oPara2 As Word.Paragraph
        Dim oPara3 As Word.Paragraph, oPara4 As Word.Paragraph
        Dim oRng As Word.Range
        Dim oShape As Word.InlineShape
        Dim oChart As Object
        Dim Pos As Double

        'Start Word and open the document template.
        oWord = CreateObject("Word.Application")
        oWord.Visible = True
        oDoc = oWord.Documents.Add

        'Insert a paragraph at the beginning of the document.
        oPara1 = oDoc.Content.Paragraphs.Add
        oPara1.SpaceBefore = 156
        oPara1.Range.Text = "Investment Proposal in Response to RFP:"
        oPara1.Range.Font.Bold = True
        oPara1.Range.Font.Name = "Calibri"
        oPara1.Range.Font.Size = 22
        oPara1.Range.Font.Color = Color.FromArgb(0, 95, 95, 95).ToArgb
        oPara1.Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter
        oPara1.Format.SpaceAfter = 0    '24 pt spacing after paragraph.
        oPara1.Range.InsertParagraphAfter()

        'Insert a paragraph at the end of the document.
        '** \endofdoc is a predefined bookmark.
        oPara2 = oDoc.Content.Paragraphs.Add(oDoc.Bookmarks.Item("\endofdoc").Range)
        oPara2.Range.Text = txtRptName.Text
        oPara2.Format.SpaceAfter = 6
        oDoc.Shapes.AddLine(0, 400, 610, 400)
        oPara2.Range.InsertParagraphAfter()

        'Insert another paragraph.
        oPara3 = oDoc.Content.Paragraphs.Add(oDoc.Bookmarks.Item("\endofdoc").Range)
        oPara3.Range.Text = "This is a sentence of normal text. Now here is a table:"
        oPara3.Range.Font.Bold = False
        oPara3.Format.SpaceAfter = 24
        oPara3.Range.InsertParagraphAfter()

        'Insert a 3 x 5 table, fill it with data, and make the first row
        'bold and italic.
        Dim r As Integer, c As Integer
        oTable = oDoc.Tables.Add(oDoc.Bookmarks.Item("\endofdoc").Range, 3, 5)
        oTable.Range.ParagraphFormat.SpaceAfter = 6
        For r = 1 To 3
            For c = 1 To 5
                oTable.Cell(r, c).Range.Text = "r" & r & "c" & c
            Next
        Next
        oTable.Rows.Item(1).Range.Font.Bold = True
        oTable.Rows.Item(1).Range.Font.Italic = True

        'Add some text after the table.
        'oTable.Range.InsertParagraphAfter()
        oPara4 = oDoc.Content.Paragraphs.Add(oDoc.Bookmarks.Item("\endofdoc").Range)
        oPara4.Range.InsertParagraphBefore()
        oPara4.Range.Text = "And here's another table:"
        oPara4.Format.SpaceAfter = 24
        oPara4.Range.InsertParagraphAfter()

        'Insert a 5 x 2 table, fill it with data, and change the column widths.
        oTable = oDoc.Tables.Add(oDoc.Bookmarks.Item("\endofdoc").Range, 5, 2)
        oTable.Range.ParagraphFormat.SpaceAfter = 6
        For r = 1 To 5
            For c = 1 To 2
                oTable.Cell(r, c).Range.Text = "r" & r & "c" & c
            Next
        Next
        oTable.Columns.Item(1).Width = oWord.InchesToPoints(2)   'Change width of columns 1 & 2
        oTable.Columns.Item(2).Width = oWord.InchesToPoints(3)

        'Keep inserting text. When you get to 7 inches from top of the
        'document, insert a hard page break.
        Pos = oWord.InchesToPoints(7)
        oDoc.Bookmarks.Item("\endofdoc").Range.InsertParagraphAfter()
        Do
            oRng = oDoc.Bookmarks.Item("\endofdoc").Range
            oRng.ParagraphFormat.SpaceAfter = 6
            oRng.InsertAfter("A line of text")
            oRng.InsertParagraphAfter()
        Loop While Pos >= oRng.Information(Word.WdInformation.wdVerticalPositionRelativeToPage)
        oRng.Collapse(Word.WdCollapseDirection.wdCollapseEnd)
        oRng.InsertBreak(Word.WdBreakType.wdPageBreak)
        oRng.Collapse(Word.WdCollapseDirection.wdCollapseEnd)
        oRng.InsertAfter("We're now on page 2. Here's my chart:")
        oRng.InsertParagraphAfter()

        'Insert a chart and change the chart.
        oShape = oDoc.Bookmarks.Item("\endofdoc").Range.InlineShapes.AddOLEObject( _
            ClassType:="MSGraph.Chart.8", FileName _
            :="", LinkToFile:=False, DisplayAsIcon:=False)
        oChart = oShape.OLEFormat.Object
        oChart.charttype = 4 'xlLine = 4
        oChart.Application.Update()
        oChart.Application.Quit()
        'If desired, you can proceed from here using the Microsoft Graph 
        'Object model on the oChart object to make additional changes to the
        'chart.
        oShape.Width = oWord.InchesToPoints(6.25)
        oShape.Height = oWord.InchesToPoints(3.57)

        'Add text after the chart.
        oRng = oDoc.Bookmarks.Item("\endofdoc").Range
        oRng.InsertParagraphAfter()
        oRng.InsertAfter("THE END.")

        'All done. Close this form.
    End Sub
End Class