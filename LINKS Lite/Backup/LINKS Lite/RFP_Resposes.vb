Public Class RFP_Resposes
    Dim thPeople As System.Threading.Thread
    Dim thProduct As System.Threading.Thread
    Dim thComposite As System.Threading.Thread
    Dim thLevel As System.Threading.Thread
    Dim thStage As System.Threading.Thread
    Dim thInsertRec As System.Threading.Thread
    Dim thUpdateRec As System.Threading.Thread

    Private Sub cbPeople_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cbPeople.CheckedChanged
        If cbPeople.Checked Then
            Label5.Visible = True
            Control.CheckForIllegalCrossThreadCalls = False
            thPeople = New System.Threading.Thread(AddressOf LoadPeople)
            thPeople.Start()
            cboPeople.Visible = True
            cboPeople.Enabled = False
        Else
            Label5.Visible = False
            cboPeople.Visible = False
        End If
    End Sub

    Public Sub LoadPeople()
        Try

            Dim conn As New OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & My.Settings.dbloc & "';Jet OLEDB:Database Password='" & My.Settings.dbpass & "';")

            Dim strSQL As String = "SELECT * FROM rfp_Contacts"
            Dim da As New OleDb.OleDbDataAdapter(strSQL, conn)
            Dim ds As New DataSet
            da.Fill(ds, "Users")

            With cboPeople
                .DataSource = ds.Tables("Users")
                .DisplayMember = "ContactName"
                .ValueMember = "ContactID"
                .SelectedIndex = 0
            End With

            Label5.Visible = False
            cboPeople.Enabled = True

            If txtID.Text = "NEW" Then

            Else
                cboPeople.SelectedValue = txtContactID.Text
            End If

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "ERROR")
        End Try
    End Sub

    Public Sub LoadComposites()
        Try

            Dim conn As New OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & My.Settings.dbloc & "';Jet OLEDB:Database Password='" & My.Settings.dbpass & "';")

            Dim strSQL As String = "SELECT * FROM rfp_Composites"
            Dim da As New OleDb.OleDbDataAdapter(strSQL, conn)
            Dim ds As New DataSet
            da.Fill(ds, "Users")

            With cboComposite
                .DataSource = ds.Tables("Users")
                .DisplayMember = "PortfolioCompositeCode"
                .ValueMember = "PortfolioCompositeID"
                .SelectedIndex = 0
            End With

            Label6.Visible = False
            cboComposite.Enabled = True

            If txtID.Text = "NEW" Then

            Else
                Dim compid As Integer
                compid = txtCompositeID.Text

                cboComposite.SelectedValue = compid
            End If

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "ERROR")
        End Try
    End Sub

    Public Sub LoadProduct()
        Try

            Dim conn As New OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & My.Settings.dbloc & "';Jet OLEDB:Database Password='" & My.Settings.dbpass & "';")

            Dim strSQL As String = "SELECT * FROM AAM_DisciplineQuery"
            Dim da As New OleDb.OleDbDataAdapter(strSQL, conn)
            Dim ds As New DataSet
            da.Fill(ds, "Users")

            With cboProduct
                .DataSource = ds.Tables("Users")
                .DisplayMember = "Discipline6"
                .ValueMember = "Discipline6"
                .SelectedIndex = 0
            End With

            Label4.Visible = False
            cboProduct.Enabled = True

            If txtID.Text = "NEW" Then
            Else
                cboProduct.SelectedValue = txtDiscipline.Text
            End If

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "ERROR")
        End Try
    End Sub

    Public Sub LoadLevel()
        Try

            Dim conn As New OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & My.Settings.dbloc & "';Jet OLEDB:Database Password='" & My.Settings.dbpass & "';")

            Dim strSQL As String = "SELECT ID, LevelName FROM rfp_Levels WHERE IsActive = -1"
            Dim da As New OleDb.OleDbDataAdapter(strSQL, conn)
            Dim ds As New DataSet
            da.Fill(ds, "Users")

            With cboLevel
                .DataSource = ds.Tables("Users")
                .DisplayMember = "LevelName"
                .ValueMember = "ID"
                .SelectedIndex = 0
            End With

            lblLevel.Visible = False
            cboLevel.Enabled = True

            If txtID.Text = "NEW" Then
            Else
                cboLevel.SelectedValue = txtLevelID.Text
            End If

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "ERROR")
        End Try
    End Sub

    Public Sub LoadStage()
        Try

            Dim conn As New OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & My.Settings.dbloc & "';Jet OLEDB:Database Password='" & My.Settings.dbpass & "';")

            Dim strSQL As String = "SELECT ID, StageName FROM rfp_QuestionStage WHERE IsActive = -1"
            Dim da As New OleDb.OleDbDataAdapter(strSQL, conn)
            Dim ds As New DataSet
            da.Fill(ds, "Users")

            With cboStage
                .DataSource = ds.Tables("Users")
                .DisplayMember = "StageName"
                .ValueMember = "ID"
                .SelectedIndex = 0
            End With

            Label9.Visible = False
            cboStage.Enabled = True

            If txtID.Text = "NEW" Then
            Else
                cboStage.SelectedValue = txtStageID.Text
            End If

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "ERROR")
        End Try
    End Sub

    Private Sub cbProduct_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cbProduct.CheckedChanged
        If cbProduct.Checked Then
            Label4.Visible = True
            Control.CheckForIllegalCrossThreadCalls = False
            thProduct = New System.Threading.Thread(AddressOf LoadProduct)
            thProduct.Start()
            cboProduct.Visible = True
            cboProduct.Enabled = False
        Else
            Label4.Visible = False
            cboProduct.Visible = False
        End If
    End Sub

    Private Sub cbComposite_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cbComposite.CheckedChanged
        If cbComposite.Checked Then
            Label6.Visible = True
            Control.CheckForIllegalCrossThreadCalls = False
            thComposite = New System.Threading.Thread(AddressOf LoadComposites)
            thComposite.Start()
            cboComposite.Visible = True
            cboComposite.Enabled = False
        Else
            Label6.Visible = False
            cboComposite.Visible = False
        End If
    End Sub

    Private Sub cbStale_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cbStale.CheckedChanged
        If cbStale.Checked Then
            lblStale.Visible = True
            DTPStale.Visible = True
        Else
            lblStale.Visible = False
            DTPStale.Visible = False
        End If
    End Sub

    Private Sub cmdSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdSave.Click
        '****BEGIN DATA VALIDATION****
        If lblLevel.Visible = True Or Label4.Visible = True Or Label5.Visible = True Or Label6.Visible = True Or Label9.Visible = True Then
            MsgBox("You must wait for the data to finish loading before you can save.", MsgBoxStyle.Information, "Data Loading")
            GoTo line1
        Else

        End If

        If IsDBNull(rtbQuestion) Or rtbQuestion.Text = "" Then
            rtbQuestion.BackColor = Color.Red
            rtbQuestion.ForeColor = Color.White
            GoTo line1
        Else
            rtbQuestion.BackColor = Color.White
            rtbQuestion.ForeColor = Color.Black
        End If

        If IsDBNull(rtbAnswer) Or rtbAnswer.Text = "" Then
            rtbAnswer.BackColor = Color.Red
            rtbAnswer.ForeColor = Color.White
            GoTo line1
        Else
            rtbAnswer.BackColor = Color.White
            rtbAnswer.ForeColor = Color.Black
        End If

        If IsDBNull(txtRptID) Or txtRptID.Text = "" Then
            txtRptID.Text = 0
        Else

        End If

        If txtID.Text = "NEW" Then
            Control.CheckForIllegalCrossThreadCalls = False
            thInsertRec = New System.Threading.Thread(AddressOf InsertRec)
            thInsertRec.Start()
        Else
            Call Tracker()
            Control.CheckForIllegalCrossThreadCalls = False
            thUpdateRec = New System.Threading.Thread(AddressOf UpdateRec)
            thUpdateRec.Start()
        End If

line1:

    End Sub

    Public Sub Tracker()
        Dim Mycn As OleDb.OleDbConnection
        Dim Command As OleDb.OleDbCommand
        Dim SQLstr As String
        Try
            Mycn = New OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & My.Settings.dbloc & "';Jet OLEDB:Database Password='" & My.Settings.dbpass & "';")

            Dim question1 As String
            question1 = rtbQuestion.Text
            question1 = Replace(question1, "'", "''")

            Dim question2 As String
            question2 = RichTextBox1.Text
            question2 = Replace(question2, "'", "''")

            Dim answer1 As String
            answer1 = rtbAnswer.Text
            answer1 = Replace(answer1, "'", "''")

            Dim answer2 As String
            answer2 = RichTextBox2.Text
            answer2 = Replace(answer2, "'", "''")

            'Mycn.Open()
            If question1 = question2 Then
                'nothing changed
            Else
                Mycn.Open()
                SQLstr = "INSERT INTO rfp_ResponseTracking([QuestionID], [Field], [OldValue], [NewValue], [ChangedBy], [DateStamp])" & _
                "VALUES(" & txtID.Text & ",'Question','" & question2 & "','" & question1 & "','" & Environ("USERNAME") & "',#" & Format(Now()) & "#)"

                Command = New OleDb.OleDbCommand(SQLstr, Mycn)
                Command.ExecuteNonQuery()

                Mycn.Close()
            End If

            If answer1 = answer2 Then
                'nothing changed
            Else
                Mycn.Open()
                SQLstr = "INSERT INTO rfp_ResponseTracking([QuestionID], [Field], [OldValue], [NewValue], [ChangedBy], [DateStamp])" & _
                "VALUES(" & txtID.Text & ",'Answer','" & answer2 & "','" & answer1 & "','" & Environ("USERNAME") & "',#" & Format(Now()) & "#)"

                Command = New OleDb.OleDbCommand(SQLstr, Mycn)
                Command.ExecuteNonQuery()

                Mycn.Close()
            End If

            If cboLevel.SelectedValue = txtLevelID.Text Then
                'nothing changed
            Else
                Mycn.Open()
                SQLstr = "INSERT INTO rfp_ResponseTracking([QuestionID], [Field], [OldValue], [NewValue], [ChangedBy], [DateStamp])" & _
                "VALUES(" & txtID.Text & ",'Level','" & txtLevelID.Text & "','" & cboLevel.SelectedValue & "','" & Environ("USERNAME") & "',#" & Format(Now()) & "#)"

                Command = New OleDb.OleDbCommand(SQLstr, Mycn)
                Command.ExecuteNonQuery()

                Mycn.Close()
            End If

            If cboPeople.SelectedValue = txtContactID.Text Then
                'nothing changed
            Else
                Mycn.Open()
                SQLstr = "INSERT INTO rfp_ResponseTracking([QuestionID], [Field], [OldValue], [NewValue], [ChangedBy], [DateStamp])" & _
                "VALUES(" & txtID.Text & ",'Contact ID','" & txtContactID.Text & "','" & cboPeople.SelectedValue & "','" & Environ("USERNAME") & "',#" & Format(Now()) & "#)"

                Command = New OleDb.OleDbCommand(SQLstr, Mycn)
                Command.ExecuteNonQuery()

                Mycn.Close()
            End If

            If cboProduct.SelectedValue = txtDiscipline.Text Then
                'nothing changed
            Else
                Mycn.Open()
                SQLstr = "INSERT INTO rfp_ResponseTracking([QuestionID], [Field], [OldValue], [NewValue], [ChangedBy], [DateStamp])" & _
                "VALUES(" & txtID.Text & ",'Product','" & txtDiscipline.Text & "','" & cboProduct.SelectedValue & "','" & Environ("USERNAME") & "',#" & Format(Now()) & "#)"

                Command = New OleDb.OleDbCommand(SQLstr, Mycn)
                Command.ExecuteNonQuery()

                Mycn.Close()
            End If

            If cboComposite.SelectedValue = txtCompositeID.Text Then
                'nothing changed
            Else
                Mycn.Open()
                SQLstr = "INSERT INTO rfp_ResponseTracking([QuestionID], [Field], [OldValue], [NewValue], [ChangedBy], [DateStamp])" & _
                "VALUES(" & txtID.Text & ",'Composite ID','" & txtCompositeID.Text & "','" & cboComposite.SelectedValue & "','" & Environ("USERNAME") & "',#" & Format(Now()) & "#)"

                Command = New OleDb.OleDbCommand(SQLstr, Mycn)
                Command.ExecuteNonQuery()

                Mycn.Close()
            End If

            If DTPStale.Text = txtStaleDate.Text Then
                'nothing changed
            Else
                Mycn.Open()
                SQLstr = "INSERT INTO rfp_ResponseTracking([QuestionID], [Field], [OldValue], [NewValue], [ChangedBy], [DateStamp])" & _
                "VALUES(" & txtID.Text & ",'Stale Date','" & txtStaleDate.Text & "','" & DTPStale.Text & "','" & Environ("USERNAME") & "',#" & Format(Now()) & "#)"

                Command = New OleDb.OleDbCommand(SQLstr, Mycn)
                Command.ExecuteNonQuery()

                Mycn.Close()
            End If

            If cbComposite.Checked = ckbComposite.CheckState Then
                'nothing changed
            Else
                Mycn.Open()
                SQLstr = "INSERT INTO rfp_ResponseTracking([QuestionID], [Field], [OldValue], [NewValue], [ChangedBy], [DateStamp])" & _
                "VALUES(" & txtID.Text & ",'Related to Composite','" & ckbComposite.CheckState & "','" & cbComposite.CheckState & "','" & Environ("USERNAME") & "',#" & Format(Now()) & "#)"

                Command = New OleDb.OleDbCommand(SQLstr, Mycn)
                Command.ExecuteNonQuery()

                Mycn.Close()
            End If

            If cbPeople.Checked = ckbPeople.CheckState Then
                'nothing changed
            Else
                Mycn.Open()
                SQLstr = "INSERT INTO rfp_ResponseTracking([QuestionID], [Field], [OldValue], [NewValue], [ChangedBy], [DateStamp])" & _
                "VALUES(" & txtID.Text & ",'Related to People','" & ckbPeople.CheckState & "','" & cbPeople.CheckState & "','" & Environ("USERNAME") & "',#" & Format(Now()) & "#)"

                Command = New OleDb.OleDbCommand(SQLstr, Mycn)
                Command.ExecuteNonQuery()

                Mycn.Close()
            End If

            If cbProduct.Checked = ckbProduct.CheckState Then
                'nothing changed
            Else
                Mycn.Open()
                SQLstr = "INSERT INTO rfp_ResponseTracking([QuestionID], [Field], [OldValue], [NewValue], [ChangedBy], [DateStamp])" & _
                "VALUES(" & txtID.Text & ",'Related to Product','" & ckbProduct.CheckState & "','" & cbProduct.CheckState & "','" & Environ("USERNAME") & "',#" & Format(Now()) & "#)"

                Command = New OleDb.OleDbCommand(SQLstr, Mycn)
                Command.ExecuteNonQuery()

                Mycn.Close()
            End If

            If cbStale.Checked = ckbDate1.CheckState Then
                'nothing changed
            Else
                Mycn.Open()
                SQLstr = "INSERT INTO rfp_ResponseTracking([QuestionID], [Field], [OldValue], [NewValue], [ChangedBy], [DateStamp])" & _
                "VALUES(" & txtID.Text & ",'Date Sensitive','" & ckbDate1.CheckState & "','" & cbStale.CheckState & "','" & Environ("USERNAME") & "',#" & Format(Now()) & "#)"

                Command = New OleDb.OleDbCommand(SQLstr, Mycn)
                Command.ExecuteNonQuery()

                Mycn.Close()
            End If

            If cboStage.SelectedValue = txtStageID.Text Then
                'nothing changed
            Else
                Mycn.Open()
                SQLstr = "INSERT INTO rfp_ResponseTracking([QuestionID], [Field], [OldValue], [NewValue], [ChangedBy], [DateStamp])" & _
                "VALUES(" & txtID.Text & ",'Question Stage','" & txtStageID.Text & "','" & cboStage.SelectedValue & "','" & Environ("USERNAME") & "',#" & Format(Now()) & "#)"

                Command = New OleDb.OleDbCommand(SQLstr, Mycn)
                Command.ExecuteNonQuery()

                Mycn.Close()
            End If

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "ERROR")

        End Try

    End Sub

    Public Sub UpdateRec()
        Dim Mycn As OleDb.OleDbConnection
        Dim Command As OleDb.OleDbCommand
        Dim SQLstr As String

        'Try
        Mycn = New OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & My.Settings.dbloc & "';Jet OLEDB:Database Password='" & My.Settings.dbpass & "';")

        Mycn.Open()

        Dim ppl As Integer
        ppl = cboPeople.SelectedValue

        Dim pro As String
        pro = cboProduct.SelectedValue


        Dim comp As Integer
        comp = cboComposite.SelectedValue


        Dim dte As Date
        If IsDBNull(DTPStale) Then
            dte = "12/31/2099"
        Else
            dte = DTPStale.Value
        End If

        Dim question As String
        question = rtbQuestion.Text
        question = Replace(question, "'", "''")

        Dim answer As String
        answer = rtbAnswer.Text
        answer = Replace(answer, "'", "''")

        SQLstr = "Update rfp_ResponseBase Set [LevelID] = " & cboLevel.SelectedValue & ", [Question] = '" & question & "', [Answer] = '" & answer & "', [LastUpdated] = #" & Format(Now()) & "#, [RelatedToPeople] = " & cbPeople.CheckState & ", [ContactID] = " & ppl & ", [RelatedToProduct] =" & cbProduct.CheckState & ", [Discipline] = '" & pro & "', [Composite] = " & cbComposite.CheckState & ", [CompositeID] = " & comp & ", [DateSensitive] = " & cbStale.CheckState & ", [StaleDate] = #" & dte & "#, [ReportQuestionID] = " & txtRptID.Text & ", [Stage] = " & cboStage.SelectedValue & "" & _
        " WHERE ID = " & txtID.Text

        Command = New OleDb.OleDbCommand(SQLstr, Mycn)
        Command.ExecuteNonQuery()

        Mycn.Close()

        Me.Close()

        'Catch ex As Exception
        'MsgBox(ex.Message, MsgBoxStyle.Critical, "ERROR")

        'End Try
    End Sub

    Public Sub InsertRec()
        Dim Mycn As OleDb.OleDbConnection
        Dim Command As OleDb.OleDbCommand
        Dim SQLstr As String

        Try
            Mycn = New OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & My.Settings.dbloc & "';Jet OLEDB:Database Password='" & My.Settings.dbpass & "';")

            Mycn.Open()

            Dim ppl As Integer
            ppl = cboPeople.SelectedValue

            Dim pro As String
            pro = cboProduct.SelectedValue


            Dim comp As Integer
            comp = cboComposite.SelectedValue


            Dim dte As Date
            If IsDBNull(DTPStale) Then
                dte = "12/31/2099"
            Else
                dte = DTPStale.Value
            End If

            Dim question As String
            question = rtbQuestion.Text
            question = Replace(question, "'", "''")

            Dim answer As String
            answer = rtbAnswer.Text
            answer = Replace(answer, "'", "''")

            SQLstr = "INSERT INTO rfp_ResponseBase([LevelID], [Question], [Answer], [DateEntered], [LastUpdated], [RelatedToPeople], [ContactID], [RelatedToProduct], [Discipline], [Composite], [CompositeID], [DateSensitive], [StaleDate], [IsActive], [ReportQuestionID], [Stage])" & _
            "VALUES(" & cboLevel.SelectedValue & ",'" & question & "','" & answer & "',#" & Format(Now()) & "#,#" & Format(Now()) & "#," & cbPeople.CheckState & "," & ppl & "," & cbProduct.CheckState & ",'" & pro & "'," & cbComposite.CheckState & "," & comp & "," & cbStale.CheckState & ",#" & dte & "#, -1, " & txtRptID.Text & "," & cboStage.SelectedValue & ")"

            Command = New OleDb.OleDbCommand(SQLstr, Mycn)
            Command.ExecuteNonQuery()

            Mycn.Close()

            Me.Close()

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "ERROR")

        End Try
    End Sub

    Private Sub RFP_Resposes_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        'Permission Check
        If Permissions.RFPQuestionView.Checked Then

        Else
            MsgBox("You do not have permission to perform this function.", MsgBoxStyle.Exclamation, "Permission Issue")
            Me.Close()
        End If

        If Permissions.RFPQuestionDelete.Checked Then
            cmdCancel.Enabled = True
        Else
            cmdCancel.Enabled = False
        End If

        If Permissions.RFPQuestionEdit.Checked Then
            cmdSave.Enabled = True
        Else
            cmdSave.Enabled = False
        End If

        lblLevel.Visible = True
        Label9.Visible = True
        Control.CheckForIllegalCrossThreadCalls = False
        thLevel = New System.Threading.Thread(AddressOf LoadLevel)
        thLevel.Start()

        thStage = New System.Threading.Thread(AddressOf LoadStage)
        thStage.Start()
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        Dim child As New rfp_ChangeLog
        child.MdiParent = Home
        child.Show()
        child.txtID.Text = txtID.Text
        Try

            Dim conn As New OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & My.Settings.dbloc & "';Jet OLEDB:Database Password='" & My.Settings.dbpass & "';")

            Dim strSQL As String = "SELECT ID, Field, OldValue As [Old Value], NewValue As [New Value], ChangedBy As [User], DateStamp As [Date Stamp] FROM rfp_ResponseTracking WHERE QuestionID = " & txtID.Text
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
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Me.Close()
    End Sub

    Private Sub cmdCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdCancel.Click
        If lblLevel.Visible = True Or Label4.Visible = True Or Label5.Visible = True Or Label6.Visible = True Then
            MsgBox("You must wait for the data to finish loading before you can save.", MsgBoxStyle.Information, "Data Loading")
            GoTo line1
        Else

        End If

        If cmdCancel.Text = "Delete" Then
            Dim ir As MsgBoxResult
            ir = MsgBox("Are you sure you want to Delete this question?", MsgBoxStyle.Critical + MsgBoxStyle.YesNo, "Confirm")

            If ir = MsgBoxResult.Yes Then
                Dim Mycn As OleDb.OleDbConnection
                Dim Command As OleDb.OleDbCommand
                Dim SQLstr As String
                Try
                    Mycn = New OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & My.Settings.dbloc & "';Jet OLEDB:Database Password='" & My.Settings.dbpass & "';")


                    Mycn.Open()
                    SQLstr = "INSERT INTO rfp_ResponseTracking([QuestionID], [Field], [OldValue], [NewValue], [ChangedBy], [DateStamp])" & _
                    "VALUES(" & txtID.Text & ",'DELETED','FALSE','TRUE','" & Environ("USERNAME") & "',#" & Format(Now()) & "#)"

                    Command = New OleDb.OleDbCommand(SQLstr, Mycn)
                    Command.ExecuteNonQuery()

                    Mycn.Close()

                    Mycn.Open()
                    SQLstr = "Update rfp_ResponseBase SET [IsActive] = Null" & _
                    " WHERE ID = " & txtID.Text

                    Command = New OleDb.OleDbCommand(SQLstr, Mycn)
                    Command.ExecuteNonQuery()

                    Mycn.Close()

                    cmdCancel.Text = "Un-Delete"
                    lblDeleted.Visible = True
                    ckbActive.Checked = False
                    cmdSave.Enabled = False

                Catch ex As Exception
                    MsgBox(ex.Message, MsgBoxStyle.Critical, "ERROR")

                End Try
            Else

            End If
        Else
            Dim Mycn As OleDb.OleDbConnection
            Dim Command As OleDb.OleDbCommand
            Dim SQLstr As String
            Try
                Mycn = New OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & My.Settings.dbloc & "';Jet OLEDB:Database Password='" & My.Settings.dbpass & "';")


                Mycn.Open()
                SQLstr = "INSERT INTO rfp_ResponseTracking([QuestionID], [Field], [OldValue], [NewValue], [ChangedBy], [DateStamp])" & _
                "VALUES(" & txtID.Text & ",'DELETED','TRUE','FALSE','" & Environ("USERNAME") & "',#" & Format(Now()) & "#)"

                Command = New OleDb.OleDbCommand(SQLstr, Mycn)
                Command.ExecuteNonQuery()

                Mycn.Close()

                Mycn.Open()
                SQLstr = "Update rfp_ResponseBase SET [IsActive] = -1" & _
                " WHERE ID = " & txtID.Text

                Command = New OleDb.OleDbCommand(SQLstr, Mycn)
                Command.ExecuteNonQuery()

                Mycn.Close()

                cmdCancel.Text = "Delete"
                lblDeleted.Visible = False
                ckbActive.Checked = True
                cmdSave.Enabled = True

            Catch ex As Exception
                MsgBox(ex.Message, MsgBoxStyle.Critical, "ERROR")

            End Try
        End If

line1:

    End Sub

    Private Sub CopyToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CopyToolStripMenuItem.Click
        If IsDBNull(rtbQuestion.SelectedText) Then

        Else
            Clipboard.SetText(rtbQuestion.SelectedText & vbNewLine & vbNewLine & My.Application.Info.Copyright)
        End If
    End Sub

    Private Sub PasteToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PasteToolStripMenuItem.Click
        If IsDBNull(Clipboard.GetText) Then

        Else
            rtbQuestion.Text = rtbQuestion.Text & Clipboard.GetText
        End If
    End Sub

    Private Sub ToolStripMenuItem1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripMenuItem1.Click
        If IsDBNull(rtbAnswer.SelectedText) Then

        Else
            Clipboard.SetText(rtbAnswer.SelectedText & vbNewLine & vbNewLine & My.Application.Info.Copyright)
        End If
    End Sub

    Private Sub ToolStripMenuItem2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripMenuItem2.Click
        If IsDBNull(Clipboard.GetText) Then

        Else
            rtbAnswer.Text = rtbAnswer.Text & Clipboard.GetText
        End If
    End Sub
End Class