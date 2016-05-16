Imports Microsoft.Office.Interop

Public Class rfp_Answers
    Dim thPeople As System.Threading.Thread
    Dim thProduct As System.Threading.Thread
    Dim thComposite As System.Threading.Thread
    Dim thLevel As System.Threading.Thread
    Dim thStage As System.Threading.Thread
    Private Sub Button6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button6.Click
        Call LoadRealQuestions()
    End Sub
    Public Sub LoadQuestions()
        Try

            Dim conn As New OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & My.Settings.dbloc & "';Jet OLEDB:Database Password='" & My.Settings.dbpass & "';")

            Dim strSQL As String = "SELECT rfp_ReportQuestions.ID, rfp_Answers.ID, rfp_ReportQuestions.Question, rfp_Answers.Answer FROM (rfp_ReportQuestions INNER JOIN rfp_QAAssignments ON rfp_ReportQuestions.ID = rfp_QAAssignments.QuestionID) INNER JOIN rfp_Answers ON rfp_QAAssignments.AnswerID = rfp_Answers.ID WHERE rfp_QAAssignments.IsActive = -1 AND rfp_Answers.IsActive = -1 AND rfp_QAAssignments.AnswerID Not In (Select AnswerID FROM rfp_QAAssignments WHERE QuestionID = " & txtID.Text & " AND rfp_QAAssignments.IsActive = -1) AND [Question] LIKE '%" & rtbQuestion.Text & "%'"

            If cbComposite.Checked Then
                strSQL = strSQL & " AND rfp_ReportQuestions.CompositeID = " & cboComposite.SelectedValue
            Else
                'do nothing
            End If

            If cbPeople.Checked Then
                strSQL = strSQL & " AND rfp_ReportQuestions.ContactID = " & cboPeople.SelectedValue
            Else
                'do nothing
            End If

            If cbProduct.Checked Then
                strSQL = strSQL & " AND rfp_ReportQuestions.Discipline = '" & cboProduct.SelectedValue & "'"
            Else
                'do nothing
            End If

            Dim da As New OleDb.OleDbDataAdapter(strSQL, conn)
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

    Public Sub LoadAllQuestions()
        Try

            Dim conn As New OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & My.Settings.dbloc & "';Jet OLEDB:Database Password='" & My.Settings.dbpass & "';")

            Dim strSQL As String = "SELECT rfp_ReportQuestions.ID, rfp_Answers.ID, rfp_ReportQuestions.Question, rfp_Answers.Answer FROM (rfp_ReportQuestions INNER JOIN rfp_QAAssignments ON rfp_ReportQuestions.ID = rfp_QAAssignments.QuestionID) INNER JOIN rfp_Answers ON rfp_QAAssignments.AnswerID = rfp_Answers.ID WHERE rfp_QAAssignments.IsActive = -1 AND rfp_Answers.IsActive = -1"

            Dim da As New OleDb.OleDbDataAdapter(strSQL, conn)
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

    Private Sub rfp_Answers_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
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

        
        If txtID.Text = "NEW" Then
            'lblLevel.Visible = True
            'Label9.Visible = True
            Control.CheckForIllegalCrossThreadCalls = False
            thLevel = New System.Threading.Thread(AddressOf LoadLevel)
            thLevel.Start()

            thStage = New System.Threading.Thread(AddressOf LoadStage)
            thStage.Start()
        Else
            'lblLevel.Visible = True
            'Label9.Visible = True
        End If

        'Call LoadAnswersGrid()
        'Call LoadQuestions()
    End Sub

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
            cboPeople.Enabled = True
            With cboPeople
                .DataSource = ds.Tables("Users")
                .DisplayMember = "ContactName"
                .ValueMember = "ContactID"
                'Me.Controls.Add(cboPeople)
                .SelectedIndex = 0
            End With

            Label5.Visible = False
            cboPeople.Enabled = True

            If txtID.Text = "NEW" Then

            Else
                If txtContactID.Text = "" Then
                Else
                    cboPeople.SelectedValue = txtContactID.Text
                End If

            End If

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "ERROR")
        End Try
    End Sub

    Public Sub LoadPeopleAnswer()
        Try

            Dim conn As New OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & My.Settings.dbloc & "';Jet OLEDB:Database Password='" & My.Settings.dbpass & "';")

            Dim strSQL As String = "SELECT * FROM rfp_Contacts"
            Dim da As New OleDb.OleDbDataAdapter(strSQL, conn)
            Dim ds As New DataSet
            da.Fill(ds, "Users")

            With cboPeopleAnswer
                .DataSource = ds.Tables("Users")
                .DisplayMember = "ContactName"
                .ValueMember = "ContactID"
                .SelectedIndex = 0
            End With

            Label12.Visible = False
            cboPeopleAnswer.Enabled = True

            If txtMasterAnswerID.Text = "NEW" Then

            Else
                If txtContactIDAnswer.Text = "" Then
                Else
                    'cboPeopleAnswer.SelectedValue = txtContactIDAnswer.Text
                End If

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
            cboComposite.Enabled = True
            With cboComposite
                .DataSource = ds.Tables("Users")
                .DisplayMember = "PortfolioCompositeCode"
                .ValueMember = "PortfolioCompositeID"
                .SelectedIndex = 0
            End With

            Label6.Visible = False


            If txtID.Text = "NEW" Then

            Else

                If txtCompositeID.Text = "" Then

                Else
                    Dim compid As Integer
                    compid = txtCompositeID.Text

                    cboComposite.SelectedValue = compid
                End If

            End If

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "ERROR")
        End Try
    End Sub

    Public Sub LoadCompositesAnswer()
        Try

            Dim conn As New OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & My.Settings.dbloc & "';Jet OLEDB:Database Password='" & My.Settings.dbpass & "';")

            Dim strSQL As String = "SELECT * FROM rfp_Composites"
            Dim da As New OleDb.OleDbDataAdapter(strSQL, conn)
            Dim ds As New DataSet
            da.Fill(ds, "Users")

            With cboCompositeAnswer
                .DataSource = ds.Tables("Users")
                .DisplayMember = "PortfolioCompositeCode"
                .ValueMember = "PortfolioCompositeID"
                .SelectedIndex = 0
            End With

            Label3.Visible = False
            cboCompositeAnswer.Enabled = True

            If txtMasterAnswerID.Text = "NEW" Then

            Else

                If txtCompositeIDAnswer.Text = "" Then

                Else
                    Dim compid As Integer
                    compid = txtCompositeIDAnswer.Text

                    'cboCompositeAnswer.SelectedValue = compid
                End If

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
            cboProduct.Enabled = True
            With cboProduct
                .DataSource = ds.Tables("Users")
                .DisplayMember = "DisplayValue"
                .ValueMember = "DisplayValue"
                .SelectedIndex = 0
            End With

            Label4.Visible = False


            If txtID.Text = "NEW" Then
            Else
                cboProduct.SelectedValue = txtDiscipline.Text
            End If

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "ERROR")
        End Try
    End Sub

    Public Sub LoadProductAnswer()
        Try

            Dim conn As New OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & My.Settings.dbloc & "';Jet OLEDB:Database Password='" & My.Settings.dbpass & "';")

            Dim strSQL As String = "SELECT * FROM AAM_DisciplineQuery"
            Dim da As New OleDb.OleDbDataAdapter(strSQL, conn)
            Dim ds As New DataSet
            da.Fill(ds, "Users")

            With cboProductAnswer
                .DataSource = ds.Tables("Users")
                .DisplayMember = "DisplayValue"
                .ValueMember = "DisplayValue"
                .SelectedIndex = 0
            End With

            Label11.Visible = False
            cboProductAnswer.Enabled = True

            If txtMasterAnswerID.Text = "NEW" Then
            Else
                'cboProductAnswer.SelectedValue = txtDisciplineAnswer.Text
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
                If txtLevelID.Text = "" Then
                Else
                    cboLevel.SelectedValue = txtLevelID.Text
                End If

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
                If txtStageID.Text = "" Then

                Else
                    cboStage.SelectedValue = txtStageID.Text
                End If

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

    Private Sub Button7_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button7.Click
        Call AddNewButton()
    End Sub

    Public Sub AddNewButton()
        If Button7.Text = "New Answer" Then
            RichTextBox2.Enabled = True
            DataGridView1.Enabled = False
            Button3.Enabled = False
            Button6.Enabled = False
            Button8.Enabled = False
            Button7.Text = "Reset"
            txtMasterAnswerID.Text = "NEW"
            RichTextBox2.Text = ""
            Button4.Enabled = True
            ckbPeopleAnswer.Checked = False
            ckbPeopleAnswer.Enabled = True
            ckbProductAnswer.Checked = False
            ckbProductAnswer.Enabled = True
            ckbCompositeAnswer.Checked = False
            ckbCompositeAnswer.Enabled = True
            cboPeopleAnswer.Visible = False
            cboProductAnswer.Visible = False
            cboCompositeAnswer.Visible = False
            ckbDateSensitiveAnswer.Checked = False
            ckbDateSensitiveAnswer.Enabled = True
            Label13.Visible = False
            dtpStaleAnswer.Visible = False
            'txtAnswerSortOrder.Text = ""
        Else
            Call ResetAnswers()
        End If
    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        If txtMasterAnswerID.Text = "NEW" Then
            MsgBox("Select a question to modify.", MsgBoxStyle.Information, "No Question Loaded.")
        Else
            If Button3.Text = "Modify" Then
                RichTextBox3.Text = RichTextBox2.Text
                RichTextBox2.Enabled = True
                DataGridView1.Enabled = False
                Button7.Enabled = False
                Button6.Enabled = False
                Button8.Enabled = False
                Button4.Enabled = True
                Button3.Text = "Reset"
            Else
                RichTextBox2.Text = RichTextBox3.Text
                Call ResetAnswers()
            End If
        End If
    End Sub

    Private Sub ckbPeopleAnswer_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ckbPeopleAnswer.CheckedChanged
        If ckbPeopleAnswer.Checked Then
            Label12.Visible = True
            Control.CheckForIllegalCrossThreadCalls = False
            thPeople = New System.Threading.Thread(AddressOf LoadPeopleAnswer)
            thPeople.Start()
            cboPeopleAnswer.Visible = True
            cboPeopleAnswer.Enabled = False
        Else
            Label12.Visible = False
            cboPeopleAnswer.Visible = False
        End If
    End Sub

    Private Sub ckbProductAnswer_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ckbProductAnswer.CheckedChanged
        If ckbProductAnswer.Checked Then
            Label11.Visible = True
            Control.CheckForIllegalCrossThreadCalls = False
            thProduct = New System.Threading.Thread(AddressOf LoadProductAnswer)
            thProduct.Start()
            cboProductAnswer.Visible = True
            cboProductAnswer.Enabled = False
        Else
            Label11.Visible = False
            cboProductAnswer.Visible = False
        End If
    End Sub

    Private Sub ckbCompositeAnswer_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ckbCompositeAnswer.CheckedChanged
        If ckbCompositeAnswer.Checked Then
            Label3.Visible = True
            Control.CheckForIllegalCrossThreadCalls = False
            thComposite = New System.Threading.Thread(AddressOf LoadCompositesAnswer)
            thComposite.Start()
            cboCompositeAnswer.Visible = True
            cboCompositeAnswer.Enabled = False
        Else
            Label3.Visible = False
            cboCompositeAnswer.Visible = False
        End If
    End Sub

    Private Sub ckbDateSensitiveAnswer_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ckbDateSensitiveAnswer.CheckedChanged
        If ckbDateSensitiveAnswer.Checked Then
            Label13.Visible = True
            dtpStaleAnswer.Visible = True
        Else
            Label13.Visible = False
            dtpStaleAnswer.Visible = False
        End If
    End Sub

    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        If txtMasterAnswerID.Text = "NEW" Then
            Call SaveNewAnswer()
        Else
            If Button3.Text = "Reset" Then
                Call SaveSlaveAnswer()
            Else
                If Button8.Text = "Reset" Then
                    Call SaveEditAnswer()
                Else
                    Call SaveAnswer()
                End If
            End If
        End If
    End Sub

    Public Sub ResetAnswers()
        RichTextBox2.Text = ""
        RichTextBox2.Enabled = False
        DataGridView1.Enabled = True
        Button7.Enabled = True
        Button6.Enabled = True
        Button8.Enabled = True
        Button3.Enabled = True
        txtMasterAnswerID.Text = "NEW"
        ckbIgnoreSort.Checked = False
        Button4.Enabled = False
        Button3.Text = "Modify"
        Button7.Text = "New Answer"
        Button8.Text = "Edit"
        ckbPeopleAnswer.Checked = False
        ckbPeopleAnswer.Enabled = False
        ckbProductAnswer.Checked = False
        ckbProductAnswer.Enabled = False
        ckbCompositeAnswer.Checked = False
        ckbCompositeAnswer.Enabled = False
        cboPeopleAnswer.Visible = False
        cboProductAnswer.Visible = False
        cboCompositeAnswer.Visible = False
        ckbDateSensitiveAnswer.Checked = False
        ckbDateSensitiveAnswer.Enabled = False
        Label13.Visible = False
        dtpStaleAnswer.Visible = False
    End Sub

    Public Sub SaveNewAnswer()
        Dim Mycn As OleDb.OleDbConnection
        Dim Command As OleDb.OleDbCommand
        Dim SQLstr As String

        Try
            Mycn = New OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & My.Settings.roledbloc & "';Jet OLEDB:Database Password='" & My.Settings.dbpass & "';")

            Dim ds1 As New DataSet
            Dim eds1 As New DataGridView
            Dim dv1 As New DataView

            Mycn.Open()

            Dim srt As Integer
            If IsDBNull(txtAnswerSortOrder) Or txtAnswerSortOrder.Text = "" Then
                srt = "9999"
            Else
                srt = txtAnswerSortOrder.Text
            End If

            Dim ppl As Integer
            ppl = cboPeopleAnswer.SelectedValue

            Dim pro As String
            pro = cboProductAnswer.SelectedValue


            Dim comp As Integer
            comp = cboCompositeAnswer.SelectedValue


            Dim dte As Date
            If IsDBNull(dtpStaleAnswer) Then
                dte = "12/31/2099"
            Else
                dte = dtpStaleAnswer.Value
            End If

            Dim answer As String
            answer = RichTextBox2.Text
            answer = Replace(answer, "'", "''")

            SQLstr = "INSERT INTO rfp_Answers([Answer], [DateEntered], [LastUpdated], [RelatedToPeople], [ContactID], [RelatedToProduct], [Discipline], [Composite], [CompositeID], [DateSensitive], [StaleDate], [IsActive])" & _
            "VALUES('" & answer & "',#" & Format(Now()) & "#,#" & Format(Now()) & "#," & ckbPeopleAnswer.CheckState & "," & ppl & "," & ckbProductAnswer.CheckState & ",'" & pro & "'," & ckbCompositeAnswer.CheckState & "," & comp & "," & ckbDateSensitiveAnswer.CheckState & ",#" & dte & "#, -1)"

            Command = New OleDb.OleDbCommand(SQLstr, Mycn)
            Command.ExecuteNonQuery()
            Mycn.Close()

            Call LoadAnswerID()
            Call LoadSortOrders()

            Mycn.Open()
            SQLstr = "INSERT INTO rfp_QAAssignments(QuestionID, AnswerID, SortOrder, IsActive, CanEdit)" & _
            "VALUES(" & txtID.Text & "," & txtMasterAnswerID.Text & "," & txtAnswerSortOrder.Text & ", -1,-1)"

            Command = New OleDb.OleDbCommand(SQLstr, Mycn)
            Command.ExecuteNonQuery()
            Mycn.Close()

            Call LoadAnswersGrid()
            Call ResetAnswers()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "ERROR")

        End Try
    End Sub

    Public Sub LoadSortOrders()
        Dim Mycn As OleDb.OleDbConnection

        Try

            Mycn = New OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & My.Settings.roledbloc & "';Jet OLEDB:Database Password='" & My.Settings.dbpass & "';")

            Mycn.Open()

            Dim sqlstring As String

            sqlstring = "SELECT Top 1 SortOrder FROM rfp_QAAssignments WHERE IsActive = -1 AND QuestionID = " & txtID.Text & "" & _
            " GROUP BY SortOrder" & _
            " ORDER BY SortOrder DESC"

            Dim queryString As String = String.Format(sqlstring)
            Dim cmd As New OleDb.OleDbCommand(queryString, Mycn)
            Dim da As New OleDb.OleDbDataAdapter(cmd)
            Dim ds As New DataSet

            da.Fill(ds, "User")
            Dim dt As DataTable = ds.Tables("User")
            If dt.Rows.Count > 0 Then

                Dim row As DataRow = dt.Rows(0)
                'lblConfirm.Visible = True
                'lblConfirm.Text = "The last question was saved with ID: " & (row("ID"))
                Dim sord As Double
                sord = (row("SortOrder")) + 1
                txtAnswerSortOrder.Text = sord
            Else
                txtAnswerSortOrder.Text = "1"
            End If
            Mycn.Close()
            Mycn.Dispose()

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "ERROR")
            'Application.Exit()
        End Try
    End Sub
    Public Sub LoadAnswerID()
        Dim Mycn As OleDb.OleDbConnection

        Try

            Mycn = New OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & My.Settings.roledbloc & "';Jet OLEDB:Database Password='" & My.Settings.dbpass & "';")

            Mycn.Open()

            Dim sqlstring As String

            sqlstring = "SELECT Top 1 ID FROM rfp_Answers" & _
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
                txtMasterAnswerID.Text = (row("ID"))
            Else

            End If
            Mycn.Close()
            Mycn.Dispose()

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "ERROR")
            'Application.Exit()
        End Try
    End Sub

    Public Sub SaveSlaveAnswer()
        Dim Mycn As OleDb.OleDbConnection
        Dim Command As OleDb.OleDbCommand
        Dim SQLstr As String

        Try
            Mycn = New OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & My.Settings.roledbloc & "';Jet OLEDB:Database Password='" & My.Settings.dbpass & "';")

            Dim ds1 As New DataSet
            Dim eds1 As New DataGridView
            Dim dv1 As New DataView

            Mycn.Open()

            Dim srt As Integer
            If IsDBNull(txtAnswerSortOrder) Or txtAnswerSortOrder.Text = "" Then
                srt = "9999"
            Else
                srt = txtAnswerSortOrder.Text
            End If

            Dim ppl As Integer
            ppl = cboPeopleAnswer.SelectedValue

            Dim pro As String
            pro = cboProductAnswer.SelectedValue


            Dim comp As Integer
            comp = cboCompositeAnswer.SelectedValue


            Dim dte As Date
            If IsDBNull(dtpStaleAnswer) Then
                dte = "12/31/2099"
            Else
                dte = dtpStaleAnswer.Value
            End If

            Dim answer As String
            answer = RichTextBox2.Text
            answer = Replace(answer, "'", "''")

            SQLstr = "INSERT INTO rfp_Answers([Answer], [DateEntered], [LastUpdated], [RelatedToPeople], [ContactID], [RelatedToProduct], [Discipline], [Composite], [CompositeID], [DateSensitive], [StaleDate], [IsActive], [Slave], [MasterID])" & _
            "VALUES('" & answer & "',#" & Format(Now()) & "#,#" & Format(Now()) & "#," & ckbPeopleAnswer.CheckState & "," & ppl & "," & ckbProductAnswer.CheckState & ",'" & pro & "'," & ckbCompositeAnswer.CheckState & "," & comp & "," & ckbDateSensitiveAnswer.CheckState & ",#" & dte & "#, -1, -1," & txtMasterAnswerID.Text & ")"

            Command = New OleDb.OleDbCommand(SQLstr, Mycn)
            Command.ExecuteNonQuery()
            Mycn.Close()

            Call LoadAnswerID()

            If ckbIgnoreSort.Checked = True Then

            Else
                Call LoadSortOrders()
            End If

            If ckbIsSlave.Checked = True Then

            Else
                Mycn.Open()
                SQLstr = "INSERT INTO rfp_QAAssignments(QuestionID, AnswerID, SortOrder, IsActive, CanEdit, IsSlave)" & _
                "VALUES(" & txtID.Text & "," & txtMasterAnswerID.Text & "," & txtAnswerSortOrder.Text & ", -1,-1, -1)"

                Command = New OleDb.OleDbCommand(SQLstr, Mycn)
                Command.ExecuteNonQuery()
                Mycn.Close()
            End If

            Call LoadAnswersGrid()
            Call ResetAnswers()

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "ERROR")

        End Try
    End Sub

    Public Sub SaveEditAnswer()
        Dim Mycn As OleDb.OleDbConnection
        Dim Command As OleDb.OleDbCommand
        Dim SQLstr As String

        Try
            Mycn = New OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & My.Settings.roledbloc & "';Jet OLEDB:Database Password='" & My.Settings.dbpass & "';")

            Dim ds1 As New DataSet
            Dim eds1 As New DataGridView
            Dim dv1 As New DataView

            Mycn.Open()

            Dim srt As Integer
            If IsDBNull(txtAnswerSortOrder) Or txtAnswerSortOrder.Text = "" Then
                srt = "9999"
            Else
                srt = txtAnswerSortOrder.Text
            End If

            Dim ppl As Integer
            ppl = cboPeopleAnswer.SelectedValue

            Dim pro As String
            pro = cboProductAnswer.SelectedValue


            Dim comp As Integer
            comp = cboCompositeAnswer.SelectedValue


            Dim dte As Date
            If IsDBNull(dtpStaleAnswer) Then
                dte = "12/31/2099"
            Else
                dte = dtpStaleAnswer.Value
            End If

            Dim answer As String
            answer = RichTextBox2.Text
            answer = Replace(answer, "'", "''")

            SQLstr = "UPDATE rfp_Answers SET [Answer] = '" & answer & "', [LastUpdated] = #" & Format(Now()) & "#, [RelatedToPeople] = " & ckbPeopleAnswer.CheckState & ", [ContactID] = " & ppl & ", [RelatedToProduct] = " & ckbProductAnswer.CheckState & ", [Discipline] = '" & pro & "', [Composite] = " & ckbCompositeAnswer.CheckState & ", [CompositeID] = " & comp & ", [DateSensitive] = " & ckbDateSensitiveAnswer.CheckState & ", [StaleDate] = #" & dte & "#, [IsActive] = -1 WHERE ID = " & txtMasterAnswerID.Text

            Command = New OleDb.OleDbCommand(SQLstr, Mycn)
            Command.ExecuteNonQuery()
            Mycn.Close()

            If ckbIgnoreSort.Checked = True Then

            Else
                Call LoadSortOrders()
            End If

            If ckbIgnoreSort.Checked Then
            Else
                Mycn.Open()
                SQLstr = "INSERT INTO rfp_QAAssignments(QuestionID, AnswerID, SortOrder, IsActive, CanEdit)" & _
                "VALUES(" & txtID.Text & "," & txtMasterAnswerID.Text & "," & txtAnswerSortOrder.Text & ", -1,-1)"

                ckbIgnoreSort.Checked = False

                Command = New OleDb.OleDbCommand(SQLstr, Mycn)
                Command.ExecuteNonQuery()
                Mycn.Close()
            End If

            Call ResetAnswers()

            Call LoadAnswersGrid()

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "ERROR")
            'Application.Exit()
        End Try
    End Sub

    Public Sub SaveAnswer()
        Dim Mycn As OleDb.OleDbConnection
        Dim Command As OleDb.OleDbCommand
        Dim SQLstr As String

        Try
            Mycn = New OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & My.Settings.roledbloc & "';Jet OLEDB:Database Password='" & My.Settings.dbpass & "';")

            Dim ds1 As New DataSet
            Dim eds1 As New DataGridView
            Dim dv1 As New DataView

            Mycn.Open()

            SQLstr = "INSERT INTO rfp_QAAssignments(QuestionID, AnswerID, SortOrder, IsActive)" & _
            "VALUES(" & txtID.Text & "," & txtMasterAnswerID.Text & "," & txtAnswerSortOrder.Text & ",-1)"

            Command = New OleDb.OleDbCommand(SQLstr, Mycn)
            Command.ExecuteNonQuery()
            Mycn.Close()


            Call LoadAnswersGrid()
            Button4.Enabled = False
            Call ResetAnswers()
            Call LoadRealQuestions()

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "ERROR")
        End Try
    End Sub

    Private Sub Button8_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button8.Click
        If txtMasterAnswerID.Text = "NEW" Then
            MsgBox("Select a question to edit.", MsgBoxStyle.Information, "No Question Loaded.")
        Else
            If Button8.Text = "Edit" Then
                RichTextBox3.Text = RichTextBox2.Text
                RichTextBox2.Enabled = True
                DataGridView1.Enabled = False
                Button7.Enabled = False
                Button6.Enabled = False
                Button3.Enabled = False
                Button8.Text = "Reset"
            Else
                RichTextBox2.Text = RichTextBox3.Text
                Call ResetAnswers()
            End If
        End If
    End Sub

    Public Sub SpellCheckAnswer()
        If RichTextBox2.Text.Length > 0 Then
            Dim wordapp As New Word.Application
            wordapp.Visible = False
            Dim doc As Word.Document = wordapp.Documents.Add
            Dim range As Word.Range
            range = doc.Range
            range.Text = RichTextBox2.Text
            doc.Activate()
            doc.CheckSpelling()
            Dim chars() As Char = {CType(vbCr, Char), CType(vbLf, Char)}
            RichTextBox2.Text = doc.Range().Text.Trim(chars)
            doc.Close(SaveChanges:=False)
            wordapp.Quit()
        End If
    End Sub

    Private Sub Button5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button5.Click
        Call SpellCheckAnswer()
    End Sub

    Public Sub LoadAnswersGrid()
        Try

            Dim conn As New OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & My.Settings.dbloc & "';Jet OLEDB:Database Password='" & My.Settings.dbpass & "';")

            Dim strSQL As String = "SELECT rfp_QAAssignments.ID, rfp_QAAssignments.AnswerID, rfp_QAAssignments.SortOrder, rfp_QAAssignments.IsSlave, rfp_QAAssignments.CanEdit, rfp_Answers.Answer" & _
            " FROM rfp_Answers INNER JOIN rfp_QAAssignments ON rfp_Answers.ID = rfp_QAAssignments.AnswerID" & _
            " WHERE(((rfp_QAAssignments.QuestionID) = " & txtID.Text & ") AND rfp_QAAssignments.IsActive = -1)" & _
            " GROUP BY rfp_QAAssignments.ID, rfp_QAAssignments.AnswerID, rfp_QAAssignments.SortOrder, rfp_Answers.Answer, rfp_QAAssignments.CanEdit,rfp_QAAssignments.IsSlave" & _
            " ORDER BY rfp_QAAssignments.SortOrder"

            Dim da As New OleDb.OleDbDataAdapter(strSQL, conn)
            Dim ds As New DataSet
            da.Fill(ds, "Users")

            With DataGridView2
                .DataSource = ds.Tables("Users")
                .Columns(0).Visible = False
                .Columns(1).Visible = False
                '.Columns(2).Visible = False
                .Columns(3).Visible = False
                .Columns(4).Visible = False
            End With

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "ERROR")
        End Try
    End Sub

    Private Sub DataGridView1_CellContentDoubleClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView1.CellContentDoubleClick
        If DataGridView1.RowCount < 1 Then
            'do nothing
        Else
            Call LoadSortOrders()
            RichTextBox2.Text = DataGridView1.SelectedCells(3).Value
            txtMasterAnswerID.Text = DataGridView1.SelectedCells(1).Value

            Dim Mycn As OleDb.OleDbConnection

            Try

                Mycn = New OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & My.Settings.roledbloc & "';Jet OLEDB:Database Password='" & My.Settings.dbpass & "';")

                Mycn.Open()

                Dim sqlstring As String

                sqlstring = "SELECT * FROM rfp_Answers WHERE ID = " & txtMasterAnswerID.Text

                Dim queryString As String = String.Format(sqlstring)
                Dim cmd As New OleDb.OleDbCommand(queryString, Mycn)
                Dim da As New OleDb.OleDbDataAdapter(cmd)
                Dim ds As New DataSet

                da.Fill(ds, "User")
                Dim dt As DataTable = ds.Tables("User")
                If dt.Rows.Count > 0 Then

                    Dim row As DataRow = dt.Rows(0)

                    Dim ppl As Integer
                    ppl = (row("RelatedToPeople"))

                    If ppl = -1 Then
                        ckbPeopleAnswer.Checked = True
                        cboPeopleAnswer.SelectedValue = (row("ContactID"))
                    Else
                        ckbPeopleAnswer.Checked = False
                        cboPeopleAnswer.Visible = False
                    End If

                    Dim prd As Integer
                    prd = (row("RelatedToProduct"))

                    If prd = -1 Then
                        ckbProductAnswer.Checked = True
                        cboProductAnswer.SelectedValue = (row("Discipline"))
                    Else
                        ckbCompositeAnswer.Checked = False
                        cboProductAnswer.Visible = False
                    End If

                    Dim cmp As Integer
                    cmp = (row("Composite"))

                    If cmp = -1 Then
                        ckbCompositeAnswer.Checked = True
                        cboCompositeAnswer.SelectedValue = (row("CompositeID"))
                    Else
                        ckbCompositeAnswer.Checked = False
                        cboCompositeAnswer.Visible = False
                    End If

                    Dim stl As Integer
                    stl = (row("DateSensitive"))

                    If stl = -1 Then
                        ckbDateSensitiveAnswer.Checked = True
                        Dim dtestl As Date
                        dtestl = (row("StaleDate"))
                        dtpStaleAnswer.Text = dtestl
                    Else
                        ckbDateSensitiveAnswer.Checked = False
                        dtpStaleAnswer.Visible = False
                    End If

                    Button4.Enabled = True

                Else
                End If
                Mycn.Close()
                Mycn.Dispose()

            Catch ex As Exception
                MsgBox(ex.Message, MsgBoxStyle.Critical, "ERROR")
            End Try
        End If
    End Sub

    Private Sub Button9_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button9.Click
        Call LoadAnswersGrid()
    End Sub

    Private Sub Button10_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button10.Click
        Call SpellCheckQuestion()
    End Sub

    Public Sub SpellCheckQuestion()
        If rtbQuestion.Text.Length > 0 Then
            Dim wordapp As New Word.Application
            wordapp.Visible = False
            Dim doc As Word.Document = wordapp.Documents.Add
            Dim range As Word.Range
            range = doc.Range
            range.Text = rtbQuestion.Text
            doc.Activate()
            doc.CheckSpelling()
            Dim chars() As Char = {CType(vbCr, Char), CType(vbLf, Char)}
            rtbQuestion.Text = doc.Range().Text.Trim(chars)
            doc.Close(SaveChanges:=False)
            wordapp.Quit()
        End If
    End Sub

    Private Sub DataGridView2_CellContentDoubleClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView2.CellContentDoubleClick
        Call EditLoadedAnswer()
    End Sub

    Private Sub RemoveToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RemoveToolStripMenuItem.Click
        Dim ir As MsgBoxResult
        ir = MsgBox("Are you sure you want to Delete this Answer?", MsgBoxStyle.YesNo + MsgBoxStyle.Critical, "Confirm Delete")
        If ir = MsgBoxResult.Yes Then
            Call RemoveAssociatedAnswer()
        Else
            'do nothing
        End If
    End Sub

    Public Sub RemoveAssociatedAnswer()
        Dim Mycn As OleDb.OleDbConnection
        Dim Command As OleDb.OleDbCommand
        Dim SQLstr As String

        Try
            Mycn = New OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & My.Settings.roledbloc & "';Jet OLEDB:Database Password='" & My.Settings.dbpass & "';")

            Dim ds1 As New DataSet
            Dim eds1 As New DataGridView
            Dim dv1 As New DataView

            Mycn.Open()

            SQLstr = "Update rfp_QAAssignments SET IsActive = 0 WHERE ID = " & DataGridView2.SelectedCells(0).Value

            Command = New OleDb.OleDbCommand(SQLstr, Mycn)
            Command.ExecuteNonQuery()
            Mycn.Close()

            Call LoadAnswersGrid()

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "ERROR")
        End Try
    End Sub

    Private Sub EditToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles EditToolStripMenuItem.Click
        Call EditLoadedAnswer()
    End Sub

    Public Sub EditLoadedAnswer()
        If DataGridView2.SelectedCells(4).Value = -1 Then
            If DataGridView2.RowCount < 1 Then
                'do nothing
            Else
                'Call LoadSortOrders()
                'RichTextBox2.Text = DataGridView1.SelectedCells(3).Value
                txtMasterAnswerID.Text = DataGridView2.SelectedCells(1).Value
                txtAnswerSortOrder.Text = DataGridView2.SelectedCells(2).Value

                Dim Mycn As OleDb.OleDbConnection

                Try

                    Mycn = New OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & My.Settings.roledbloc & "';Jet OLEDB:Database Password='" & My.Settings.dbpass & "';")

                    Mycn.Open()

                    Dim sqlstring As String

                    sqlstring = "SELECT * FROM rfp_Answers WHERE ID = " & DataGridView2.SelectedCells(1).Value

                    Dim queryString As String = String.Format(sqlstring)
                    Dim cmd As New OleDb.OleDbCommand(queryString, Mycn)
                    Dim da As New OleDb.OleDbDataAdapter(cmd)
                    Dim ds As New DataSet

                    da.Fill(ds, "User")
                    Dim dt As DataTable = ds.Tables("User")
                    If dt.Rows.Count > 0 Then

                        Dim row As DataRow = dt.Rows(0)

                        RichTextBox2.Text = (row("Answer"))

                        Dim ppl As Integer
                        ppl = (row("RelatedToPeople"))

                        If ppl = -1 Then
                            ckbPeopleAnswer.Checked = True
                            cboPeopleAnswer.SelectedValue = (row("ContactID"))
                        Else
                            ckbPeopleAnswer.Checked = False
                            cboPeopleAnswer.Visible = False
                        End If

                        Dim prd As Integer
                        prd = (row("RelatedToProduct"))

                        If prd = -1 Then
                            ckbProductAnswer.Checked = True
                            cboProductAnswer.SelectedValue = (row("Discipline"))
                        Else
                            ckbCompositeAnswer.Checked = False
                            cboProductAnswer.Visible = False
                        End If

                        Dim cmp As Integer
                        cmp = (row("Composite"))

                        If cmp = -1 Then
                            ckbCompositeAnswer.Checked = True
                            cboCompositeAnswer.SelectedValue = (row("CompositeID"))
                        Else
                            ckbCompositeAnswer.Checked = False
                            cboCompositeAnswer.Visible = False
                        End If

                        Dim stl As Integer
                        stl = (row("DateSensitive"))

                        If stl = -1 Then
                            ckbDateSensitiveAnswer.Checked = True
                            Dim dtestl As Date
                            dtestl = (row("StaleDate"))
                            dtpStaleAnswer.Text = dtestl
                        Else
                            ckbDateSensitiveAnswer.Checked = False
                            dtpStaleAnswer.Visible = False
                        End If

                        'Button4.Enabled = True
                        'Button8.Text = "Reset"
                        ckbIgnoreSort.Checked = True

                        If DataGridView2.SelectedCells(3).Value = -1 Then
                            ckbIsSlave.Checked = True
                            Button4.Enabled = True
                            Button3.Text = "Reset"
                            Button7.Enabled = False
                            Button8.Enabled = False
                        Else
                            ckbIsSlave.Checked = False
                            Button4.Enabled = True
                            Button8.Text = "Reset"
                            Button3.Enabled = False
                            Button7.Enabled = False
                            ckbPeopleAnswer.Checked = False
                            ckbPeopleAnswer.Enabled = True
                            ckbProductAnswer.Checked = False
                            ckbProductAnswer.Enabled = True
                            ckbCompositeAnswer.Checked = False
                            ckbCompositeAnswer.Enabled = True
                            cboPeopleAnswer.Visible = False
                            cboProductAnswer.Visible = False
                            cboCompositeAnswer.Visible = False
                            ckbDateSensitiveAnswer.Checked = False
                            ckbDateSensitiveAnswer.Enabled = True
                            Label13.Visible = False
                            dtpStaleAnswer.Visible = False
                        End If


                        RichTextBox2.Enabled = True

                    Else
                    End If
                    Mycn.Close()
                    Mycn.Dispose()

                Catch ex As Exception
                    MsgBox(ex.Message, MsgBoxStyle.Critical, "ERROR")
                End Try
            End If
        Else
            Dim ir As MsgBoxResult
            ir = MsgBox("This Answer cannot be edited.  Would you like to remove this answer?" & vbNewLine & "Once removed the answer will appear in the Related Questions section.", MsgBoxStyle.YesNo + MsgBoxStyle.Information, "Cannot Edit")
            If ir = MsgBoxResult.Yes Then
                Call RemoveAssociatedAnswer()
            Else
                'do nothing
            End If
        End If
    End Sub

    Private Sub ckbLoadAll_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ckbLoadAll.CheckedChanged
        Call LoadRealQuestions()
    End Sub

    Public Sub LoadRealQuestions()
        If ckbLoadAll.Checked Then
            Call LoadAllQuestions()
        Else
            Call LoadQuestions()
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

        If IsDBNull(txtRptID) Or txtRptID.Text = "" Then
            txtRptID.Text = 0
        Else

        End If

        If txtID.Text = "NEW" Then
            Call SaveNew()
        Else
            Call SaveOld()
        End If

line1:
    End Sub

    Public Sub SaveNew()
        Dim Mycn As OleDb.OleDbConnection
        Dim Command As OleDb.OleDbCommand
        Dim SQLstr As String

        Try
            Mycn = New OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & My.Settings.roledbloc & "';Jet OLEDB:Database Password='" & My.Settings.dbpass & "';")

            Dim ds1 As New DataSet
            Dim eds1 As New DataGridView
            Dim dv1 As New DataView

            Mycn.Open()

            Dim ppl As Integer
            ppl = cboPeople.SelectedValue

            Dim pro As String
            pro = cboProduct.SelectedValue

            Dim pplid As Integer
            If IsDBNull(cboPeople.SelectedValue) Then
                pplid = 0
            Else
                pplid = cboPeople.SelectedValue
            End If

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

            SQLstr = "INSERT INTO rfp_ReportQuestions([Question], [AddedBy], [AddedDateStamp], [IsActive], [RelatedToPeople], [ContactID], [RelatedToProduct], [Discipline], [Composite], [CompositeID], [DateSensitive], [StaleDate], [NoReport], [LevelID], [StageID])" & _
            "VALUES('" & question & "'," & My.Settings.userid & ",#" & Format(Now()) & "#," & ckbActive.CheckState & "," & cbPeople.CheckState & "," & pplid & "," & cbProduct.CheckState & ",'" & pro & "'," & cbComposite.CheckState & "," & comp & "," & cbStale.CheckState & ",#" & dte & "#,-1," & cboLevel.SelectedValue & "," & cboStage.SelectedValue & ")"

            Command = New OleDb.OleDbCommand(SQLstr, Mycn)
            Command.ExecuteNonQuery()
            Mycn.Close()

            Call LoadQuestionID()
            MsgBox("Question Saved!", MsgBoxStyle.Information, "Saved")

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "ERROR")
        End Try
    End Sub

    Public Sub LoadQuestionID()
        Dim Mycn As OleDb.OleDbConnection

        Try

            Mycn = New OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & My.Settings.roledbloc & "';Jet OLEDB:Database Password='" & My.Settings.dbpass & "';")

            Mycn.Open()

            Dim sqlstring As String

            sqlstring = "SELECT Top 1 ID FROM rfp_ReportQuestions" & _
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
            Else

            End If
            Mycn.Close()
            Mycn.Dispose()

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "ERROR")
            'Application.Exit()
        End Try
    End Sub

    Public Sub SaveOld()
        Dim Mycn As OleDb.OleDbConnection
        Dim Command As OleDb.OleDbCommand
        Dim SQLstr As String

        Call QuestionTracker()
        'Try
        Mycn = New OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & My.Settings.roledbloc & "';Jet OLEDB:Database Password='" & My.Settings.dbpass & "';")

        Dim ds1 As New DataSet
        Dim eds1 As New DataGridView
        Dim dv1 As New DataView

        Mycn.Open()

        Dim ppl As Integer
        ppl = cboPeople.SelectedValue

        Dim pro As String
        pro = cboProduct.SelectedValue


        Dim comp As Integer
        comp = cboComposite.SelectedValue

        Dim pplid As Integer
        If IsDBNull(cboPeople.SelectedValue) Then
            pplid = 0
        Else
            pplid = cboPeople.SelectedValue
        End If

        Dim dte As Date
        If IsDBNull(DTPStale) Then
            dte = "12/31/2099"
        Else
            dte = DTPStale.Value
        End If

        Dim question As String
        question = rtbQuestion.Text
        question = Replace(question, "'", "''")

        SQLstr = "Update rfp_ReportQuestions SET [Question] = '" & question & "', [IsActive] = " & ckbActive.CheckState & ", [RelatedToPeople]=" & cbPeople.CheckState & ", [ContactID]=" & pplid & ", [RelatedToProduct]=" & cbProduct.CheckState & ", [Discipline]='" & pro & "', [Composite]=" & cbComposite.CheckState & ", [CompositeID]=" & comp & ", [DateSensitive]=" & cbStale.CheckState & ", [StaleDate]=#" & dte & "#, [LevelID]=" & cboLevel.SelectedValue & ", [StageID]=" & cboStage.SelectedValue & " WHERE ID = " & txtID.Text


        Command = New OleDb.OleDbCommand(SQLstr, Mycn)
        Command.ExecuteNonQuery()
        Mycn.Close()

        MsgBox("Question Saved!", MsgBoxStyle.Information, "Saved")
        'Call LoadQuestionID()

        'Catch ex As Exception
        'MsgBox(ex.Message, MsgBoxStyle.Critical, "ERROR")
        'End Try
    End Sub

    Private Sub cmdCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdCancel.Click
        If txtID.Text = "NEW" Then
            MsgBox("You must save the record before it can be deleted", MsgBoxStyle.Information, "Save Record")
        Else
            If ckbActive.Checked Then
                Dim ir As MsgBoxResult
                ir = MsgBox("Are you sure you want to delete this question?", MsgBoxStyle.YesNo + MsgBoxStyle.Critical, "Confirm Delete")
                If ir = MsgBoxResult.Yes Then
                    ckbActive.Checked = False
                    lblDeleted.Visible = True
                    cmdSave.Enabled = False
                    cmdCancel.Text = "Un Delete"
                    Call DeleteRecord()
                Else

                End If
            Else
                ckbActive.Checked = True
                lblDeleted.Visible = False
                cmdSave.Enabled = True
                cmdCancel.Text = "Delete"
                Call DeleteRecord()
            End If
        End If
    End Sub

    Public Sub DeleteRecord()
        Dim Mycn As OleDb.OleDbConnection
        Dim Command As OleDb.OleDbCommand
        Dim SQLstr As String

        Try
            Mycn = New OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & My.Settings.roledbloc & "';Jet OLEDB:Database Password='" & My.Settings.dbpass & "';")

            If ckbActive.Checked = True Then
                Mycn.Open()
                SQLstr = "Update rfp_ReportQuestions SET [IsActive] = True WHERE ID = " & txtID.Text
                Command = New OleDb.OleDbCommand(SQLstr, Mycn)
                Command.ExecuteNonQuery()
                Mycn.Close()

                Mycn.Open()
                SQLstr = "INSERT INTO rfp_ResponseTracking([QuestionID], [Field], [NewValue], [ChangedBy], [DateStamp])" & _
                "VALUES(" & txtID.Text & ",'Question Restored','" & ckbActive.CheckState & "','" & Environ("USERNAME") & "',#" & Format(Now()) & "#)"
                Command = New OleDb.OleDbCommand(SQLstr, Mycn)
                Command.ExecuteNonQuery()
                Mycn.Close()
            Else
                Mycn.Open()
                SQLstr = "Update rfp_ReportQuestions SET [IsActive] = False WHERE ID = " & txtID.Text
                Command = New OleDb.OleDbCommand(SQLstr, Mycn)
                Command.ExecuteNonQuery()
                Mycn.Close()

                Mycn.Open()
                SQLstr = "INSERT INTO rfp_ResponseTracking([QuestionID], [Field], [NewValue], [ChangedBy], [DateStamp])" & _
                "VALUES(" & txtID.Text & ",'Question Deleted','" & ckbActive.CheckState & "','" & Environ("USERNAME") & "',#" & Format(Now()) & "#)"
                Command = New OleDb.OleDbCommand(SQLstr, Mycn)
                Command.ExecuteNonQuery()
                Mycn.Close()
            End If

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "ERROR")
        End Try
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click

        If txtID.Text = "NEW" Then
            Dim ir As MsgBoxResult
            ir = MsgBox("Are you sure you want to close without saving?", MsgBoxStyle.YesNo + MsgBoxStyle.Information, "Confirm Close")
            If ir = MsgBoxResult.Yes Then
                Me.Close()
            Else

            End If
        Else


            Dim question1 As String
            question1 = rtbQuestion.Text
            question1 = Replace(question1, "'", "''")

            Dim question2 As String
            question2 = RichTextBox1.Text
            question2 = Replace(question2, "'", "''")


            'Mycn.Open()
            If question1 = question2 Then
                'nothing changed
                GoTo line1
            Else
                Dim ir As MsgBoxResult
                ir = MsgBox("Changes detected.  Are you sure you want to close without saving?", MsgBoxStyle.YesNo + MsgBoxStyle.Information, "Changes Detected")
                If ir = MsgBoxResult.Yes Then
                    Me.Close()
                Else

                End If
            End If

            If cboLevel.SelectedValue = txtLevelID.Text Then
                'nothing changed
                GoTo line1
            Else
                Dim ir As MsgBoxResult
                ir = MsgBox("Changes detected.  Are you sure you want to close without saving?", MsgBoxStyle.YesNo + MsgBoxStyle.Information, "Changes Detected")
                If ir = MsgBoxResult.Yes Then
                    Me.Close()
                Else

                End If
            End If

            If cboPeople.SelectedValue = txtContactID.Text Then
                'nothing changed
                GoTo line1
            Else
                Dim ir As MsgBoxResult
                ir = MsgBox("Changes detected.  Are you sure you want to close without saving?", MsgBoxStyle.YesNo + MsgBoxStyle.Information, "Changes Detected")
                If ir = MsgBoxResult.Yes Then
                    Me.Close()
                Else

                End If
            End If

            If cboProduct.SelectedValue = txtDiscipline.Text Then
                'nothing changed
                GoTo line1
            Else
                Dim ir As MsgBoxResult
                ir = MsgBox("Changes detected.  Are you sure you want to close without saving?", MsgBoxStyle.YesNo + MsgBoxStyle.Information, "Changes Detected")
                If ir = MsgBoxResult.Yes Then
                    Me.Close()
                Else

                End If
            End If

            If cboComposite.SelectedValue = txtCompositeID.Text Then
                'nothing changed
                GoTo line1
            Else
                Dim ir As MsgBoxResult
                ir = MsgBox("Changes detected.  Are you sure you want to close without saving?", MsgBoxStyle.YesNo + MsgBoxStyle.Information, "Changes Detected")
                If ir = MsgBoxResult.Yes Then
                    Me.Close()
                Else

                End If
            End If

            If DTPStale.Text = txtStaleDate.Text Then
                'nothing changed
                GoTo line1
            Else
                Dim ir As MsgBoxResult
                ir = MsgBox("Changes detected.  Are you sure you want to close without saving?", MsgBoxStyle.YesNo + MsgBoxStyle.Information, "Changes Detected")
                If ir = MsgBoxResult.Yes Then
                    Me.Close()
                Else

                End If
            End If

            If cbComposite.Checked = ckbComposite.CheckState Then
                'nothing changed
                GoTo line1
            Else
                Dim ir As MsgBoxResult
                ir = MsgBox("Changes detected.  Are you sure you want to close without saving?", MsgBoxStyle.YesNo + MsgBoxStyle.Information, "Changes Detected")
                If ir = MsgBoxResult.Yes Then
                    Me.Close()
                Else

                End If
            End If

            If cbPeople.Checked = ckbPeople.CheckState Then
                'nothing changed
                GoTo line1
            Else
                Dim ir As MsgBoxResult
                ir = MsgBox("Changes detected.  Are you sure you want to close without saving?", MsgBoxStyle.YesNo + MsgBoxStyle.Information, "Changes Detected")
                If ir = MsgBoxResult.Yes Then
                    Me.Close()
                Else

                End If
            End If

            If cbProduct.Checked = ckbProduct.CheckState Then
                'nothing changed
                GoTo line1
            Else
                Dim ir As MsgBoxResult
                ir = MsgBox("Changes detected.  Are you sure you want to close without saving?", MsgBoxStyle.YesNo + MsgBoxStyle.Information, "Changes Detected")
                If ir = MsgBoxResult.Yes Then
                    Me.Close()
                Else

                End If
            End If

            If cbStale.Checked = ckbDate1.CheckState Then
                'nothing changed
                GoTo line1
            Else
                Dim ir As MsgBoxResult
                ir = MsgBox("Changes detected.  Are you sure you want to close without saving?", MsgBoxStyle.YesNo + MsgBoxStyle.Information, "Changes Detected")
                If ir = MsgBoxResult.Yes Then
                    Me.Close()
                Else

                End If
            End If

            If cboStage.SelectedValue = txtStageID.Text Then
                'nothing changed
                GoTo line1
            Else
                Dim ir As MsgBoxResult
                ir = MsgBox("Changes detected.  Are you sure you want to close without saving?", MsgBoxStyle.YesNo + MsgBoxStyle.Information, "Changes Detected")
                If ir = MsgBoxResult.Yes Then
                    Me.Close()
                Else

                End If
            End If
        End If

line1:
        Me.Close()
line2:


    End Sub

    Public Sub LoadForm()
        Dim Mycn As OleDb.OleDbConnection

        Try

            Mycn = New OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & My.Settings.roledbloc & "';Jet OLEDB:Database Password='" & My.Settings.dbpass & "';")

            Mycn.Open()

            Dim sqlstring As String

            sqlstring = "SELECT * FROM rfp_ReportQuestions WHERE ID = " & txtID.Text

            Dim queryString As String = String.Format(sqlstring)
            Dim cmd As New OleDb.OleDbCommand(queryString, Mycn)
            Dim da As New OleDb.OleDbDataAdapter(cmd)
            Dim ds As New DataSet

            da.Fill(ds, "User")
            Dim dt As DataTable = ds.Tables("User")
            If dt.Rows.Count > 0 Then

                Dim row As DataRow = dt.Rows(0)
                txtLevelID.Text = (row("LevelID"))
                Call LoadLevel()
                'cboLevel.SelectedValue = (row("LevelID"))
                txtStageID.Text = (row("StageID"))
                Call LoadStage()
                'cboStage.SelectedValue = (row("StageID"))
                RichTextBox1.Text = rtbQuestion.Text
                cbPeople.Checked = (row("RelatedToPeople"))
                ckbPeople.Checked = (row("RelatedToPeople"))
                Dim cid As Integer
                cid = (row("ContactID"))
                If cid = 0 Then
                Else
                    txtContactID.Text = (row("ContactID"))
                    Call LoadPeople()
                    'cboPeople.SelectedValue = (row("ContactID"))

                End If
                cbProduct.Checked = (row("RelatedToProduct"))
                ckbProduct.Checked = (row("RelatedToProduct"))
                If IsDBNull(row("Discipline")) Then
                Else
                    txtDiscipline.Text = (row("Discipline"))
                    Call LoadProduct()
                    'cboProduct.SelectedValue = (row("Discipline"))

                End If
                cbComposite.Checked = (row("Composite"))
                ckbComposite.Checked = (row("Composite"))
                If IsDBNull(row("CompositeID")) Then
                Else
                    txtCompositeID.Text = (row("CompositeID"))
                    Call LoadComposites()
                    'cboComposite.SelectedValue = (row("CompositeID"))

                End If
                cbStale.Checked = (row("DateSensitive"))
                ckbDate1.Checked = (row("DateSensitive"))
                DTPStale.Text = (row("StaleDate"))
                txtStaleDate.Text = (row("StaleDate"))

            Else

            End If
            Mycn.Close()
            Mycn.Dispose()

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "ERROR")
        End Try
    End Sub

    Public Sub QuestionTracker()
        Dim Mycn As OleDb.OleDbConnection
        Dim Command As OleDb.OleDbCommand
        Dim SQLstr As String
        'Try
        Mycn = New OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & My.Settings.dbloc & "';Jet OLEDB:Database Password='" & My.Settings.dbpass & "';")

        Dim question1 As String
        question1 = rtbQuestion.Text
        question1 = Replace(question1, "'", "''")

        Dim question2 As String
        question2 = RichTextBox1.Text
        question2 = Replace(question2, "'", "''")


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

        Dim cid As Double
        'cid = txtContactID.Text
        If txtContactID.Text = "" Then
            cid = 0
        Else
            cid = txtContactID.Text
        End If
        If cboPeople.SelectedValue = cid Then
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

        'Catch ex As Exception
        'MsgBox(ex.Message, MsgBoxStyle.Critical, "ERROR")

        'End Try
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        If txtID.Text = "NEW" Then
        Else

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

        End If
    End Sub
End Class