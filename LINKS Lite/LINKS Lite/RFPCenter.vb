Public Class RFPCenter
    Dim thread1 As System.Threading.Thread
    Dim thPeople As System.Threading.Thread
    Dim thProduct As System.Threading.Thread
    Dim thComposite As System.Threading.Thread

    Dim thWorkingRFPs As System.Threading.Thread
    Dim thNewRFPs As System.Threading.Thread
    Dim thFinishedRFPs As System.Threading.Thread

    Dim thWorkingRFPReport As System.Threading.Thread
    Dim thNewRFPReport As System.Threading.Thread
    Dim thFinishedRFPReport As System.Threading.Thread

    Private Sub RFPCenter_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles Me.KeyDown
        If e.KeyData = Keys.Control + Keys.C Then
            Clipboard.SetText(RichTextBox1.SelectedText & vbNewLine & vbNewLine & My.Application.Info.Copyright)
        Else

        End If
    End Sub

    Private Sub RFPCenter_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        'Validate Permissions
        If Permissions.RFPQuestionView.Checked Then
            Label5.Visible = True

            Control.CheckForIllegalCrossThreadCalls = False
            thread1 = New System.Threading.Thread(AddressOf LoadLevel)
            thread1.Start()
        Else
            TabControl1.TabPages.Remove(TabPage3)
            'GoTo Line1
        End If

        If Permissions.RFPWork.Checked Then
            Control.CheckForIllegalCrossThreadCalls = False
            thWorkingRFPs = New System.Threading.Thread(AddressOf LoadWorkingRFP)
            thWorkingRFPs.Start()

            thNewRFPs = New System.Threading.Thread(AddressOf LoadNewRFP)
            thNewRFPs.Start()

            thFinishedRFPs = New System.Threading.Thread(AddressOf LoadFinishedRFP)
            thFinishedRFPs.Start()

        Else
            TabControl1.TabPages.Remove(TabPage4)
            'GoTo Line1
        End If

        If Permissions.RFPQuestionView.Checked Then
            Button2.Visible = True
        Else
            Button2.Visible = False
        End If

        If Permissions.RFPQuestionEdit.Checked Then
            Button1.Visible = True
        Else
            Button1.Visible = False
        End If

        'ComboBox1.Visible = False
        

Line1:

    End Sub

    Public Sub LoadWorkingRFP()
        Try

            DataGridView2.Enabled = False
            cmdReloadWorkingRFP.Enabled = False
            lblWorkingLoading.Visible = True

            Dim conn As New OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & My.Settings.dbloc & "';Jet OLEDB:Database Password='" & My.Settings.dbpass & "';")

            Dim strSQL As String = "SELECT * FROM rfp_WorkingReports"
            Dim da As New OleDb.OleDbDataAdapter(strSQL, conn)
            Dim ds As New DataSet
            da.Fill(ds, "Users")

            With DataGridView2
                .DataSource = ds.Tables("Users")
                .Columns(0).Visible = False
            End With

            Label9.Text = "Open/Working (" & DataGridView2.RowCount & "):"

            DataGridView2.Enabled = True
            cmdReloadWorkingRFP.Enabled = True
            lblWorkingLoading.Visible = False

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "ERROR")
            'Application.Exit()
        End Try
    End Sub

    Public Sub LoadNewRFP()
        Try

            DataGridView3.Enabled = False
            cmdReloadNewRFP.Enabled = False
            lblNewLoading.Visible = True

            Dim conn As New OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & My.Settings.dbloc & "';Jet OLEDB:Database Password='" & My.Settings.dbpass & "';")

            Dim strSQL As String = "SELECT * FROM rfp_NewReports"
            Dim da As New OleDb.OleDbDataAdapter(strSQL, conn)
            Dim ds As New DataSet
            da.Fill(ds, "Users")

            With DataGridView3
                .DataSource = ds.Tables("Users")
                .Columns(0).Visible = False
            End With

            Label10.Text = "New/Not Started (" & DataGridView3.RowCount & "):"

            DataGridView3.Enabled = True
            cmdReloadNewRFP.Enabled = True
            lblNewLoading.Visible = False

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "ERROR")
            'Application.Exit()
        End Try
    End Sub

    Public Sub LoadFinishedRFP()
        Try

            DataGridView4.Enabled = False
            cmdReloadFinishedRFP.Enabled = False
            lblFinishedLoading.Visible = True

            Dim conn As New OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & My.Settings.dbloc & "';Jet OLEDB:Database Password='" & My.Settings.dbpass & "';")

            Dim strSQL As String = "SELECT * FROM rfp_FinishedReports"
            Dim da As New OleDb.OleDbDataAdapter(strSQL, conn)
            Dim ds As New DataSet
            da.Fill(ds, "Users")

            With DataGridView4
                .DataSource = ds.Tables("Users")
                .Columns(0).Visible = False
            End With

            Label11.Text = "Finished/Submitted (" & DataGridView4.RowCount & "):"

            DataGridView4.Enabled = True
            cmdReloadFinishedRFP.Enabled = True
            lblFinishedLoading.Visible = False

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "ERROR")
            'Application.Exit()
        End Try
    End Sub

    Public Sub LoadLevel()
        Try

            Dim conn As New OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & My.Settings.dbloc & "';Jet OLEDB:Database Password='" & My.Settings.dbpass & "';")

            Dim strSQL As String = "SELECT ID, LevelName FROM rfp_Levels WHERE IsActive = -1"
            Dim da As New OleDb.OleDbDataAdapter(strSQL, conn)
            Dim ds As New DataSet
            da.Fill(ds, "Users")

            With ComboBox1
                .DataSource = ds.Tables("Users")
                .DisplayMember = "LevelName"
                .ValueMember = "ID"
                .SelectedIndex = 0
            End With

            Label5.Visible = False
            'ComboBox1.Visible = True

            'ComboBox1.SelectedItem = "View All"

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "ERROR")
            'Application.Exit()
        End Try
    End Sub

    Private Sub CheckBox1_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckBox1.CheckedChanged
        If CheckBox1.Checked Then
            ComboBox1.Enabled = True
        Else
            ComboBox1.Enabled = False
        End If
    End Sub

    Private Sub OK_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles OK.Click
        Call LoadQuestions()
    End Sub

    Public Sub LoadQuestions()
        Try

            Dim conn As New OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & My.Settings.dbloc & "';Jet OLEDB:Database Password='" & My.Settings.dbpass & "';")

            'If CheckBox1.Checked = False Then
            Dim strSQL As String = "SELECT ID, Question, ReportID FROM rfp_ReportQuestions WHERE [Question] LIKE '%" & TextBox3.Text & "%'"

            If CheckBox1.Checked Then
                strSQL = strSQL & " AND LevelID = " & ComboBox1.SelectedValue
            Else
                'do nothing
            End If

            If cbComposite.Checked Then
                strSQL = strSQL & " AND CompositeID = " & cboComposite.SelectedValue
            Else
                'do nothing
            End If

            If cbPeople.Checked Then
                strSQL = strSQL & " AND ContactID = " & cboPeople.SelectedValue
            Else
                'do nothing
            End If

            If cbProduct.Checked Then
                strSQL = strSQL & " AND Discipline = '" & cboProduct.SelectedValue & "'"
            Else
                'do nothing
            End If

            If cbDeleted.Checked Then
                strSQL = strSQL & " AND IsActive <> -1"
            Else
                strSQL = strSQL & " AND IsActive = -1"
            End If

            Dim da As New OleDb.OleDbDataAdapter(strSQL, conn)
            Dim ds As New DataSet
            da.Fill(ds, "Users")

            With DataGridView1
                .DataSource = ds.Tables("Users")
                .Columns(0).Visible = False
                .Columns(2).Visible = False
            End With

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "ERROR")
        End Try
    End Sub

    Public Sub LoadAnswer()

        Dim Mycn1 As OleDb.OleDbConnection

        'Try

        Mycn1 = New OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & My.Settings.dbloc & "';Jet OLEDB:Database Password='" & My.Settings.dbpass & "';")

        Mycn1.Open()

        Dim sqlstring As String

        'sqlstring = "SELECT * FROM rfp_ResponseBase WHERE ID = " & DataGridView1.SelectedCells(0).Value

        'sqlstring = "SELECT rfp_Answers.ID, rfp_QAAssignments.QuestionID, rfp_ReportQuestions.ReportID, rfp_ReportQuestions.StageID, rfp_Answers.IsActive, rfp_ReportQuestions.Question, rfp_Answers.Answer, rfp_Answers.DateEntered, rfp_Answers.LastUpdated, rfp_Answers.CompositeID, rfp_Answers.ContactID, rfp_ReportQuestions.LevelID, rfp_Answers.Discipline, rfp_Answers.StaleDate, rfp_Answers.Composite, rfp_Answers.DateSensitive, rfp_Answers.RelatedToProduct, rfp_Answers.RelatedToPeople, rfp_Answers.Stale" & _
        '" FROM (rfp_ReportQuestions INNER JOIN rfp_QAAssignments ON rfp_ReportQuestions.ID = rfp_QAAssignments.QuestionID) INNER JOIN rfp_Answers ON rfp_QAAssignments.AnswerID = rfp_Answers.ID" & _
        '" WHERE(((rfp_QAAssignments.IsSlave) = False) And ((rfp_QAAssignments.CanEdit) = True) And ((rfp_ReportQuestions.ID) = " & DataGridView1.SelectedCells(0).Value & "))" & _
        '" GROUP BY rfp_Answers.ID, rfp_QAAssignments.QuestionID, rfp_Answers.IsActive, rfp_ReportQuestions.StageID, rfp_ReportQuestions.ReportID,rfp_ReportQuestions.Question, rfp_Answers.Answer, rfp_Answers.DateEntered, rfp_Answers.LastUpdated, rfp_Answers.CompositeID, rfp_Answers.ContactID, rfp_ReportQuestions.LevelID, rfp_Answers.Discipline, rfp_Answers.StaleDate, rfp_Answers.Composite, rfp_Answers.DateSensitive, rfp_Answers.RelatedToProduct, rfp_Answers.RelatedToPeople, rfp_Answers.Stale"

        sqlstring = "SELECT Top 1 rfp_Answers.ID, rfp_QAAssignments.QuestionID, rfp_ReportQuestions.ReportID, rfp_ReportQuestions.StageID, rfp_ReportQuestions.IsActive, rfp_ReportQuestions.Question, rfp_Answers.Answer, rfp_Answers.DateEntered, rfp_Answers.LastUpdated, rfp_Answers.CompositeID, rfp_Answers.ContactID, rfp_ReportQuestions.LevelID, rfp_Answers.Discipline, rfp_Answers.StaleDate, rfp_Answers.Composite, rfp_Answers.DateSensitive, rfp_Answers.RelatedToProduct, rfp_Answers.RelatedToPeople, rfp_Answers.Stale" & _
       " FROM (rfp_ReportQuestions INNER JOIN rfp_QAAssignments ON rfp_ReportQuestions.ID = rfp_QAAssignments.QuestionID) INNER JOIN rfp_Answers ON rfp_QAAssignments.AnswerID = rfp_Answers.ID" & _
       " WHERE(((rfp_ReportQuestions.ID) = " & DataGridView1.SelectedCells(0).Value & "))" & _
       " GROUP BY rfp_Answers.ID, rfp_QAAssignments.QuestionID, rfp_ReportQuestions.IsActive, rfp_ReportQuestions.StageID, rfp_ReportQuestions.ReportID,rfp_ReportQuestions.Question, rfp_Answers.Answer, rfp_Answers.DateEntered, rfp_Answers.LastUpdated, rfp_Answers.CompositeID, rfp_Answers.ContactID, rfp_ReportQuestions.LevelID, rfp_Answers.Discipline, rfp_Answers.StaleDate, rfp_Answers.Composite, rfp_Answers.DateSensitive, rfp_Answers.RelatedToProduct, rfp_Answers.RelatedToPeople, rfp_Answers.Stale, rfp_QAAssignments.SortOrder" & _
       " ORDER BY rfp_QAAssignments.SortOrder;"


        Dim queryString As String = String.Format(sqlstring)
        Dim cmd As New OleDb.OleDbCommand(queryString, Mycn1)
        Dim da As New OleDb.OleDbDataAdapter(cmd)
        Dim ds As New DataSet

        da.Fill(ds, "User")
        Dim dt As DataTable = ds.Tables("User")
        If dt.Rows.Count > 0 Then

            Dim row As DataRow = dt.Rows(0)

            RichTextBox1.Text = "Answer ID: " & (row("ID")) & vbNewLine & vbNewLine
            txtID.Text = (row("ID"))
            txtQID.Text = (row("QuestionID"))

            If IsDBNull(row("ReportID")) Then
                txtRID.Text = "NEW"
            Else
                txtRID.Text = (row("ReportID"))
            End If

            If IsDBNull(row("StageID")) Then
                txtStageID.Text = 1
            Else
                txtStageID.Text = (row("StageID"))
            End If

            If (row("IsActive")) = -1 Then

            Else
                RichTextBox1.Text = RichTextBox1.Text & "**** THIS QUESTION HAS BEEN DELETED ****" & vbNewLine & vbNewLine
            End If

            If IsDBNull(row("Question")) Then
                RichTextBox1.Text = RichTextBox1.Text & "Question: NULL" & vbNewLine & vbNewLine & "ERROR.  Please correct or deactive this response."
                GoTo Line1
            Else
                RichTextBox1.Text = RichTextBox1.Text & "Question:" & vbNewLine & (row("Question"))
                rtbQuestion.Text = (row("Question"))
            End If

            If IsDBNull(row("Answer")) Then
                RichTextBox1.Text = RichTextBox1.Text & vbNewLine & vbNewLine & "Answer:" & vbNewLine & "**THIS QUESTION HAS NOT BEEN ANSWERED**"
            Else
                RichTextBox1.Text = RichTextBox1.Text & vbNewLine & vbNewLine & "Answer:" & vbNewLine & (row("Answer"))
                rtbAnswer.Text = (row("Answer"))
            End If

            If IsDBNull(row("DateEntered")) Then
                RichTextBox1.Text = RichTextBox1.Text & vbNewLine & vbNewLine & "Date Entered: NULL"
            Else
                RichTextBox1.Text = RichTextBox1.Text & vbNewLine & vbNewLine & "Date Entered:" & vbNewLine & (row("DateEntered"))
            End If

            If IsDBNull(row("LastUpdated")) Then
                RichTextBox1.Text = RichTextBox1.Text & vbNewLine & vbNewLine & "Last Updated: NULL"
            Else
                RichTextBox1.Text = RichTextBox1.Text & vbNewLine & vbNewLine & "Last Updated:" & vbNewLine & (row("LastUpdated"))
            End If

            txtCompositeID.Text = (row("CompositeID"))
            txtContactID.Text = (row("ContactID"))
            txtLevelID.Text = (row("LevelID"))
            txtDiscipline.Text = (row("Discipline"))
            txtStaleDate.Text = (row("StaleDate"))

            ckbActive.Checked = (row("IsActive"))
            ckbComposite.Checked = (row("Composite"))
            ckbDate1.Checked = (row("DateSensitive"))
            ckbPeople.Checked = (row("RelatedToPeople"))
            ckbProduct.Checked = (row("RelatedToProduct"))
            ckbStale.Checked = (row("Stale"))

        Else
            txtID.Text = "NEW"
            txtQID.Text = DataGridView1.SelectedCells(0).Value
            If IsDBNull(DataGridView1.SelectedCells(2).Value) Then
                txtRID.Text = "NEW"
            Else
                txtRID.Text = DataGridView1.SelectedCells(2).Value
            End If
            RichTextBox1.Text = "Answer ID: NEW" & vbNewLine & vbNewLine
            If IsDBNull(DataGridView1.SelectedCells(1)) Or DataGridView1.SelectedCells(1).Value = "" Then
                RichTextBox1.Text = RichTextBox1.Text & "Question:" & vbNewLine & "NULL"
                rtbQuestion.Text = ""
            Else
                RichTextBox1.Text = RichTextBox1.Text & "Question:" & vbNewLine & (DataGridView1.SelectedCells(1).Value)
                'RichTextBox1.Text = RichTextBox1.Text & DataGridView1.SelectedCells(1).Value
                rtbQuestion.Text = DataGridView1.SelectedCells(1).Value
            End If

            'GoTo Line1
            RichTextBox1.Text = RichTextBox1.Text & vbNewLine & vbNewLine & "Answer:" & vbNewLine & "**THIS QUESTION HAS NOT BEEN ANSWERED**"

            txtStageID.Text = 1
            txtLevelID.Text = 11

        End If
        Mycn1.Close()
        Mycn1.Dispose()


        'Catch ex As Exception
        'MsgBox(ex.Message, MsgBoxStyle.Critical, "ERROR")
        'End Try
Line1:

    End Sub

    Private Sub DataGridView1_CellContentClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs)

    End Sub

    Private Sub DataGridView1_CellContentDoubleClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs)
        Call LoadAnswer()
    End Sub

    Private Sub TextBox3_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles TextBox3.Enter
        Me.AcceptButton = OK
    End Sub

    Private Sub TextBox3_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox3.TextChanged

    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Dim child As New rfp_ReportQuestion
        child.MdiParent = Home
        child.txtID.Text = "NEW"
        child.txtRFPID.Text = "NEW"
        child.Show()
        'child.cmdCancel.Enabled = False
    End Sub

    Public Sub ReLoadAns()
        Dim Mycn1 As OleDb.OleDbConnection

        Try

            Mycn1 = New OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & My.Settings.dbloc & "';Jet OLEDB:Database Password='" & My.Settings.dbpass & "';")

            Mycn1.Open()

            Dim sqlstring As String

            sqlstring = "SELECT * FROM rfp_ResponseBase WHERE ID = " & txtID.Text

            Dim queryString As String = String.Format(sqlstring)
            Dim cmd As New OleDb.OleDbCommand(queryString, Mycn1)
            Dim da As New OleDb.OleDbDataAdapter(cmd)
            Dim ds As New DataSet

            da.Fill(ds, "User")
            Dim dt As DataTable = ds.Tables("User")
            If dt.Rows.Count > 0 Then

                Dim row As DataRow = dt.Rows(0)

                RichTextBox1.Text = "Question ID: " & (row("ID")) & vbNewLine & vbNewLine
                txtID.Text = (row("ID"))

                If (row("IsActive")) = -1 Then

                Else
                    RichTextBox1.Text = RichTextBox1.Text & "**** THIS QUESTION HAS BEEN DELETED ****" & vbNewLine & vbNewLine
                End If

                If IsDBNull(row("Question")) Then
                    RichTextBox1.Text = RichTextBox1.Text & "Question: NULL" & vbNewLine & vbNewLine & "ERROR.  Please correct or deactive this response."
                    GoTo Line1
                Else
                    RichTextBox1.Text = RichTextBox1.Text & "Question:" & vbNewLine & (row("Question"))
                    rtbQuestion.Text = (row("Question"))
                End If

                If IsDBNull(row("Answer")) Then
                    RichTextBox1.Text = RichTextBox1.Text & vbNewLine & vbNewLine & "Answer:" & vbNewLine & "**THIS QUESTION HAS NOT BEEN ANSWERED**"
                Else
                    RichTextBox1.Text = RichTextBox1.Text & vbNewLine & vbNewLine & "Answer:" & vbNewLine & (row("Answer"))
                    rtbAnswer.Text = (row("Answer"))
                End If

                If IsDBNull(row("DateEntered")) Then
                    RichTextBox1.Text = RichTextBox1.Text & vbNewLine & vbNewLine & "Date Entered: NULL"
                Else
                    RichTextBox1.Text = RichTextBox1.Text & vbNewLine & vbNewLine & "Date Entered:" & vbNewLine & (row("DateEntered"))
                End If

                If IsDBNull(row("LastUpdated")) Then
                    RichTextBox1.Text = RichTextBox1.Text & vbNewLine & vbNewLine & "Last Updated: NULL"
                Else
                    RichTextBox1.Text = RichTextBox1.Text & vbNewLine & vbNewLine & "Last Updated:" & vbNewLine & (row("LastUpdated"))
                End If

                txtCompositeID.Text = (row("CompositeID"))
                txtContactID.Text = (row("ContactID"))
                txtLevelID.Text = (row("LevelID"))
                txtDiscipline.Text = (row("Discipline"))
                txtStaleDate.Text = (row("StaleDate"))

                ckbActive.Checked = (row("IsActive"))
                ckbComposite.Checked = (row("Composite"))
                ckbDate1.Checked = (row("DateSensitive"))
                ckbPeople.Checked = (row("RelatedToPeople"))
                ckbProduct.Checked = (row("RelatedToProduct"))
                ckbStale.Checked = (row("Stale"))

            Else

            End If
            Mycn1.Close()
            Mycn1.Dispose()


        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "ERROR")
        End Try

Line1:

    End Sub


    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        If IsDBNull(txtID.Text) Or txtID.Text = "" Then
            MsgBox("You must load a question before you use this feature.", MsgBoxStyle.Information, "ERROR")
        Else
            'Call ReLoadAns()

            Dim child As New rfp_Answers
            child.MdiParent = Home

            child.txtID.Text = txtQID.Text
            child.txtRptID.Text = txtRID.Text
            child.rtbQuestion.Text = rtbQuestion.Text
            child.RichTextBox1.Text = rtbQuestion.Text
            child.cbComposite.Checked = ckbComposite.CheckState
            child.ckbComposite.Checked = ckbComposite.CheckState
            child.cbPeople.Checked = ckbPeople.CheckState
            child.ckbPeople.Checked = ckbPeople.CheckState
            child.cbProduct.Checked = ckbProduct.CheckState
            child.ckbProduct.Checked = ckbProduct.CheckState
            child.cbStale.Checked = ckbDate1.CheckState
            child.ckbDate1.Checked = ckbDate1.CheckState
            child.txtStageID.Text = txtStageID.Text

            child.txtLevelID.Text = txtLevelID.Text
            child.txtContactID.Text = txtContactID.Text
            child.txtDiscipline.Text = txtDiscipline.Text
            child.txtCompositeID.Text = txtCompositeID.Text
            child.DTPStale.Text = txtStaleDate.Text
            child.txtStaleDate.Text = txtStaleDate.Text
            child.lblLevel.Visible = False
            child.Label9.Visible = False
            child.ckbActive.Checked = ckbActive.CheckState
            Call child.LoadStage()
            Call child.LoadLevel()
            child.Show()
            'Dim child As New rfp_Answers
            'child.MdiParent = Home
            'child.Show()
            'child.txtID.Text = txtQID.Text
            'child.txtRptID.Text = txtID.Text
            'child.rtbQuestion.Text = rtbQuestion.Text
            'Call child.LoadForm()

            If ckbActive.Checked Then
                child.ckbActive.Checked = True
            Else
                child.cmdCancel.Text = "Un-Delete"
                child.cmdSave.Enabled = False
                child.lblDeleted.Visible = True
                child.ckbActive.Checked = False
            End If
        End If
    End Sub

    Private Sub cbProduct_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cbProduct.CheckedChanged
        If cbProduct.Checked Then
            Label7.Visible = True
            Control.CheckForIllegalCrossThreadCalls = False
            thProduct = New System.Threading.Thread(AddressOf LoadProduct)
            thProduct.Start()
            cboProduct.Visible = True
            cboProduct.Enabled = False
        Else
            Label7.Visible = False
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

    Private Sub cbPeople_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cbPeople.CheckedChanged
        If cbPeople.Checked Then
            Label8.Visible = True
            Control.CheckForIllegalCrossThreadCalls = False
            thPeople = New System.Threading.Thread(AddressOf LoadPeople)
            thPeople.Start()
            cboPeople.Visible = True
            cboPeople.Enabled = False
        Else
            Label8.Visible = False
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

            Label8.Visible = False
            cboPeople.Enabled = True

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

            Label7.Visible = False
            cboProduct.Enabled = True

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "ERROR")
        End Try
    End Sub

    Private Sub DataGridView1_CellContentClick_1(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick

    End Sub

    Private Sub DataGridView1_CellContentDoubleClick1(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView1.CellContentDoubleClick
        Call LoadAnswer()
    End Sub

    Private Sub CopyToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CopyToolStripMenuItem.Click
        If IsDBNull(RichTextBox1.SelectedText) Then

        Else
            Clipboard.SetText(RichTextBox1.SelectedText & vbNewLine & vbNewLine & My.Application.Info.Copyright)
        End If


    End Sub

    Private Sub RichTextBox1_KeyDown1(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles RichTextBox1.KeyDown
        If e.KeyData = Keys.Control + Keys.C Then
            Clipboard.SetText(RichTextBox1.SelectedText & vbNewLine & vbNewLine & My.Application.Info.Copyright)
        Else

        End If
    End Sub

    Private Sub RichTextBox1_Keydown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles RichTextBox1.KeyPress

    End Sub

    Private Sub RichTextBox1_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RichTextBox1.TextChanged

    End Sub

    Private Sub cmdAddRFP_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Dim child As New rfp_Report
        child.MdiParent = Home
        child.Show()
        child.txtID.Text = "NEW"
    End Sub

    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        Dim child As New rfp_Report
        child.MdiParent = Home
        child.Show()
        child.txtID.Text = "NEW"
    End Sub

    Private Sub cmdReloadWorkingRFP_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdReloadWorkingRFP.Click
        Control.CheckForIllegalCrossThreadCalls = False
        thWorkingRFPs = New System.Threading.Thread(AddressOf LoadWorkingRFP)
        thWorkingRFPs.Start()
    End Sub

    Private Sub cmdReloadNewRFP_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdReloadNewRFP.Click
        Control.CheckForIllegalCrossThreadCalls = False
        thNewRFPs = New System.Threading.Thread(AddressOf LoadNewRFP)
        thNewRFPs.Start()
    End Sub

    Private Sub cmdReloadFinishedRFP_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdReloadFinishedRFP.Click
        Control.CheckForIllegalCrossThreadCalls = False
        thFinishedRFPs = New System.Threading.Thread(AddressOf LoadFinishedRFP)
        thFinishedRFPs.Start()
    End Sub

    Private Sub DataGridView2_CellContentClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView2.CellContentClick

    End Sub

    Private Sub DataGridView2_CellDoubleClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView2.CellDoubleClick
        'Control.CheckForIllegalCrossThreadCalls = False
        'thWorkingRFPReport = New System.Threading.Thread(AddressOf LoadWorkingRFPReport)
        'thWorkingRFPReport.Start()

        Call LoadWorkingRFPReport()
    End Sub

    Public Sub LoadWorkingRFPReport()
        If DataGridView2.RowCount < 1 Then

        Else

            Dim Mycn1 As OleDb.OleDbConnection

            Try

                Mycn1 = New OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & My.Settings.dbloc & "';Jet OLEDB:Database Password='" & My.Settings.dbpass & "';")

                Mycn1.Open()

                Dim sqlstring As String

                sqlstring = "SELECT * FROM rfp_Reports WHERE ID = " & DataGridView2.SelectedCells(0).Value

                Dim queryString As String = String.Format(sqlstring)
                Dim cmd As New OleDb.OleDbCommand(queryString, Mycn1)
                Dim da As New OleDb.OleDbDataAdapter(cmd)
                Dim ds As New DataSet

                da.Fill(ds, "User")
                Dim dt As DataTable = ds.Tables("User")
                If dt.Rows.Count > 0 Then

                    Dim row As DataRow = dt.Rows(0)

                    Dim child As New rfp_Report
                    child.MdiParent = Home
                    child.Show()

                    child.txtlID.Text = (row("ID"))
                    child.txtID.Text = (row("ID"))

                    child.txtLrtpnme.Text = (row("Title"))
                    child.txtRptName.Text = child.txtLrtpnme.Text

                    child.txtlContactID.Text = (row("ContactID"))

                    child.txtlDueDate.Text = (row("DueDate"))
                    child.DateTimePicker1.Text = child.txtlDueDate.Text

                    child.txtlRqstBy.Text = (row("RequestedBy"))

                    'Reason can be Null
                    If IsDBNull(row("Reason")) Then

                    Else
                        child.txtlReason.Text = (row("Reason"))
                    End If

                    'Stage can be Null - but shouldnt be...
                    If IsDBNull(row("Stage")) Then

                    Else
                        child.txtlStage.Text = (row("Stage"))
                    End If


                    child.CheckBox2.Checked = (row("Working"))
                    child.CheckBox1.Checked = (row("Working"))

                    child.txtDateStarted.Text = (row("DateStarted"))

                    child.txtDateFinished.Text = (row("DateFinished"))

                    'worked by can be null
                    If IsDBNull(row("WorkedBy")) Then

                    Else
                        child.txtWorkedBy.Text = (row("WorkedBy"))
                    End If

                    If IsDBNull(row("CompanyName")) Then
                    Else
                        child.txtCompanyName.Text = (row("CompanyName"))
                        child.txtCompanyNameCopy.Text = (row("CompanyName"))
                    End If

                    If IsDBNull(row("PrimaryContact")) Then
                    Else
                        child.txtPrimaryContact.Text = (row("PrimaryContact"))
                        child.txtPrimaryContactCopy.Text = (row("PrimaryContact"))
                    End If

                    child.ckbPriContact.Checked = (row("PriContactShow"))
                    child.ckbPriContactCopy.Checked = (row("PriContactShow"))

                    If IsDBNull(row("Address")) Then
                    Else
                        child.txtAddress.Text = (row("Address"))
                        child.txtAddressCopy.Text = (row("Address"))
                    End If

                    If IsDBNull(row("City")) Then
                    Else
                        child.txtCity.Text = (row("City"))
                        child.txtCityCopy.Text = (row("City"))
                    End If

                    If IsDBNull(row("State")) Then
                    Else
                        child.txtState.Text = (row("State"))
                        child.txtStateCopy.Text = (row("State"))
                    End If

                    If IsDBNull(row("Zip")) Then
                    Else
                        child.txtZip.Text = (row("Zip"))
                        child.txtZipCopy.Text = (row("Zip"))
                    End If

                    If IsDBNull(row("Phone")) Then
                    Else
                        child.txtPhone.Text = (row("Phone"))
                        child.txtPhoneCopy.Text = (row("Phone"))
                    End If

                    If IsDBNull(row("Fax")) Then
                    Else
                        child.txtFax.Text = (row("Fax"))
                        child.txtFaxCopy.Text = (row("Fax"))
                    End If

                    If IsDBNull(row("Website")) Then
                    Else
                        child.txtWebsite.Text = (row("Website"))
                        child.txtWebsiteCopy.Text = (row("Website"))
                    End If

                    If IsDBNull(row("Email")) Then
                    Else
                        child.txtEmail.Text = (row("Email"))
                        child.txtEmailCopy.Text = (row("Email"))
                    End If

                    If IsDBNull(row("ProposedStrategy")) Then
                    Else
                        child.cboDiscipline.SelectedValue = (row("ProposedStrategy"))
                        child.txtDisciplineCopy.Text = (row("ProposedStrategy"))
                    End If

                    If IsDBNull(row("CustomizationNotes")) Then
                    Else
                        child.rtbCustomization.Text = (row("CustomizationNotes"))
                        child.rtbCustomizationCopy.Text = (row("CustomizationNotes"))
                    End If

                    If IsDBNull(row("DeliveryMethod")) Then
                    Else
                        child.cboDeliveryMethod.SelectedValue = (row("DeliveryMethod"))
                        child.txtDeliveryMethodCopy.Text = (row("DeliveryMethod"))
                    End If

                    If IsDBNull(row("TrackingNumber")) Then
                    Else
                        child.txtTrackingNumber.Text = (row("TrackingNumber"))
                        child.txtTrackingNumberCopy.Text = (row("TrackingNumber"))
                    End If

                    child.ckbLockData.Checked = (row("DataLock"))

                    If IsDBNull(row("AUM")) Then
                    Else
                        child.txtAUM.Text = (row("AUM"))
                        child.txtAUMCopy.Text = (row("AUM"))
                    End If

                    child.ckbFlatRate.Checked = (row("FlatRate"))
                    child.ckbFlatRateCopy.Checked = (row("FlatRate"))
                    If IsDBNull(row("FRate")) Then
                    Else
                        child.txtFlatRate.Text = (row("FRate"))
                        child.txtFlatRateCopy.Text = (row("FRate"))
                    End If

                    If IsDBNull(row("MinFee")) Then
                    Else
                        child.txtMinFeeCopy.Text = (row("MinFee"))
                        child.txtMinFee.Text = (row("MinFee"))
                    End If

                    If IsDBNull(row("B1To")) Then
                    Else
                        child.txtF1To.Text = (row("B1To"))
                        child.txtF1ToCopy.Text = (row("B1To"))
                    End If

                    If IsDBNull(row("B1Rate")) Then
                    Else
                        child.txtF1Rate.Text = (row("B1Rate"))
                        child.txtF1RateCopy.Text = (row("B1Rate"))
                    End If

                    If IsDBNull(row("B2To")) Then
                    Else
                        child.txtF2To.Text = (row("B2To"))
                        child.txtF2ToCopy.Text = (row("B2To"))
                    End If

                    If IsDBNull(row("B2Rate")) Then
                    Else
                        child.txtF2Rate.Text = (row("B2Rate"))
                        child.txtF2RateCopy.Text = (row("B2Rate"))
                    End If

                    If IsDBNull(row("B3To")) Then
                    Else
                        child.txtF3To.Text = (row("B3To"))
                        child.txtF3ToCopy.Text = (row("B3To"))
                    End If

                    If IsDBNull(row("B3Rate")) Then
                    Else
                        child.txtF3Rate.Text = (row("B3Rate"))
                        child.txtF3RateCopy.Text = (row("B3Rate"))
                    End If

                    If IsDBNull(row("B4To")) Then
                    Else
                        child.txtF4To.Text = (row("B4To"))
                        child.txtF4ToCopy.Text = (row("B4To"))
                    End If

                    If IsDBNull(row("B4Rate")) Then
                    Else
                        child.txtF4Rate.Text = (row("B4Rate"))
                        child.txtF4RateCopy.Text = (row("B4Rate"))
                    End If

                    If IsDBNull(row("B5To")) Then
                    Else
                        child.txtF5To.Text = (row("B5To"))
                        child.txtF5ToCopy.Text = (row("B5To"))
                    End If

                    If IsDBNull(row("B5Rate")) Then
                    Else
                        child.txtF5Rate.Text = (row("B5Rate"))
                        child.txtF5RateCopy.Text = (row("B5Rate"))
                    End If

                    If IsDBNull(row("B6Rate")) Then
                    Else
                        child.txtF6Rate.Text = (row("B6Rate"))
                        child.txtF6RateCopy.Text = (row("B6Rate"))
                    End If

                    child.ckbHasImage.Checked = (row("HasImage"))
                    child.ckbHasImageCopy.Checked = (row("HasImage"))

                    If IsDBNull(row("ImageLocation")) Then
                    Else
                        child.txtImageLocation.Text = (row("ImageLocation"))
                        child.txtImageLocationCopy.Text = (row("ImageLocation"))
                        Call child.loadimage()
                    End If

                    child.ckbDisclosure.Checked = (row("DefaultDisclosure"))
                    child.ckbDisclosureCopy.Checked = (row("DefaultDisclosure"))
                    Call child.pulldisclosure()

                    If IsDBNull(row("Narrative")) Then
                    Else
                        child.rtbNarrative.Text = (row("Narrative"))
                        'If IsDBNull(row("NarrativeFont")) Then
                        'Else
                        'Dim fnt As Font
                        'fnt = (row("NarrativeFont"))
                        'child.rtbNarrative.Font = fnt
                        'End If

                    End If

                    If IsDBNull(row("PBFirm")) Then
                    Else
                        child.txtPBFirm.Text = (row("PBFirm"))
                    End If

                    If IsDBNull(row("PBPerson")) Then
                    Else
                        child.txtPBPerson.Text = (row("PBPerson"))
                    End If

                    If IsDBNull(row("PBAddress")) Then
                    Else
                        child.txtPBAddress.Text = (row("PBAddress"))
                    End If

                    If IsDBNull(row("PBCity")) Then
                    Else
                        child.txtPBCity.Text = (row("PBCity"))
                    End If

                    If IsDBNull(row("PBState")) Then
                    Else
                        child.txtPBState.Text = (row("PBState"))
                    End If

                    If IsDBNull(row("PBZip")) Then
                    Else
                        child.txtPBZip.Text = (row("PBZip"))
                    End If

                    If IsDBNull(row("PBPhone")) Then
                    Else
                        child.txtPBPhone.Text = (row("PBPhone"))
                    End If

                Else
                End If
                    Mycn1.Close()
                    Mycn1.Dispose()


            Catch ex As Exception
                MsgBox(ex.Message, MsgBoxStyle.Critical, "ERROR")
            End Try

        End If
Line1:
    End Sub

    Private Sub DataGridView3_CellContentClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView3.CellContentClick

    End Sub

    Private Sub DataGridView3_CellContentDoubleClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView3.CellContentDoubleClick
        If DataGridView3.RowCount < 1 Then

        Else

            Dim Mycn1 As OleDb.OleDbConnection

            'Try

            Mycn1 = New OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & My.Settings.dbloc & "';Jet OLEDB:Database Password='" & My.Settings.dbpass & "';")

            Mycn1.Open()

            Dim sqlstring As String

            sqlstring = "SELECT * FROM rfp_Reports WHERE ID = " & DataGridView3.SelectedCells(0).Value

            Dim queryString As String = String.Format(sqlstring)
            Dim cmd As New OleDb.OleDbCommand(queryString, Mycn1)
            Dim da As New OleDb.OleDbDataAdapter(cmd)
            Dim ds As New DataSet

            da.Fill(ds, "User")
            Dim dt As DataTable = ds.Tables("User")
            If dt.Rows.Count > 0 Then

                Dim row As DataRow = dt.Rows(0)

                Dim child As New rfp_Report
                child.MdiParent = Home
                child.Show()

                child.txtlID.Text = (row("ID"))
                child.txtID.Text = (row("ID"))

                child.txtLrtpnme.Text = (row("Title"))
                child.txtRptName.Text = child.txtLrtpnme.Text

                child.txtlContactID.Text = (row("ContactID"))

                child.txtlDueDate.Text = (row("DueDate"))
                child.DateTimePicker1.Text = child.txtlDueDate.Text

                child.txtlRqstBy.Text = (row("RequestedBy"))

                'Reason can be Null
                If IsDBNull(row("Reason")) Then

                Else
                    child.txtlReason.Text = (row("Reason"))
                End If

                'Stage can be Null - but shouldnt be...
                If IsDBNull(row("Stage")) Then

                Else
                    child.txtlStage.Text = (row("Stage"))
                End If


                child.CheckBox2.Checked = (row("Working"))
                child.CheckBox1.Checked = (row("Working"))

                child.txtDateStarted.Text = (row("DateStarted"))

                child.txtDateFinished.Text = (row("DateFinished"))

                'worked by can be null
                If IsDBNull(row("WorkedBy")) Then

                Else
                    child.txtWorkedBy.Text = (row("WorkedBy"))
                End If


            Else

            End If
            Mycn1.Close()
            Mycn1.Dispose()


            'Catch ex As Exception
            '   MsgBox(ex.Message, MsgBoxStyle.Critical, "ERROR")
            'End Try

        End If
Line1:
    End Sub

    Private Sub DataGridView4_CellContentClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView4.CellContentClick

    End Sub

    Private Sub DataGridView4_CellContentDoubleClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView4.CellContentDoubleClick
        If DataGridView4.RowCount < 1 Then

        Else

            Dim Mycn1 As OleDb.OleDbConnection

            Try

                Mycn1 = New OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & My.Settings.dbloc & "';Jet OLEDB:Database Password='" & My.Settings.dbpass & "';")

                Mycn1.Open()

                Dim sqlstring As String

                sqlstring = "SELECT * FROM rfp_Reports WHERE ID = " & DataGridView4.SelectedCells(0).Value

                Dim queryString As String = String.Format(sqlstring)
                Dim cmd As New OleDb.OleDbCommand(queryString, Mycn1)
                Dim da As New OleDb.OleDbDataAdapter(cmd)
                Dim ds As New DataSet

                da.Fill(ds, "User")
                Dim dt As DataTable = ds.Tables("User")
                If dt.Rows.Count > 0 Then

                    Dim row As DataRow = dt.Rows(0)

                    Dim child As New rfp_Report
                    child.MdiParent = Home
                    child.Show()

                    child.txtlID.Text = (row("ID"))
                    child.txtID.Text = (row("ID"))

                    child.txtLrtpnme.Text = (row("Title"))
                    child.txtRptName.Text = child.txtLrtpnme.Text

                    child.txtlContactID.Text = (row("ContactID"))

                    child.txtlDueDate.Text = (row("DueDate"))
                    child.DateTimePicker1.Text = child.txtlDueDate.Text

                    child.txtlRqstBy.Text = (row("RequestedBy"))

                    'Reason can be Null
                    If IsDBNull(row("Reason")) Then

                    Else
                        child.txtlReason.Text = (row("Reason"))
                    End If

                    'Stage can be Null - but shouldnt be...
                    If IsDBNull(row("Stage")) Then

                    Else
                        child.txtlStage.Text = (row("Stage"))
                    End If


                    child.CheckBox2.Checked = (row("Working"))
                    child.CheckBox1.Checked = (row("Working"))

                    child.txtDateStarted.Text = (row("DateStarted"))

                    child.txtDateFinished.Text = (row("DateFinished"))

                    'worked by can be null
                    If IsDBNull(row("WorkedBy")) Then

                    Else
                        child.txtWorkedBy.Text = (row("WorkedBy"))
                    End If


                Else

                End If
                Mycn1.Close()
                Mycn1.Dispose()


            Catch ex As Exception
                MsgBox(ex.Message, MsgBoxStyle.Critical, "ERROR")
            End Try

        End If
Line1:
    End Sub

    Private Sub TextBox1_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox1.TextChanged

    End Sub

    Private Sub DataGridView1_CellDoubleClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView1.CellDoubleClick
        Call LoadAnswer()
    End Sub
End Class