Imports Microsoft.Office.Interop

Public Class rfp_ReportQuestion
    Dim thInsert As System.Threading.Thread

    Private Sub cmdNew_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdNew.Click
        If ckbAutoSpell.Checked Then
            Call SpellCheckRTB()
        Else

        End If

        Control.CheckForIllegalCrossThreadCalls = False
        thInsert = New System.Threading.Thread(AddressOf InsertRec)
        thInsert.Start()
        'Call InsertRec()
    End Sub

    Public Sub InsertRec()
        Dim Mycn As OleDb.OleDbConnection
        Dim Command As OleDb.OleDbCommand
        Dim SQLstr As String

        Try
            cmdNew.Enabled = False
            Button1.Enabled = False
            Mycn = New OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & My.Settings.roledbloc & "';Jet OLEDB:Database Password='" & My.Settings.dbpass & "';")

            Dim ds1 As New DataSet
            Dim eds1 As New DataGridView
            Dim dv1 As New DataView

            Mycn.Open()

            Dim srt As Integer
            If IsDBNull(txtSortOrder) Or txtSortOrder.Text = "" Then
                srt = "0"
            Else
                srt = txtSortOrder.Text
            End If

            Dim question As String
            question = rtbQuestion.Text
            question = Replace(question, "'", "''")

            Dim rid As Integer
            Dim norpt As Integer
            If txtRFPID.Text = "NEW" Then
                rid = 0
                norpt = -1
            Else
                rid = txtRFPID.Text
                norpt = 0
            End If

            SQLstr = "INSERT INTO rfp_ReportQuestions([SortOrder], [Question], [ReportID], [AddedBy], [AddedDateStamp], [IsActive], [LevelID], [StageID], [StaleDate], [ContactID], [CompositeID], [NoReport])" & _
            "VALUES(" & srt & ",'" & question & "'," & rid & ",'" & Environ("USERNAME") & "',#" & Format(Now()) & "#, -1,11,1,#" & Format(Now()) & "#,0,0," & norpt & ")"

            Command = New OleDb.OleDbCommand(SQLstr, Mycn)
            Command.ExecuteNonQuery()

            Mycn.Close()

            Call LoadID()

            Mycn.Open()
            SQLstr = "INSERT INTO rfp_ResponseTracking([QuestionID], [Field], [NewValue], [ChangedBy], [DateStamp])" & _
            "VALUES(" & txtID.Text & ",'Question Added','" & txtID.Text & "','" & Environ("USERNAME") & "',#" & Format(Now()) & "#)"
            Command = New OleDb.OleDbCommand(SQLstr, Mycn)
            Command.ExecuteNonQuery()
            Mycn.Close()

            Mycn.Open()
            SQLstr = "INSERT INTO rfp_ResponseTracking([QuestionID], [Field], [NewValue], [ChangedBy], [DateStamp])" & _
            "VALUES(" & txtID.Text & ",'Question','" & question & "','" & Environ("USERNAME") & "',#" & Format(Now()) & "#)"
            Command = New OleDb.OleDbCommand(SQLstr, Mycn)
            Command.ExecuteNonQuery()
            Mycn.Close()

            Mycn.Open()
            SQLstr = "INSERT INTO rfp_ResponseTracking([QuestionID], [Field], [NewValue], [ChangedBy], [DateStamp])" & _
            "VALUES(" & txtID.Text & ",'Report ID','" & txtRFPID.Text & "','" & Environ("USERNAME") & "',#" & Format(Now()) & "#)"
            Command = New OleDb.OleDbCommand(SQLstr, Mycn)
            Command.ExecuteNonQuery()
            Mycn.Close()

            txtSortOrder.Text = srt + 1

            rtbQuestion.Text = ""


            txtID.Text = "NEW"
            cmdNew.Enabled = True
            Button1.Enabled = True

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "ERROR")

        End Try
    End Sub

    Public Sub LoadID()
        Dim Mycn As OleDb.OleDbConnection

        'Try

        Mycn = New OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & My.Settings.roledbloc & "';Jet OLEDB:Database Password='" & My.Settings.dbpass & "';")

        Mycn.Open()

        Dim sqlstring As String

        Dim title As String
        title = rtbQuestion.Text
        title = Replace(title, "'", "''")

        Dim rid As Integer
        If txtRFPID.Text = "NEW" Then
            rid = 0
        Else
            rid = txtRFPID.Text
        End If

        sqlstring = "SELECT Top 1 ID, SortOrder FROM rfp_ReportQuestions WHERE IsActive = -1 AND ReportID = " & rid & _
        " GROUP BY ID, SortOrder" & _
        " ORDER BY ID DESC"

        Dim queryString As String = String.Format(sqlstring)
        Dim cmd As New OleDb.OleDbCommand(queryString, Mycn)
        Dim da As New OleDb.OleDbDataAdapter(cmd)
        Dim ds As New DataSet

        da.Fill(ds, "User")
        Dim dt As DataTable = ds.Tables("User")
        If dt.Rows.Count > 0 Then

            Dim row As DataRow = dt.Rows(0)
            lblConfirm.Visible = True
            lblConfirm.Text = "The last question was saved with ID: " & (row("ID"))
            txtID.Text = (row("ID"))
            Dim sord As Double
            sord = (row("SortOrder")) + 1
            txtSortOrder.Text = sord
        Else

        End If
        Mycn.Close()
        Mycn.Dispose()

        'Catch ex As Exception
        'MsgBox(ex.Message, MsgBoxStyle.Critical, "ERROR")
        'End Try
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        If ckbAutoSpell.Checked Then
            Call SpellCheckRTB()
        Else

        End If

        Control.CheckForIllegalCrossThreadCalls = False
        thInsert = New System.Threading.Thread(AddressOf InsertRecClose)
        thInsert.Start()
    End Sub

    Public Sub InsertRecClose()
        Dim Mycn As OleDb.OleDbConnection
        Dim Command As OleDb.OleDbCommand
        Dim SQLstr As String

        Try
            Button1.Enabled = False
            cmdNew.Enabled = False
            Mycn = New OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & My.Settings.roledbloc & "';Jet OLEDB:Database Password='" & My.Settings.dbpass & "';")

            Dim ds1 As New DataSet
            Dim eds1 As New DataGridView
            Dim dv1 As New DataView

            Mycn.Open()

            Dim srt As Integer
            If IsDBNull(txtSortOrder) Or txtSortOrder.Text = "" Then
                srt = "0"
            Else
                srt = txtSortOrder.Text
            End If

            Dim question As String
            question = rtbQuestion.Text
            question = Replace(question, "'", "''")

            Dim rid As Integer
            Dim norpt As Integer
            If txtRFPID.Text = "NEW" Then
                rid = 0
                norpt = -1
            Else
                rid = txtRFPID.Text
                norpt = 0
            End If

            SQLstr = "INSERT INTO rfp_ReportQuestions([SortOrder], [Question], [ReportID], [AddedBy], [AddedDateStamp], [IsActive], [LevelID], [StageID], [StaleDate], [ContactID], [CompositeID], [NoReport])" & _
            "VALUES(" & srt & ",'" & question & "'," & rid & ",'" & Environ("USERNAME") & "',#" & Format(Now()) & "#, -1,11,1,#" & Format(Now()) & "#,0,0," & norpt & ")"

            Command = New OleDb.OleDbCommand(SQLstr, Mycn)
            Command.ExecuteNonQuery()

            Mycn.Close()

            Call LoadID()

            Mycn.Open()
            SQLstr = "INSERT INTO rfp_ResponseTracking([QuestionID], [Field], [NewValue], [ChangedBy], [DateStamp])" & _
            "VALUES(" & txtID.Text & ",'Question Added','" & txtID.Text & "','" & Environ("USERNAME") & "',#" & Format(Now()) & "#)"
            Command = New OleDb.OleDbCommand(SQLstr, Mycn)
            Command.ExecuteNonQuery()
            Mycn.Close()

            Mycn.Open()
            SQLstr = "INSERT INTO rfp_ResponseTracking([QuestionID], [Field], [NewValue], [ChangedBy], [DateStamp])" & _
            "VALUES(" & txtID.Text & ",'Question','" & question & "','" & Environ("USERNAME") & "',#" & Format(Now()) & "#)"
            Command = New OleDb.OleDbCommand(SQLstr, Mycn)
            Command.ExecuteNonQuery()
            Mycn.Close()

            Mycn.Open()
            SQLstr = "INSERT INTO rfp_ResponseTracking([QuestionID], [Field], [NewValue], [ChangedBy], [DateStamp])" & _
            "VALUES(" & txtID.Text & ",'Report ID','" & txtRFPID.Text & "','" & Environ("USERNAME") & "',#" & Format(Now()) & "#)"
            Command = New OleDb.OleDbCommand(SQLstr, Mycn)
            Command.ExecuteNonQuery()
            Mycn.Close()

            'txtSortOrder.Text = srt + 1

            'rtbQuestion.Text = ""
            Button1.Enabled = True
            cmdNew.Enabled = True
            Me.Close()
            'txtID.Text = "NEW"

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "ERROR")

        End Try
    End Sub

    Private Sub rtbQuestion_KeyDown(ByVal sender As Object, ByVal e As KeyEventArgs) Handles rtbQuestion.KeyDown
        'MsgBox("1")
        If Keys.Control And Keys.V Then
            MsgBox("2")
            Using box As New RichTextBox
                box.SelectAll()
                box.SelectedRtf = Clipboard.GetText(TextDataFormat.Rtf)
                box.SelectAll()
                box.SelectionBackColor = Color.White
                box.SelectionColor = Color.Black
                rtbQuestion.SelectedRtf = box.SelectedRtf
            End Using

            'Dim iData As IDataObject = Clipboard.GetDataObject()

            ' Determines whether the data is in a format you can use.
            'If iData.GetDataPresent(DataFormats.Text) Then
            ' Yes it is, so display it in a text box.
            'rtbQuestion.SelectedText = CType(iData.GetData(DataFormats.Text), String)
            'Else
            ' No it is not.
            'rtbQuestion.SelectedText = "Could not retrieve data off the clipboard."
            'End If

            e.Handled = True
        End If
    End Sub

    Private Sub rfp_ReportQuestion_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        If Button1.Enabled = False Or cmdNew.Enabled = False Then
            MsgBox("Please wait for the form to save before closing.", MsgBoxStyle.Information, "Saving.")
            e.Cancel = True
        Else
            e.Cancel = False
        End If
    End Sub

    Private Sub rfp_ReportQuestion_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Dim rtbContextMenu As New ContextMenu()
        Dim mymenuitem As MenuItem
        mymenuitem = rtbContextMenu.MenuItems.Add("Copy")
        AddHandler mymenuitem.Click, AddressOf myCopy_Click


        mymenuitem = rtbContextMenu.MenuItems.Add("Paste")
        AddHandler mymenuitem.Click, AddressOf myPaste_Click

        rtbQuestion.ContextMenu = rtbContextMenu

        Call LoadSortOrders()
    End Sub

    Public Sub LoadSortOrders()
        If txtRFPID.Text = "NEW" Then
        Else
            Dim Mycn As OleDb.OleDbConnection

            'Try

            Mycn = New OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & My.Settings.roledbloc & "';Jet OLEDB:Database Password='" & My.Settings.dbpass & "';")

            Mycn.Open()

            Dim sqlstring As String

            Dim title As String
            title = rtbQuestion.Text
            title = Replace(title, "'", "''")

            sqlstring = "SELECT Top 1 SortOrder FROM rfp_ReportQuestions WHERE IsActive = -1 AND ReportID = " & txtRFPID.Text & "" & _
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
                txtSortOrder.Text = sord
            Else
                txtSortOrder.Text = "1"
            End If
            Mycn.Close()
            Mycn.Dispose()

            'Catch ex As Exception
            'MsgBox(ex.Message, MsgBoxStyle.Critical, "ERROR")
            'End Try
        End If
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        Call SpellCheckRTB()
    End Sub

    Public Sub SpellCheckRTB()
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

    Private Sub rtbQuestion_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles rtbQuestion.LostFocus

    End Sub

    Private Sub myCopy_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Clipboard.SetDataObject(rtbQuestion.SelectedText)
    End Sub
    Private Sub myPaste_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        ' Declares an IDataObject to hold the data returned from the clipboard.
        ' Retrieves the data from the clipboard.
        Dim iData As IDataObject = Clipboard.GetDataObject()

        ' Determines whether the data is in a format you can use.
        If iData.GetDataPresent(DataFormats.Text) Then
            ' Yes it is, so display it in a text box.
            rtbQuestion.SelectedText = CType(iData.GetData(DataFormats.Text), String)
        Else
            ' No it is not.
            rtbQuestion.SelectedText = "Could not retrieve data off the clipboard."
        End If


    End Sub

    Private Sub rtbQuestion_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rtbQuestion.TextChanged

    End Sub
End Class