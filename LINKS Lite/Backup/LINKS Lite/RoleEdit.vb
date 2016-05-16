Public Class RoleEdit
    Dim thread1 As System.Threading.Thread

    Public Sub CurrentRole()

        Dim Mycn As OleDb.OleDbConnection

        Try

            Mycn = New OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & My.Settings.roledbloc & "';Jet OLEDB:Database Password='" & My.Settings.dbpass & "';")

            Mycn.Open()

            Dim sqlstring As String

            sqlstring = "SELECT * FROM sys_Roles WHERE ID = " & ID.Text & ";"

            Dim queryString As String = String.Format(sqlstring)
            Dim cmd As New OleDb.OleDbCommand(queryString, Mycn)
            Dim da As New OleDb.OleDbDataAdapter(cmd)
            Dim ds As New DataSet

            da.Fill(ds, "User")
            Dim dt As DataTable = ds.Tables("User")
            If dt.Rows.Count > 0 Then

                Dim row As DataRow = dt.Rows(0)

                If IsDBNull(row("RoleName")) Then
                    RoleName.Text = "*UNKNOWN*"
                Else
                    RoleName.Text = (row("RoleName"))
                End If

                Active.Checked = (row("Active"))
                Launch.Checked = (row("Launch"))
                Login.Checked = (row("Login"))
                ViewAccount.Checked = (row("ViewAccount"))
                ViewPositions.Checked = (row("AccountPositions"))
                ViewTeam.Checked = (row("ViewTeam"))
                ViewFirms.Checked = (row("ViewFirms"))
                ViewReps.Checked = (row("ViewReps"))
                ViewDist.Checked = (row("ViewDist"))
                SpoofUser.Checked = (row("SpoofUser"))
                ViewUsers.Checked = (row("ViewUser"))
                EditUser.Checked = (row("EditUser"))
                AllAccountAccess.Checked = (row("AllAccountAccess"))
                SetDefaultFee.Checked = (row("SetDefaultFee"))
                EditFirmFee.Checked = (row("EditFirmFee"))
                ViewAPX.Checked = (row("ViewAPX"))
                ViewAdmin.Checked = (row("AdminView"))
                AddUser.Checked = (row("AddUser"))
                ViewOps.Checked = (row("ViewOps"))
                ViewCompositeVer.Checked = (row("ViewCompositeVer"))
                EditCompositeVer.Checked = (row("EditCompositeVer"))
                EditCompositeReasons.Checked = (row("EditCompositeReasons"))
                FinishRecon.Checked = (row("FinishRecon"))
                RFPView.Checked = (row("RFPView"))
                RFPSubmit.Checked = (row("RFPSubmit"))
                RFPWork.Checked = (row("RFPWork"))
                RFPQuestionView.Checked = (row("RFPQuestionView"))
                RFPQuestionEdit.Checked = (row("RFPQuestionEdit"))
                RFPQuestionDelete.Checked = (row("RFPQuestionDelete"))
                RFPAddFirm.Checked = (row("RFPAddFirm"))
                UMAImport.Checked = (row("UMAImport"))
                UMALaunch.Checked = (row("UMALaunch"))
                UMASkipCheck.Checked = (row("UMASkipCheck"))
                UMATranslate.Checked = (row("UMATranslate"))
                UMAPortfolioFunctions.Checked = (row("UMAPortfolioFunctions"))
                UMATradeFunctions.Checked = (row("UMATradeFunctions"))
                UMASystemFunctions.Checked = (row("UMASystemFunctions"))
                UMAAutoRecon.Checked = (row("UMAAutoRecon"))
                MasterMappingCenterView.Checked = (row("MasterMappingCenterView"))
                MAPAddAgreement.Checked = (row("MAPAddAgreement"))
                MAPAddFirm.Checked = (row("MAPAddFirm"))
                MAPAddProducts.Checked = (row("MAPAddProducts"))
                MAPAddProductType.Checked = (row("MAPAddProductType"))
                MAPAddObjective.Checked = (row("MAPAddObjective"))
                MAPEditProducts.Checked = (row("MAPEditProducts"))
                MAPEditProductTypes.Checked = (row("MAPEditProductTypes"))
                MAPEditFirms.Checked = (row("MAPEditFirms"))
                MAPEditAgreements.Checked = (row("MAPEditAgreements"))
                MAPEditObjective.Checked = (row("MAPEditObjective"))
                MAPAddPlatform.Checked = (row("MAPAddPlatform"))
                MAPEditPlatform.Checked = (row("MAPEditPlatform"))
                MAPAddWLPlatform.Checked = (row("MAPAddWLPlatform"))
                MAPEditWLPlatform.Checked = (row("MAPEditWLPlatform"))
                MAPAssociatePlatform.Checked = (row("MAPAssociatePlatform"))
                MAPAddDatabase.Checked = (row("MAPAddDatabase"))
                MAPRefreshApprovals.Checked = (row("MAPRefreshApprovals"))
                MAPGenerateReports.Checked = (row("MAPGenerateReports"))
                ViewReports.Checked = (row("ViewReports"))
                ETFPriceImport.Checked = (row("ETFPriceImport"))
                LaunchRevenueCenter.Checked = (row("LaunchRevenueCenter"))
                ViewRepBusinessReport.Checked = (row("ViewRepBusinessReport"))
                TeamManagerAccess.Checked = (row("TeamManagerAccess"))
                SalesManagementAccess.Checked = (row("SalesManagementAccess"))
                LeadManagerAccess.Checked = (row("LeadManagerAccess"))
                AMPAccess.Checked = (row("AMPAccess"))
                AMPSecCodes.Checked = (row("AMPSecCodes"))
                ViewAccountMoves.Checked = (row("ViewAccountMoves"))
                EditAccountMovesSettings.Checked = (row("EditAccountMovesSettings"))
                EditWhatsNew.Checked = (row("EditWhatsNew"))

            Else

            End If
            Mycn.Close()
            Mycn.Dispose()

        Catch ex As Exception
            MsgBox(ex.Message & vbNewLine & vbNewLine & "This program will now exit.", MsgBoxStyle.Critical, "ERROR")
            'Application.Exit()
        End Try
    End Sub


    Private Sub cmdCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdCancel.Click
        Me.Close()
    End Sub

    Private Sub RoleEdit_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

    End Sub

    Private Sub cmdSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdSave.Click
        If ID.Text = "NEW" Then
            'Value Check
            If IsDBNull(RoleName) Or RoleName.Text = "" Then
                RoleName.BackColor = Color.Red
                RoleName.ForeColor = Color.White
                GoTo Line1
            Else
                RoleName.BackColor = Color.White
                RoleName.ForeColor = Color.Black
            End If

            'Insert new record
            cmdSave.Text = "SAVING..."
            cmdSave.Enabled = False
            Dim Mycn As OleDb.OleDbConnection

            Try

                Mycn = New OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & My.Settings.roledbloc & "';Jet OLEDB:Database Password='" & My.Settings.dbpass & "';")

                Mycn.Open()

                Dim sqlstring As String

                sqlstring = "SELECT * FROM sys_Roles WHERE [RoleName] = '" & RoleName.Text & "'"


                Dim queryString As String = String.Format(sqlstring)
                Dim cmd As New OleDb.OleDbCommand(queryString, Mycn)
                Dim da As New OleDb.OleDbDataAdapter(cmd)
                Dim ds As New DataSet

                da.Fill(ds, "User")
                Dim dt As DataTable = ds.Tables("User")
                If dt.Rows.Count = 0 Then
                    Control.CheckForIllegalCrossThreadCalls = False
                    thread1 = New System.Threading.Thread(AddressOf InsertRec)
                    thread1.Start()
                Else
                    RoleName.BackColor = Color.Red
                    RoleName.ForeColor = Color.White
                    MsgBox("A profile already exsists with that Role Name.  You cannot save a duplicate record.", MsgBoxStyle.Information, "DUPLICATE RECORD DETECTED")
                    cmdSave.Enabled = True
                    cmdSave.Text = "Save"
                    GoTo Line1
                End If

            Catch ex As Exception
                MsgBox(ex.Message)
            End Try
        Else
            'Update record
            cmdSave.Text = "SAVING..."
            cmdSave.Enabled = False
            Control.CheckForIllegalCrossThreadCalls = False
            thread1 = New System.Threading.Thread(AddressOf UpdateRec)
            thread1.Start()
        End If

Line1:
    End Sub

    Public Sub InsertRec()
        Dim Mycn As OleDb.OleDbConnection
        Dim Command As OleDb.OleDbCommand
        Dim SQLstr As String

        Try
            Mycn = New OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & My.Settings.roledbloc & "';Jet OLEDB:Database Password='" & My.Settings.dbpass & "';")

            Dim ds1 As New DataSet
            Dim eds1 As New DataGridView
            Dim dv1 As New DataView

            Mycn.Open()

            SQLstr = "INSERT INTO sys_Roles([RoleName], [Active], [Launch], [Login], [ViewAccount], [AccountPositions], [ViewTeam], [ViewFirms], [ViewReps], [ViewDist], [SpoofUser], [ViewUser], [EditUser], [AllAccountAccess], [EditFirmFee], [ViewAPX], [SetDefaultFee], [AdminView], [AddUser], [ViewOps], [ViewCompositeVer], [EditCompositeVer], [EditCompositeReasons], [FinishRecon], [RFPView], [RFPSubmit], [RFPWork], [RFPQuestionView], [RFPQuestionEdit], [RFPQuestionDelete], [RFPAddFirm], [UMALaunch], [UMAImport], [UMATranslate], [UMASkipCheck], [UMAPortfolioFunctions], [UMATradeFunctions], [UMASystemFunctions], [UMAAutoRecon], [MasterMappingCenterView],MAPAddAgreement,MAPAddFirm,MAPAddProducts,MAPAddProductType,MAPAddObjective,MAPEditProducts,MAPEditProductTypes,MAPEditFirms,MAPEditAgreements,MAPEditObjective,MAPAddPlatform,MAPEditPlatform,MAPAddWLPlatform,MAPEditWLPlatform,MAPAssociatePlatform,MAPAddDatabase,MAPRefreshApprovals,MAPGenerateReports,ViewReports,ETFPriceImport, LaunchRevenueCenter, ViewRepBusinessReport, TeamManagerAccess, SalesManagementAccess, LeadManagerAccess, AMPAccess, AMPSecCodes, ViewAccountMoves, EditAccountMovesSettings, EditWhatsNew)" & _
            "VALUES('" & RoleName.Text & "'," & Active.Checked & "," & Launch.Checked & "," & Login.Checked & "," & ViewAccount.Checked & "," & ViewPositions.Checked & "," & ViewTeam.Checked & "," & ViewFirms.Checked & "," & ViewReps.Checked & "," & ViewDist.Checked & "," & SpoofUser.Checked & "," & ViewUsers.Checked & "," & EditUser.Checked & "," & AllAccountAccess.Checked & "," & EditFirmFee.Checked & "," & ViewAPX.Checked & "," & SetDefaultFee.Checked & "," & ViewAdmin.Checked & "," & AddUser.Checked & "," & ViewOps.Checked & "," & ViewCompositeVer.Checked & "," & EditCompositeVer.Checked & "," & EditCompositeReasons.Checked & "," & FinishRecon.Checked & "," & RFPView.Checked & "," & RFPSubmit.Checked & "," & RFPWork.Checked & "," & RFPQuestionView.Checked & "," & RFPQuestionEdit.Checked & "," & RFPQuestionDelete.Checked & "," & RFPAddFirm.Checked & "," & UMALaunch.Checked & "," & UMAImport.Checked & "," & UMATranslate.Checked & "," & UMASkipCheck.Checked & "," & UMAPortfolioFunctions.Checked & "," & UMATradeFunctions.Checked & "," & UMASystemFunctions.Checked & "," & UMAAutoRecon.Checked & "," & MasterMappingCenterView.Checked & "," & MAPAddAgreement.Checked & "," & MAPAddFirm.Checked & "," & MAPAddProducts.Checked & "," & MAPAddProductType.Checked & "," & MAPAddObjective.Checked & "," & MAPEditProducts.Checked & "," & MAPEditProductTypes.Checked & "," & MAPEditFirms.Checked & "," & MAPEditAgreements.Checked & "," & MAPEditObjective.Checked & "," & MAPAddPlatform.Checked & "," & MAPEditPlatform.Checked & "," & MAPAddWLPlatform.Checked & "," & MAPEditWLPlatform.Checked & "," & MAPAssociatePlatform.Checked & "," & MAPAddDatabase.Checked & "," & MAPRefreshApprovals.Checked & "," & MAPGenerateReports.Checked & "," & ViewReports.Checked & "," & ETFPriceImport.Checked & "," & LaunchRevenueCenter.Checked & "," & ViewRepBusinessReport.Checked & "," & TeamManagerAccess.Checked & "," & SalesManagementAccess.Checked & "," & LeadManagerAccess.Checked & "," & AMPAccess.Checked & "," & AMPSecCodes.Checked & "," & ViewAccountMoves.Checked & "," & EditAccountMovesSettings.Checked & "," & EditWhatsNew.Checked & ")"

            Command = New OleDb.OleDbCommand(SQLstr, Mycn)
            Command.ExecuteNonQuery()

            Mycn.Close()

            Me.Close()

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "ERROR")

        End Try
    End Sub

    Public Sub UpdateRec()
        Dim Mycn As OleDb.OleDbConnection
        Dim Command As OleDb.OleDbCommand
        Dim SQLstr As String

        Try
            Mycn = New OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & My.Settings.roledbloc & "';Jet OLEDB:Database Password='" & My.Settings.dbpass & "';")

            Dim ds1 As New DataSet
            Dim eds1 As New DataGridView
            Dim dv1 As New DataView

            Mycn.Open()

            SQLstr = "Update sys_Roles" & _
            " SET [RoleName] = '" & RoleName.Text & "', [Active] = " & Active.Checked & ", [Launch] = " & Launch.Checked & ", [Login] = " & Login.Checked & ", [ViewAccount] = " & ViewAccount.Checked & ", [AccountPositions] = " & ViewPositions.Checked & ", [ViewTeam] = " & ViewTeam.Checked & ", [ViewFirms] = " & ViewFirms.Checked & ", [ViewReps] = " & ViewReps.Checked & ", [ViewDist] = " & ViewDist.Checked & ", [SpoofUser] = " & SpoofUser.Checked & ", [ViewUser] = " & ViewUsers.Checked & ", [EditUser] = " & EditUser.Checked & ", [AllAccountAccess] = " & AllAccountAccess.Checked & ", [EditFirmFee] = " & EditFirmFee.Checked & ", [ViewAPX] =" & ViewAPX.Checked & ", [SetDefaultFee] = " & SetDefaultFee.Checked & ", [AdminView] = " & ViewAdmin.Checked & ", [AddUser] = " & AddUser.Checked & ", [ViewOps] = " & ViewOps.Checked & ", [ViewCompositeVer] = " & ViewCompositeVer.Checked & ", [EditCompositeVer] = " & EditCompositeVer.Checked & ", [EditCompositeReasons] = " & EditCompositeReasons.Checked & ", [FinishRecon] = " & FinishRecon.Checked & ", [RFPView] = " & RFPView.Checked & ", [RFPSubmit] = " & RFPSubmit.Checked & ", [RFPWork] = " & RFPWork.Checked & ", [RFPQuestionView] = " & RFPQuestionView.Checked & ", [RFPQuestionEdit] = " & RFPQuestionEdit.Checked & ", [RFPQuestionDelete] = " & RFPQuestionDelete.Checked & ", [RFPAddFirm] = " & RFPAddFirm.Checked & ", [UMALaunch] = " & UMALaunch.Checked & ", [UMAImport] = " & UMAImport.Checked & ", [UMATranslate] = " & UMATranslate.Checked & ", [UMASkipCheck] = " & UMASkipCheck.Checked & ", [UMAPortfolioFunctions] = " & UMAPortfolioFunctions.Checked & ", [UMATradeFunctions] = " & UMATradeFunctions.Checked & ", [UMASystemFunctions] = " & UMASystemFunctions.Checked & ", [UMAAutoRecon] = " & UMAAutoRecon.Checked & ", [MasterMappingCenterView] = " & MasterMappingCenterView.Checked & ", MAPAddAgreement = " & MAPAddAgreement.Checked & ", MAPAddFirm = " & MAPAddFirm.Checked & ", MAPAddProducts = " & MAPAddProducts.Checked & ", MAPAddProductType = " & MAPAddProductType.Checked & ", MAPAddObjective = " & MAPAddObjective.Checked & ", MAPEditProducts = " & MAPEditProducts.Checked & ", MAPEditProductTypes = " & MAPEditProductTypes.Checked & ", MAPEditFirms = " & MAPEditFirms.Checked & ", MAPEditAgreements = " & MAPEditAgreements.Checked & ", MAPEditObjective = " & MAPEditObjective.Checked & ", MAPAddPlatform = " & MAPAddPlatform.Checked & ", MAPEditPlatform = " & MAPEditPlatform.Checked & ", MAPAddWLPlatform = " & MAPAddWLPlatform.Checked & ", MAPEditWLPlatform = " & MAPEditWLPlatform.Checked & ", MAPAssociatePlatform = " & MAPAssociatePlatform.Checked & ", MAPAddDatabase = " & MAPAddDatabase.Checked & ", MAPRefreshApprovals = " & MAPRefreshApprovals.Checked & ", MAPGenerateReports = " & MAPGenerateReports.Checked & ", ViewReports = " & ViewReports.Checked & ", ETFPriceImport = " & ETFPriceImport.Checked & ", LaunchRevenueCenter = " & LaunchRevenueCenter.Checked & ", ViewRepBusinessReport = " & ViewRepBusinessReport.Checked & ", TeamManagerAccess = " & TeamManagerAccess.Checked & ", SalesManagementAccess = " & SalesManagementAccess.Checked & ", LeadManagerAccess = " & LeadManagerAccess.Checked & ", AMPAccess = " & AMPAccess.Checked & ", AMPSecCodes = " & AMPSecCodes.Checked & ", ViewAccountMoves = " & ViewAccountMoves.Checked & ", EditAccountMovesSettings = " & EditAccountMovesSettings.Checked & ",EditWhatsNew = " & EditWhatsNew.Checked & _
            " WHERE ID = " & ID.Text

            Command = New OleDb.OleDbCommand(SQLstr, Mycn)
            Command.ExecuteNonQuery()

            Mycn.Close()

            Me.Close()

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "ERROR")

        End Try
    End Sub

    Private Sub ViewUsers_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ViewUsers.CheckedChanged
        If ViewUsers.Checked Then
            ViewAdmin.Checked = True
        Else
            'do nothing
        End If
    End Sub

    Private Sub EditUser_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles EditUser.CheckedChanged
        If EditUser.Checked Then
            ViewAdmin.Checked = True
        Else
            'do nothing
        End If
    End Sub

    Private Sub AddUser_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles AddUser.CheckedChanged
        If AddUser.Checked Then
            ViewAdmin.Checked = True
        Else
            'do nothing
        End If
    End Sub

    Private Sub MAPAddPlatform_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MAPAddPlatform.CheckedChanged

    End Sub

    Private Sub TabPage5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TabPage5.Click

    End Sub
End Class