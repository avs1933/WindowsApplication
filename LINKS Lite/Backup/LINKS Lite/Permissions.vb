Public Class Permissions
    Dim thread1 As System.Threading.Thread

    Private Sub Permissions_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
       
    End Sub

    Public Sub loadrole()

    End Sub

    Public Sub Populate()
        Dim Mycn As OleDb.OleDbConnection

        Try

            Mycn = New OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & My.Settings.roledbloc & "';Jet OLEDB:Database Password='" & My.Settings.dbpass & "';")

            Mycn.Open()

            Dim sqlstring As String

            sqlstring = "SELECT * FROM sys_Roles WHERE ID = (SELECT RoleID from sys_Users WHERE ID = " & My.Settings.userid & ");"

            Dim queryString As String = String.Format(sqlstring)
            Dim cmd As New OleDb.OleDbCommand(queryString, Mycn)
            Dim da As New OleDb.OleDbDataAdapter(cmd)
            Dim ds As New DataSet

            da.Fill(ds, "User")
            Dim dt As DataTable = ds.Tables("User")
            If dt.Rows.Count > 0 Then

                Dim row As DataRow = dt.Rows(0)
                'This block loads the users rights to access the program
                ActiveCheckBox.Checked = (row("Active"))
                LoginCheckBox.Checked = (row("Login"))
                LaunchCheckBox.Checked = (row("Launch"))

                RoleNameTextBox.Text = (row("RoleName")).ToString

                'This block loads data access rights.
                ViewAccountCheckBox.Checked = (row("ViewAccount"))
                ViewTeamCheckBox.Checked = (row("ViewTeam"))
                ViewFirmsCheckBox.Checked = (row("ViewFirms"))
                ViewRepsCheckBox.Checked = (row("ViewReps"))
                ViewDistCheckBox.Checked = (row("ViewDist"))
                SpoofUser.Checked = (row("SpoofUser"))
                EditFirmFee.Checked = (row("EditFirmFee"))
                ViewAPX.Checked = (row("ViewAPX"))
                SetDefaultFee.Checked = (row("SetDefaultFee"))
                ViewUser.Checked = (row("ViewUser"))
                EditUser.Checked = (row("EditUser"))
                AddUser.Checked = (row("AddUser"))
                AdminView.Checked = (row("AdminView"))
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

                ProgressBar1.Value = 50
                Call permissioncheck()
            Else
                MsgBox("Invalid Login", MsgBoxStyle.Critical, "Permissions Issue")
            End If
            Mycn.Close()
            Mycn.Dispose()

        Catch ex As Exception
            MsgBox(ex.Message & vbNewLine & vbNewLine & "This program will now exit.", MsgBoxStyle.Critical, "ERROR")
            Application.Exit()
        End Try
    End Sub

    Public Sub permissioncheck()
        'This section looks for active user
        If ActiveCheckBox.Checked Then
            Login.Close()
        Else
            MsgBox("Your user account appears to be inactive.  Please contact the Asset Management Division for help.", MsgBoxStyle.Critical, "Permission Issue")
            Login.Show()
            Home.Close()
            Me.Close()
        End If

        'This section check if User can Launch Program
        If LaunchCheckBox.Checked Then
            Login.Close()
        Else
            MsgBox("You do not have permission to launch this program.  Please contact the Asset Management Division for help.", MsgBoxStyle.Critical, "Permission Issue")
            Application.Exit()
        End If

        'This section checks if User can Login.  If true enable home.
        If LoginCheckBox.Checked Then
            Home.Enabled = True
            Login.Close()
        Else
            MsgBox("You do not have permissions to Login.  Please contact the Asset Management Division for help.", MsgBoxStyle.Critical, "Permission Issue")
            Login.Show()
            Home.Close()
            Me.Close()
        End If

        'Set Home Status Bar
        Try
            Dim conn As New OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & My.Settings.roledbloc & "';Jet OLEDB:Database Password='" & My.Settings.dbpass & "';")

            Dim strSQL As String = "SELECT * FROM sys_Users WHERE ID = " & My.Settings.userid & ";"
            Dim da As New OleDb.OleDbDataAdapter(strSQL, conn)
            Dim ds As New DataSet
            'da.Fill(ds, "Users")
            da.Fill(ds, "User")
            Dim dt As DataTable = ds.Tables("User")
            'ds.Tables("Users")
            Dim row As DataRow = dt.Rows(0)
            Home.tssUserName.Text = (row("FullName")) & " has successfully logged in."
            My.Settings.teamnum = (row("Team"))
        Catch ex As Exception

        End Try
        Home.tssDomain.Text = "Current Domain: " & Environment.UserDomainName
        Home.tssComputer.Text = "Current Workstation: " & My.Computer.Name
        If My.Settings.dbloc <> "\\monumentco1\data\ToolboxDBs\SalesInterface2.accdb" Then
            Home.Text = "Welcome to the " & Application.ProductName & " system. Version Number: " & Application.ProductVersion & ". ***NOT CONNECTED TO PRODUCTION DATABASE***"
        Else
            Home.Text = "Welcome to the " & Application.ProductName & " system. Version Number: " & Application.ProductVersion & "."
        End If


        ProgressBar1.Value = 65

        If ViewAccountCheckBox.Checked Then
            Home.ToolStripDropDownButton2.Visible = True
        Else
            Home.ToolStripDropDownButton2.Visible = False
        End If

        If ViewFirmsCheckBox.Checked Then
            Home.ToolStripDropDownButton3.Visible = True
        Else
            Home.ToolStripDropDownButton3.Visible = False
        End If

        If SpoofUser.Checked Or My.Settings.IsSpoof = "TRUE" Then
            Home.tsSpoof.Visible = True
        Else
            Home.tsSpoof.Visible = False
        End If

        If AdminView.Checked Then
            Home.ToolStripDropDownButton5.Visible = True
        Else
            Home.ToolStripDropDownButton5.Visible = False
        End If

        If ((ViewUser.Checked) Or (EditUser.Checked)) Then
            Home.UserControlToolStripMenuItem.Visible = True
        Else
            Home.UserControlToolStripMenuItem.Visible = False
        End If

        If AddUser.Checked Then
            Home.AddUserToolStripMenuItem.Visible = True
        Else
            Home.AddUserToolStripMenuItem.Visible = False
        End If

        If ViewOps.Checked Then
            Home.ToolStripDropDownButton7.Visible = True
        Else
            Home.ToolStripDropDownButton7.Visible = False
        End If

        If ViewCompositeVer.Checked Then
            Home.CompositeVerificationToolStripMenuItem.Visible = True
        Else
            Home.CompositeVerificationToolStripMenuItem.Visible = False
        End If

        If FinishRecon.Checked Then
            Home.FinishReconToolStripMenuItem.Visible = True
        Else
            Home.FinishReconToolStripMenuItem.Visible = False
        End If

        If RFPView.Checked Then
            Home.RFPCenterToolStripMenuItem.Visible = True
        Else
            Home.RFPCenterToolStripMenuItem.Visible = False
        End If

        If UMALaunch.Checked Then
            Home.UMAToolStripMenuItem.Visible = True
        Else
            Home.UMAToolStripMenuItem.Visible = False
        End If

        If MasterMappingCenterView.Checked Then
            Home.MasterMappingCenterToolStripMenuItem.Visible = True
        Else
            Home.MasterMappingCenterToolStripMenuItem.Visible = False
        End If

        If ViewReports.Checked Then
            Home.ToolStripDropDownButton4.Visible = True
        Else
            Home.ToolStripDropDownButton4.Visible = False
        End If

        If MAPGenerateReports.Checked Then
            Home.PositionByAccountToolStripMenuItem.Visible = True
            Home.TransactionReportToolStripMenuItem.Visible = True
        Else
            Home.PositionByAccountToolStripMenuItem.Visible = False
            Home.TransactionReportToolStripMenuItem.Visible = False
        End If

        If ETFPriceImport.Checked Then
            Home.ToolStripMenuItem6.Visible = True
        Else
            Home.ToolStripMenuItem6.Visible = False
        End If

        If LaunchRevenueCenter.Checked Then
            Home.RevenueCenterToolStripMenuItem.Visible = True
        Else
            Home.RevenueCenterToolStripMenuItem.Visible = False
        End If

        If ViewRepBusinessReport.Checked Then
            Home.RepBusinessBreakdownToolStripMenuItem.Visible = True
        Else
            Home.RepBusinessBreakdownToolStripMenuItem.Visible = False
        End If

        If TeamManagerAccess.Checked Then
            Home.TeamManagerToolStripMenuItem.Visible = True
        Else
            Home.TeamManagerToolStripMenuItem.Visible = False
        End If

        If SalesManagementAccess.Checked Then
            Home.ToolStripDropDownButton9.Visible = True
        Else
            Home.ToolStripDropDownButton9.Visible = False
        End If

        If LeadManagerAccess.Checked Then
            Home.LeadManagerToolStripMenuItem.Visible = True
        Else
            Home.LeadManagerToolStripMenuItem.Visible = False
        End If

        If AMPAccess.Checked Then
            Home.AMToolStripMenuItem.Visible = True
        Else
            Home.AMToolStripMenuItem.Visible = False
        End If

        If ViewAccountMoves.Checked Then
            Home.AccountMovesToolStripMenuItem.Visible = True
        Else
            Home.AccountMovesToolStripMenuItem.Visible = False
        End If

        If EditWhatsNew.Checked Then
            Home.EditWhatsNewToolStripMenuItem.Visible = True
        Else
            Home.EditWhatsNewToolStripMenuItem.Visible = False
        End If

        'Finally hide the permissions screen.  Must remain loaded to retain permission data in real time.
        Me.Visible = False
    End Sub

    Private Sub setpermissions()

    End Sub
End Class