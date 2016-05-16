Public Class map_LoadFirmApprovals
    Dim findfirms As System.Threading.Thread
    Dim mymax As Integer = 37
    Dim pb As ProgressBar

    Private Sub map_LoadFirmApprovals_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Control.CheckForIllegalCrossThreadCalls = False
        findfirms = New System.Threading.Thread(AddressOf StartLoad)
        findfirms.Start()
    End Sub

    Public Sub StartLoad()
        pb = ProgressBar1
        pb.Maximum = mymax
        pb.Value = 0
        Label2.Text = "Priming Tables..."
        Call CleanTempTables()
        Label2.Text = "Tables Primed..."
        Call LoadAllInhouseApprovedAgreementsDefaultFee()
        Call LoadAllInhouseApprovedAgreementsDefaultFeeSOLICITED()
        Call LoadAllInhouseApprovedAgreementsDefaultFeeADVISED()
        Call LoadPlatformApprovals()
        Call WLPlatformCustomFeeOnlyUMA()

        Label2.Text = "FINISHED!"
        CheckBox1.Checked = True

    End Sub

    Public Sub CleanTempTables()

        Dim Mycn As OleDb.OleDbConnection
        Dim Command As OleDb.OleDbCommand
        Dim SQLstr As String
        Try
            Mycn = New OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & My.Settings.dbloc & "';Jet OLEDB:Database Password='" & My.Settings.dbpass & "';")
            Mycn.Open()

            SQLstr = "DELETE * FROM map_TempAssignments"

            Command = New OleDb.OleDbCommand(SQLstr, Mycn)
            Command.ExecuteNonQuery()

            Mycn.Close()

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "ERROR")
        End Try

    End Sub

    Private Sub Timer1_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer1.Tick
        If Label1.Text = "Working" Then
            Label1.Text = "Working."
        Else
            If Label1.Text = "Working." Then
                Label1.Text = "Working.."
            Else
                If Label1.Text = "Working.." Then
                    Label1.Text = "Working..."
                Else
                    If Label1.Text = "Working..." Then
                        Label1.Text = "Working"
                    End If
                End If
            End If
        End If

        If CheckBox1.Checked Then
            Timer2.Enabled = True
        Else
            'do nothing
        End If
    End Sub

    Public Sub LoadPlatformApprovals()
        'QID 1001
        Label2.Text = "QID:1001 - Looking for Platform Approvals..."
        Dim Mycn As OleDb.OleDbConnection
        Dim Command As OleDb.OleDbCommand
        Dim SQLstr As String
        Try
            Mycn = New OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & My.Settings.dbloc & "';Jet OLEDB:Database Password='" & My.Settings.dbpass & "';")
            Mycn.Open()

            SQLstr = "INSERT INTO map_TempAssignments([FirmID], [ProductID], [TypeOfProductID], [TypeID], [WLID], [PlatformID], [ManagingFirmID], [BreakpointFrom], [BreakpointTo], [Fee], [MaxRepFee], [SMAOffered], [UMAOffered], [PlatformApproved], [QID])" & _
            " SELECT map_PlatformAssignments.FirmID, map_Products.ID, map_Products.TypeOfProductID, 4, 47, map_PlatformListings.PlatformID, map_Products.ManagingFirmID, map_PlatformFees.BreakpointFrom, map_PlatformFees.BreakpointTo, map_PlatformFees.Fee, map_PlatformFees.MaxRepFee, map_PlatformListings.SMAOffered, map_PlatformListings.UMAOffered, map_PlatformListings.PlatformApproved, 1001" & _
            " FROM map_PlatformAssignments INNER JOIN ((map_PlatformListings INNER JOIN map_Products ON map_PlatformListings.ProductID = map_Products.ID) INNER JOIN map_PlatformFees ON (map_PlatformFees.PlatformID = map_PlatformListings.PlatformID) AND (map_PlatformListings.ProductID = map_PlatformFees.ProductID)) ON map_PlatformAssignments.PlatformID = map_PlatformListings.PlatformID" & _
            " WHERE(((map_PlatformListings.Active) = True) AND map_PlatformListings.WLID Is Null AND map_PlatformFees.WLID Is Null AND map_PlatformAssignments.Active = True And ((map_PlatformFees.Active) = True))" & _
            " GROUP BY map_PlatformAssignments.FirmID, map_Products.ID, map_Products.TypeOfProductID, map_PlatformListings.PlatformID, map_Products.ManagingFirmID, map_PlatformFees.BreakpointFrom, map_PlatformFees.BreakpointTo, map_PlatformFees.Fee, map_PlatformFees.MaxRepFee, map_PlatformListings.SMAOffered, map_PlatformListings.UMAOffered, map_PlatformListings.PlatformApproved;"

            Command = New OleDb.OleDbCommand(SQLstr, Mycn)
            Command.ExecuteNonQuery()

            Label2.Text = "Loaded..."
            pb.Value = pb.Value + 1
            Mycn.Close()

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "ERROR")
        End Try
    End Sub

    Public Sub WLPlatformCustomFeeOnlyUMA()
        'QID 1002
        Label2.Text = "QID:1002 - Looking for WL Platforms Only UMA Custom Fee..."
        Dim Mycn As OleDb.OleDbConnection
        Dim Command As OleDb.OleDbCommand
        Dim SQLstr As String
        Try
            Mycn = New OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & My.Settings.dbloc & "';Jet OLEDB:Database Password='" & My.Settings.dbpass & "';")
            Mycn.Open()

            SQLstr = "INSERT INTO map_TempAssignments([FirmID], [ProductID], [TypeOfProductID], [TypeID], [PlatformID], [WLID], [ManagingFirmID], [BreakpointFrom], [BreakpointTo], [Fee], [MaxRepFee], [SMAOffered], [UMAOffered], [PlatformApproved], [QID])" & _
            " SELECT map_PlatformAssignments.FirmID, map_PlatformFees.ProductID, map_Products.TypeOfProductID, 5, 7, map_PlatformsWL.ID, map_Products.ManagingFirmID, map_PlatformFees.BreakpointFrom, map_PlatformFees.BreakpointTo, map_PlatformFees.Fee, map_PlatformFees.MaxRepFee, map_PlatformListings.SMAOffered, map_PlatformListings.UMAOffered, map_PlatformListings.PlatformApproved, 1002" & _
            " FROM ((map_PlatformAssignments INNER JOIN map_PlatformsWL ON map_PlatformAssignments.WLID = map_PlatformsWL.ID) INNER JOIN (map_PlatformFees INNER JOIN map_Products ON map_PlatformFees.ProductID = map_Products.ID) ON (map_PlatformsWL.PlatformDriverID = map_PlatformFees.PlatformID) AND (map_PlatformsWL.ID = map_PlatformFees.WLID)) INNER JOIN map_PlatformListings ON (map_PlatformListings.ProductID = map_PlatformFees.ProductID) AND (map_PlatformsWL.PlatformDriverID = map_PlatformListings.PlatformID)" & _
            " WHERE(((map_PlatformsWL.CustomFee) = True) And ((map_PlatformsWL.OfferUMA) = True) AND ((map_PlatformListings.UMAOffered) = True) And ((map_PlatformsWL.AdditionalApproval) = False) And ((map_PlatformsWL.OfferSMA) = False) And ((map_PlatformFees.Active) = True) And ((map_PlatformsWL.Active) = True) And ((map_PlatformAssignments.Active) = True) AND map_PlatformAssignments.UseWL = True AND  map_PlatformsWL.RequirePlatformApproval = FALSE) AND ((map_PlatformListings.Active)=True) AND ((map_PlatformListings.WLID) Is Null)" & _
            " GROUP BY map_PlatformAssignments.FirmID, map_PlatformFees.ProductID, map_Products.TypeOfProductID, map_PlatformsWL.ID, map_Products.ManagingFirmID, map_PlatformFees.BreakpointFrom, map_PlatformFees.BreakpointTo, map_PlatformFees.Fee, map_PlatformFees.MaxRepFee, map_PlatformListings.SMAOffered, map_PlatformListings.UMAOffered, map_PlatformListings.PlatformApproved;"

            Command = New OleDb.OleDbCommand(SQLstr, Mycn)
            Command.ExecuteNonQuery()
            pb.Value = pb.Value + 1
            Label2.Text = "Loaded..."
            Call WLPlatformCustomFeeOnlySMA()

            Mycn.Close()

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "ERROR")
        End Try
    End Sub

    Public Sub WLPlatformCustomFeeOnlySMA()
        'QID 1003
        Label2.Text = "QID:1003 - Looking for WL Platforms Only SMA Custom Fee..."
        Dim Mycn As OleDb.OleDbConnection
        Dim Command As OleDb.OleDbCommand
        Dim SQLstr As String
        Try
            Mycn = New OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & My.Settings.dbloc & "';Jet OLEDB:Database Password='" & My.Settings.dbpass & "';")
            Mycn.Open()

            SQLstr = "INSERT INTO map_TempAssignments([FirmID], [ProductID], [TypeOfProductID], [TypeID], [PlatformID], [WLID], [ManagingFirmID], [BreakpointFrom], [BreakpointTo], [Fee], [MaxRepFee], [SMAOffered], [UMAOffered], [PlatformApproved], [QID])" & _
            " SELECT map_PlatformAssignments.FirmID, map_PlatformFees.ProductID, map_Products.TypeOfProductID, 5, 7, map_PlatformsWL.ID, map_Products.ManagingFirmID, map_PlatformFees.BreakpointFrom, map_PlatformFees.BreakpointTo, map_PlatformFees.Fee, map_PlatformFees.MaxRepFee, map_PlatformListings.SMAOffered, map_PlatformListings.UMAOffered, map_PlatformListings.PlatformApproved, 1003" & _
            " FROM ((map_PlatformAssignments INNER JOIN map_PlatformsWL ON map_PlatformAssignments.WLID = map_PlatformsWL.ID) INNER JOIN (map_PlatformFees INNER JOIN map_Products ON map_PlatformFees.ProductID = map_Products.ID) ON (map_PlatformsWL.PlatformDriverID = map_PlatformFees.PlatformID) AND (map_PlatformsWL.ID = map_PlatformFees.WLID)) INNER JOIN map_PlatformListings ON (map_PlatformListings.ProductID = map_PlatformFees.ProductID) AND (map_PlatformsWL.PlatformDriverID = map_PlatformListings.PlatformID)" & _
            " WHERE(((map_PlatformsWL.CustomFee) = True) And ((map_PlatformsWL.OfferUMA) = False) AND ((map_PlatformListings.SMAOffered) = True) And ((map_PlatformsWL.AdditionalApproval) = False) And ((map_PlatformsWL.OfferSMA) = True) And ((map_PlatformFees.Active) = True) And ((map_PlatformsWL.Active) = True) And ((map_PlatformAssignments.Active) = True) AND map_PlatformAssignments.UseWL = True AND  map_PlatformsWL.RequirePlatformApproval = FALSE) AND ((map_PlatformListings.Active)=True) AND ((map_PlatformListings.WLID) Is Null)" & _
            " GROUP BY map_PlatformAssignments.FirmID, map_PlatformFees.ProductID, map_Products.TypeOfProductID, map_PlatformsWL.ID, map_Products.ManagingFirmID, map_PlatformFees.BreakpointFrom, map_PlatformFees.BreakpointTo, map_PlatformFees.Fee, map_PlatformFees.MaxRepFee, map_PlatformListings.SMAOffered, map_PlatformListings.UMAOffered, map_PlatformListings.PlatformApproved;"

            Command = New OleDb.OleDbCommand(SQLstr, Mycn)
            Command.ExecuteNonQuery()
            pb.Value = pb.Value + 1
            Label2.Text = "Loaded..."
            Call WLPlatformCustomFeeSMAUMA()

            Mycn.Close()

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "ERROR")
        End Try
    End Sub

    Public Sub WLPlatformCustomFeeSMAUMA()
        'QID 1004
        Label2.Text = "QID:1004 - Looking for WL Platforms Only SMA Custom Fee..."
        Dim Mycn As OleDb.OleDbConnection
        Dim Command As OleDb.OleDbCommand
        Dim SQLstr As String
        Try
            Mycn = New OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & My.Settings.dbloc & "';Jet OLEDB:Database Password='" & My.Settings.dbpass & "';")
            Mycn.Open()

            SQLstr = "INSERT INTO map_TempAssignments([FirmID], [ProductID], [TypeOfProductID], [TypeID], [WLID], [PlatformID], [ManagingFirmID], [BreakpointFrom], [BreakpointTo], [Fee], [MaxRepFee], [SMAOffered], [UMAOffered], [PlatformApproved], [QID])" & _
            " SELECT map_PlatformAssignments.FirmID, map_PlatformFees.ProductID, map_Products.TypeOfProductID, 5, 7, map_PlatformsWL.ID, map_Products.ManagingFirmID, map_PlatformFees.BreakpointFrom, map_PlatformFees.BreakpointTo, map_PlatformFees.Fee, map_PlatformFees.MaxRepFee, map_PlatformListings.SMAOffered, map_PlatformListings.UMAOffered, map_PlatformListings.PlatformApproved, 1004" & _
            " FROM ((map_PlatformAssignments INNER JOIN map_PlatformsWL ON map_PlatformAssignments.WLID = map_PlatformsWL.ID) INNER JOIN (map_PlatformFees INNER JOIN map_Products ON map_PlatformFees.ProductID = map_Products.ID) ON (map_PlatformsWL.PlatformDriverID = map_PlatformFees.PlatformID) AND (map_PlatformsWL.ID = map_PlatformFees.WLID)) INNER JOIN map_PlatformListings ON (map_PlatformListings.ProductID = map_PlatformFees.ProductID) AND (map_PlatformsWL.PlatformDriverID = map_PlatformListings.PlatformID)" & _
            " WHERE(((map_PlatformsWL.CustomFee) = True) And ((map_PlatformsWL.OfferUMA) = True) And ((map_PlatformsWL.AdditionalApproval) = False) And ((map_PlatformsWL.OfferSMA) = True) And ((map_PlatformFees.Active) = True) And ((map_PlatformsWL.Active) = True) And ((map_PlatformAssignments.Active) = True) AND map_PlatformAssignments.UseWL = True AND map_PlatformsWL.RequirePlatformApproval = FALSE) AND ((map_PlatformListings.Active)=True) AND ((map_PlatformListings.WLID) Is Null)" & _
            " GROUP BY map_PlatformAssignments.FirmID, map_PlatformFees.ProductID, map_Products.TypeOfProductID, map_PlatformsWL.ID, map_Products.ManagingFirmID, map_PlatformFees.BreakpointFrom, map_PlatformFees.BreakpointTo, map_PlatformFees.Fee, map_PlatformFees.MaxRepFee, map_PlatformListings.SMAOffered, map_PlatformListings.UMAOffered, map_PlatformListings.PlatformApproved;"

            Command = New OleDb.OleDbCommand(SQLstr, Mycn)
            Command.ExecuteNonQuery()
            pb.Value = pb.Value + 1
            Label2.Text = "Loaded..."
            Call WLPlatformCustomFeeOnlyUMAPLApproved()

            Mycn.Close()

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "ERROR")
        End Try
    End Sub

    Public Sub WLPlatformCustomFeeOnlyUMAPLApproved()
        'QID 1005
        Label2.Text = "QID:1005 - Looking for WL Platforms Approved Only UMA Custom Fee..."
        Dim Mycn As OleDb.OleDbConnection
        Dim Command As OleDb.OleDbCommand
        Dim SQLstr As String
        Try
            Mycn = New OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & My.Settings.dbloc & "';Jet OLEDB:Database Password='" & My.Settings.dbpass & "';")
            Mycn.Open()

            SQLstr = "INSERT INTO map_TempAssignments([FirmID], [ProductID], [TypeOfProductID], [TypeID], [PlatformID], [WLID], [ManagingFirmID], [BreakpointFrom], [BreakpointTo], [Fee], [MaxRepFee], [SMAOffered], [UMAOffered], [PlatformApproved], [QID])" & _
            " SELECT map_PlatformAssignments.FirmID, map_PlatformFees.ProductID, map_Products.TypeOfProductID, 5, 7, map_PlatformsWL.ID, map_Products.ManagingFirmID, map_PlatformFees.BreakpointFrom, map_PlatformFees.BreakpointTo, map_PlatformFees.Fee, map_PlatformFees.MaxRepFee, map_PlatformListings.SMAOffered, map_PlatformListings.UMAOffered, map_PlatformListings.PlatformApproved, 1005" & _
            " FROM ((map_PlatformAssignments INNER JOIN map_PlatformsWL ON map_PlatformAssignments.WLID = map_PlatformsWL.ID) INNER JOIN (map_PlatformFees INNER JOIN map_Products ON map_PlatformFees.ProductID = map_Products.ID) ON (map_PlatformsWL.PlatformDriverID = map_PlatformFees.PlatformID) AND (map_PlatformsWL.ID = map_PlatformFees.WLID)) INNER JOIN map_PlatformListings ON (map_PlatformListings.ProductID = map_PlatformFees.ProductID) AND (map_PlatformsWL.PlatformDriverID = map_PlatformListings.PlatformID)" & _
            " WHERE(((map_PlatformsWL.CustomFee) = True) And ((map_PlatformsWL.OfferUMA) = True) AND ((map_PlatformListings.UMAOffered) = True) And ((map_PlatformsWL.AdditionalApproval) = False) And ((map_PlatformsWL.OfferSMA) = False) And ((map_PlatformFees.Active) = True) And ((map_PlatformsWL.Active) = True) And ((map_PlatformAssignments.Active) = True) AND map_PlatformAssignments.UseWL = True AND map_PlatformsWL.RequirePlatformApproval = TRUE AND map_PlatformListings.PlatformApproved = True) AND ((map_PlatformListings.Active)=True) AND ((map_PlatformListings.WLID) Is Null)" & _
            " GROUP BY map_PlatformAssignments.FirmID, map_PlatformFees.ProductID, map_Products.TypeOfProductID, map_PlatformsWL.ID, map_Products.ManagingFirmID, map_PlatformFees.BreakpointFrom, map_PlatformFees.BreakpointTo, map_PlatformFees.Fee, map_PlatformFees.MaxRepFee, map_PlatformListings.SMAOffered, map_PlatformListings.UMAOffered, map_PlatformListings.PlatformApproved;"

            Command = New OleDb.OleDbCommand(SQLstr, Mycn)
            Command.ExecuteNonQuery()
            pb.Value = pb.Value + 1
            Label2.Text = "Loaded..."
            Call WLPlatformCustomFeeOnlySMAPLApproved()

            Mycn.Close()

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "ERROR")
        End Try
    End Sub

    Public Sub WLPlatformCustomFeeOnlySMAPLApproved()
        'QID 1006
        Label2.Text = "QID:1006 - Looking for WL Platforms Approved Only SMA Custom Fee..."
        Dim Mycn As OleDb.OleDbConnection
        Dim Command As OleDb.OleDbCommand
        Dim SQLstr As String
        Try
            Mycn = New OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & My.Settings.dbloc & "';Jet OLEDB:Database Password='" & My.Settings.dbpass & "';")
            Mycn.Open()

            SQLstr = "INSERT INTO map_TempAssignments([FirmID], [ProductID], [TypeOfProductID], [TypeID], [WLID], [PlatformID], [ManagingFirmID], [BreakpointFrom], [BreakpointTo], [Fee], [MaxRepFee], [SMAOffered], [UMAOffered], [PlatformApproved], [QID])" & _
            " SELECT map_PlatformAssignments.FirmID, map_PlatformFees.ProductID, map_Products.TypeOfProductID, 5, 7, map_PlatformsWL.ID, map_Products.ManagingFirmID, map_PlatformFees.BreakpointFrom, map_PlatformFees.BreakpointTo, map_PlatformFees.Fee, map_PlatformFees.MaxRepFee, map_PlatformListings.SMAOffered, map_PlatformListings.UMAOffered, map_PlatformListings.PlatformApproved, 1006" & _
            " FROM ((map_PlatformAssignments INNER JOIN map_PlatformsWL ON map_PlatformAssignments.WLID = map_PlatformsWL.ID) INNER JOIN (map_PlatformFees INNER JOIN map_Products ON map_PlatformFees.ProductID = map_Products.ID) ON (map_PlatformsWL.PlatformDriverID = map_PlatformFees.PlatformID) AND (map_PlatformsWL.ID = map_PlatformFees.WLID)) INNER JOIN map_PlatformListings ON (map_PlatformListings.ProductID = map_PlatformFees.ProductID) AND (map_PlatformsWL.PlatformDriverID = map_PlatformListings.PlatformID)" & _
            " WHERE(((map_PlatformsWL.CustomFee) = True) And ((map_PlatformsWL.OfferUMA) = False) AND ((map_PlatformListings.SMAOffered) = True) And ((map_PlatformsWL.AdditionalApproval) = False) And ((map_PlatformsWL.OfferSMA) = True) And ((map_PlatformFees.Active) = True) And ((map_PlatformsWL.Active) = True) And ((map_PlatformAssignments.Active) = True) AND map_PlatformAssignments.UseWL = True AND map_PlatformsWL.RequirePlatformApproval = TRUE AND map_PlatformListings.PlatformApproved = True) AND ((map_PlatformListings.Active)=True) AND ((map_PlatformListings.WLID) Is Null)" & _
            " GROUP BY map_PlatformAssignments.FirmID, map_PlatformFees.ProductID, map_Products.TypeOfProductID, map_PlatformsWL.ID, map_Products.ManagingFirmID, map_PlatformFees.BreakpointFrom, map_PlatformFees.BreakpointTo, map_PlatformFees.Fee, map_PlatformFees.MaxRepFee, map_PlatformListings.SMAOffered, map_PlatformListings.UMAOffered, map_PlatformListings.PlatformApproved;"

            Command = New OleDb.OleDbCommand(SQLstr, Mycn)
            Command.ExecuteNonQuery()
            pb.Value = pb.Value + 1
            Label2.Text = "Loaded..."
            Call WLPlatformCustomFeeSMAUMAPLApproved()

            Mycn.Close()

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "ERROR")
        End Try
    End Sub

    Public Sub WLPlatformCustomFeeSMAUMAPLApproved()
        'QID 1007
        Label2.Text = "QID:1007 - Looking for WL Platforms Only UMA/SMA Custom Fee..."
        Dim Mycn As OleDb.OleDbConnection
        Dim Command As OleDb.OleDbCommand
        Dim SQLstr As String
        Try
            Mycn = New OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & My.Settings.dbloc & "';Jet OLEDB:Database Password='" & My.Settings.dbpass & "';")
            Mycn.Open()

            SQLstr = "INSERT INTO map_TempAssignments([FirmID], [ProductID], [TypeOfProductID], [TypeID], [WLID], [PlatformID], [ManagingFirmID], [BreakpointFrom], [BreakpointTo], [Fee], [MaxRepFee], [SMAOffered], [UMAOffered], [PlatformApproved], [QID])" & _
            " SELECT map_PlatformAssignments.FirmID, map_PlatformFees.ProductID, map_Products.TypeOfProductID, 5, 7, map_PlatformsWL.ID, map_Products.ManagingFirmID, map_PlatformFees.BreakpointFrom, map_PlatformFees.BreakpointTo, map_PlatformFees.Fee, map_PlatformFees.MaxRepFee, map_PlatformListings.SMAOffered, map_PlatformListings.UMAOffered, map_PlatformListings.PlatformApproved, 1007" & _
            " FROM ((map_PlatformAssignments INNER JOIN map_PlatformsWL ON map_PlatformAssignments.WLID = map_PlatformsWL.ID) INNER JOIN (map_PlatformFees INNER JOIN map_Products ON map_PlatformFees.ProductID = map_Products.ID) ON (map_PlatformsWL.PlatformDriverID = map_PlatformFees.PlatformID) AND (map_PlatformsWL.ID = map_PlatformFees.WLID)) INNER JOIN map_PlatformListings ON (map_PlatformListings.ProductID = map_PlatformFees.ProductID) AND (map_PlatformsWL.PlatformDriverID = map_PlatformListings.PlatformID)" & _
            " WHERE(((map_PlatformsWL.CustomFee) = True) And ((map_PlatformsWL.OfferUMA) = True) And ((map_PlatformsWL.AdditionalApproval) = False) And ((map_PlatformsWL.OfferSMA) = True) And ((map_PlatformFees.Active) = True) And ((map_PlatformsWL.Active) = True) And ((map_PlatformAssignments.Active) = True) AND map_PlatformAssignments.UseWL = True AND map_PlatformsWL.RequirePlatformApproval = TRUE AND map_PlatformListings.PlatformApproved = True) AND ((map_PlatformListings.Active)=True) AND ((map_PlatformListings.WLID) Is Null)" & _
            " GROUP BY map_PlatformAssignments.FirmID, map_PlatformFees.ProductID, map_Products.TypeOfProductID, map_PlatformsWL.ID, map_Products.ManagingFirmID, map_PlatformFees.BreakpointFrom, map_PlatformFees.BreakpointTo, map_PlatformFees.Fee, map_PlatformFees.MaxRepFee, map_PlatformListings.SMAOffered, map_PlatformListings.UMAOffered, map_PlatformListings.PlatformApproved;"

            Command = New OleDb.OleDbCommand(SQLstr, Mycn)
            Command.ExecuteNonQuery()
            pb.Value = pb.Value + 1
            Label2.Text = "Loaded..."
            Call WLPlatformDefaultFeeOnlyUMA()

            Mycn.Close()

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "ERROR")
        End Try
    End Sub

    Public Sub WLPlatformDefaultFeeOnlyUMA()
        'QID 1008
        Label2.Text = "QID:1008 - Looking for WL Platforms Only UMA Default Fee..."
        Dim Mycn As OleDb.OleDbConnection
        Dim Command As OleDb.OleDbCommand
        Dim SQLstr As String
        Try
            Mycn = New OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & My.Settings.dbloc & "';Jet OLEDB:Database Password='" & My.Settings.dbpass & "';")
            Mycn.Open()

            SQLstr = "INSERT INTO map_TempAssignments([FirmID], [ProductID], [TypeOfProductID], [TypeID], [PlatformID], [WLID], [ManagingFirmID], [BreakpointFrom], [BreakpointTo], [Fee], [MaxRepFee], [SMAOffered], [UMAOffered], [PlatformApproved], [QID])" & _
            " SELECT map_PlatformAssignments.FirmID, map_PlatformFees.ProductID, map_Products.TypeOfProductID, 5, 7, map_PlatformsWL.ID, map_Products.ManagingFirmID, map_PlatformFees.BreakpointFrom, map_PlatformFees.BreakpointTo, map_PlatformFees.Fee, map_PlatformFees.MaxRepFee, map_PlatformListings.SMAOffered, map_PlatformListings.UMAOffered, map_PlatformListings.PlatformApproved, 1008" & _
            " FROM ((map_PlatformAssignments INNER JOIN map_PlatformsWL ON map_PlatformAssignments.WLID = map_PlatformsWL.ID) INNER JOIN (map_PlatformFees INNER JOIN map_Products ON map_PlatformFees.ProductID = map_Products.ID) ON map_PlatformsWL.PlatformDriverID = map_PlatformFees.PlatformID) INNER JOIN map_PlatformListings ON (map_PlatformsWL.PlatformDriverID = map_PlatformListings.PlatformID) AND (map_PlatformFees.ProductID = map_PlatformListings.ProductID)" & _
            " WHERE (((map_PlatformsWL.CustomFee)=False) AND ((map_PlatformsWL.OfferUMA)=True) AND ((map_PlatformListings.UMAOffered) = True) And ((map_PlatformsWL.AdditionalApproval) = False) AND ((map_PlatformsWL.OfferSMA)=False) AND ((map_PlatformFees.Active)=True) AND ((map_PlatformsWL.Active)=True) AND ((map_PlatformAssignments.Active)=True) AND ((map_PlatformAssignments.UseWL)=True) AND ((map_PlatformsWL.RequirePlatformApproval)=False) AND ((map_PlatformFees.WLID) Is Null)) AND ((map_PlatformListings.Active)=True) AND ((map_PlatformListings.WLID) Is Null)" & _
            " GROUP BY map_PlatformAssignments.FirmID, map_PlatformFees.ProductID, map_Products.TypeOfProductID, map_PlatformsWL.ID, map_Products.ManagingFirmID, map_PlatformFees.BreakpointFrom, map_PlatformFees.BreakpointTo, map_PlatformFees.Fee, map_PlatformFees.MaxRepFee, map_PlatformListings.SMAOffered, map_PlatformListings.UMAOffered, map_PlatformListings.PlatformApproved;"

            Command = New OleDb.OleDbCommand(SQLstr, Mycn)
            Command.ExecuteNonQuery()
            pb.Value = pb.Value + 1
            Label2.Text = "Loaded..."
            Call WLPlatformDefaultFeeOnlySMA()

            Mycn.Close()

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "ERROR")
        End Try
    End Sub

    Public Sub WLPlatformDefaultFeeOnlySMA()
        'QID 1009
        Label2.Text = "QID:1009 - Looking for WL Platforms Only SMA Default Fee..."
        Dim Mycn As OleDb.OleDbConnection
        Dim Command As OleDb.OleDbCommand
        Dim SQLstr As String
        Try
            Mycn = New OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & My.Settings.dbloc & "';Jet OLEDB:Database Password='" & My.Settings.dbpass & "';")
            Mycn.Open()

            SQLstr = "INSERT INTO map_TempAssignments([FirmID], [ProductID], [TypeOfProductID], [TypeID], [PlatformID], [WLID], [ManagingFirmID], [BreakpointFrom], [BreakpointTo], [Fee], [MaxRepFee], [SMAOffered], [UMAOffered], [PlatformApproved], [QID])" & _
            " SELECT map_PlatformAssignments.FirmID, map_PlatformFees.ProductID, map_Products.TypeOfProductID, 5, 7, map_PlatformsWL.ID, map_Products.ManagingFirmID, map_PlatformFees.BreakpointFrom, map_PlatformFees.BreakpointTo, map_PlatformFees.Fee, map_PlatformFees.MaxRepFee, map_PlatformListings.SMAOffered, map_PlatformListings.UMAOffered, map_PlatformListings.PlatformApproved, 1009" & _
            " FROM ((map_PlatformAssignments INNER JOIN map_PlatformsWL ON map_PlatformAssignments.WLID = map_PlatformsWL.ID) INNER JOIN (map_PlatformFees INNER JOIN map_Products ON map_PlatformFees.ProductID = map_Products.ID) ON map_PlatformsWL.PlatformDriverID = map_PlatformFees.PlatformID) INNER JOIN map_PlatformListings ON (map_PlatformsWL.PlatformDriverID = map_PlatformListings.PlatformID) AND (map_PlatformFees.ProductID = map_PlatformListings.ProductID)" & _
            " WHERE (((map_PlatformsWL.CustomFee)=False) AND ((map_PlatformsWL.OfferSMA)=True) AND ((map_PlatformListings.SMAOffered) = True) And ((map_PlatformsWL.AdditionalApproval) = False) AND ((map_PlatformsWL.OfferUMA)=False) AND ((map_PlatformFees.Active)=True) AND ((map_PlatformsWL.Active)=True) AND ((map_PlatformAssignments.Active)=True) AND ((map_PlatformAssignments.UseWL)=True) AND ((map_PlatformsWL.RequirePlatformApproval)=False) AND ((map_PlatformFees.WLID) Is Null)) AND ((map_PlatformListings.Active)=True) AND ((map_PlatformListings.WLID) Is Null)" & _
            " GROUP BY map_PlatformAssignments.FirmID, map_PlatformFees.ProductID, map_Products.TypeOfProductID, map_PlatformsWL.ID, map_Products.ManagingFirmID, map_PlatformFees.BreakpointFrom, map_PlatformFees.BreakpointTo, map_PlatformFees.Fee, map_PlatformFees.MaxRepFee, map_PlatformListings.SMAOffered, map_PlatformListings.UMAOffered, map_PlatformListings.PlatformApproved;"

            Command = New OleDb.OleDbCommand(SQLstr, Mycn)
            Command.ExecuteNonQuery()
            pb.Value = pb.Value + 1
            Label2.Text = "Loaded..."
            Call WLPlatformDefaultFeeUMASMA()

            Mycn.Close()

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "ERROR")
        End Try
    End Sub

    Public Sub WLPlatformDefaultFeeUMASMA()
        'QID 1010
        Label2.Text = "QID:1010 - Looking for WL Platforms Only UMA/SMA Default Fee..."
        Dim Mycn As OleDb.OleDbConnection
        Dim Command As OleDb.OleDbCommand
        Dim SQLstr As String
        Try
            Mycn = New OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & My.Settings.dbloc & "';Jet OLEDB:Database Password='" & My.Settings.dbpass & "';")
            Mycn.Open()

            SQLstr = "INSERT INTO map_TempAssignments([FirmID], [ProductID], [TypeOfProductID], [TypeID], [PlatformID], [WLID], [ManagingFirmID], [BreakpointFrom], [BreakpointTo], [Fee], [MaxRepFee], [SMAOffered], [UMAOffered], [PlatformApproved], [QID])" & _
            " SELECT map_PlatformAssignments.FirmID, map_PlatformFees.ProductID, map_Products.TypeOfProductID, 5, 7, map_PlatformsWL.ID, map_Products.ManagingFirmID, map_PlatformFees.BreakpointFrom, map_PlatformFees.BreakpointTo, map_PlatformFees.Fee, map_PlatformFees.MaxRepFee, map_PlatformListings.SMAOffered, map_PlatformListings.UMAOffered, map_PlatformListings.PlatformApproved, 1010" & _
            " FROM ((map_PlatformAssignments INNER JOIN map_PlatformsWL ON map_PlatformAssignments.WLID = map_PlatformsWL.ID) INNER JOIN (map_PlatformFees INNER JOIN map_Products ON map_PlatformFees.ProductID = map_Products.ID) ON map_PlatformsWL.PlatformDriverID = map_PlatformFees.PlatformID) INNER JOIN map_PlatformListings ON (map_PlatformsWL.PlatformDriverID = map_PlatformListings.PlatformID) AND (map_PlatformFees.ProductID = map_PlatformListings.ProductID)" & _
            " WHERE (((map_PlatformsWL.CustomFee)=False) AND ((map_PlatformsWL.OfferSMA)=True) And ((map_PlatformsWL.AdditionalApproval) = False) AND ((map_PlatformsWL.OfferUMA)=True) AND ((map_PlatformFees.Active)=True) AND ((map_PlatformsWL.Active)=True) AND ((map_PlatformAssignments.Active)=True) AND ((map_PlatformAssignments.UseWL)=True) AND ((map_PlatformsWL.RequirePlatformApproval)=False) AND ((map_PlatformFees.WLID) Is Null)) AND ((map_PlatformListings.Active)=True) AND ((map_PlatformListings.WLID) Is Null)" & _
            " GROUP BY map_PlatformAssignments.FirmID, map_PlatformFees.ProductID, map_Products.TypeOfProductID, map_PlatformsWL.ID, map_Products.ManagingFirmID, map_PlatformFees.BreakpointFrom, map_PlatformFees.BreakpointTo, map_PlatformFees.Fee, map_PlatformFees.MaxRepFee, map_PlatformListings.SMAOffered, map_PlatformListings.UMAOffered, map_PlatformListings.PlatformApproved;"

            Command = New OleDb.OleDbCommand(SQLstr, Mycn)
            Command.ExecuteNonQuery()
            pb.Value = pb.Value + 1
            Label2.Text = "Loaded..."
            Call WLPlatformDefaultFeeOnlyUMAPLApproved()

            Mycn.Close()

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "ERROR")
        End Try
    End Sub

    Public Sub WLPlatformDefaultFeeOnlyUMAPLApproved()
        'QID 1011
        Label2.Text = "QID:1011 - Looking for WL Platforms Approved Only UMA Default Fee..."
        Dim Mycn As OleDb.OleDbConnection
        Dim Command As OleDb.OleDbCommand
        Dim SQLstr As String
        Try
            Mycn = New OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & My.Settings.dbloc & "';Jet OLEDB:Database Password='" & My.Settings.dbpass & "';")
            Mycn.Open()

            SQLstr = "INSERT INTO map_TempAssignments([FirmID], [ProductID], [TypeOfProductID], [TypeID], [PlatformID], [WLID], [ManagingFirmID], [BreakpointFrom], [BreakpointTo], [Fee], [MaxRepFee], [SMAOffered], [UMAOffered], [PlatformApproved], [QID])" & _
            " SELECT map_PlatformAssignments.FirmID, map_PlatformFees.ProductID, map_Products.TypeOfProductID, 5, 7, map_PlatformsWL.ID, map_Products.ManagingFirmID, map_PlatformFees.BreakpointFrom, map_PlatformFees.BreakpointTo, map_PlatformFees.Fee, map_PlatformFees.MaxRepFee, map_PlatformListings.SMAOffered, map_PlatformListings.UMAOffered, map_PlatformListings.PlatformApproved, 1011" & _
            " FROM ((map_PlatformAssignments INNER JOIN map_PlatformsWL ON map_PlatformAssignments.WLID = map_PlatformsWL.ID) INNER JOIN (map_PlatformFees INNER JOIN map_Products ON map_PlatformFees.ProductID = map_Products.ID) ON map_PlatformsWL.PlatformDriverID = map_PlatformFees.PlatformID) INNER JOIN map_PlatformListings ON (map_PlatformsWL.PlatformDriverID = map_PlatformListings.PlatformID) AND (map_PlatformFees.ProductID = map_PlatformListings.ProductID)" & _
            " WHERE(((map_PlatformsWL.CustomFee) = False) And ((map_PlatformsWL.OfferUMA) = True) AND ((map_PlatformListings.UMAOffered) = True) And ((map_PlatformsWL.AdditionalApproval) = False) And ((map_PlatformsWL.OfferSMA) = False) And ((map_PlatformFees.Active) = True) And ((map_PlatformsWL.Active) = True) And ((map_PlatformAssignments.Active) = True) And ((map_PlatformAssignments.UseWL) = True) And ((map_PlatformsWL.RequirePlatformApproval) = True) And ((map_PlatformListings.PlatformApproved) = True) And ((map_PlatformFees.WLID) Is Null))" & _
            " GROUP BY map_PlatformAssignments.FirmID, map_PlatformFees.ProductID, map_Products.TypeOfProductID, map_PlatformsWL.ID, map_Products.ManagingFirmID, map_PlatformFees.BreakpointFrom, map_PlatformFees.BreakpointTo, map_PlatformFees.Fee, map_PlatformFees.MaxRepFee, map_PlatformListings.SMAOffered, map_PlatformListings.UMAOffered, map_PlatformListings.PlatformApproved;"


            Command = New OleDb.OleDbCommand(SQLstr, Mycn)
            Command.ExecuteNonQuery()
            pb.Value = pb.Value + 1
            Label2.Text = "Loaded..."
            Call WLPlatformDefaultFeeOnlySMAPLApproved()

            Mycn.Close()

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "ERROR")
        End Try
    End Sub

    Public Sub WLPlatformDefaultFeeOnlySMAPLApproved()
        'QID 1012
        Label2.Text = "QID:1012 - Looking for WL Platforms Approved Only SMA Default Fee..."
        Dim Mycn As OleDb.OleDbConnection
        Dim Command As OleDb.OleDbCommand
        Dim SQLstr As String
        Try
            Mycn = New OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & My.Settings.dbloc & "';Jet OLEDB:Database Password='" & My.Settings.dbpass & "';")
            Mycn.Open()

            SQLstr = "INSERT INTO map_TempAssignments([FirmID], [ProductID], [TypeOfProductID], [TypeID], [PlatformID], [WLID], [ManagingFirmID], [BreakpointFrom], [BreakpointTo], [Fee], [MaxRepFee], [SMAOffered], [UMAOffered], [PlatformApproved], [QID])" & _
            " SELECT map_PlatformAssignments.FirmID, map_PlatformFees.ProductID, map_Products.TypeOfProductID, 5, 7, map_PlatformsWL.ID, map_Products.ManagingFirmID, map_PlatformFees.BreakpointFrom, map_PlatformFees.BreakpointTo, map_PlatformFees.Fee, map_PlatformFees.MaxRepFee, map_PlatformListings.SMAOffered, map_PlatformListings.UMAOffered, map_PlatformListings.PlatformApproved, 1012" & _
            " FROM ((map_PlatformAssignments INNER JOIN map_PlatformsWL ON map_PlatformAssignments.WLID = map_PlatformsWL.ID) INNER JOIN (map_PlatformFees INNER JOIN map_Products ON map_PlatformFees.ProductID = map_Products.ID) ON map_PlatformsWL.PlatformDriverID = map_PlatformFees.PlatformID) INNER JOIN map_PlatformListings ON (map_PlatformsWL.PlatformDriverID = map_PlatformListings.PlatformID) AND (map_PlatformFees.ProductID = map_PlatformListings.ProductID)" & _
            " WHERE(((map_PlatformsWL.CustomFee) = False) And ((map_PlatformsWL.OfferSMA) = True) AND ((map_PlatformListings.SMAOffered) = True) And ((map_PlatformsWL.AdditionalApproval) = False) And ((map_PlatformsWL.OfferUMA) = False) And ((map_PlatformFees.Active) = True) And ((map_PlatformsWL.Active) = True) And ((map_PlatformAssignments.Active) = True) And ((map_PlatformAssignments.UseWL) = True) And ((map_PlatformsWL.RequirePlatformApproval) = True) And ((map_PlatformListings.PlatformApproved) = True) And ((map_PlatformFees.WLID) Is Null))" & _
            " GROUP BY map_PlatformAssignments.FirmID, map_PlatformFees.ProductID, map_Products.TypeOfProductID, map_PlatformsWL.ID, map_Products.ManagingFirmID, map_PlatformFees.BreakpointFrom, map_PlatformFees.BreakpointTo, map_PlatformFees.Fee, map_PlatformFees.MaxRepFee, map_PlatformListings.SMAOffered, map_PlatformListings.UMAOffered, map_PlatformListings.PlatformApproved;"


            Command = New OleDb.OleDbCommand(SQLstr, Mycn)
            Command.ExecuteNonQuery()
            pb.Value = pb.Value + 1
            Label2.Text = "Loaded..."
            Call WLPlatformDefaultFeeOnlySMAUMAPLApproved()

            Mycn.Close()

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "ERROR")
        End Try
    End Sub

    Public Sub WLPlatformDefaultFeeOnlySMAUMAPLApproved()
        'QID 1013
        Label2.Text = "QID:1013 - Looking for WL Platforms Approved Only SMA Default Fee..."
        Dim Mycn As OleDb.OleDbConnection
        Dim Command As OleDb.OleDbCommand
        Dim SQLstr As String
        Try
            Mycn = New OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & My.Settings.dbloc & "';Jet OLEDB:Database Password='" & My.Settings.dbpass & "';")
            Mycn.Open()

            SQLstr = "INSERT INTO map_TempAssignments([FirmID], [ProductID], [TypeOfProductID], [TypeID], [PlatformID], [WLID], [ManagingFirmID], [BreakpointFrom], [BreakpointTo], [Fee], [MaxRepFee], [SMAOffered], [UMAOffered], [PlatformApproved], [QID])" & _
            " SELECT map_PlatformAssignments.FirmID, map_PlatformFees.ProductID, map_Products.TypeOfProductID, 5, 7, map_PlatformsWL.ID, map_Products.ManagingFirmID, map_PlatformFees.BreakpointFrom, map_PlatformFees.BreakpointTo, map_PlatformFees.Fee, map_PlatformFees.MaxRepFee, map_PlatformListings.SMAOffered, map_PlatformListings.UMAOffered, map_PlatformListings.PlatformApproved, 1013" & _
            " FROM ((map_PlatformAssignments INNER JOIN map_PlatformsWL ON map_PlatformAssignments.WLID = map_PlatformsWL.ID) INNER JOIN (map_PlatformFees INNER JOIN map_Products ON map_PlatformFees.ProductID = map_Products.ID) ON map_PlatformsWL.PlatformDriverID = map_PlatformFees.PlatformID) INNER JOIN map_PlatformListings ON (map_PlatformsWL.PlatformDriverID = map_PlatformListings.PlatformID) AND (map_PlatformFees.ProductID = map_PlatformListings.ProductID)" & _
            " WHERE(((map_PlatformsWL.CustomFee) = False) And ((map_PlatformsWL.OfferSMA) = True) And ((map_PlatformsWL.AdditionalApproval) = False) And ((map_PlatformsWL.OfferUMA) = True) And ((map_PlatformFees.Active) = True) And ((map_PlatformsWL.Active) = True) And ((map_PlatformAssignments.Active) = True) And ((map_PlatformAssignments.UseWL) = True) And ((map_PlatformsWL.RequirePlatformApproval) = True) And ((map_PlatformListings.PlatformApproved) = True) And ((map_PlatformFees.WLID) Is Null))" & _
            " GROUP BY map_PlatformAssignments.FirmID, map_PlatformFees.ProductID, map_Products.TypeOfProductID, map_PlatformsWL.ID, map_Products.ManagingFirmID, map_PlatformFees.BreakpointFrom, map_PlatformFees.BreakpointTo, map_PlatformFees.Fee, map_PlatformFees.MaxRepFee, map_PlatformListings.SMAOffered, map_PlatformListings.UMAOffered, map_PlatformListings.PlatformApproved;"


            Command = New OleDb.OleDbCommand(SQLstr, Mycn)
            Command.ExecuteNonQuery()
            pb.Value = pb.Value + 1
            Label2.Text = "Loaded..."
            Call RemovedAutoApprovedRecords()

            Mycn.Close()

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "ERROR")
        End Try
    End Sub

    Public Sub RemovedAutoApprovedRecords()
        Label2.Text = "Removing Auto Assigned Products"
        Dim Mycn As OleDb.OleDbConnection
        Dim Command As OleDb.OleDbCommand
        Dim SQLstr As String
        Try
            Mycn = New OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & My.Settings.dbloc & "';Jet OLEDB:Database Password='" & My.Settings.dbpass & "';")
            Mycn.Open()

            SQLstr = "DELETE * FROM map_PlatformListings WHERE AutoAdded = True AND Active = True AND WLID In (SELECT map_PlatformsWL.ID" & _
            " FROM map_PlatformAssignments INNER JOIN map_PlatformsWL ON map_PlatformAssignments.WLID = map_PlatformsWL.ID" & _
            " WHERE(((map_PlatformsWL.RequirePlatformApproval) = True) And ((map_PlatformsWL.AdditionalApproval) = True))" & _
            " GROUP BY map_PlatformsWL.ID)"

            Command = New OleDb.OleDbCommand(SQLstr, Mycn)
            Command.ExecuteNonQuery()


            Mycn.Close()
            Label2.Text = "Removed..."
            Call InsertPlatformApproved()

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "ERROR")
        End Try
    End Sub

    Public Sub InsertPlatformApproved()
        Label2.Text = "Loading Auto Approve Products"
        Dim Mycn As OleDb.OleDbConnection
        Dim Command As OleDb.OleDbCommand
        Dim SQLstr As String
        Try
            Mycn = New OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & My.Settings.dbloc & "';Jet OLEDB:Database Password='" & My.Settings.dbpass & "';")
            Mycn.Open()

            SQLstr = "INSERT INTO map_PlatformListings(PlatformID, WLID, ProductID, PlatformApproved, WLApproved, Active, SMAOffered, UMAOffered, AutoAdded)" & _
            "SELECT map_PlatformsWL.PlatformDriverID, map_PlatformsWL.ID, map_PlatformListings.ProductID, map_PlatformListings.PlatformApproved, map_PlatformListings.WLApproved, map_PlatformListings.Active, map_PlatformListings.SMAOffered, map_PlatformListings.UMAOffered, -1" & _
            " FROM (map_PlatformAssignments INNER JOIN map_PlatformsWL ON map_PlatformAssignments.WLID = map_PlatformsWL.ID) INNER JOIN map_PlatformListings ON map_PlatformsWL.PlatformDriverID = map_PlatformListings.PlatformID" & _
            " WHERE(((map_PlatformsWL.RequirePlatformApproval) = True) And ((map_PlatformsWL.AdditionalApproval) = True) And ((map_PlatformListings.PlatformApproved) = True) And ((map_PlatformListings.Active) = True) And ((map_PlatformAssignments.Active) = True) And ((map_PlatformListings.WLID) Is Null))" & _
            " GROUP BY map_PlatformsWL.PlatformDriverID, map_PlatformsWL.ID, map_PlatformListings.ProductID, map_PlatformListings.PlatformApproved, map_PlatformListings.WLApproved, map_PlatformListings.Active, map_PlatformListings.SMAOffered, map_PlatformListings.UMAOffered;"

            Command = New OleDb.OleDbCommand(SQLstr, Mycn)
            Command.ExecuteNonQuery()

            Mycn.Close()
            Label2.Text = "Loaded..."
            Call WLPlatformCustomFeeOnlyUMACustomApproval()

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "ERROR")
        End Try
    End Sub

    Public Sub WLPlatformCustomFeeOnlyUMACustomApproval()
        'QID 1014
        Label2.Text = "QID:1014 - Looking for WL Platforms Only UMA Custom Fee Custom Approval..."
        Dim Mycn As OleDb.OleDbConnection
        Dim Command As OleDb.OleDbCommand
        Dim SQLstr As String
        Try
            Mycn = New OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & My.Settings.dbloc & "';Jet OLEDB:Database Password='" & My.Settings.dbpass & "';")
            Mycn.Open()

            SQLstr = "INSERT INTO map_TempAssignments([FirmID], [ProductID], [TypeOfProductID], [TypeID], [PlatformID], [WLID], [ManagingFirmID], [BreakpointFrom], [BreakpointTo], [Fee], [MaxRepFee], [SMAOffered], [UMAOffered], [PlatformApproved], [QID])" & _
            " SELECT map_PlatformAssignments.FirmID, map_PlatformFees.ProductID, map_Products.TypeOfProductID, 5, 7, map_PlatformsWL.ID, map_Products.ManagingFirmID, map_PlatformFees.BreakpointFrom, map_PlatformFees.BreakpointTo, map_PlatformFees.Fee, map_PlatformFees.MaxRepFee, map_PlatformsWL.OfferSMA, map_PlatformsWL.OfferUMA, map_PlatformListings.PlatformApproved, 1014" & _
            " FROM ((map_PlatformAssignments INNER JOIN map_PlatformsWL ON map_PlatformAssignments.WLID = map_PlatformsWL.ID) INNER JOIN map_PlatformListings ON map_PlatformsWL.ID = map_PlatformListings.WLID) INNER JOIN (map_PlatformFees INNER JOIN map_Products ON map_PlatformFees.ProductID = map_Products.ID) ON (map_PlatformListings.ProductID = map_PlatformFees.ProductID) AND (map_PlatformListings.WLID = map_PlatformFees.WLID)" & _
            " WHERE(((map_PlatformsWL.CustomFee) = True) And ((map_PlatformsWL.OfferUMA) = True) AND ((map_PlatformListings.UMAOffered) = True) And ((map_PlatformsWL.AdditionalApproval) = True) And ((map_PlatformsWL.OfferSMA) = False) And ((map_PlatformFees.Active) = True) And ((map_PlatformsWL.Active) = True) And ((map_PlatformAssignments.Active) = True) And ((map_PlatformAssignments.UseWL) = True))" & _
            " GROUP BY map_PlatformAssignments.FirmID, map_PlatformFees.ProductID, map_Products.TypeOfProductID, map_PlatformsWL.ID, map_Products.ManagingFirmID, map_PlatformFees.BreakpointFrom, map_PlatformFees.BreakpointTo, map_PlatformFees.Fee, map_PlatformFees.MaxRepFee, map_PlatformsWL.OfferSMA, map_PlatformsWL.OfferUMA, map_PlatformListings.PlatformApproved;"

            Command = New OleDb.OleDbCommand(SQLstr, Mycn)
            Command.ExecuteNonQuery()
            pb.Value = pb.Value + 1
            Label2.Text = "Loaded..."
            Call WLPlatformCustomFeeOnlySMACustomApproval()

            Mycn.Close()

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "ERROR")
        End Try
    End Sub

    Public Sub WLPlatformCustomFeeOnlySMACustomApproval()
        'QID 1015
        Label2.Text = "QID:1015 - Looking for WL Platforms Only SMA Custom Fee Custom Approval..."
        Dim Mycn As OleDb.OleDbConnection
        Dim Command As OleDb.OleDbCommand
        Dim SQLstr As String
        Try
            Mycn = New OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & My.Settings.dbloc & "';Jet OLEDB:Database Password='" & My.Settings.dbpass & "';")
            Mycn.Open()

            SQLstr = "INSERT INTO map_TempAssignments([FirmID], [ProductID], [TypeOfProductID], [TypeID], [PlatformID], [WLID], [ManagingFirmID], [BreakpointFrom], [BreakpointTo], [Fee], [MaxRepFee], [SMAOffered], [UMAOffered], [PlatformApproved], [QID])" & _
            " SELECT map_PlatformAssignments.FirmID, map_PlatformFees.ProductID, map_Products.TypeOfProductID, 5, 7, map_PlatformsWL.ID, map_Products.ManagingFirmID, map_PlatformFees.BreakpointFrom, map_PlatformFees.BreakpointTo, map_PlatformFees.Fee, map_PlatformFees.MaxRepFee, map_PlatformsWL.OfferSMA, map_PlatformsWL.OfferUMA, map_PlatformListings.PlatformApproved, 1015" & _
            " FROM ((map_PlatformAssignments INNER JOIN map_PlatformsWL ON map_PlatformAssignments.WLID = map_PlatformsWL.ID) INNER JOIN map_PlatformListings ON map_PlatformsWL.ID = map_PlatformListings.WLID) INNER JOIN (map_PlatformFees INNER JOIN map_Products ON map_PlatformFees.ProductID = map_Products.ID) ON (map_PlatformListings.ProductID = map_PlatformFees.ProductID) AND (map_PlatformListings.WLID = map_PlatformFees.WLID)" & _
            " WHERE(((map_PlatformsWL.CustomFee) = True) And ((map_PlatformsWL.OfferSMA) = True) AND ((map_PlatformListings.SMAOffered) = True) And ((map_PlatformsWL.AdditionalApproval) = True) And ((map_PlatformsWL.OfferUMA) = False) And ((map_PlatformFees.Active) = True) And ((map_PlatformsWL.Active) = True) And ((map_PlatformAssignments.Active) = True) And ((map_PlatformAssignments.UseWL) = True))" & _
            " GROUP BY map_PlatformAssignments.FirmID, map_PlatformFees.ProductID, map_Products.TypeOfProductID, map_PlatformsWL.ID, map_Products.ManagingFirmID, map_PlatformFees.BreakpointFrom, map_PlatformFees.BreakpointTo, map_PlatformFees.Fee, map_PlatformFees.MaxRepFee, map_PlatformsWL.OfferSMA, map_PlatformsWL.OfferUMA, map_PlatformListings.PlatformApproved;"

            Command = New OleDb.OleDbCommand(SQLstr, Mycn)
            Command.ExecuteNonQuery()
            pb.Value = pb.Value + 1
            Label2.Text = "Loaded..."
            Call WLPlatformCustomFeeOnlySMAUMACustomApproval()

            Mycn.Close()

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "ERROR")
        End Try
    End Sub

    Public Sub WLPlatformCustomFeeOnlySMAUMACustomApproval()
        'QID 1016
        Label2.Text = "QID:1016 - Looking for WL Platforms UMA/SMA Custom Fee Custom Approval..."
        Dim Mycn As OleDb.OleDbConnection
        Dim Command As OleDb.OleDbCommand
        Dim SQLstr As String
        Try
            Mycn = New OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & My.Settings.dbloc & "';Jet OLEDB:Database Password='" & My.Settings.dbpass & "';")
            Mycn.Open()

            SQLstr = "INSERT INTO map_TempAssignments([FirmID], [ProductID], [TypeOfProductID], [TypeID], [PlatformID], [WLID], [ManagingFirmID], [BreakpointFrom], [BreakpointTo], [Fee], [MaxRepFee], [SMAOffered], [UMAOffered], [PlatformApproved], [QID])" & _
            " SELECT map_PlatformAssignments.FirmID, map_PlatformFees.ProductID, map_Products.TypeOfProductID, 5, 7, map_PlatformsWL.ID, map_Products.ManagingFirmID, map_PlatformFees.BreakpointFrom, map_PlatformFees.BreakpointTo, map_PlatformFees.Fee, map_PlatformFees.MaxRepFee, map_PlatformsWL.OfferSMA, map_PlatformsWL.OfferUMA, map_PlatformListings.PlatformApproved, 1016" & _
            " FROM ((map_PlatformAssignments INNER JOIN map_PlatformsWL ON map_PlatformAssignments.WLID = map_PlatformsWL.ID) INNER JOIN map_PlatformListings ON map_PlatformsWL.ID = map_PlatformListings.WLID) INNER JOIN (map_PlatformFees INNER JOIN map_Products ON map_PlatformFees.ProductID = map_Products.ID) ON (map_PlatformListings.ProductID = map_PlatformFees.ProductID) AND (map_PlatformListings.WLID = map_PlatformFees.WLID)" & _
            " WHERE(((map_PlatformsWL.CustomFee) = True) And ((map_PlatformsWL.OfferSMA) = True) And ((map_PlatformsWL.AdditionalApproval) = True) And ((map_PlatformsWL.OfferUMA) = True) And ((map_PlatformFees.Active) = True) And ((map_PlatformsWL.Active) = True) And ((map_PlatformAssignments.Active) = True) And ((map_PlatformAssignments.UseWL) = True))" & _
            " GROUP BY map_PlatformAssignments.FirmID, map_PlatformFees.ProductID, map_Products.TypeOfProductID, map_PlatformsWL.ID, map_Products.ManagingFirmID, map_PlatformFees.BreakpointFrom, map_PlatformFees.BreakpointTo, map_PlatformFees.Fee, map_PlatformFees.MaxRepFee, map_PlatformsWL.OfferSMA, map_PlatformsWL.OfferUMA, map_PlatformListings.PlatformApproved;"

            Command = New OleDb.OleDbCommand(SQLstr, Mycn)
            Command.ExecuteNonQuery()
            pb.Value = pb.Value + 1
            Label2.Text = "Loaded..."
            Call WLPlatformDefaultFeeOnlyUMACustomApproval()

            Mycn.Close()

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "ERROR")
        End Try
    End Sub

    Public Sub WLPlatformDefaultFeeOnlyUMACustomApproval()
        'QID 1017
        Label2.Text = "QID:1017 - Looking for WL Platforms UMA/SMA Custom Fee Custom Approval..."
        Dim Mycn As OleDb.OleDbConnection
        Dim Command As OleDb.OleDbCommand
        Dim SQLstr As String
        Try
            Mycn = New OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & My.Settings.dbloc & "';Jet OLEDB:Database Password='" & My.Settings.dbpass & "';")
            Mycn.Open()

            SQLstr = "INSERT INTO map_TempAssignments([FirmID], [ProductID], [TypeOfProductID], [TypeID], [PlatformID], [WLID], [ManagingFirmID], [BreakpointFrom], [BreakpointTo], [Fee], [MaxRepFee], [SMAOffered], [UMAOffered], [PlatformApproved], [QID])" & _
            " SELECT map_PlatformAssignments.FirmID, map_PlatformFees.ProductID, map_Products.TypeOfProductID, 5, 7, map_PlatformsWL.ID, map_Products.ManagingFirmID, map_PlatformFees.BreakpointFrom, map_PlatformFees.BreakpointTo, map_PlatformFees.Fee, map_PlatformFees.MaxRepFee, map_PlatformsWL.OfferSMA, map_PlatformsWL.OfferUMA, map_PlatformListings.PlatformApproved, 1017" & _
            " FROM ((map_PlatformAssignments INNER JOIN map_PlatformsWL ON map_PlatformAssignments.WLID = map_PlatformsWL.ID) INNER JOIN map_PlatformListings ON map_PlatformsWL.ID = map_PlatformListings.WLID) INNER JOIN (map_PlatformFees INNER JOIN map_Products ON map_PlatformFees.ProductID = map_Products.ID) ON (map_PlatformListings.ProductID = map_Products.ID) AND (map_PlatformListings.PlatformID = map_PlatformFees.PlatformID)" & _
            " WHERE(((map_PlatformsWL.CustomFee) = False) And ((map_PlatformsWL.OfferUMA) = True) AND ((map_PlatformListings.UMAOffered) = True) And ((map_PlatformsWL.AdditionalApproval) = True) And ((map_PlatformsWL.OfferSMA) = False) And ((map_PlatformFees.Active) = True) And ((map_PlatformsWL.Active) = True) And ((map_PlatformAssignments.Active) = True) And ((map_PlatformAssignments.UseWL) = True) And ((map_PlatformFees.WLID) Is Null))" & _
            " GROUP BY map_PlatformAssignments.FirmID, map_PlatformFees.ProductID, map_Products.TypeOfProductID, map_PlatformsWL.ID, map_Products.ManagingFirmID, map_PlatformFees.BreakpointFrom, map_PlatformFees.BreakpointTo, map_PlatformFees.Fee, map_PlatformFees.MaxRepFee, map_PlatformsWL.OfferSMA, map_PlatformsWL.OfferUMA, map_PlatformListings.PlatformApproved;"


            Command = New OleDb.OleDbCommand(SQLstr, Mycn)
            Command.ExecuteNonQuery()
            pb.Value = pb.Value + 1
            Label2.Text = "Loaded..."
            Call WLPlatformDefaultFeeOnlySMACustomApproval()

            Mycn.Close()

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "ERROR")
        End Try
    End Sub

    Public Sub WLPlatformDefaultFeeOnlySMACustomApproval()
        'QID 1018
        Label2.Text = "QID:1018 - Looking for WL Platforms UMA/SMA Custom Fee Custom Approval..."
        Dim Mycn As OleDb.OleDbConnection
        Dim Command As OleDb.OleDbCommand
        Dim SQLstr As String
        Try
            Mycn = New OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & My.Settings.dbloc & "';Jet OLEDB:Database Password='" & My.Settings.dbpass & "';")
            Mycn.Open()

            SQLstr = "INSERT INTO map_TempAssignments([FirmID], [ProductID], [TypeOfProductID], [TypeID], [PlatformID], [WLID], [ManagingFirmID], [BreakpointFrom], [BreakpointTo], [Fee], [MaxRepFee], [SMAOffered], [UMAOffered], [PlatformApproved], [QID])" & _
            " SELECT map_PlatformAssignments.FirmID, map_PlatformFees.ProductID, map_Products.TypeOfProductID, 5, 7, map_PlatformsWL.ID, map_Products.ManagingFirmID, map_PlatformFees.BreakpointFrom, map_PlatformFees.BreakpointTo, map_PlatformFees.Fee, map_PlatformFees.MaxRepFee, map_PlatformsWL.OfferSMA, map_PlatformsWL.OfferUMA, map_PlatformListings.PlatformApproved, 1018" & _
            " FROM ((map_PlatformAssignments INNER JOIN map_PlatformsWL ON map_PlatformAssignments.WLID = map_PlatformsWL.ID) INNER JOIN map_PlatformListings ON map_PlatformsWL.ID = map_PlatformListings.WLID) INNER JOIN (map_PlatformFees INNER JOIN map_Products ON map_PlatformFees.ProductID = map_Products.ID) ON (map_PlatformListings.ProductID = map_Products.ID) AND (map_PlatformListings.PlatformID = map_PlatformFees.PlatformID)" & _
            " WHERE(((map_PlatformsWL.CustomFee) = False) And ((map_PlatformsWL.OfferSMA) = True) AND ((map_PlatformListings.SMAOffered) = True) And ((map_PlatformsWL.AdditionalApproval) = True) And ((map_PlatformsWL.OfferUMA) = False) And ((map_PlatformFees.Active) = True) And ((map_PlatformsWL.Active) = True) And ((map_PlatformAssignments.Active) = True) And ((map_PlatformAssignments.UseWL) = True) And ((map_PlatformFees.WLID) Is Null))" & _
            " GROUP BY map_PlatformAssignments.FirmID, map_PlatformFees.ProductID, map_Products.TypeOfProductID, map_PlatformsWL.ID, map_Products.ManagingFirmID, map_PlatformFees.BreakpointFrom, map_PlatformFees.BreakpointTo, map_PlatformFees.Fee, map_PlatformFees.MaxRepFee, map_PlatformsWL.OfferSMA, map_PlatformsWL.OfferUMA, map_PlatformListings.PlatformApproved;"

            Command = New OleDb.OleDbCommand(SQLstr, Mycn)
            Command.ExecuteNonQuery()
            pb.Value = pb.Value + 1
            Label2.Text = "Loaded..."
            Call WLPlatformDefaultFeeOnlyUMASMACustomApproval()

            Mycn.Close()

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "ERROR")
        End Try
    End Sub

    Public Sub WLPlatformDefaultFeeOnlyUMASMACustomApproval()
        'QID 1019
        Label2.Text = "QID:1019 - Looking for WL Platforms UMA/SMA Custom Fee Custom Approval..."
        Dim Mycn As OleDb.OleDbConnection
        Dim Command As OleDb.OleDbCommand
        Dim SQLstr As String
        Try
            Mycn = New OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & My.Settings.dbloc & "';Jet OLEDB:Database Password='" & My.Settings.dbpass & "';")
            Mycn.Open()

            SQLstr = "INSERT INTO map_TempAssignments([FirmID], [ProductID], [TypeOfProductID], [TypeID], [PlatformID], [WLID], [ManagingFirmID], [BreakpointFrom], [BreakpointTo], [Fee], [MaxRepFee], [SMAOffered], [UMAOffered], [PlatformApproved], [QID])" & _
            " SELECT map_PlatformAssignments.FirmID, map_PlatformFees.ProductID, map_Products.TypeOfProductID, 5, 7, map_PlatformsWL.ID, map_Products.ManagingFirmID, map_PlatformFees.BreakpointFrom, map_PlatformFees.BreakpointTo, map_PlatformFees.Fee, map_PlatformFees.MaxRepFee, map_PlatformsWL.OfferSMA, map_PlatformsWL.OfferUMA, map_PlatformListings.PlatformApproved, 1019" & _
            " FROM ((map_PlatformAssignments INNER JOIN map_PlatformsWL ON map_PlatformAssignments.WLID = map_PlatformsWL.ID) INNER JOIN map_PlatformListings ON map_PlatformsWL.ID = map_PlatformListings.WLID) INNER JOIN (map_PlatformFees INNER JOIN map_Products ON map_PlatformFees.ProductID = map_Products.ID) ON (map_PlatformListings.ProductID = map_Products.ID) AND (map_PlatformListings.PlatformID = map_PlatformFees.PlatformID)" & _
            " WHERE(((map_PlatformsWL.CustomFee) = False) And ((map_PlatformsWL.OfferSMA) = True) And ((map_PlatformsWL.AdditionalApproval) = True) And ((map_PlatformsWL.OfferUMA) = True) And ((map_PlatformFees.Active) = True) And ((map_PlatformsWL.Active) = True) And ((map_PlatformAssignments.Active) = True) And ((map_PlatformAssignments.UseWL) = True) And ((map_PlatformFees.WLID) Is Null))" & _
            " GROUP BY map_PlatformAssignments.FirmID, map_PlatformFees.ProductID, map_Products.TypeOfProductID, map_PlatformsWL.ID, map_Products.ManagingFirmID, map_PlatformFees.BreakpointFrom, map_PlatformFees.BreakpointTo, map_PlatformFees.Fee, map_PlatformFees.MaxRepFee, map_PlatformsWL.OfferSMA, map_PlatformsWL.OfferUMA, map_PlatformListings.PlatformApproved;"

            Command = New OleDb.OleDbCommand(SQLstr, Mycn)
            Command.ExecuteNonQuery()
            pb.Value = pb.Value + 1
            Label2.Text = "Loaded..."

            Mycn.Close()

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "ERROR")
        End Try
    End Sub

    Public Sub LoadAllInhouseApprovedAgreementsDefaultFee()
        'QID 1
        Label2.Text = "QID:1 - Looking for All InHouse Approved Default Fee..."
        Dim Mycn As OleDb.OleDbConnection
        Dim Command As OleDb.OleDbCommand
        Dim SQLstr As String
        Try
            Mycn = New OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & My.Settings.dbloc & "';Jet OLEDB:Database Password='" & My.Settings.dbpass & "';")
            Mycn.Open()

            SQLstr = "INSERT INTO map_TempAssignments([PlatformID], [WLID], [FirmID], [ProductID], [TypeOfProductID], [TypeID], [ManagingFirmID], [BreakpointFrom], [BreakpointTo], [Fee], [MaxRepFee], [SMAOffered], [QID])" & _
            " SELECT 7, 47, map_Agreements.FirmID, map_Products.ID, map_Products.TypeOfProductID, map_Agreements.TypeID, map_Products.ManagingFirmID, map_Fees.BreakpointFrom, map_Fees.BreakpointTo, map_Fees.Fee, map_Fees.MaxRepFee, -1, 1" & _
            " FROM map_Fees INNER JOIN (map_Products INNER JOIN map_Agreements ON map_Products.Active = map_Agreements.AllInHouseSMAApproved) ON (map_Agreements.TypeID = map_Fees.AgreementTypeID) AND (map_Fees.ProductID = map_Products.ID)" & _
            " WHERE(((map_Products.ManagingFirmID) = 1) And ((map_Fees.Default) = True) And ((map_Fees.Active) = True) And ((map_Agreements.UseDefaultFee) = True) And ((map_Agreements.Active) = True) And ((map_Agreements.TypeID) = 1))" & _
            " GROUP BY map_Agreements.FirmID, map_Products.ID, map_Products.TypeOfProductID, map_Agreements.TypeID, map_Products.ManagingFirmID, map_Fees.BreakpointFrom, map_Fees.BreakpointTo, map_Fees.Fee, map_Fees.MaxRepFee;"

            Command = New OleDb.OleDbCommand(SQLstr, Mycn)
            Command.ExecuteNonQuery()
            pb.Value = pb.Value + 1
            Label2.Text = "Loaded..."
            Call LoadAllOutsideApprovedAgreementsDefaultFee()

            Mycn.Close()

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "ERROR")
        End Try
    End Sub

    Public Sub LoadAllInhouseApprovedAgreementsDefaultFeeSOLICITED()
        'QID 2
        Label2.Text = "QID:2 - Looking for All InHouse Approved Default Fee..."
        Dim Mycn As OleDb.OleDbConnection
        Dim Command As OleDb.OleDbCommand
        Dim SQLstr As String
        Try
            Mycn = New OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & My.Settings.dbloc & "';Jet OLEDB:Database Password='" & My.Settings.dbpass & "';")
            Mycn.Open()

            SQLstr = "INSERT INTO map_TempAssignments([PlatformID], [WLID], [FirmID], [ProductID], [TypeOfProductID], [TypeID], [ManagingFirmID], [BreakpointFrom], [BreakpointTo], [Fee], [MaxRepFee], [SMAOffered], [QID])" & _
            " SELECT 7,47,map_Agreements.FirmID, map_Products.ID, map_Products.TypeOfProductID, map_Agreements.TypeID, map_Products.ManagingFirmID, map_Fees.BreakpointFrom, map_Fees.BreakpointTo, map_Fees.Fee, map_Fees.MaxRepFee,-1, 2" & _
            " FROM map_Fees INNER JOIN (map_Products INNER JOIN map_Agreements ON map_Products.Active = map_Agreements.AllInHouseSMAApproved) ON (map_Agreements.TypeID = map_Fees.AgreementTypeID) AND (map_Fees.ProductID = map_Products.ID)" & _
            " WHERE(((map_Products.ManagingFirmID) = 1) And ((map_Fees.Default) = True) And ((map_Fees.Active) = True) And ((map_Agreements.UseDefaultFee) = True) And ((map_Agreements.Active) = True) And ((map_Agreements.TypeID) = 2))" & _
            " GROUP BY map_Agreements.FirmID, map_Products.ID, map_Products.TypeOfProductID, map_Agreements.TypeID, map_Products.ManagingFirmID, map_Fees.BreakpointFrom, map_Fees.BreakpointTo, map_Fees.Fee, map_Fees.MaxRepFee;"

            Command = New OleDb.OleDbCommand(SQLstr, Mycn)
            Command.ExecuteNonQuery()
            pb.Value = pb.Value + 1
            Label2.Text = "Loaded..."
            Call LoadAllOutsideApprovedAgreementsDefaultFeeSOLICITED()

            Mycn.Close()

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "ERROR")
        End Try
    End Sub

    Public Sub LoadAllInhouseApprovedAgreementsDefaultFeeADVISED()
        'QID 13
        Label2.Text = "QID:13 - Looking for All InHouse Approved Default Fee..."
        Dim Mycn As OleDb.OleDbConnection
        Dim Command As OleDb.OleDbCommand
        Dim SQLstr As String
        Try
            Mycn = New OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & My.Settings.dbloc & "';Jet OLEDB:Database Password='" & My.Settings.dbpass & "';")
            Mycn.Open()

            SQLstr = "INSERT INTO map_TempAssignments([PlatformID], [WLID], [FirmID], [ProductID], [TypeOfProductID], [TypeID], [ManagingFirmID], [BreakpointFrom], [BreakpointTo], [Fee], [MaxRepFee], [SMAOffered], [QID])" & _
            " SELECT 7,47,map_Agreements.FirmID, map_Products.ID, map_Products.TypeOfProductID, map_Agreements.TypeID, map_Products.ManagingFirmID, map_Fees.BreakpointFrom, map_Fees.BreakpointTo, map_Fees.Fee, map_Fees.MaxRepFee,-1, 13" & _
            " FROM map_Fees INNER JOIN (map_Products INNER JOIN map_Agreements ON map_Products.Active = map_Agreements.AllInHouseSMAApproved) ON (map_Agreements.TypeID = map_Fees.AgreementTypeID) AND (map_Fees.ProductID = map_Products.ID)" & _
            " WHERE(((map_Products.ManagingFirmID) = 1) And ((map_Fees.Default) = True) And ((map_Fees.Active) = True) And ((map_Agreements.UseDefaultFee) = True) And ((map_Agreements.Active) = True) And ((map_Agreements.TypeID) = 3))" & _
            " GROUP BY map_Agreements.FirmID, map_Products.ID, map_Products.TypeOfProductID, map_Agreements.TypeID, map_Products.ManagingFirmID, map_Fees.BreakpointFrom, map_Fees.BreakpointTo, map_Fees.Fee, map_Fees.MaxRepFee;"

            Command = New OleDb.OleDbCommand(SQLstr, Mycn)
            Command.ExecuteNonQuery()
            pb.Value = pb.Value + 1
            Label2.Text = "Loaded..."
            Call LoadAllOutsideApprovedAgreementsDefaultFeeADVISED()

            Mycn.Close()

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "ERROR")
        End Try
    End Sub

    Public Sub LoadAllOutsideApprovedAgreementsDefaultFee()
        'QID 3
        Label2.Text = "QID:3 - Looking for All Outside Approved Default Fee..."
        Dim Mycn As OleDb.OleDbConnection
        Dim Command As OleDb.OleDbCommand
        Dim SQLstr As String
        Try
            Mycn = New OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & My.Settings.dbloc & "';Jet OLEDB:Database Password='" & My.Settings.dbpass & "';")
            Mycn.Open()

            SQLstr = "INSERT INTO map_TempAssignments([PlatformID], [WLID], [FirmID], [ProductID], [TypeOfProductID], [TypeID], [ManagingFirmID], [BreakpointFrom], [BreakpointTo], [Fee], [MaxRepFee], [SMAOffered], [QID])" & _
            " SELECT 7,47,map_Agreements.FirmID, map_Products.ID, map_Products.TypeOfProductID, map_Agreements.TypeID, map_Products.ManagingFirmID, map_Fees.BreakpointFrom, map_Fees.BreakpointTo, map_Fees.Fee, map_Fees.MaxRepFee,-1, 3" & _
            " FROM map_Agreements INNER JOIN (map_Fees INNER JOIN map_Products ON map_Fees.ProductID = map_Products.ID) ON (map_Products.Active = map_Agreements.AllOutsideSMAApproved) AND (map_Agreements.TypeID = map_Fees.AgreementTypeID)" & _
            " WHERE(((map_Products.ManagingFirmID) = 1) And ((map_Fees.Default) = True) And ((map_Fees.Active) = True) And ((map_Agreements.UseDefaultFee) = True) And ((map_Agreements.Active) = True) And ((map_Agreements.TypeID) = 1))" & _
            " GROUP BY map_Agreements.FirmID, map_Products.ID, map_Products.TypeOfProductID, map_Agreements.TypeID, map_Products.ManagingFirmID, map_Fees.BreakpointFrom, map_Fees.BreakpointTo, map_Fees.Fee, map_Fees.MaxRepFee;"

            Command = New OleDb.OleDbCommand(SQLstr, Mycn)
            Command.ExecuteNonQuery()
            pb.Value = pb.Value + 1
            Label2.Text = "Loaded..."
            Call LoadAllInhouseApprovedAgreementsCustomFee()

            Mycn.Close()

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "ERROR")
        End Try
    End Sub

    Public Sub LoadAllOutsideApprovedAgreementsDefaultFeeSOLICITED()
        'QID 4
        Label2.Text = "QID:4 - Looking for All Outside Approved Default Fee..."
        Dim Mycn As OleDb.OleDbConnection
        Dim Command As OleDb.OleDbCommand
        Dim SQLstr As String
        Try
            Mycn = New OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & My.Settings.dbloc & "';Jet OLEDB:Database Password='" & My.Settings.dbpass & "';")
            Mycn.Open()

            SQLstr = "INSERT INTO map_TempAssignments([PlatformID], [WLID], [FirmID], [ProductID], [TypeOfProductID], [TypeID], [ManagingFirmID], [BreakpointFrom], [BreakpointTo], [Fee], [MaxRepFee], [SMAOffered], [QID])" & _
            " SELECT 7,47,map_Agreements.FirmID, map_Products.ID, map_Products.TypeOfProductID, map_Agreements.TypeID, map_Products.ManagingFirmID, map_Fees.BreakpointFrom, map_Fees.BreakpointTo, map_Fees.Fee, map_Fees.MaxRepFee,-1, 4" & _
            " FROM map_Agreements INNER JOIN (map_Fees INNER JOIN map_Products ON map_Fees.ProductID = map_Products.ID) ON (map_Products.Active = map_Agreements.AllOutsideSMAApproved) AND (map_Agreements.TypeID = map_Fees.AgreementTypeID)" & _
            " WHERE(((map_Products.ManagingFirmID) = 1) And ((map_Fees.Default) = True) And ((map_Fees.Active) = True) And ((map_Agreements.UseDefaultFee) = True) And ((map_Agreements.Active) = True) And ((map_Agreements.TypeID) = 2))" & _
            " GROUP BY map_Agreements.FirmID, map_Products.ID, map_Products.TypeOfProductID, map_Agreements.TypeID, map_Products.ManagingFirmID, map_Fees.BreakpointFrom, map_Fees.BreakpointTo, map_Fees.Fee, map_Fees.MaxRepFee;"


            Command = New OleDb.OleDbCommand(SQLstr, Mycn)
            Command.ExecuteNonQuery()
            pb.Value = pb.Value + 1
            Label2.Text = "Loaded..."
            Call LoadAllInhouseApprovedAgreementsCustomFeeSOLICITED()

            Mycn.Close()

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "ERROR")
        End Try
    End Sub

    Public Sub LoadAllOutsideApprovedAgreementsDefaultFeeADVISED()
        'QID 14
        Label2.Text = "QID:14 - Looking for All Outside Approved Default Fee..."
        Dim Mycn As OleDb.OleDbConnection
        Dim Command As OleDb.OleDbCommand
        Dim SQLstr As String
        Try
            Mycn = New OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & My.Settings.dbloc & "';Jet OLEDB:Database Password='" & My.Settings.dbpass & "';")
            Mycn.Open()

            SQLstr = "INSERT INTO map_TempAssignments([PlatformID], [WLID], [FirmID], [ProductID], [TypeOfProductID], [TypeID], [ManagingFirmID], [BreakpointFrom], [BreakpointTo], [Fee], [MaxRepFee], [SMAOffered], [QID])" & _
            " SELECT 7,47,map_Agreements.FirmID, map_Products.ID, map_Products.TypeOfProductID, map_Agreements.TypeID, map_Products.ManagingFirmID, map_Fees.BreakpointFrom, map_Fees.BreakpointTo, map_Fees.Fee, map_Fees.MaxRepFee,-1, 14" & _
            " FROM map_Agreements INNER JOIN (map_Fees INNER JOIN map_Products ON map_Fees.ProductID = map_Products.ID) ON (map_Products.Active = map_Agreements.AllOutsideSMAApproved) AND (map_Agreements.TypeID = map_Fees.AgreementTypeID)" & _
            " WHERE(((map_Products.ManagingFirmID) = 1) And ((map_Fees.Default) = True) And ((map_Fees.Active) = True) And ((map_Agreements.UseDefaultFee) = True) And ((map_Agreements.Active) = True) And ((map_Agreements.TypeID) = 3))" & _
            " GROUP BY map_Agreements.FirmID, map_Products.ID, map_Products.TypeOfProductID, map_Agreements.TypeID, map_Products.ManagingFirmID, map_Fees.BreakpointFrom, map_Fees.BreakpointTo, map_Fees.Fee, map_Fees.MaxRepFee;"


            Command = New OleDb.OleDbCommand(SQLstr, Mycn)
            Command.ExecuteNonQuery()
            pb.Value = pb.Value + 1
            Label2.Text = "Loaded..."
            Call LoadAllInhouseApprovedAgreementsCustomFeeADVISED()

            Mycn.Close()

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "ERROR")
        End Try
    End Sub

    Public Sub LoadAllInhouseApprovedAgreementsCustomFee()
        'QID 5
        Label2.Text = "QID:5 - Looking for All InHouse Approved Custom Fee..."
        Dim Mycn As OleDb.OleDbConnection
        Dim Command As OleDb.OleDbCommand
        Dim SQLstr As String
        Try
            Mycn = New OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & My.Settings.dbloc & "';Jet OLEDB:Database Password='" & My.Settings.dbpass & "';")
            Mycn.Open()

            SQLstr = "INSERT INTO map_TempAssignments([PlatformID], [WLID], [FirmID], [ProductID], [TypeOfProductID], [TypeID], [ManagingFirmID], [BreakpointFrom], [BreakpointTo], [Fee], [MaxRepFee], [SMAOffered], [QID])" & _
            " SELECT 7,47,map_Agreements.FirmID,map_Products.ID, map_Products.TypeOfProductID, map_Fees.AgreementTypeID, map_Products.ManagingFirmID, map_Fees.BreakpointFrom, map_Fees.BreakpointTo, map_Fees.Fee, map_Fees.MaxRepFee,-1, 5" & _
            " FROM (map_Fees INNER JOIN map_Agreements ON map_Fees.AgreementID = map_Agreements.ID) INNER JOIN map_Products ON map_Fees.ProductID = map_Products.ID" & _
            " WHERE(((map_Fees.Active) = True) And ((map_Agreements.UseDefaultFee) = False)  And ((map_Agreements.TypeID) = 1) And ((map_Agreements.AllInHouseSMAApproved) = True) And map_Products.ManagingFirmID = 1 And ((map_Agreements.Active) = True))" & _
            " GROUP BY map_Agreements.FirmID, map_Products.ID, map_Products.TypeOfProductID, map_Fees.AgreementTypeID, map_Products.ManagingFirmID, map_Fees.BreakpointFrom, map_Fees.BreakpointTo, map_Fees.Fee, map_Fees.MaxRepFee"

            Command = New OleDb.OleDbCommand(SQLstr, Mycn)
            Command.ExecuteNonQuery()
            pb.Value = pb.Value + 1
            Label2.Text = "Loaded..."
            Call LoadAllOutsideApprovedAgreementsCustomFee()

            Mycn.Close()

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "ERROR")
        End Try
    End Sub

    Public Sub LoadAllInhouseApprovedAgreementsCustomFeeSOLICITED()
        'QID 6
        Label2.Text = "QID:6 - Looking for All InHouse Approved Custom Fee..."
        Dim Mycn As OleDb.OleDbConnection
        Dim Command As OleDb.OleDbCommand
        Dim SQLstr As String
        Try
            Mycn = New OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & My.Settings.dbloc & "';Jet OLEDB:Database Password='" & My.Settings.dbpass & "';")
            Mycn.Open()

            SQLstr = "INSERT INTO map_TempAssignments([PlatformID], [WLID], [FirmID], [ProductID], [TypeOfProductID], [TypeID], [ManagingFirmID], [BreakpointFrom], [BreakpointTo], [Fee], [MaxRepFee], [SMAOffered], [QID])" & _
            " SELECT 7,47,map_Agreements.FirmID,map_Products.ID, map_Products.TypeOfProductID, map_Fees.AgreementTypeID, map_Products.ManagingFirmID, map_Fees.BreakpointFrom, map_Fees.BreakpointTo, map_Fees.Fee, map_Fees.MaxRepFee,-1, 6" & _
            " FROM (map_Fees INNER JOIN map_Agreements ON map_Fees.AgreementID = map_Agreements.ID) INNER JOIN map_Products ON map_Fees.ProductID = map_Products.ID" & _
            " WHERE(((map_Fees.Active) = True) And ((map_Agreements.UseDefaultFee) = False)  And ((map_Agreements.TypeID) = 2) And ((map_Agreements.AllInHouseSMAApproved) = True) And map_Products.ManagingFirmID = 1 And ((map_Agreements.Active) = True))" & _
            " GROUP BY map_Agreements.FirmID, map_Products.ID, map_Products.TypeOfProductID, map_Fees.AgreementTypeID, map_Products.ManagingFirmID, map_Fees.BreakpointFrom, map_Fees.BreakpointTo, map_Fees.Fee, map_Fees.MaxRepFee"


            Command = New OleDb.OleDbCommand(SQLstr, Mycn)
            Command.ExecuteNonQuery()
            pb.Value = pb.Value + 1
            Label2.Text = "Loaded..."
            Call LoadAllOutsideApprovedAgreementsCustomFeeSOLICITED()

            Mycn.Close()

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "ERROR")
        End Try
    End Sub

    Public Sub LoadAllInhouseApprovedAgreementsCustomFeeADVISED()
        'QID 15
        Label2.Text = "QID:15 - Looking for All InHouse Approved Custom Fee..."
        Dim Mycn As OleDb.OleDbConnection
        Dim Command As OleDb.OleDbCommand
        Dim SQLstr As String
        Try
            Mycn = New OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & My.Settings.dbloc & "';Jet OLEDB:Database Password='" & My.Settings.dbpass & "';")
            Mycn.Open()

            SQLstr = "INSERT INTO map_TempAssignments([PlatformID], [WLID], [FirmID], [ProductID], [TypeOfProductID], [TypeID], [ManagingFirmID], [BreakpointFrom], [BreakpointTo], [Fee], [MaxRepFee], [SMAOffered], [QID])" & _
            " SELECT 7,47,map_Agreements.FirmID,map_Products.ID, map_Products.TypeOfProductID, map_Fees.AgreementTypeID, map_Products.ManagingFirmID, map_Fees.BreakpointFrom, map_Fees.BreakpointTo, map_Fees.Fee, map_Fees.MaxRepFee,-1, 15" & _
            " FROM (map_Fees INNER JOIN map_Agreements ON map_Fees.AgreementID = map_Agreements.ID) INNER JOIN map_Products ON map_Fees.ProductID = map_Products.ID" & _
            " WHERE(((map_Fees.Active) = True) And ((map_Agreements.UseDefaultFee) = False)  And ((map_Agreements.TypeID) = 3) And ((map_Agreements.AllInHouseSMAApproved) = True) And map_Products.ManagingFirmID = 1 And ((map_Agreements.Active) = True))" & _
            " GROUP BY map_Agreements.FirmID, map_Products.ID, map_Products.TypeOfProductID, map_Fees.AgreementTypeID, map_Products.ManagingFirmID, map_Fees.BreakpointFrom, map_Fees.BreakpointTo, map_Fees.Fee, map_Fees.MaxRepFee"


            Command = New OleDb.OleDbCommand(SQLstr, Mycn)
            Command.ExecuteNonQuery()
            pb.Value = pb.Value + 1
            Label2.Text = "Loaded..."
            Call LoadAllOutsideApprovedAgreementsCustomFeeADVISED()

            Mycn.Close()

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "ERROR")
        End Try
    End Sub

    Public Sub LoadAllOutsideApprovedAgreementsCustomFee()
        'QID 7
        Label2.Text = "QID:7 - Looking for All Outside Approved Custom Fee..."
        Dim Mycn As OleDb.OleDbConnection
        Dim Command As OleDb.OleDbCommand
        Dim SQLstr As String
        Try
            Mycn = New OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & My.Settings.dbloc & "';Jet OLEDB:Database Password='" & My.Settings.dbpass & "';")
            Mycn.Open()

            SQLstr = "INSERT INTO map_TempAssignments([PlatformID], [WLID], [FirmID], [ProductID], [TypeOfProductID], [TypeID], [ManagingFirmID], [BreakpointFrom], [BreakpointTo], [Fee], [MaxRepFee], [SMAOffered], [QID])" & _
            " SELECT 7,47,map_Agreements.FirmID,map_Products.ID, map_Products.TypeOfProductID, map_Fees.AgreementTypeID, map_Products.ManagingFirmID, map_Fees.BreakpointFrom, map_Fees.BreakpointTo, map_Fees.Fee, map_Fees.MaxRepFee,-1, 7" & _
            " FROM (map_Fees INNER JOIN map_Agreements ON map_Fees.AgreementID = map_Agreements.ID) INNER JOIN map_Products ON map_Fees.ProductID = map_Products.ID" & _
            " WHERE(((map_Fees.Active) = True) And ((map_Agreements.UseDefaultFee) = False)  And ((map_Agreements.TypeID) = 1) And ((map_Agreements.AllInHouseSMAApproved) = True) And map_Products.ManagingFirmID <> 1 And ((map_Agreements.TypeID) = 1) And ((map_Agreements.Active) = True))" & _
            " GROUP BY map_Agreements.FirmID, map_Products.ID, map_Products.TypeOfProductID, map_Fees.AgreementTypeID, map_Products.ManagingFirmID, map_Fees.BreakpointFrom, map_Fees.BreakpointTo, map_Fees.Fee, map_Fees.MaxRepFee"

            Command = New OleDb.OleDbCommand(SQLstr, Mycn)
            Command.ExecuteNonQuery()
            pb.Value = pb.Value + 1
            Label2.Text = "Loaded..."
            Call LoadCustomApprovedAgreementsCustomFee()

            Mycn.Close()

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "ERROR")
        End Try
    End Sub

    Public Sub LoadAllOutsideApprovedAgreementsCustomFeeSOLICITED()
        'QID 8
        Label2.Text = "QID:8 - Looking for All Outside Approved Custom Fee..."
        Dim Mycn As OleDb.OleDbConnection
        Dim Command As OleDb.OleDbCommand
        Dim SQLstr As String
        Try
            Mycn = New OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & My.Settings.dbloc & "';Jet OLEDB:Database Password='" & My.Settings.dbpass & "';")
            Mycn.Open()

            SQLstr = "INSERT INTO map_TempAssignments([PlatformID], [WLID], [FirmID], [ProductID], [TypeOfProductID], [TypeID], [ManagingFirmID], [BreakpointFrom], [BreakpointTo], [Fee], [MaxRepFee], [SMAOffered], [QID])" & _
            " SELECT 7,47,map_Agreements.FirmID,map_Products.ID, map_Products.TypeOfProductID, map_Fees.AgreementTypeID, map_Products.ManagingFirmID, map_Fees.BreakpointFrom, map_Fees.BreakpointTo, map_Fees.Fee, map_Fees.MaxRepFee,-1, 8" & _
            " FROM (map_Fees INNER JOIN map_Agreements ON map_Fees.AgreementID = map_Agreements.ID) INNER JOIN map_Products ON map_Fees.ProductID = map_Products.ID" & _
            " WHERE(((map_Fees.Active) = True) And ((map_Agreements.UseDefaultFee) = False)  And ((map_Agreements.TypeID) = 2) And ((map_Agreements.AllInHouseSMAApproved) = True) And map_Products.ManagingFirmID <> 1 And ((map_Agreements.TypeID) = 2) And ((map_Agreements.Active) = True))" & _
            " GROUP BY map_Agreements.FirmID, map_Products.ID, map_Products.TypeOfProductID, map_Fees.AgreementTypeID, map_Products.ManagingFirmID, map_Fees.BreakpointFrom, map_Fees.BreakpointTo, map_Fees.Fee, map_Fees.MaxRepFee"


            Command = New OleDb.OleDbCommand(SQLstr, Mycn)
            Command.ExecuteNonQuery()
            pb.Value = pb.Value + 1
            Label2.Text = "Loaded..."
            Call LoadCustomApprovedAgreementsCustomFeeSOLICITED()

            Mycn.Close()

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "ERROR")
        End Try
    End Sub

    Public Sub LoadAllOutsideApprovedAgreementsCustomFeeADVISED()
        'QID 16
        Label2.Text = "QID:16 - Looking for All Outside Approved Custom Fee..."
        Dim Mycn As OleDb.OleDbConnection
        Dim Command As OleDb.OleDbCommand
        Dim SQLstr As String
        Try
            Mycn = New OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & My.Settings.dbloc & "';Jet OLEDB:Database Password='" & My.Settings.dbpass & "';")
            Mycn.Open()

            SQLstr = "INSERT INTO map_TempAssignments([PlatformID], [WLID], [FirmID], [ProductID], [TypeOfProductID], [TypeID], [ManagingFirmID], [BreakpointFrom], [BreakpointTo], [Fee], [MaxRepFee], [SMAOffered], [QID])" & _
            " SELECT 7,47,map_Agreements.FirmID,map_Products.ID, map_Products.TypeOfProductID, map_Fees.AgreementTypeID, map_Products.ManagingFirmID, map_Fees.BreakpointFrom, map_Fees.BreakpointTo, map_Fees.Fee, map_Fees.MaxRepFee,-1, 16" & _
            " FROM (map_Fees INNER JOIN map_Agreements ON map_Fees.AgreementID = map_Agreements.ID) INNER JOIN map_Products ON map_Fees.ProductID = map_Products.ID" & _
            " WHERE(((map_Fees.Active) = True) And ((map_Agreements.UseDefaultFee) = False)  And ((map_Agreements.TypeID) = 3) And ((map_Agreements.AllInHouseSMAApproved) = True) And map_Products.ManagingFirmID <> 1 And ((map_Agreements.TypeID) = 3) And ((map_Agreements.Active) = True))" & _
            " GROUP BY map_Agreements.FirmID, map_Products.ID, map_Products.TypeOfProductID, map_Fees.AgreementTypeID, map_Products.ManagingFirmID, map_Fees.BreakpointFrom, map_Fees.BreakpointTo, map_Fees.Fee, map_Fees.MaxRepFee"


            Command = New OleDb.OleDbCommand(SQLstr, Mycn)
            Command.ExecuteNonQuery()
            pb.Value = pb.Value + 1
            Label2.Text = "Loaded..."
            Call LoadCustomApprovedAgreementsCustomFeeADVISED()

            Mycn.Close()

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "ERROR")
        End Try
    End Sub

    Public Sub LoadCustomApprovedAgreementsCustomFee()
        'QID 9
        Label2.Text = "QID:9 - Looking for Custom Approved Custom Fee..."
        Dim Mycn As OleDb.OleDbConnection
        Dim Command As OleDb.OleDbCommand
        Dim SQLstr As String
        Try
            Mycn = New OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & My.Settings.dbloc & "';Jet OLEDB:Database Password='" & My.Settings.dbpass & "';")
            Mycn.Open()

            SQLstr = "INSERT INTO map_TempAssignments([PlatformID], [WLID], [FirmID], [ProductID], [TypeOfProductID], [TypeID], [ManagingFirmID], [BreakpointFrom], [BreakpointTo], [Fee], [MaxRepFee], [SMAOffered], [QID])" & _
            " SELECT 7,47,map_Agreements.FirmID,map_Products.ID, map_Products.TypeOfProductID, map_Fees.AgreementTypeID, map_Products.ManagingFirmID, map_Fees.BreakpointFrom, map_Fees.BreakpointTo, map_Fees.Fee, map_Fees.MaxRepFee,-1, 9" & _
            " FROM ((map_Fees INNER JOIN map_ProductListing ON map_Fees.ProductListingID = map_ProductListing.ID) INNER JOIN map_Products ON map_Fees.ProductID = map_Products.ID) INNER JOIN map_Agreements ON map_ProductListing.AgreementID = map_Agreements.ID" & _
            " WHERE(((map_Fees.Active) = True) And ((map_Agreements.Active) = True) And ((map_Agreements.UseDefaultFee) = False) And ((map_Agreements.AllInHouseSMAApproved) = False) And ((map_Agreements.AllOutsideSMAApproved) = False) And ((map_Agreements.Active) = True) And ((map_Agreements.TypeID) = 1))" & _
            " GROUP BY map_Agreements.FirmID, map_Products.ID, map_Products.TypeOfProductID, map_Fees.AgreementTypeID, map_Products.ManagingFirmID, map_Fees.BreakpointFrom, map_Fees.BreakpointTo, map_Fees.Fee, map_Fees.MaxRepFee;"

            Command = New OleDb.OleDbCommand(SQLstr, Mycn)
            Command.ExecuteNonQuery()
            pb.Value = pb.Value + 1
            Label2.Text = "Loaded..."
            Call LoadAllCustomAgreementsDefualtFee()

            Mycn.Close()

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "ERROR")
        End Try
    End Sub

    Public Sub LoadCustomApprovedAgreementsCustomFeeSOLICITED()
        'QID 10
        Label2.Text = "QID:10 - Looking for Custom Approved Custom Fee..."
        Dim Mycn As OleDb.OleDbConnection
        Dim Command As OleDb.OleDbCommand
        Dim SQLstr As String
        Try
            Mycn = New OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & My.Settings.dbloc & "';Jet OLEDB:Database Password='" & My.Settings.dbpass & "';")
            Mycn.Open()

            SQLstr = "INSERT INTO map_TempAssignments([PlatformID], [WLID], [FirmID], [ProductID], [TypeOfProductID], [TypeID], [ManagingFirmID], [BreakpointFrom], [BreakpointTo], [Fee], [MaxRepFee], [SMAOffered], [QID])" & _
            " SELECT 7,47,map_Agreements.FirmID,map_Products.ID, map_Products.TypeOfProductID, map_Fees.AgreementTypeID, map_Products.ManagingFirmID, map_Fees.BreakpointFrom, map_Fees.BreakpointTo, map_Fees.Fee, map_Fees.MaxRepFee,-1, 10" & _
            " FROM ((map_Fees INNER JOIN map_ProductListing ON map_Fees.ProductListingID = map_ProductListing.ID) INNER JOIN map_Products ON map_Fees.ProductID = map_Products.ID) INNER JOIN map_Agreements ON map_ProductListing.AgreementID = map_Agreements.ID" & _
            " WHERE(((map_Fees.Active) = True) And ((map_Agreements.Active) = True) And ((map_Agreements.UseDefaultFee) = False) And ((map_Agreements.AllInHouseSMAApproved) = False) And ((map_Agreements.AllOutsideSMAApproved) = False) And ((map_Agreements.Active) = True) And ((map_Agreements.TypeID) = 2))" & _
            " GROUP BY map_Agreements.FirmID, map_Products.ID, map_Products.TypeOfProductID, map_Fees.AgreementTypeID, map_Products.ManagingFirmID, map_Fees.BreakpointFrom, map_Fees.BreakpointTo, map_Fees.Fee, map_Fees.MaxRepFee;"


            Command = New OleDb.OleDbCommand(SQLstr, Mycn)
            Command.ExecuteNonQuery()
            pb.Value = pb.Value + 1
            Label2.Text = "Loaded..."
            Call LoadAllCustomAgreementsDefualtFeeSOLICITED()

            Mycn.Close()

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "ERROR")
        End Try
    End Sub

    Public Sub LoadCustomApprovedAgreementsCustomFeeADVISED()
        'QID 17
        Label2.Text = "QID:17 - Looking for Custom Approved Custom Fee..."
        Dim Mycn As OleDb.OleDbConnection
        Dim Command As OleDb.OleDbCommand
        Dim SQLstr As String
        Try
            Mycn = New OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & My.Settings.dbloc & "';Jet OLEDB:Database Password='" & My.Settings.dbpass & "';")
            Mycn.Open()

            SQLstr = "INSERT INTO map_TempAssignments([PlatformID], [WLID], [FirmID], [ProductID], [TypeOfProductID], [TypeID], [ManagingFirmID], [BreakpointFrom], [BreakpointTo], [Fee], [MaxRepFee], [SMAOffered], [QID])" & _
            " SELECT 7,47,map_Agreements.FirmID, map_Products.ID, map_Products.TypeOfProductID, map_Fees.AgreementTypeID, map_Products.ManagingFirmID, map_Fees.BreakpointFrom, map_Fees.BreakpointTo, map_Fees.Fee, map_Fees.MaxRepFee,-1, 17" & _
            " FROM ((map_Fees INNER JOIN map_ProductListing ON map_Fees.ProductListingID = map_ProductListing.ID) INNER JOIN map_Products ON map_Fees.ProductID = map_Products.ID) INNER JOIN map_Agreements ON map_ProductListing.AgreementID = map_Agreements.ID" & _
            " WHERE(((map_Fees.Active) = True) And ((map_Agreements.Active) = True) And ((map_Agreements.UseDefaultFee) = False) And ((map_Agreements.AllInHouseSMAApproved) = False) And ((map_Agreements.AllOutsideSMAApproved) = False) And ((map_Agreements.Active) = True) And ((map_Agreements.TypeID) = 3))" & _
            " GROUP BY map_Agreements.FirmID, map_Products.ID, map_Products.TypeOfProductID, map_Fees.AgreementTypeID, map_Products.ManagingFirmID, map_Fees.BreakpointFrom, map_Fees.BreakpointTo, map_Fees.Fee, map_Fees.MaxRepFee;"


            Command = New OleDb.OleDbCommand(SQLstr, Mycn)
            Command.ExecuteNonQuery()
            pb.Value = pb.Value + 1
            Label2.Text = "Loaded..."
            Call LoadAllCustomAgreementsDefualtFeeADVISED()

            Mycn.Close()

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "ERROR")
        End Try
    End Sub

    Public Sub LoadAllCustomAgreementsDefualtFee()
        'QID 11
        Label2.Text = "QID:11 - Looking for Custom Approved Default Fee..."
        Dim Mycn As OleDb.OleDbConnection
        Dim Command As OleDb.OleDbCommand
        Dim SQLstr As String
        Try
            Mycn = New OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & My.Settings.dbloc & "';Jet OLEDB:Database Password='" & My.Settings.dbpass & "';")
            Mycn.Open()

            SQLstr = "INSERT INTO map_TempAssignments([PlatformID], [WLID], [FirmID], [ProductID], [TypeOfProductID], [TypeID], [ManagingFirmID], [BreakpointFrom], [BreakpointTo], [Fee], [MaxRepFee], [SMAOffered], [QID])" & _
            " SELECT 7,47,map_Agreements.FirmID, map_Products.ID, map_Products.TypeOfProductID, map_Fees.AgreementTypeID, map_Products.ManagingFirmID, map_Fees.BreakpointFrom, map_Fees.BreakpointTo, map_Fees.Fee, map_Fees.MaxRepFee,-1, 11" & _
            " FROM (map_ProductListing INNER JOIN map_Agreements ON map_ProductListing.AgreementID = map_Agreements.ID) INNER JOIN (map_Fees INNER JOIN map_Products ON map_Fees.ProductID = map_Products.ID) ON (map_Agreements.TypeID = map_Fees.AgreementTypeID) AND (map_ProductListing.ProductID = map_Products.ID)" & _
            " WHERE(((map_Fees.Active) = True) And ((map_Agreements.Active) = True) And ((map_Agreements.UseDefaultFee) = True) And ((map_Agreements.AllInHouseSMAApproved) = False) And ((map_Agreements.AllOutsideSMAApproved) = False) And ((map_Agreements.AllMFApproved) = False) And ((map_Agreements.AllUITApproved) = False) And ((map_Agreements.TypeID) = 1) And ((map_Fees.Default) = True))" & _
            " GROUP BY map_Agreements.FirmID, map_Products.ID, map_Products.TypeOfProductID, map_Fees.AgreementTypeID, map_Products.ManagingFirmID, map_Fees.BreakpointFrom, map_Fees.BreakpointTo, map_Fees.Fee, map_Fees.MaxRepFee;"

            Command = New OleDb.OleDbCommand(SQLstr, Mycn)
            Command.ExecuteNonQuery()
            pb.Value = pb.Value + 1
            Label2.Text = "Loaded..."

            Mycn.Close()

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "ERROR")
        End Try
    End Sub

    Public Sub LoadAllCustomAgreementsDefualtFeeSOLICITED()
        'QID 12
        Label2.Text = "QID:12 - Looking for Custom Approved Default Fee..."
        Dim Mycn As OleDb.OleDbConnection
        Dim Command As OleDb.OleDbCommand
        Dim SQLstr As String
        Try
            Mycn = New OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & My.Settings.dbloc & "';Jet OLEDB:Database Password='" & My.Settings.dbpass & "';")
            Mycn.Open()

            SQLstr = "INSERT INTO map_TempAssignments([PlatformID], [WLID], [FirmID], [ProductID], [TypeOfProductID], [TypeID], [ManagingFirmID], [BreakpointFrom], [BreakpointTo], [Fee], [MaxRepFee], [SMAOffered], [QID])" & _
            " SELECT 7,47,map_Agreements.FirmID, map_Products.ID, map_Products.TypeOfProductID, map_Fees.AgreementTypeID, map_Products.ManagingFirmID, map_Fees.BreakpointFrom, map_Fees.BreakpointTo, map_Fees.Fee, map_Fees.MaxRepFee,-1, 12" & _
            " FROM (map_ProductListing INNER JOIN map_Agreements ON map_ProductListing.AgreementID = map_Agreements.ID) INNER JOIN (map_Fees INNER JOIN map_Products ON map_Fees.ProductID = map_Products.ID) ON (map_Agreements.TypeID = map_Fees.AgreementTypeID) AND (map_ProductListing.ProductID = map_Products.ID)" & _
            " WHERE(((map_Fees.Active) = True) And ((map_Agreements.Active) = True) And ((map_Agreements.UseDefaultFee) = True) And ((map_Agreements.AllInHouseSMAApproved) = False) And ((map_Agreements.AllOutsideSMAApproved) = False) And ((map_Agreements.AllMFApproved) = False) And ((map_Agreements.AllUITApproved) = False) And ((map_Agreements.TypeID) = 2) And ((map_Fees.Default) = True))" & _
            " GROUP BY map_Agreements.FirmID, map_Products.ID, map_Products.TypeOfProductID, map_Fees.AgreementTypeID, map_Products.ManagingFirmID, map_Fees.BreakpointFrom, map_Fees.BreakpointTo, map_Fees.Fee, map_Fees.MaxRepFee;"

            '" FROM (map_ProductListing INNER JOIN map_Agreements ON map_ProductListing.AgreementID = map_Agreements.ID) INNER JOIN (map_Fees INNER JOIN map_Products ON map_Fees.ProductID = map_Products.ID) ON map_ProductListing.ProductID = map_Products.ID" & _

            Command = New OleDb.OleDbCommand(SQLstr, Mycn)
            Command.ExecuteNonQuery()
            pb.Value = pb.Value + 1
            Label2.Text = "Loaded..."

            Mycn.Close()

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "ERROR")
        End Try
    End Sub

    Public Sub LoadAllCustomAgreementsDefualtFeeADVISED()
        'QID 18
        Label2.Text = "QID:18 - Looking for Custom Approved Default Fee..."
        Dim Mycn As OleDb.OleDbConnection
        Dim Command As OleDb.OleDbCommand
        Dim SQLstr As String
        Try
            Mycn = New OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & My.Settings.dbloc & "';Jet OLEDB:Database Password='" & My.Settings.dbpass & "';")
            Mycn.Open()

            SQLstr = "INSERT INTO map_TempAssignments([PlatformID], [WLID], [FirmID], [ProductID], [TypeOfProductID], [TypeID], [ManagingFirmID], [BreakpointFrom], [BreakpointTo], [Fee], [MaxRepFee], [SMAOffered], [QID])" & _
            " SELECT 7,47,map_Agreements.FirmID, map_Products.ID, map_Products.TypeOfProductID, map_Fees.AgreementTypeID, map_Products.ManagingFirmID, map_Fees.BreakpointFrom, map_Fees.BreakpointTo, map_Fees.Fee, map_Fees.MaxRepFee,-1, 18" & _
            " FROM (map_ProductListing INNER JOIN map_Agreements ON map_ProductListing.AgreementID = map_Agreements.ID) INNER JOIN (map_Fees INNER JOIN map_Products ON map_Fees.ProductID = map_Products.ID) ON (map_Agreements.TypeID = map_Fees.AgreementTypeID) AND (map_ProductListing.ProductID = map_Products.ID)" & _
            " WHERE(((map_Fees.Active) = True) And ((map_Agreements.Active) = True) And ((map_Agreements.UseDefaultFee) = True) And ((map_Agreements.AllInHouseSMAApproved) = False) And ((map_Agreements.AllOutsideSMAApproved) = False) And ((map_Agreements.AllMFApproved) = False) And ((map_Agreements.AllUITApproved) = False) And ((map_Agreements.TypeID) = 3) And ((map_Fees.Default) = True))" & _
            " GROUP BY map_Agreements.FirmID,map_Products.ID, map_Products.TypeOfProductID, map_Fees.AgreementTypeID, map_Products.ManagingFirmID, map_Fees.BreakpointFrom, map_Fees.BreakpointTo, map_Fees.Fee, map_Fees.MaxRepFee;"

            '" FROM (map_ProductListing INNER JOIN map_Agreements ON map_ProductListing.AgreementID = map_Agreements.ID) INNER JOIN (map_Fees INNER JOIN map_Products ON map_Fees.ProductID = map_Products.ID) ON map_ProductListing.ProductID = map_Products.ID" & _

            Command = New OleDb.OleDbCommand(SQLstr, Mycn)
            Command.ExecuteNonQuery()
            pb.Value = pb.Value + 1
            Label2.Text = "Loaded..."

            Mycn.Close()

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "ERROR")
        End Try
    End Sub

    Private Sub Timer2_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer2.Tick
        If CheckBox1.Checked Then
            Me.Close()
        Else

        End If
    End Sub
End Class