Public Class MasterMapping

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Dim child As New mmc_Database
        child.MdiParent = Home
        child.Show()
        child.txtID.Text = "NEW"
    End Sub

    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        Dim child As New map_EditProductType
        child.MdiParent = Me.MdiParent
        child.Show()
        child.TypeID.Text = "NEW"
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        Dim child As New map_SMA
        child.MdiParent = Me.MdiParent
        child.Show()
        child.ProductID.Text = "NEW"
    End Sub

    Private Sub Button5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button5.Click
        Dim child As New map_ProductTypesEdit
        child.MdiParent = Home
        child.Show()
    End Sub

    Private Sub Button7_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button7.Click
        Dim child As New map_EditMFirms
        child.MdiParent = Home
        child.Show()
        child.ID.Text = "NEW"
    End Sub

    Private Sub Button6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button6.Click
        Dim child As New map_ViewFirms
        child.MdiParent = Home
        child.Show()
    End Sub

    Private Sub Button8_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button8.Click
        Dim child As New map_EditObjective
        child.MdiParent = Home
        child.Show()
        child.ID.Text = "NEW"
    End Sub

    Private Sub Button9_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button9.Click
        Dim child As New map_ViewObjectives
        child.MdiParent = Home
        child.Show()
    End Sub

    Private Sub Button10_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button10.Click
        Dim child As New map_ViewProducts
        child.MdiParent = Home
        child.Show()
    End Sub

    Private Sub Button11_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button11.Click
        Dim child As New map_EditAgreement
        child.MdiParent = Home
        child.Show()
        child.ID.Text = "NEW"
    End Sub

    Private Sub Button12_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button12.Click
        Dim child As New map_SelectAgreement
        child.MdiParent = Home
        child.Show()
    End Sub

    Private Sub Button13_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button13.Click
        Dim child As New map_EditPlatform
        child.MdiParent = Home
        child.Show()
        child.ID.Text = "NEW"
    End Sub

    Private Sub Button14_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button14.Click
        Dim child As New map_ViewPlatforms
        child.MdiParent = Home
        child.Show()
    End Sub

    Private Sub Button15_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button15.Click
        Dim child As New map_EditWLPlatform
        child.MdiParent = Home
        child.Show()
        child.ID.Text = "NEW"
    End Sub

    Private Sub Button16_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button16.Click
        Dim child As New map_ViewWLPlatform
        child.MdiParent = Home
        child.Show()
    End Sub

    Private Sub Button17_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button17.Click
        Dim child As New map_PlatformAssignment
        child.MdiParent = Home
        child.Show()
        child.ID.Text = "NEW"
    End Sub

    Private Sub Button18_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button18.Click
        Dim child As New map_ViewPlatformAssignments
        child.MdiParent = Home
        child.Show()
    End Sub

    Private Sub Button19_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button19.Click
        Dim child As New map_LoadFirmApprovals
        child.MdiParent = Home
        child.Show()
    End Sub

    Private Sub MasterMapping_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Call LoadPermissions()
    End Sub

    Public Sub LoadPermissions()
        If Permissions.MAPAddAgreement.Checked Then
            Button11.Enabled = True
        Else
            Button11.Enabled = False
        End If

        If Permissions.MAPAddFirm.Checked Then
            Button7.Enabled = True
        Else
            Button7.Enabled = False
        End If

        If Permissions.MAPAddProducts.Checked Then
            Button2.Enabled = True
        Else
            Button2.Enabled = False
        End If

        If Permissions.MAPAddProductType.Checked Then
            Button4.Enabled = True
        Else
            Button4.Enabled = False
        End If

        If Permissions.MAPAddObjective.Checked Then
            Button8.Enabled = True
        Else
            Button8.Enabled = False
        End If

        If Permissions.MAPEditProducts.Checked Then
            Button10.Enabled = True
        Else
            Button10.Enabled = False
        End If

        If Permissions.MAPEditProductTypes.Checked Then
            Button5.Enabled = True
        Else
            Button5.Enabled = False
        End If

        If Permissions.MAPEditFirms.Checked Then
            Button6.Enabled = True
        Else
            Button6.Enabled = False
        End If

        If Permissions.MAPEditAgreements.Checked Then
            Button12.Enabled = True
        Else
            Button12.Enabled = False
        End If

        If Permissions.MAPEditObjective.Checked Then
            Button9.Enabled = True
        Else
            Button9.Enabled = False
        End If

        If Permissions.MAPAddPlatform.Checked Then
            Button13.Enabled = True
        Else
            Button13.Enabled = False
        End If

        If Permissions.MAPEditPlatform.Checked Then
            Button14.Enabled = True
        Else
            Button14.Enabled = False
        End If

        If Permissions.MAPAddWLPlatform.Checked Then
            Button15.Enabled = True
        Else
            Button15.Enabled = False
        End If

        If Permissions.MAPEditWLPlatform.Checked Then
            Button16.Enabled = True
        Else
            Button16.Enabled = False
        End If

        If Permissions.MAPAssociatePlatform.Checked Then
            Button17.Enabled = True
        Else
            Button17.Enabled = False
        End If

        If Permissions.MAPAddDatabase.Checked Then
            Button1.Enabled = True
        Else
            Button1.Enabled = False
        End If

        If Permissions.MAPRefreshApprovals.Checked Then
            Button19.Enabled = True
        Else
            Button19.Enabled = False
        End If
    End Sub
End Class