Public Class WhatsNew

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Call LoadGrid()
    End Sub

    Public Sub LoadGrid()
        Try

            Dim strSQL As String
            Dim conn As New OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & My.Settings.dbloc & "';Jet OLEDB:Database Password='" & My.Settings.dbpass & "';")

            strSQL = "SELECT sys_SoftwareUpdates.ID, sys_SoftwareUpdates.MainFunction AS Function, sys_SoftwareUpdates.ChangesMade AS [Changes Made], sys_Users.FullName AS [Made By], sys_SoftwareUpdates.DateOfRelease AS [Date of Relese]" & _
            " FROM sys_SoftwareUpdates INNER JOIN sys_Users ON sys_SoftwareUpdates.MadeBy = sys_Users.ID" & _
            " WHERE VersionNumber = '" & My.Application.Info.Version.Major & "." & My.Application.Info.Version.Minor & "." & My.Application.Info.Version.Build & "." & My.Application.Info.Version.Revision & "'" & _
            " GROUP BY sys_SoftwareUpdates.ID, sys_SoftwareUpdates.MainFunction, sys_SoftwareUpdates.ChangesMade, sys_Users.FullName, sys_SoftwareUpdates.DateOfRelease;"

            Dim da As New OleDb.OleDbDataAdapter(strSQL, conn)
            Dim ds As New DataSet
            da.Fill(ds, "Users")

            With DataGridView1
                .DataSource = ds.Tables("Users")
                .Columns(0).Visible = False
            End With

            lblEnhancements.Text = "Found " & DataGridView1.RowCount & " enhancements."

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "ERROR")
        End Try
    End Sub

    Private Sub btnReset_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub

    Private Sub WhatsNew_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        txtHeader.Text = "See whats new in the " & My.Application.Info.ProductName & ", Version Number: " & My.Application.Info.Version.Major & "." & My.Application.Info.Version.Minor & "." & My.Application.Info.Version.Build & "."
        Call LoadGrid()

    End Sub

    Private Sub DataGridView1_CellContentClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick

    End Sub

    Private Sub DataGridView1_CellDoubleClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView1.CellDoubleClick
        txtID.Text = DataGridView1.SelectedCells(0).Value
        txtFunction.Text = DataGridView1.SelectedCells(1).Value
        rtbChanges.Text = DataGridView1.SelectedCells(2).Value
        txtMadeBy.Text = DataGridView1.SelectedCells(3).Value
        dtpRelease.Text = DataGridView1.SelectedCells(4).Value
    End Sub
End Class