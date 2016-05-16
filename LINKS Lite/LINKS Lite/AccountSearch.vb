Public Class AccountSearch
    Dim thread1 As System.Threading.Thread

    Private Sub DataGridView1_CellContentClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick

    End Sub

    Private Sub AccountSearch_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Me.Text = Application.ProductName & " Account Search"
    End Sub

    Private Sub cmdSearch_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdSearch.Click
        ToolStripProgressBar1.Enabled = True
        Control.CheckForIllegalCrossThreadCalls = False
        thread1 = New System.Threading.Thread(AddressOf LoadAccounts)
        thread1.Start()
    End Sub

    Public Sub LoadAccounts()

    End Sub
End Class