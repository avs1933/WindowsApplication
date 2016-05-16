Public Class CommCalc
    'Thread1 controls At-a-Glance
    Dim thread1 As System.Threading.Thread

    Private Sub Timer1_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer1.Tick
        Call docalc()
    End Sub

    Public Sub docalc()

        Dim tmv As Double
        Dim fee1 As Double
        Dim fee2 As Double

        tmv = TextBox1.Text
        fee1 = ComboBox1.SelectedItem
        fee2 = ComboBox2.SelectedItem

        Dim ourfee As Double
        Dim extfee As Double

        ourfee = ((tmv * fee1) / 100)
        extfee = ((tmv * fee2) / 100)

        Dim totfee As Double

        totfee = ourfee + extfee

        TextBox3.Text = Format(ourfee, "$#,###.00")
        TextBox4.Text = Format(extfee, "$#,###.00")
        TextBox2.Text = Format(totfee, "$#,###.00")

        Dim perfee1 As Double
        Dim perfee2 As Double
        Dim totfee1 As Double

        perfee1 = ourfee / tmv
        perfee2 = extfee / tmv
        totfee1 = totfee / tmv

        TextBox6.Text = Format(perfee1, "##.00%")
        TextBox5.Text = Format(perfee2, "##.00%")
        TextBox7.Text = Format(totfee1, "##.00%")

    End Sub

    Private Sub CommCalc_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        TextBox1.Text = "1,000,000.00"
        ComboBox1.SelectedItem = "0.45"
        ComboBox2.SelectedItem = "0.00"
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Call docalc()
    End Sub
End Class