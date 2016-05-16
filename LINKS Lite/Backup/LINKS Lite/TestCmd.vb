Public Class TestCmd

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        'Shell("cmd.exe /k" + "CD C:\Program Files (x86)\Advent\ApxClient\4.0", "AdvScriptRunner ApxIX -i -a\\aamapxapps01\APX$\imp\ -Ara -f\\aamapxapps01\APX$\imp031814.pri -tcsv4 -u -N_268")
        'Shell("cmd.exe /k CD C:\")
        Dim p As New ProcessStartInfo
        p.FileName = "cmd.exe"
        p.Arguments = "CD C:\Program Files (x86)\Advent\ApxClient\4.0" & " " & "AdvScriptRunner ApxIX -i -a\\aamapxapps01\APX$\imp\ -Ara -f\\aamapxapps01\APX$\imp031814.pri -tcsv4 -u -N_268"
        p.WindowStyle = ProcessWindowStyle.Normal
        Process.Start(p)
        'End Using




        'Dim info As New System.Diagnostics.ProcessStartInfo
        'info.FileName = "cmd.exe"
        'info.Arguments = "ApxIX -i -a\\aamapxapps01\APX$\imp\ -Ara -f\\aamapxapps01\APX$\imp031814.pri -tcsv4 -u -N_268"
        'Dim process As New System.Diagnostics.Process
        'process.StartInfo = info
        'process.Start("cmd.exe", "CD C:\Program Files (x86)\Advent\ApxClient\4.0") ', "AdvScriptRunner ApxIX -i -a\\aamapxapps01\APX$\imp\ -Ara -f\\aamapxapps01\APX$\imp031814.pri -tcsv4 -u -N_268")
        'MessageBox.Show(info.Arguments.ToString())
        'process.Close()
    End Sub
End Class