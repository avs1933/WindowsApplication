Public Class UMATradeEdit

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Dim ir As MsgBoxResult
        ir = MsgBox("You are about to modify this trade.  Are you sure you want to continue?", MsgBoxStyle.YesNo + MsgBoxStyle.Critical, "Confirm")
        If ir = MsgBoxResult.Yes Then

            Call CheckChange()
            Call UpdateTrade()
            'Dim chd As UMA.MdiParent = Home
            UMA.MdiParent = Home

            Call UMA.LoadPendingTradesGrid()

            'MdiParent.MdiChildren.UMA.LoadPendingTradesGrid()
            Me.Close()
        Else

        End If
    End Sub

    Public Sub UpdateTrade()
        Dim Mycn As OleDb.OleDbConnection
        Dim Command As OleDb.OleDbCommand
        Dim SQLstr As String
        Try
            Mycn = New OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & My.Settings.dbloc & "';Jet OLEDB:Database Password='" & My.Settings.dbpass & "';")
            Dim ds1 As New DataSet
            Dim eds1 As New DataGridView
            Dim dv1 As New DataView
            Mycn.Open()

            SQLstr = "Update uma_envestnet_TradePending SET Field1 = '" & Field1.Text & "',Field2 = '" & Field2.Text & "',Field3 = '" & Field3.Text & "',Field4 = '" & Field4.Text & "',Field5 = '" & Field5.Text & "',Field6 = '" & Field6.Text & "',Field7 = '" & Field7.Text & "',Field8 = '" & Field8.Text & "',Field9 = '" & Field9.Text & "',Field10 = '" & Field10.Text & "',Field11 = '" & Field11.Text & "',Field12 = '" & Field12.Text & "',Field13 = '" & Field13.Text & "',Field14 = '" & Field14.Text & "',Field15 = '" & Field15.Text & "',Field16 = '" & Field16.Text & "',Field17 = '" & Field17.Text & "',Field18 = '" & Field18.Text & "' WHERE ID = " & ID.Text

            Command = New OleDb.OleDbCommand(SQLstr, Mycn)
            Command.ExecuteNonQuery()



            Mycn.Close()

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "ERROR")
        End Try
    End Sub

    Public Sub UpdateApproveTrade()

        If Field14.Text = "" Or IsDBNull(Field14) Then
            MsgBox("You must set a trade status.", MsgBoxStyle.Critical, "No Staus Set")

        Else

            Dim Mycn As OleDb.OleDbConnection
            Dim Command As OleDb.OleDbCommand
            Dim SQLstr As String
            Try
                Mycn = New OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & My.Settings.dbloc & "';Jet OLEDB:Database Password='" & My.Settings.dbpass & "';")
                Dim ds1 As New DataSet
                Dim eds1 As New DataGridView
                Dim dv1 As New DataView
                Mycn.Open()

                SQLstr = "Update uma_envestnet_TradePending SET Field1 = '" & Field1.Text & "',Field2 = '" & Field2.Text & "',Field3 = '" & Field3.Text & "',Field4 = '" & Field4.Text & "',Field5 = '" & Field5.Text & "',Field6 = '" & Field6.Text & "',Field7 = '" & Field7.Text & "',Field8 = '" & Field8.Text & "',Field9 = '" & Field9.Text & "',Field10 = '" & Field10.Text & "',Field11 = '" & Field11.Text & "',Field12 = '" & Field12.Text & "',Field13 = '" & Field13.Text & "',Field14 = '" & Field14.Text & "',Field15 = '" & Field15.Text & "',Field16 = '" & Field16.Text & "',Field17 = '" & Field17.Text & "',Field18 = '" & Field18.Text & "', TradeReady = -1 WHERE ID = " & ID.Text

                Command = New OleDb.OleDbCommand(SQLstr, Mycn)
                Command.ExecuteNonQuery()

                Mycn.Close()

            Catch ex As Exception
                MsgBox(ex.Message, MsgBoxStyle.Critical, "ERROR")
            End Try
        End If

    End Sub

    Public Sub CheckChange()
        Dim Mycn As OleDb.OleDbConnection
        Dim Command As OleDb.OleDbCommand
        Dim Command1 As OleDb.OleDbCommand
        Dim Command2 As OleDb.OleDbCommand
        Dim Command3 As OleDb.OleDbCommand
        Dim Command4 As OleDb.OleDbCommand
        Dim Command5 As OleDb.OleDbCommand
        Dim Command6 As OleDb.OleDbCommand
        Dim Command7 As OleDb.OleDbCommand
        Dim Command8 As OleDb.OleDbCommand
        Dim Command9 As OleDb.OleDbCommand
        Dim Command10 As OleDb.OleDbCommand
        Dim Command11 As OleDb.OleDbCommand
        Dim Command12 As OleDb.OleDbCommand
        Dim Command13 As OleDb.OleDbCommand
        Dim Command14 As OleDb.OleDbCommand
        Dim Command15 As OleDb.OleDbCommand
        Dim Command16 As OleDb.OleDbCommand
        Dim Command17 As OleDb.OleDbCommand
        Dim SQLstr As String
        Try
            Mycn = New OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & My.Settings.dbloc & "';Jet OLEDB:Database Password='" & My.Settings.dbpass & "';")
            Mycn.Open()

            If Field1.Text = Old1.Text Then
                'nothing changed
            Else
                SQLstr = "INSERT INTO uma_tradechangelog(MoxyOrderID, Field, OldValue, NewValue, ChangeBy, ChangeDate)" & _
                "VALUES(" & MoxyOrderID.Text & ",'Field1','" & Old1.Text & "', '" & Field1.Text & "','" & Environ("USERNAME") & "',#" & Format(Now()) & "#)"

                Command = New OleDb.OleDbCommand(SQLstr, Mycn)
                Command.ExecuteNonQuery()
            End If

            If Field2.Text = Old2.Text Then
                'nothing changed
            Else
                SQLstr = "INSERT INTO uma_tradechangelog(MoxyOrderID, Field, OldValue, NewValue, ChangeBy, ChangeDate)" & _
                "VALUES(" & MoxyOrderID.Text & ",'Field2','" & Old2.Text & "', '" & Field2.Text & "','" & Environ("USERNAME") & "',#" & Format(Now()) & "#)"

                Command1 = New OleDb.OleDbCommand(SQLstr, Mycn)
                Command1.ExecuteNonQuery()
            End If

            If Field3.Text = Old3.Text Then
                'nothing changed
            Else
                SQLstr = "INSERT INTO uma_tradechangelog(MoxyOrderID, Field, OldValue, NewValue, ChangeBy, ChangeDate)" & _
                "VALUES(" & MoxyOrderID.Text & ",'Field3','" & Old3.Text & "', '" & Field3.Text & "','" & Environ("USERNAME") & "',#" & Format(Now()) & "#)"

                Command2 = New OleDb.OleDbCommand(SQLstr, Mycn)
                Command2.ExecuteNonQuery()
            End If

            If Field4.Text = Old4.Text Then
                'nothing changed
            Else
                SQLstr = "INSERT INTO uma_tradechangelog(MoxyOrderID, Field, OldValue, NewValue, ChangeBy, ChangeDate)" & _
                "VALUES(" & MoxyOrderID.Text & ",'Field4','" & Old4.Text & "', '" & Field4.Text & "','" & Environ("USERNAME") & "',#" & Format(Now()) & "#)"

                Command3 = New OleDb.OleDbCommand(SQLstr, Mycn)
                Command3.ExecuteNonQuery()
            End If

            If Field5.Text = Old5.Text Then
                'nothing changed
            Else
                SQLstr = "INSERT INTO uma_tradechangelog(MoxyOrderID, Field, OldValue, NewValue, ChangeBy, ChangeDate)" & _
                "VALUES(" & MoxyOrderID.Text & ",'Field5','" & Old5.Text & "', '" & Field5.Text & "','" & Environ("USERNAME") & "',#" & Format(Now()) & "#)"

                Command4 = New OleDb.OleDbCommand(SQLstr, Mycn)
                Command4.ExecuteNonQuery()
            End If

            If Field6.Text = Old6.Text Then
                'nothing changed
            Else
                SQLstr = "INSERT INTO uma_tradechangelog(MoxyOrderID, Field, OldValue, NewValue, ChangeBy, ChangeDate)" & _
                "VALUES(" & MoxyOrderID.Text & ",'Field6','" & Old6.Text & "', '" & Field6.Text & "','" & Environ("USERNAME") & "',#" & Format(Now()) & "#)"

                Command5 = New OleDb.OleDbCommand(SQLstr, Mycn)
                Command5.ExecuteNonQuery()
            End If

            If Field7.Text = Old7.Text Then
                'nothing changed
            Else
                SQLstr = "INSERT INTO uma_tradechangelog(MoxyOrderID, Field, OldValue, NewValue, ChangeBy, ChangeDate)" & _
                "VALUES(" & MoxyOrderID.Text & ",'Field7','" & Old7.Text & "', '" & Field7.Text & "','" & Environ("USERNAME") & "',#" & Format(Now()) & "#)"

                Command6 = New OleDb.OleDbCommand(SQLstr, Mycn)
                Command6.ExecuteNonQuery()
            End If

            If Field8.Text = Old8.Text Then
                'nothing changed
            Else
                SQLstr = "INSERT INTO uma_tradechangelog(MoxyOrderID, Field, OldValue, NewValue, ChangeBy, ChangeDate)" & _
                "VALUES(" & MoxyOrderID.Text & ",'Field8','" & Old8.Text & "', '" & Field8.Text & "','" & Environ("USERNAME") & "',#" & Format(Now()) & "#)"

                Command7 = New OleDb.OleDbCommand(SQLstr, Mycn)
                Command7.ExecuteNonQuery()
            End If

            If Field9.Text = Old9.Text Then
                'nothing changed
            Else
                SQLstr = "INSERT INTO uma_tradechangelog(MoxyOrderID, Field, OldValue, NewValue, ChangeBy, ChangeDate)" & _
                "VALUES(" & MoxyOrderID.Text & ",'Field9','" & Old9.Text & "', '" & Field9.Text & "','" & Environ("USERNAME") & "',#" & Format(Now()) & "#)"

                Command8 = New OleDb.OleDbCommand(SQLstr, Mycn)
                Command8.ExecuteNonQuery()
            End If

            If Field10.Text = Old10.Text Then
                'nothing changed
            Else
                SQLstr = "INSERT INTO uma_tradechangelog(MoxyOrderID, Field, OldValue, NewValue, ChangeBy, ChangeDate)" & _
                "VALUES(" & MoxyOrderID.Text & ",'Field10','" & Old10.Text & "', '" & Field10.Text & "','" & Environ("USERNAME") & "',#" & Format(Now()) & "#)"

                Command9 = New OleDb.OleDbCommand(SQLstr, Mycn)
                Command9.ExecuteNonQuery()
            End If

            If Field11.Text = Old11.Text Then
                'nothing changed
            Else
                SQLstr = "INSERT INTO uma_tradechangelog(MoxyOrderID, Field, OldValue, NewValue, ChangeBy, ChangeDate)" & _
                "VALUES(" & MoxyOrderID.Text & ",'Field11','" & Old11.Text & "', '" & Field11.Text & "','" & Environ("USERNAME") & "',#" & Format(Now()) & "#)"

                Command10 = New OleDb.OleDbCommand(SQLstr, Mycn)
                Command10.ExecuteNonQuery()
            End If

            If Field12.Text = Old12.Text Then
                'nothing changed
            Else
                SQLstr = "INSERT INTO uma_tradechangelog(MoxyOrderID, Field, OldValue, NewValue, ChangeBy, ChangeDate)" & _
                "VALUES(" & MoxyOrderID.Text & ",'Field12','" & Old12.Text & "', '" & Field12.Text & "','" & Environ("USERNAME") & "',#" & Format(Now()) & "#)"

                Command11 = New OleDb.OleDbCommand(SQLstr, Mycn)
                Command11.ExecuteNonQuery()
            End If

            If Field13.Text = Old13.Text Then
                'nothing changed
            Else
                SQLstr = "INSERT INTO uma_tradechangelog(MoxyOrderID, Field, OldValue, NewValue, ChangeBy, ChangeDate)" & _
                "VALUES(" & MoxyOrderID.Text & ",'Field13','" & Old13.Text & "', '" & Field13.Text & "','" & Environ("USERNAME") & "',#" & Format(Now()) & "#)"

                Command12 = New OleDb.OleDbCommand(SQLstr, Mycn)
                Command12.ExecuteNonQuery()
            End If

            If Field14.Text = Old14.Text Then
                'nothing changed
            Else
                SQLstr = "INSERT INTO uma_tradechangelog(MoxyOrderID, Field, OldValue, NewValue, ChangeBy, ChangeDate)" & _
                "VALUES(" & MoxyOrderID.Text & ",'Field14','" & Old14.Text & "', '" & Field14.Text & "','" & Environ("USERNAME") & "',#" & Format(Now()) & "#)"

                Command13 = New OleDb.OleDbCommand(SQLstr, Mycn)
                Command13.ExecuteNonQuery()
            End If

            If Field15.Text = Old15.Text Then
                'nothing changed
            Else
                SQLstr = "INSERT INTO uma_tradechangelog(MoxyOrderID, Field, OldValue, NewValue, ChangeBy, ChangeDate)" & _
                "VALUES(" & MoxyOrderID.Text & ",'Field15','" & Old15.Text & "', '" & Field15.Text & "','" & Environ("USERNAME") & "',#" & Format(Now()) & "#)"

                Command14 = New OleDb.OleDbCommand(SQLstr, Mycn)
                Command14.ExecuteNonQuery()
            End If

            If Field16.Text = Old16.Text Then
                'nothing changed
            Else
                SQLstr = "INSERT INTO uma_tradechangelog(MoxyOrderID, Field, OldValue, NewValue, ChangeBy, ChangeDate)" & _
                "VALUES(" & MoxyOrderID.Text & ",'Field16','" & Old16.Text & "', '" & Field16.Text & "','" & Environ("USERNAME") & "',#" & Format(Now()) & "#)"

                Command15 = New OleDb.OleDbCommand(SQLstr, Mycn)
                Command15.ExecuteNonQuery()
            End If

            If Field17.Text = Old17.Text Then
                'nothing changed
            Else
                SQLstr = "INSERT INTO uma_tradechangelog(MoxyOrderID, Field, OldValue, NewValue, ChangeBy, ChangeDate)" & _
                "VALUES(" & MoxyOrderID.Text & ",'Field17','" & Old17.Text & "', '" & Field17.Text & "','" & Environ("USERNAME") & "',#" & Format(Now()) & "#)"

                Command16 = New OleDb.OleDbCommand(SQLstr, Mycn)
                Command16.ExecuteNonQuery()
            End If

            If Field18.Text = Old18.Text Then
                'nothing changed
            Else
                SQLstr = "INSERT INTO uma_tradechangelog(MoxyOrderID, Field, OldValue, NewValue, ChangeBy, ChangeDate)" & _
                "VALUES(" & MoxyOrderID.Text & ",'Field18','" & Old18.Text & "', '" & Field18.Text & "','" & Environ("USERNAME") & "',#" & Format(Now()) & "#)"

                Command17 = New OleDb.OleDbCommand(SQLstr, Mycn)
                Command17.ExecuteNonQuery()
            End If

            Mycn.Close()

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "ERROR")
        End Try
    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        Dim ir As MsgBoxResult
        ir = MsgBox("You are about to modify and approve this trade.  Are you sure you want to continue?", MsgBoxStyle.YesNo + MsgBoxStyle.Critical, "Confirm")
        If ir = MsgBoxResult.Yes Then

            Call CheckChange()
            Call UpdateApproveTrade()
            Dim actchd As Form = Home.ActiveMdiChild

            UMA.MdiParent = Home

            Call UMA.LoadPendingTradesGrid()
            'Call UMA.LoadPendingTradesGrid()
            Me.Close()
        Else

        End If
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        Dim ir As MsgBoxResult
        ir = MsgBox("Are you sure you want to exit without saving?", MsgBoxStyle.YesNo + MsgBoxStyle.Critical, "Confirm")
        If ir = MsgBoxResult.Yes Then
            Me.Close()
        Else

        End If
    End Sub
End Class