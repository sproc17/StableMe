Public Class Form3

    Dim provider As String
    Dim dataFile As String
    Dim connString As String
    Dim myConnection As OleDbConnection = New OleDbConnection
    Dim em1 As Integer = 0
    Dim em2 As Integer = 0
    Dim em3 As Integer = 0
    Dim em4 As Integer = 0
    Dim em5 As Integer = 0
    Dim sit1 As Integer = 0
    Dim sit2 As Integer = 0
    Dim sit3 As Integer = 0
    Dim phys1 As Integer = 0
    Dim phys2 As Integer = 0
    Dim phys3 As Integer = 0
    Dim phys4 As Integer = 0
    Dim phys5 As Integer = 0


    Private Sub Label1_Click(sender As Object, e As EventArgs) Handles Label1.Click
        My.Forms.StableMe1.Show()
        Me.Close()
    End Sub

    Private Sub PictureBox2_Click(sender As Object, e As EventArgs) Handles PictureBox2.Click
        If PictureBox2.BorderStyle = BorderStyle.None Then
            PictureBox2.BorderStyle = BorderStyle.FixedSingle
            If em1 = 0 Then
                em1 = 14
            ElseIf em1 <> 0 And em2 = 0 Then
                em2 = 14
            ElseIf em2 <> 0 And em3 = 0 Then
                em3 = 14
            ElseIf em3 <> 0 And em4 = 0 Then
                em4 = 14
            ElseIf em4 <> 0 And em5 = 0 Then
                em5 = 14
            Else
                MsgBox("I know things can be rough, but try to stick to just 5 emotions, ok?", MsgBoxStyle.OkOnly, "Error")

            End If
        Else
            PictureBox2.BorderStyle = BorderStyle.None
        End If
    End Sub

    Private Sub PictureBox3_Click(sender As Object, e As EventArgs) Handles PictureBox3.Click
        If PictureBox3.BorderStyle = BorderStyle.None Then
            PictureBox3.BorderStyle = BorderStyle.FixedSingle
            If em1 = 0 Then
                em1 = 6
            ElseIf em1 <> 0 And em2 = 0 Then
                em2 = 6
            ElseIf em2 <> 0 And em3 = 0 Then
                em3 = 6
            ElseIf em3 <> 0 And em4 = 0 Then
                em4 = 6
            ElseIf em4 <> 0 And em5 = 0 Then
                em5 = 6
            Else
                MsgBox("I know things can be rough, but try to stick to just 5 for, ok?", MsgBoxStyle.OkOnly, "Error")

            End If
        Else
            PictureBox3.BorderStyle = BorderStyle.None
        End If
    End Sub

    Private Sub PictureBox4_Click(sender As Object, e As EventArgs) Handles PictureBox4.Click
        If PictureBox4.BorderStyle = BorderStyle.None Then
            PictureBox4.BorderStyle = BorderStyle.FixedSingle
            If em1 = 0 Then
                em1 = 8
            ElseIf em1 <> 0 And em2 = 0 Then
                em2 = 8
            ElseIf em2 <> 0 And em3 = 0 Then
                em3 = 8
            ElseIf em3 <> 0 And em4 = 0 Then
                em4 = 8
            ElseIf em4 <> 0 And em5 = 0 Then
                em5 = 8
            Else
                MsgBox("I know things can be rough, but try to stick to just 5 now, ok?", MsgBoxStyle.OkOnly, "Error")

            End If
        Else
            PictureBox4.BorderStyle = BorderStyle.None
        End If
    End Sub

    Private Sub PictureBox5_Click(sender As Object, e As EventArgs) Handles PictureBox5.Click
        If PictureBox5.BorderStyle = BorderStyle.None Then
            PictureBox5.BorderStyle = BorderStyle.FixedSingle
            If em1 = 0 Then
                em1 = 5
            ElseIf em1 <> 0 And em2 = 0 Then
                em2 = 5
            ElseIf em2 <> 0 And em3 = 0 Then
                em3 = 5
            ElseIf em3 <> 0 And em4 = 0 Then
                em4 = 5
            ElseIf em4 <> 0 And em5 = 0 Then
                em5 = 5
            Else
                MsgBox("I know things can be rough, but try to stick to just 5 emotions, ok?", MsgBoxStyle.OkOnly, "Error")

            End If
        Else
            PictureBox5.BorderStyle = BorderStyle.None
        End If
    End Sub

    Private Sub PictureBox6_Click(sender As Object, e As EventArgs) Handles PictureBox6.Click
        If PictureBox6.BorderStyle = BorderStyle.None Then
            PictureBox6.BorderStyle = BorderStyle.FixedSingle
            If em1 = 0 Then
                em1 = 7
            ElseIf em1 <> 0 And em2 = 0 Then
                em2 = 7
            ElseIf em2 <> 0 And em3 = 0 Then
                em3 = 7
            ElseIf em3 <> 0 And em4 = 0 Then
                em4 = 7
            ElseIf em4 <> 0 And em5 = 0 Then
                em5 = 7
            Else
                MsgBox("I know things can be rough, but try to stick to just 5 emotions, ok?", MsgBoxStyle.OkOnly, "Error")

            End If
        Else
            PictureBox6.BorderStyle = BorderStyle.None
        End If
    End Sub

    Private Sub PictureBox7_Click(sender As Object, e As EventArgs) Handles PictureBox7.Click
        If PictureBox7.BorderStyle = BorderStyle.None Then
            PictureBox7.BorderStyle = BorderStyle.FixedSingle
            If em1 = 0 Then
                em1 = 12
            ElseIf em1 <> 0 And em2 = 0 Then
                em2 = 12
            ElseIf em2 <> 0 And em3 = 0 Then
                em3 = 12
            ElseIf em3 <> 0 And em4 = 0 Then
                em4 = 12
            ElseIf em4 <> 0 And em5 = 0 Then
                em5 = 12
            Else
                MsgBox("I know things can be rough, but try to stick to just 5 emotions, ok?", MsgBoxStyle.OkOnly, "Error")

            End If
        Else
            PictureBox7.BorderStyle = BorderStyle.None
        End If
    End Sub

    Private Sub PictureBox9_Click(sender As Object, e As EventArgs) Handles PictureBox9.Click
        If PictureBox9.BorderStyle = BorderStyle.None Then
            PictureBox9.BorderStyle = BorderStyle.FixedSingle
            If em1 = 0 Then
                em1 = 9
            ElseIf em1 <> 0 And em2 = 0 Then
                em2 = 9
            ElseIf em2 <> 0 And em3 = 0 Then
                em3 = 9
            ElseIf em3 <> 0 And em4 = 0 Then
                em4 = 9
            ElseIf em4 <> 0 And em5 = 0 Then
                em5 = 9
            Else
                MsgBox("I know things can be rough, but try to stick to just 5 emotions, ok?", MsgBoxStyle.OkOnly, "Error")

            End If
        Else
            PictureBox9.BorderStyle = BorderStyle.None
        End If
    End Sub

    Private Sub PictureBox12_Click(sender As Object, e As EventArgs) Handles PictureBox12.Click
        If PictureBox12.BorderStyle = BorderStyle.None Then
            PictureBox12.BorderStyle = BorderStyle.FixedSingle
            If em1 = 0 Then
                em1 = 4
            ElseIf em1 <> 0 And em2 = 0 Then
                em2 = 4
            ElseIf em2 <> 0 And em3 = 0 Then
                em3 = 4
            ElseIf em3 <> 0 And em4 = 0 Then
                em4 = 4
            ElseIf em4 <> 0 And em5 = 0 Then
                em5 = 4
            Else
                MsgBox("I know things can be rough, but try to stick to just 5 emotions, ok?", MsgBoxStyle.OkOnly, "Error")

            End If
        Else
            PictureBox12.BorderStyle = BorderStyle.None
        End If
    End Sub

    Private Sub PictureBox13_Click(sender As Object, e As EventArgs) Handles PictureBox13.Click
        If PictureBox13.BorderStyle = BorderStyle.None Then
            PictureBox13.BorderStyle = BorderStyle.FixedSingle
            If em1 = 0 Then
                em1 = 1
            ElseIf em1 <> 0 And em2 = 0 Then
                em2 = 1
            ElseIf em2 <> 0 And em3 = 0 Then
                em3 = 1
            ElseIf em3 <> 0 And em4 = 0 Then
                em4 = 1
            ElseIf em4 <> 0 And em5 = 0 Then
                em5 = 1
            Else
                MsgBox("I know things can be rough, but try to stick to just 5 emotions, ok?", MsgBoxStyle.OkOnly, "Error")

            End If
        Else
            PictureBox13.BorderStyle = BorderStyle.None
        End If
    End Sub

    Private Sub PictureBox17_Click(sender As Object, e As EventArgs) Handles PictureBox17.Click
        If PictureBox17.BorderStyle = BorderStyle.None Then
            PictureBox17.BorderStyle = BorderStyle.FixedSingle
            If em1 = 0 Then
                em1 = 3
            ElseIf em1 <> 0 And em2 = 0 Then
                em2 = 3
            ElseIf em2 <> 0 And em3 = 0 Then
                em3 = 3
            ElseIf em3 <> 0 And em4 = 0 Then
                em4 = 3
            ElseIf em4 <> 0 And em5 = 0 Then
                em5 = 3
            Else
                MsgBox("I know things can be rough, but try to stick to just 5 emotions, ok?", MsgBoxStyle.OkOnly, "Error")

            End If
        Else
            PictureBox17.BorderStyle = BorderStyle.None
        End If
    End Sub

    Private Sub PictureBox20_Click(sender As Object, e As EventArgs) Handles PictureBox20.Click
        If PictureBox20.BorderStyle = BorderStyle.None Then
            PictureBox20.BorderStyle = BorderStyle.FixedSingle
            If em1 = 0 Then
                em1 = 11
            ElseIf em1 <> 0 And em2 = 0 Then
                em2 = 11
            ElseIf em2 <> 0 And em3 = 0 Then
                em3 = 11
            ElseIf em3 <> 0 And em4 = 0 Then
                em4 = 11
            ElseIf em4 <> 0 And em5 = 0 Then
                em5 = 11
            Else
                MsgBox("I know things can be rough, but try to stick to just 5 emotions, ok?", MsgBoxStyle.OkOnly, "Error")

            End If
        Else
            PictureBox20.BorderStyle = BorderStyle.None
        End If
    End Sub

    Private Sub PictureBox24_Click(sender As Object, e As EventArgs) Handles PictureBox24.Click
        If PictureBox24.BorderStyle = BorderStyle.None Then
            PictureBox24.BorderStyle = BorderStyle.FixedSingle
            If em1 = 0 Then
                em1 = 2
            ElseIf em1 <> 0 And em2 = 0 Then
                em2 = 2
            ElseIf em2 <> 0 And em3 = 0 Then
                em3 = 2
            ElseIf em3 <> 0 And em4 = 0 Then
                em4 = 2
            ElseIf em4 <> 0 And em5 = 0 Then
                em5 = 2
            Else
                MsgBox("I know things can be rough, but try to stick to just 5 emotions, ok?", MsgBoxStyle.OkOnly, "Error")

            End If
        Else
            PictureBox24.BorderStyle = BorderStyle.None
        End If
    End Sub

    Private Sub PictureBox11_Click(sender As Object, e As EventArgs) Handles PictureBox11.Click
        If PictureBox11.BorderStyle = BorderStyle.None Then
            PictureBox11.BorderStyle = BorderStyle.FixedSingle
            If sit1 = 0 Then
                sit1 = 4
            ElseIf sit1 <> 0 And sit2 = 0 Then
                sit2 = 4
            ElseIf sit2 <> 0 And sit3 = 0 Then
                sit3 = 4
            Else
                MsgBox("I know theres a lot going on, but try to stick to just 3 situations, ok?", MsgBoxStyle.OkOnly, "Error")

            End If
        Else
            PictureBox11.BorderStyle = BorderStyle.None
        End If
    End Sub

    Private Sub PictureBox15_Click(sender As Object, e As EventArgs) Handles PictureBox15.Click
        If PictureBox15.BorderStyle = BorderStyle.None Then
            PictureBox15.BorderStyle = BorderStyle.FixedSingle
            If sit1 = 0 Then
                sit1 = 2
            ElseIf sit1 <> 0 And sit2 = 0 Then
                sit2 = 2
            ElseIf sit2 <> 0 And sit3 = 0 Then
                sit3 = 2
            Else
                MsgBox("I know theres a lot going on, but try to stick to just 3 situations, ok?", MsgBoxStyle.OkOnly, "Error")

            End If
        Else
            PictureBox15.BorderStyle = BorderStyle.None
        End If
    End Sub

    Private Sub PictureBox18_Click(sender As Object, e As EventArgs) Handles PictureBox18.Click
        If PictureBox18.BorderStyle = BorderStyle.None Then
            PictureBox18.BorderStyle = BorderStyle.FixedSingle
            If sit1 = 0 Then
                sit1 = 8
            ElseIf sit1 <> 0 And sit2 = 0 Then
                sit2 = 8
            ElseIf sit2 <> 0 And sit3 = 0 Then
                sit3 = 8
            Else
                MsgBox("I know theres a lot going on, but try to stick to just 3 situations, ok?", MsgBoxStyle.OkOnly, "Error")

            End If
        Else
            PictureBox18.BorderStyle = BorderStyle.None
        End If
    End Sub

    Private Sub PictureBox25_Click(sender As Object, e As EventArgs) Handles PictureBox25.Click
        If PictureBox25.BorderStyle = BorderStyle.None Then
            PictureBox25.BorderStyle = BorderStyle.FixedSingle
            If sit1 = 0 Then
                sit1 = 6
            ElseIf sit1 <> 0 And sit2 = 0 Then
                sit2 = 6
            ElseIf sit2 <> 0 And sit3 = 0 Then
                sit3 = 6
            Else
                MsgBox("I know theres a lot going on, but try to stick to just 3 situations, ok?", MsgBoxStyle.OkOnly, "Error")

            End If
        Else
            PictureBox25.BorderStyle = BorderStyle.None
        End If
    End Sub

    Private Sub PictureBox27_Click(sender As Object, e As EventArgs) Handles PictureBox27.Click
        If PictureBox27.BorderStyle = BorderStyle.None Then
            PictureBox27.BorderStyle = BorderStyle.FixedSingle
            If sit1 = 0 Then
                sit1 = 5
            ElseIf sit1 <> 0 And sit2 = 0 Then
                sit2 = 5
            ElseIf sit2 <> 0 And sit3 = 0 Then
                sit3 = 5
            Else
                MsgBox("I know theres a lot going on, but try to stick to just 3 situations, ok?", MsgBoxStyle.OkOnly, "Error")

            End If
        Else
            PictureBox27.BorderStyle = BorderStyle.None
        End If
    End Sub

    Private Sub PictureBox31_Click(sender As Object, e As EventArgs) Handles PictureBox31.Click
        If PictureBox31.BorderStyle = BorderStyle.None Then
            PictureBox31.BorderStyle = BorderStyle.FixedSingle
            If sit1 = 0 Then
                sit1 = 7
            ElseIf sit1 <> 0 And sit2 = 0 Then
                sit2 = 7
            ElseIf sit2 <> 0 And sit3 = 0 Then
                sit3 = 7
            Else
                MsgBox("I know theres a lot going on, but try to stick to just 3 situations, ok?", MsgBoxStyle.OkOnly, "Error")

            End If
        Else
            PictureBox31.BorderStyle = BorderStyle.None
        End If
    End Sub

    Private Sub PictureBox28_Click(sender As Object, e As EventArgs) Handles PictureBox28.Click
        If PictureBox28.BorderStyle = BorderStyle.None Then
            PictureBox28.BorderStyle = BorderStyle.FixedSingle
            If sit1 = 0 Then
                sit1 = 3
            ElseIf sit1 <> 0 And sit2 = 0 Then
                sit2 = 3
            ElseIf sit2 <> 0 And sit3 = 0 Then
                sit3 = 3
            Else
                MsgBox("I know theres a lot going on, but try to stick to just 3 situations, ok?", MsgBoxStyle.OkOnly, "Error")

            End If
        Else
            PictureBox28.BorderStyle = BorderStyle.None
        End If
    End Sub

    Private Sub PictureBox8_Click(sender As Object, e As EventArgs) Handles PictureBox8.Click
        If PictureBox8.BorderStyle = BorderStyle.None Then
            PictureBox8.BorderStyle = BorderStyle.FixedSingle
            If phys1 = 0 Then
                phys1 = 2
            ElseIf phys1 <> 0 And phys2 = 0 Then
                phys2 = 2
            ElseIf phys2 <> 0 And phys3 = 0 Then
                phys3 = 2
            ElseIf phys3 <> 0 And phys4 = 0 Then
                phys4 = 2
            ElseIf phys4 <> 0 And phys5 = 0 Then
                phys5 = 2
            Else
                MsgBox("I know things can be rough, but try to stick to just 5 physical symptoms for now, ok?", MsgBoxStyle.OkOnly, "Error")

            End If
        Else
            PictureBox8.BorderStyle = BorderStyle.None
        End If
    End Sub

    Private Sub PictureBox14_Click(sender As Object, e As EventArgs) Handles PictureBox14.Click
        If PictureBox14.BorderStyle = BorderStyle.None Then
            PictureBox14.BorderStyle = BorderStyle.FixedSingle
            If phys1 = 0 Then
                phys1 = 3
            ElseIf phys1 <> 0 And phys2 = 0 Then
                phys2 = 3
            ElseIf phys2 <> 0 And phys3 = 0 Then
                phys3 = 3
            ElseIf phys3 <> 0 And phys4 = 0 Then
                phys4 = 3
            ElseIf phys4 <> 0 And phys5 = 0 Then
                phys5 = 3
            Else
                MsgBox("I know things can be rough, but try to stick to just 5 physical symtoms for now, ok?", MsgBoxStyle.OkOnly, "Error")

            End If
        Else
            PictureBox14.BorderStyle = BorderStyle.None
        End If
    End Sub

    Private Sub PictureBox16_Click(sender As Object, e As EventArgs) Handles PictureBox16.Click
        If PictureBox16.BorderStyle = BorderStyle.None Then
            PictureBox16.BorderStyle = BorderStyle.FixedSingle
            If phys1 = 0 Then
                phys1 = 11
            ElseIf phys1 <> 0 And phys2 = 0 Then
                phys2 = 11
            ElseIf phys2 <> 0 And phys3 = 0 Then
                phys3 = 11
            ElseIf phys3 <> 0 And phys4 = 0 Then
                phys4 = 11
            ElseIf phys4 <> 0 And phys5 = 0 Then
                phys5 = 11
            Else
                MsgBox("I know things can be rough, but try to stick to just 5 physical symptoms for now, ok?", MsgBoxStyle.OkOnly, "Error")

            End If
        Else
            PictureBox16.BorderStyle = BorderStyle.None
        End If
    End Sub

    Private Sub PictureBox21_Click(sender As Object, e As EventArgs) Handles PictureBox21.Click
        If PictureBox21.BorderStyle = BorderStyle.None Then
            PictureBox21.BorderStyle = BorderStyle.FixedSingle
            If phys1 = 0 Then
                phys1 = 11
            ElseIf phys1 <> 0 And phys2 = 0 Then
                phys2 = 11
            ElseIf phys2 <> 0 And phys3 = 0 Then
                phys3 = 11
            ElseIf phys3 <> 0 And phys4 = 0 Then
                phys4 = 11
            ElseIf phys4 <> 0 And phys5 = 0 Then
                phys5 = 11
            Else
                MsgBox("I know things can be rough, but try to stick to just 5 physical symptoms for now, ok?", MsgBoxStyle.OkOnly, "Error")

            End If
        Else
            PictureBox21.BorderStyle = BorderStyle.None
        End If
    End Sub

    Private Sub PictureBox23_Click(sender As Object, e As EventArgs) Handles PictureBox23.Click
        If PictureBox23.BorderStyle = BorderStyle.None Then
            PictureBox23.BorderStyle = BorderStyle.FixedSingle
            If phys1 = 0 Then
                phys1 = 12
            ElseIf phys1 <> 0 And phys2 = 0 Then
                phys2 = 12
            ElseIf phys2 <> 0 And phys3 = 0 Then
                phys3 = 12
            ElseIf phys3 <> 0 And phys4 = 0 Then
                phys4 = 12
            ElseIf phys4 <> 0 And phys5 = 0 Then
                phys5 = 12
            Else
                MsgBox("I know things can be rough, but try to stick to just 5 physical symptoms for now, ok?", MsgBoxStyle.OkOnly, "Error")

            End If
        Else
            PictureBox23.BorderStyle = BorderStyle.None
        End If
    End Sub

    Private Sub PictureBox19_Click(sender As Object, e As EventArgs) Handles PictureBox19.Click
        If PictureBox19.BorderStyle = BorderStyle.None Then
            PictureBox19.BorderStyle = BorderStyle.FixedSingle
            If phys1 = 0 Then
                phys1 = 6
            ElseIf phys1 <> 0 And phys2 = 0 Then
                phys2 = 6
            ElseIf phys2 <> 0 And phys3 = 0 Then
                phys3 = 6
            ElseIf phys3 <> 0 And phys4 = 0 Then
                phys4 = 6
            ElseIf phys4 <> 0 And phys5 = 0 Then
                phys5 = 6
            Else
                MsgBox("I know things can be rough, but try to stick to just 5 physical symptoms for now, ok?", MsgBoxStyle.OkOnly, "Error")

            End If
        Else
            PictureBox19.BorderStyle = BorderStyle.None
        End If
    End Sub

    Private Sub PictureBox10_Click(sender As Object, e As EventArgs) Handles PictureBox10.Click
        If PictureBox10.BorderStyle = BorderStyle.None Then
            PictureBox10.BorderStyle = BorderStyle.FixedSingle
            If phys1 = 0 Then
                phys1 = 5
            ElseIf phys1 <> 0 And phys2 = 0 Then
                phys2 = 5
            ElseIf phys2 <> 0 And phys3 = 0 Then
                phys3 = 5
            ElseIf phys3 <> 0 And phys4 = 0 Then
                phys4 = 5
            ElseIf phys4 <> 0 And phys5 = 0 Then
                phys5 = 5
            Else
                MsgBox("I know things can be rough, but try to stick to just 5 physical symptoms for now, ok?", MsgBoxStyle.OkOnly, "Error")

            End If
        Else
            PictureBox10.BorderStyle = BorderStyle.None
        End If
    End Sub

    Private Sub PictureBox30_Click(sender As Object, e As EventArgs) Handles PictureBox30.Click
        If PictureBox30.BorderStyle = BorderStyle.None Then
            PictureBox30.BorderStyle = BorderStyle.FixedSingle
            If phys1 = 0 Then
                phys1 = 4
            ElseIf phys1 <> 0 And phys2 = 0 Then
                phys2 = 4
            ElseIf phys2 <> 0 And phys3 = 0 Then
                phys3 = 4
            ElseIf phys3 <> 0 And phys4 = 0 Then
                phys4 = 4
            ElseIf phys4 <> 0 And phys5 = 0 Then
                phys5 = 4
            Else
                MsgBox("I know things can be rough, but try to stick to just 5 physical symptoms for now, ok?", MsgBoxStyle.OkOnly, "Error")

            End If
        Else
            PictureBox30.BorderStyle = BorderStyle.None
        End If
    End Sub

    Private Sub PictureBox22_Click(sender As Object, e As EventArgs) Handles PictureBox22.Click
        If PictureBox22.BorderStyle = BorderStyle.None Then
            PictureBox22.BorderStyle = BorderStyle.FixedSingle
            If phys1 = 0 Then
                phys1 = 8
            ElseIf phys1 <> 0 And phys2 = 0 Then
                phys2 = 8
            ElseIf phys2 <> 0 And phys3 = 0 Then
                phys3 = 8
            ElseIf phys3 <> 0 And phys4 = 0 Then
                phys4 = 8
            ElseIf phys4 <> 0 And phys5 = 0 Then
                phys5 = 8
            Else
                MsgBox("I know things can be rough, but try to stick to just 5 physical symptoms for now, ok?", MsgBoxStyle.OkOnly, "Error")

            End If
        Else
            PictureBox22.BorderStyle = BorderStyle.None
        End If
    End Sub

    Private Sub PictureBox29_Click(sender As Object, e As EventArgs) Handles PictureBox29.Click
        If PictureBox29.BorderStyle = BorderStyle.None Then
            PictureBox29.BorderStyle = BorderStyle.FixedSingle
            If phys1 = 0 Then
                phys1 = 9
            ElseIf phys1 <> 0 And phys2 = 0 Then
                phys2 = 9
            ElseIf phys2 <> 0 And phys3 = 0 Then
                phys3 = 9
            ElseIf phys3 <> 0 And phys4 = 0 Then
                phys4 = 9
            ElseIf phys4 <> 0 And phys5 = 0 Then
                phys5 = 9
            Else
                MsgBox("I know things can be rough, but try to stick to just 5 physical symptoms for now, ok?", MsgBoxStyle.OkOnly, "Error")

            End If
        Else
            PictureBox29.BorderStyle = BorderStyle.None
        End If
    End Sub

    Private Sub PictureBox26_Click(sender As Object, e As EventArgs) Handles PictureBox26.Click
        If PictureBox26.BorderStyle = BorderStyle.None Then
            PictureBox26.BorderStyle = BorderStyle.FixedSingle
            If phys1 = 0 Then
                phys1 = 7
            ElseIf phys1 <> 0 And phys2 = 0 Then
                phys2 = 7
            ElseIf phys2 <> 0 And phys3 = 0 Then
                phys3 = 7
            ElseIf phys3 <> 0 And phys4 = 0 Then
                phys4 = 7
            ElseIf phys4 <> 0 And phys5 = 0 Then
                phys5 = 7
            Else
                MsgBox("I know things can be rough, but try to stick to just 5 physical symptoms for now, ok?", MsgBoxStyle.OkOnly, "Error")

            End If
        Else
            PictureBox26.BorderStyle = BorderStyle.None
        End If
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If PictureBox26.BorderStyle = BorderStyle.FixedSingle Then
            MsgBox("You said you've been self harming. Please call the Suicide Hotline at 1-800-273-8255 if you need someone to talk to. We're here for you.", MsgBoxStyle.OkOnly, "StableMe")
        ElseIf PictureBox35.BorderStyle = BorderStyle.FixedSingle Then
            MsgBox("You said you have been having suicidal thought. Please call the National Suicide Hotline at 1-800-273-8255 if you need someone to talk to. We're here for you.", MsgBoxStyle.OkOnly, "StableMe")
        Else
            My.Forms.StableMe1.Show()
            Me.Close()
        End If
    End Sub

    Private Sub PictureBox33_Click(sender As Object, e As EventArgs) Handles PictureBox33.Click
        If PictureBox33.BorderStyle = BorderStyle.None Then
            PictureBox33.BorderStyle = BorderStyle.FixedSingle
            If em1 = 0 Then
                em1 = 9
            ElseIf em1 <> 0 And em2 = 0 Then
                em2 = 9
            ElseIf em2 <> 0 And em3 = 0 Then
                em3 = 9
            ElseIf em3 <> 0 And em4 = 0 Then
                em4 = 9
            ElseIf em4 <> 0 And em5 = 0 Then
                em5 = 9
            Else
                MsgBox("I know things can be rough, but try to stick to just 5 emotions, ok?", MsgBoxStyle.OkOnly, "Error")

            End If
        Else
            PictureBox33.BorderStyle = BorderStyle.None
        End If
    End Sub

    Private Sub PictureBox32_Click(sender As Object, e As EventArgs) Handles PictureBox32.Click
        If PictureBox32.BorderStyle = BorderStyle.None Then
            PictureBox32.BorderStyle = BorderStyle.FixedSingle
            If em1 = 0 Then
                em1 = 13
            ElseIf em1 <> 0 And em2 = 0 Then
                em2 = 13
            ElseIf em2 <> 0 And em3 = 0 Then
                em3 = 13
            ElseIf em3 <> 0 And em4 = 0 Then
                em4 = 13
            ElseIf em4 <> 0 And em5 = 0 Then
                em5 = 13
            Else
                MsgBox("I know things can be rough, but try to stick to just 5 emotions, ok?", MsgBoxStyle.OkOnly, "Error")

            End If
        Else
            PictureBox32.BorderStyle = BorderStyle.None
        End If
    End Sub

    Private Sub PictureBox34_Click(sender As Object, e As EventArgs) Handles PictureBox34.Click
        If PictureBox34.BorderStyle = BorderStyle.None Then
            PictureBox34.BorderStyle = BorderStyle.FixedSingle
            If sit1 = 0 Then
                sit1 = 1
            ElseIf sit1 <> 0 And sit2 = 0 Then
                sit2 = 1
            ElseIf sit2 <> 0 And sit3 = 0 Then
                sit3 = 1
            Else
                MsgBox("I know theres a lot going on, but try to stick to just 3 situations, ok?", MsgBoxStyle.OkOnly, "Error")

            End If
        Else
            PictureBox34.BorderStyle = BorderStyle.None
        End If
    End Sub

    Private Sub PictureBox35_Click(sender As Object, e As EventArgs) Handles PictureBox35.Click
        If PictureBox35.BorderStyle = BorderStyle.None Then
            PictureBox35.BorderStyle = BorderStyle.FixedSingle
            If phys1 = 0 Then
                phys1 = 1
            ElseIf phys1 <> 0 And phys2 = 0 Then
                phys2 = 1
            ElseIf phys2 <> 0 And phys3 = 0 Then
                phys3 = 1
            ElseIf phys3 <> 0 And phys4 = 0 Then
                phys4 = 1
            ElseIf phys4 <> 0 And phys5 = 0 Then
                phys5 = 1
            Else
                MsgBox("I know things can be rough, but try to stick to just 5 physical, ok?", MsgBoxStyle.OkOnly, "Error")

            End If
        Else
            PictureBox35.BorderStyle = BorderStyle.None
        End If
    End Sub

    Private Sub PictureBox36_Click(sender As Object, e As EventArgs) Handles PictureBox36.Click
        If PictureBox36.BorderStyle = BorderStyle.None Then
            PictureBox36.BorderStyle = BorderStyle.FixedSingle
            If phys1 = 0 Then
                phys1 = 13
            ElseIf phys1 <> 0 And phys2 = 0 Then
                phys2 = 13
            ElseIf phys2 <> 0 And phys3 = 0 Then
                phys3 = 13
            ElseIf phys3 <> 0 And phys4 = 0 Then
                phys4 = 13
            ElseIf phys4 <> 0 And phys5 = 0 Then
                phys5 = 13
            Else
                MsgBox("I know things can be rough, but try to stick to just 5 emotions, ok?", MsgBoxStyle.OkOnly, "Error")

            End If
        Else
            PictureBox36.BorderStyle = BorderStyle.None
        End If
    End Sub
End Class