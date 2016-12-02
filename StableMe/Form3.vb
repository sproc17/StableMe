Public Class Form3
    Private Access As New DBControl

    Dim em1 As String = ""
    Dim em2 As String = ""
    Dim em3 As String = ""
    Dim em4 As String = ""
    Dim em5 As String = ""
    Dim sit1 As String = ""
    Dim sit2 As String = ""
    Dim sit3 As String = ""
    Dim phys1 As String = ""
    Dim phys2 As String = ""
    Dim phys3 As String = ""
    Dim phys4 As String = ""
    Dim phys5 As String = ""
    Dim note As String = ""
    Dim eID1 As Integer = 0
    Dim eID2 As Integer = 0
    Dim eID3 As Integer = 0
    Dim eID4 As Integer = 0
    Dim eID5 As Integer = 0
    Dim sID1 As Integer = 0
    Dim sID2 As Integer = 0
    Dim sID3 As Integer = 0

    Private Sub Label1_Click(sender As Object, e As EventArgs) Handles Label1.Click
        My.Forms.StableMe1.Show()
        Me.Close()
    End Sub

    Private Sub PictureBox2_Click(sender As Object, e As EventArgs) Handles PictureBox2.Click
        If PictureBox2.BorderStyle = BorderStyle.None Then
            PictureBox2.BorderStyle = BorderStyle.FixedSingle
            If em1 = "" Then
                em1 = "Angry"
                eID1 = 14
            ElseIf em1 <> "" And em2 = "" Then
                em2 = "Angry"
                eID2 = 14
            ElseIf em2 <> "" And em3 = "" Then
                em3 = "Angry"
                eID3 = 14
            ElseIf em3 <> "" And em4 = "" Then
                em4 = "Angry"
                eID4 = 14
            ElseIf em4 <> "" And em5 = "" Then
                em5 = "Angry"
                eID5 = 14
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
            If em1 = "" Then
                em1 = "Emotional"
                eID1 = 6
            ElseIf em1 <> "" And em2 = "" Then
                em2 = "Emotional"
                eID2 = 6
            ElseIf em2 <> "" And em3 = "" Then
                em3 = "Emotional"
                eID3 = 6
            ElseIf em3 <> "" And em4 = "" Then
                em4 = "Emotional"
                eID4 = 6
            ElseIf em4 <> "" And em5 = "" Then
                em5 = "Emotional"
                eID5 = 6
            Else
                MsgBox("I know things can be rough, but try to stick to just 5 emotions, ok?", MsgBoxStyle.OkOnly, "Error")
            End If
        Else
            PictureBox3.BorderStyle = BorderStyle.None
        End If
    End Sub

    Private Sub PictureBox4_Click(sender As Object, e As EventArgs) Handles PictureBox4.Click
        If PictureBox4.BorderStyle = BorderStyle.None Then
            PictureBox4.BorderStyle = BorderStyle.FixedSingle
            If em1 = "" Then
                em1 = "Motivated"
                eID1 = 8
            ElseIf em1 <> "" And em2 = "" Then
                em2 = "Motivated"
                eID2 = 8
            ElseIf em2 <> "" And em3 = "" Then
                em3 = "Motivated"
                eID3 = 8
            ElseIf em3 <> "" And em4 = "" Then
                em4 = "Motivated"
                eID4 = 8
            ElseIf em4 <> "" And em5 = "" Then
                em5 = "Motivated"
                eID5 = 8
            Else
                MsgBox("I know things can be rough, but try to stick to just 5 emotions, ok?", MsgBoxStyle.OkOnly, "Error")
            End If
        Else
            PictureBox4.BorderStyle = BorderStyle.None
        End If
    End Sub

    Private Sub PictureBox5_Click(sender As Object, e As EventArgs) Handles PictureBox5.Click
        If PictureBox5.BorderStyle = BorderStyle.None Then
            PictureBox5.BorderStyle = BorderStyle.FixedSingle
            If em1 = "" Then
                em1 = "Content"
                eID1 = 5
            ElseIf em1 <> "" And em2 = "" Then
                em2 = "Content"
                eID2 = 5
            ElseIf em2 <> "" And em3 = "" Then
                em3 = "Content"
                eID3 = 5
            ElseIf em3 <> "" And em4 = "" Then
                em4 = "Content"
                eID4 = 5
            ElseIf em4 <> "" And em5 = "" Then
                em5 = "Content"
                eID5 = 5
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
            If em1 = "" Then
                em1 = "Anxious"
                eID1 = 7
            ElseIf em1 <> "" And em2 = "" Then
                em2 = "Anxious"
                eID2 = 7
            ElseIf em2 <> "" And em3 = "" Then
                em3 = "Anxious"
                eID3 = 7
            ElseIf em3 <> "" And em4 = "" Then
                em4 = "Anxious"
                eID4 = 7
            ElseIf em4 <> "" And em5 = "" Then
                em5 = "Anxious"
                eID5 = 7
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
            If em1 = "" Then
                em1 = "Stressed"
                eID1 = 12
            ElseIf em1 <> "" And em2 = "" Then
                em2 = "Stressed"
                eID2 = 12
            ElseIf em2 <> "" And em3 = "" Then
                em3 = "Stressed"
                eID3 = 12
            ElseIf em3 <> "" And em4 = "" Then
                em4 = "Stressed"
                eID4 = 12
            ElseIf em4 <> "" And em5 = "" Then
                em5 = "Stressed"
                eID5 = 12
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
            If em1 = "" Then
                em1 = "Excited"
                eID1 = 9
            ElseIf em1 <> "" And em2 = "" Then
                em2 = "Excited"
                eID2 = 9
            ElseIf em2 <> "" And em3 = "" Then
                em3 = "Excited"
                eID3 = 9
            ElseIf em3 <> "" And em4 = "" Then
                em4 = "Excited"
                eID4 = 9
            ElseIf em4 <> "" And em5 = "" Then
                em5 = "Excited"
                eID5 = 9
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
            If em1 = "" Then
                em1 = "Frustrated"
                eID1 = 4
            ElseIf em1 <> "" And em2 = "" Then
                em2 = "Frustrated"
                eID2 = 4
            ElseIf em2 <> "" And em3 = "" Then
                em3 = "Frustrated"
                eID3 = 4
            ElseIf em3 <> "" And em4 = "" Then
                em4 = "Frustrated"
                eID4 = 4
            ElseIf em4 <> "" And em5 = "" Then
                em5 = "Frustrated"
                eID5 = 4
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
            If em1 = "" Then
                em1 = "Happy"
                eID1 = 1
            ElseIf em1 <> "" And em2 = "" Then
                em2 = "Happy"
                eID2 = 1
            ElseIf em2 <> "" And em3 = "" Then
                em3 = "Happy"
                eID3 = 1
            ElseIf em3 <> "" And em4 = "" Then
                em4 = "Happy"
                eID4 = 1
            ElseIf em4 <> "" And em5 = "" Then
                em5 = "Happy"
                eID5 = 1
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
            If em1 = "" Then
                em1 = "Empty"
                eID1 = 3
            ElseIf em1 <> "" And em2 = "" Then
                em2 = "Empty"
                eID2 = 3
            ElseIf em2 <> "" And em3 = "" Then
                em3 = "Empty"
                eID3 = 3
            ElseIf em3 <> "" And em4 = "" Then
                em4 = "Empty"
                eID4 = 3
            ElseIf em4 <> "" And em5 = "" Then
                em5 = "Empty"
                eID5 = 3
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
            If em1 = "" Then
                em1 = "Confidant"
                eID1 = 11
            ElseIf em1 <> "" And em2 = "" Then
                em2 = "Confidant"
                eID2 = 11
            ElseIf em2 <> "" And em3 = "" Then
                em3 = "Confidant"
                eID3 = 11
            ElseIf em3 <> "" And em4 = "" Then
                em4 = "Confidant"
                eID4 = 11
            ElseIf em4 <> "" And em5 = "" Then
                em5 = "Confidant"
                eID5 = 11
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
            If em1 = "" Then
                em1 = "Sad"
                eID1 = 2
            ElseIf em1 <> "" And em2 = "" Then
                em2 = "Sad"
                eID2 = 2
            ElseIf em2 <> "" And em3 = "" Then
                em3 = "Sad"
                eID3 = 2
            ElseIf em3 <> "" And em4 = "" Then
                em4 = "Sad"
                eID4 = 2
            ElseIf em4 <> "" And em5 = "" Then
                em5 = "Sad"
                eID4 = 2
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
            If sit1 = "" Then
                sit1 = "Friends"
                sID1 = 4
            ElseIf sit1 <> "" And sit2 = "" Then
                sit2 = "Friends"
                sID2 = 4
            ElseIf sit2 <> "" And sit3 = "" Then
                sit3 = "Friends"
                sID3 = 4
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
            If sit1 = "" Then
                sit1 = "Family"
                sID1 = 2
            ElseIf sit1 <> "" And sit2 = "" Then
                sit2 = "Family"
                sID2 = 2
            ElseIf sit2 <> "" And sit3 = "" Then
                sit3 = "Family"
                sID3 = 2
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
            If sit1 = "" Then
                sit1 = "Love Life"
                sID1 = 8
            ElseIf sit1 <> "" And sit2 = "" Then
                sit2 = "Love Life"
                sID2 = 8
            ElseIf sit2 <> "" And sit3 = "" Then
                sit3 = "Love Life"
                sID3 = 8
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
            If sit1 = "" Then
                sit1 = "School"
                sID1 = 6
            ElseIf sit1 <> "" And sit2 = "" Then
                sit2 = "School"
                sID2 = 6
            ElseIf sit2 <> "" And sit3 = "" Then
                sit3 = "School"
                sID3 = 6
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
            If sit1 = "" Then
                sit1 = "Illness"
                sID1 = 5
            ElseIf sit1 <> "" And sit2 = "" Then
                sit2 = "Illness"
                sID2 = 5
            ElseIf sit2 <> "" And sit3 = "" Then
                sit3 = "Illness"
                sID3 = 5
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
            If sit1 = "" Then
                sit1 = "Work"
                sID1 = 7
            ElseIf sit1 <> "" And sit2 = "" Then
                sit2 = "Work"
                sID2 = 7
            ElseIf sit2 <> "" And sit3 = "" Then
                sit3 = "Work"
                sID3 = 7
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
            If sit1 = "" Then
                sit1 = "Unsure"
                sID1 = 3
            ElseIf sit1 <> "" And sit2 = "" Then
                sit2 = "Unsure"
                sID2 = 3
            ElseIf sit2 <> "" And sit3 = "" Then
                sit3 = "Unsure"
                sID3 = 3
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
            If phys1 = "" Then
                phys1 = "Alcohol/Drug Abuse"
            ElseIf phys1 <> "" And phys2 = "" Then
                phys2 = "Alcohol/Drug Abuse"
            ElseIf phys2 <> "" And phys3 = "" Then
                phys3 = "Alcohol/Drug Abuse"
            ElseIf phys3 <> "" And phys4 = "" Then
                phys4 = "Alcohol/Drug Abuse"
            ElseIf phys4 <> "" And phys5 = "" Then
                phys5 = "Alcohol/Drug Abuse"
            Else
                MsgBox("I know things can be rough, but try to stick to just 5 physical symptoms, ok?", MsgBoxStyle.OkOnly, "Error")
            End If
        Else
            PictureBox8.BorderStyle = BorderStyle.None
        End If
    End Sub

    Private Sub PictureBox14_Click(sender As Object, e As EventArgs) Handles PictureBox14.Click
        If PictureBox14.BorderStyle = BorderStyle.None Then
            PictureBox14.BorderStyle = BorderStyle.FixedSingle
            If phys1 = "" Then
                phys1 = "Headache"
            ElseIf phys1 <> "" And phys2 = "" Then
                phys2 = "Headache"
            ElseIf phys2 <> "" And phys3 = "" Then
                phys3 = "Headache"
            ElseIf phys3 <> "" And phys4 = "" Then
                phys4 = "Headache"
            ElseIf phys4 <> "" And phys5 = "" Then
                phys5 = "Headache"
            Else
                MsgBox("I know things can be rough, but try to stick to just 5 physical symtoms, ok?", MsgBoxStyle.OkOnly, "Error")
            End If
        Else
            PictureBox14.BorderStyle = BorderStyle.None
        End If
    End Sub

    Private Sub PictureBox16_Click(sender As Object, e As EventArgs) Handles PictureBox16.Click
        If PictureBox16.BorderStyle = BorderStyle.None Then
            PictureBox16.BorderStyle = BorderStyle.FixedSingle
            If phys1 = "" Then
                phys1 = "Increased Hunger"
            ElseIf phys1 <> "" And phys2 = "" Then
                phys2 = "Increased Hunger"
            ElseIf phys2 <> "" And phys3 = "" Then
                phys3 = "Increased Hunger"
            ElseIf phys3 <> "" And phys4 = "" Then
                phys4 = "Increased Hunger"
            ElseIf phys4 <> "" And phys5 = "" Then
                phys5 = "Increased Hunger"
            Else
                MsgBox("I know things can be rough, but try to stick to just 5 physical symptoms, ok?", MsgBoxStyle.OkOnly, "Error")
            End If
        Else
            PictureBox16.BorderStyle = BorderStyle.None
        End If
    End Sub

    Private Sub PictureBox21_Click(sender As Object, e As EventArgs) Handles PictureBox21.Click
        If PictureBox21.BorderStyle = BorderStyle.None Then
            PictureBox21.BorderStyle = BorderStyle.FixedSingle
            If phys1 = "" Then
                phys1 = "Decreased Hunger"
            ElseIf phys1 <> "" And phys2 = "" Then
                phys2 = "Decreased Hunger"
            ElseIf phys2 <> "" And phys3 = "" Then
                phys3 = "Decreased Hunger"
            ElseIf phys3 <> "" And phys4 = "" Then
                phys4 = "Decreased Hunger"
            ElseIf phys4 <> "" And phys5 = "" Then
                phys5 = "Decreased Hunger"
            Else
                MsgBox("I know things can be rough, but try to stick to just 5 physical symptoms, ok?", MsgBoxStyle.OkOnly, "Error")
            End If
        Else
            PictureBox21.BorderStyle = BorderStyle.None
        End If
    End Sub

    Private Sub PictureBox23_Click(sender As Object, e As EventArgs) Handles PictureBox23.Click
        If PictureBox23.BorderStyle = BorderStyle.None Then
            PictureBox23.BorderStyle = BorderStyle.FixedSingle
            If phys1 = "" Then
                phys1 = "Paranoia"
            ElseIf phys1 <> "" And phys2 = "" Then
                phys2 = "Paranoia"
            ElseIf phys2 <> "" And phys3 = "" Then
                phys3 = "Paranoia"
            ElseIf phys3 <> "" And phys4 = "" Then
                phys4 = "Paranoia"
            ElseIf phys4 <> "" And phys5 = "" Then
                phys5 = "Paranoia"
            Else
                MsgBox("I know things can be rough, but try to stick to just 5 physical symptoms, ok?", MsgBoxStyle.OkOnly, "Error")
            End If
        Else
            PictureBox23.BorderStyle = BorderStyle.None
        End If
    End Sub

    Private Sub PictureBox19_Click(sender As Object, e As EventArgs) Handles PictureBox19.Click
        If PictureBox19.BorderStyle = BorderStyle.None Then
            PictureBox19.BorderStyle = BorderStyle.FixedSingle
            If phys1 = "" Then
                phys1 = "Mood Swing"
            ElseIf phys1 <> "" And phys2 = "" Then
                phys2 = "Mood Swing"
            ElseIf phys2 <> "" And phys3 = "" Then
                phys3 = "Mood Swing"
            ElseIf phys3 <> "" And phys4 = "" Then
                phys4 = "Mood Swing"
            ElseIf phys4 <> "" And phys5 = "" Then
                phys5 = "Mood Swing"
            Else
                MsgBox("I know things can be rough, but try to stick to just 5 physical symptoms, ok?", MsgBoxStyle.OkOnly, "Error")
            End If
        Else
            PictureBox19.BorderStyle = BorderStyle.None
        End If
    End Sub

    Private Sub PictureBox10_Click(sender As Object, e As EventArgs) Handles PictureBox10.Click
        If PictureBox10.BorderStyle = BorderStyle.None Then
            PictureBox10.BorderStyle = BorderStyle.FixedSingle
            If phys1 = "" Then
                phys1 = "Fatigue"
            ElseIf phys1 <> "" And phys2 = "" Then
                phys2 = "Fatigue"
            ElseIf phys2 <> "" And phys3 = "" Then
                phys3 = "Fatigue"
            ElseIf phys3 <> "" And phys4 = "" Then
                phys4 = "Fatigue"
            ElseIf phys4 <> "" And phys5 = "" Then
                phys5 = "Fatigue"
            Else
                MsgBox("I know things can be rough, but try to stick to just 5 physical symptoms, ok?", MsgBoxStyle.OkOnly, "Error")
            End If
        Else
            PictureBox10.BorderStyle = BorderStyle.None
        End If
    End Sub

    Private Sub PictureBox30_Click(sender As Object, e As EventArgs) Handles PictureBox30.Click
        If PictureBox30.BorderStyle = BorderStyle.None Then
            PictureBox30.BorderStyle = BorderStyle.FixedSingle
            If phys1 = "" Then
                phys1 = "Withdraw"
            ElseIf phys1 <> "" And phys2 = "" Then
                phys2 = "Withdraw"
            ElseIf phys2 <> "" And phys3 = "" Then
                phys3 = "Withdraw"
            ElseIf phys3 <> "" And phys4 = "" Then
                phys4 = "Withdraw"
            ElseIf phys4 <> "" And phys5 = "" Then
                phys5 = "Withdraw"
            Else
                MsgBox("I know things can be rough, but try to stick to just 5 physical symptoms, ok?", MsgBoxStyle.OkOnly, "Error")
            End If
        Else
            PictureBox30.BorderStyle = BorderStyle.None
        End If
    End Sub

    Private Sub PictureBox22_Click(sender As Object, e As EventArgs) Handles PictureBox22.Click
        If PictureBox22.BorderStyle = BorderStyle.None Then
            PictureBox22.BorderStyle = BorderStyle.FixedSingle
            If phys1 = "" Then
                phys1 = "Decreased Sex Drive"
            ElseIf phys1 <> "" And phys2 = "" Then
                phys2 = "Decreased Sex Drive"
            ElseIf phys2 <> "" And phys3 = "" Then
                phys3 = "Decreased Sex Drive"
            ElseIf phys3 <> "" And phys4 = "" Then
                phys4 = "Decreased Sex Drive"
            ElseIf phys4 <> "" And phys5 = "" Then
                phys5 = "Decreased Sex Drive"
            Else
                MsgBox("I know things can be rough, but try to stick to just 5 physical symptoms, ok?", MsgBoxStyle.OkOnly, "Error")
            End If
        Else
            PictureBox22.BorderStyle = BorderStyle.None
        End If
    End Sub

    Private Sub PictureBox29_Click(sender As Object, e As EventArgs) Handles PictureBox29.Click
        If PictureBox29.BorderStyle = BorderStyle.None Then
            PictureBox29.BorderStyle = BorderStyle.FixedSingle
            If phys1 = "" Then
                phys1 = "Weight Gain"
            ElseIf phys1 <> "" And phys2 = "" Then
                phys2 = "Weight Gain"
            ElseIf phys2 <> "" And phys3 = "" Then
                phys3 = "Weight Gain"
            ElseIf phys3 <> "" And phys4 = "" Then
                phys4 = "Weight Gain"
            ElseIf phys4 <> "" And phys5 = "" Then
                phys5 = "Weight Gain"
            Else
                MsgBox("I know things can be rough, but try to stick to just 5 physical symptoms, ok?", MsgBoxStyle.OkOnly, "Error")
            End If
        Else
            PictureBox29.BorderStyle = BorderStyle.None
        End If
    End Sub

    Private Sub PictureBox26_Click(sender As Object, e As EventArgs) Handles PictureBox26.Click
        If PictureBox26.BorderStyle = BorderStyle.None Then
            PictureBox26.BorderStyle = BorderStyle.FixedSingle
            If phys1 = "" Then
                phys1 = "Self Harm"
            ElseIf phys1 <> "" And phys2 = "" Then
                phys2 = "Self Harm"
            ElseIf phys2 <> "" And phys3 = "" Then
                phys3 = "Self Harm"
            ElseIf phys3 <> "" And phys4 = "" Then
                phys4 = "Self Harm"
            ElseIf phys4 <> "" And phys5 = "" Then
                phys5 = "Self Harm"
            Else
                MsgBox("I know things can be rough, but try to stick to just 5 physical symptoms, ok?", MsgBoxStyle.OkOnly, "Error")
            End If
        Else
            PictureBox26.BorderStyle = BorderStyle.None
        End If
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click 'HERE IS BUTTON
        If PictureBox26.BorderStyle = BorderStyle.FixedSingle Then
            MsgBox("You said you've been self harming. Please call the Suicide Hotline at 1-800-273-8255 if you need someone to talk to. We're here for you.", MsgBoxStyle.Information, "StableMe")
        ElseIf PictureBox35.BorderStyle = BorderStyle.FixedSingle Then
            MsgBox("You said you have been having suicidal thought. Please call the National Suicide Hotline at 1-800-273-8255 if you need someone to talk to. We're here for you.", MsgBoxStyle.Information, "StableMe")
        Else
            AddLog()
            My.Forms.StableMe1.Show()
            Me.Close()
        End If
    End Sub

    Private Sub PictureBox33_Click(sender As Object, e As EventArgs) Handles PictureBox33.Click
        If PictureBox33.BorderStyle = BorderStyle.None Then
            PictureBox33.BorderStyle = BorderStyle.FixedSingle
            If em1 = "" Then
                em1 = "Drained"
                eID1 = 10
            ElseIf em1 <> "" And em2 = "" Then
                em2 = "Drained"
                eID2 = 10
            ElseIf em2 <> "" And em3 = "" Then
                em3 = "Drained"
                eID3 = 10
            ElseIf em3 <> "" And em4 = "" Then
                em4 = "Drained"
                eID4 = 10
            ElseIf em4 <> "" And em5 = "" Then
                em5 = "Drained"
                eID5 = 10
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
            If em1 = "" Then
                em1 = "Other"
                eID1 = 13
            ElseIf em1 <> "" And em2 = "" Then
                em2 = "Other"
                eID2 = 13
            ElseIf em2 <> "" And em3 = "" Then
                em3 = "Other"
                eID3 = 13
            ElseIf em3 <> "" And em4 = "" Then
                em4 = "Other"
                eID4 = 13
            ElseIf em4 <> "" And em5 = "" Then
                em5 = "Other"
                eID5 = 13
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
            If sit1 = "" Then
                sit1 = "Other"
                sID1 = 1
            ElseIf sit1 <> "" And sit2 = "" Then
                sit2 = "Other"
                sID2 = 1
            ElseIf sit2 <> "" And sit3 = "" Then
                sit3 = "Other"
                sID3 = 1
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
            If phys1 = "" Then
                phys1 = "Suicidal Thoughts"
            ElseIf phys1 <> "" And phys2 = "" Then
                phys2 = "Suicidal Thoughts"
            ElseIf phys2 <> "" And phys3 = "" Then
                phys3 = "Suicidal Thoughts"
            ElseIf phys3 <> "" And phys4 = "" Then
                phys4 = "Suicidal Thoughts"
            ElseIf phys4 <> "" And phys5 = "" Then
                phys5 = "Suicidal Thoughts"
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
            If phys1 = "" Then
                phys1 = "Other"
            ElseIf phys1 <> "" And phys2 = "" Then
                phys2 = "Other"
            ElseIf phys2 <> "" And phys3 = "" Then
                phys3 = "Other"
            ElseIf phys3 <> "" And phys4 = "" Then
                phys4 = "Other"
            ElseIf phys4 <> "" And phys5 = "" Then
                phys5 = "Other"
            Else
                MsgBox("I know things can be rough, but try to stick to just 5 emotions, ok?", MsgBoxStyle.OkOnly, "Error")
            End If
        Else
            PictureBox36.BorderStyle = BorderStyle.None
        End If
    End Sub

    Private Sub AddLog()
        Access.AddParams("@em1", em1)
        Access.AddParams("@em2", em2)
        Access.AddParams("@em3", em3)
        Access.AddParams("@em4", em4)
        Access.AddParams("@em5", em5)
        Access.AddParams("@sit1", sit1)
        Access.AddParams("@sit2", sit2)
        Access.AddParams("@sit3", sit3)
        Access.AddParams("@phys1", phys1)
        Access.AddParams("@phys2", phys2)
        Access.AddParams("@phys3", phys3)
        Access.AddParams("@phys4", phys4)
        Access.AddParams("@phys5", phys5)
        Access.AddParams("@note", note)
        Access.AddParams("@eID1", eID1)
        Access.AddParams("@eID2", eID2)
        Access.AddParams("@eID3", eID3)
        Access.AddParams("@eID4", eID4)
        Access.AddParams("@eID5", eID5)
        Access.AddParams("@sID1", sID1)
        Access.AddParams("@sID2", sID2)
        Access.AddParams("@sID3", sID3)

        Access.ExecQuery("INSERT INTO EmotionLogDB(emotion1, emotion2, emotion3, emotion4, emotion5, " & _
                         "situation1, situation2, situation3, physical1, physical2, physical3, physical4, physical5, " & _
                         "notes, emID1, emID2, emID3, emID4, emID5, sitID1, sitID2, sitID3) " & _
                         "VALUES (@em1, @em2, @em3, @em4, @em5, @sit1, @sit2, @sit3, @phys1, @phys2, @phys3, @phys4, @phys5, " & _
                         " @note, @eID1, @eID2, @eID3, @eID4, @eID5, @sID1, @sID2, @sID3); ")

        If Not String.IsNullOrEmpty(Access.exception) Then
            MsgBox(Access.exception)
        End If
    End Sub
End Class