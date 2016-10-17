Public Class StableMe1

    Private access As New DBControl

    Private Sub Form2_Shown(sender As Object, e As EventArgs) Handles Me.Shown
        access.ExecQuery("SELECT emotion1,emotion2,emotion3,emotion4,emotion5, " & _
                         "situation1,situation2,situation3, " & _
                         "physical1,physical2,physical3,physical4,physical5 " & _
                         "FROM emotionlogdb ORDER BY id DESC")
        If Not String.IsNullOrEmpty(access.exception) Then
            MsgBox(access.exception) : Exit Sub
        End If
        DataGridView1.DataSource = access.DBDT

    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        My.Forms.Form3.Show()
        Me.Close()
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        My.Forms.Form4.Show()
        Me.Close()
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        My.Forms.Form5.Show()
        Me.Close()
    End Sub

    Private Sub PictureBox2_Click(sender As Object, e As EventArgs) Handles PictureBox2.Click
        My.Forms.Form7.Show()
        Me.Close()
    End Sub

    Private Sub PictureBox4_Click(sender As Object, e As EventArgs) Handles PictureBox4.Click
        My.Forms.Form10.Show()
        Me.Close()
    End Sub

    Private Sub PictureBox5_Click(sender As Object, e As EventArgs) Handles PictureBox5.Click
        My.Forms.Form6.Show()
        Me.Close()
    End Sub
End Class
