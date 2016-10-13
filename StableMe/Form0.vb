Public Class Form0

    Private Access As New DBControl

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        My.Forms.Form1.Show()
        Me.Close()
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If TextBox1.Text = "" Then
            MsgBox("You must enter a username.", MsgBoxStyle.OkOnly, "Error")
        ElseIf TextBox2.Text = "" Then
            MsgBox("You must enter a password.", MsgBoxStyle.OkOnly, "Error")
        Else
            My.Forms.StableMe1.Show()
            Me.Close()
        End If
    End Sub

    Private Sub TextBoxCheck(sender As Object, e As EventArgs) Handles TextBox1.TextChanged, TextBox2.TextChanged
        If Not String.IsNullOrWhiteSpace(TextBox1.Text) AndAlso Not String.IsNullOrWhiteSpace(TextBox2.Text) Then
            Button1.Enabled = True
        End If
    End Sub
End Class