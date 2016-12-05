Public Class Form1

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If TextBox1.Text = "" Then
            MsgBox("You must enter a username.", MsgBoxStyle.OkOnly, "Error")
        ElseIf TextBox2.Text = "" Then
            MsgBox("You must enter a password.", MsgBoxStyle.OkOnly, "Error")
        ElseIf TextBox3.Text = "" Then
            MsgBox("You must confirm your password.", MsgBoxStyle.OkOnly, "Error")
        Else
            My.Settings.User = TextBox1.Text
            My.Settings.Pass = TextBox2.Text
            My.Settings.Code = False
            My.Settings.Save()
            My.Forms.Form0.Show()
            Me.Close()
        End If
    End Sub

    Private Sub TextBoxCheck(sender As Object, e As EventArgs) Handles TextBox1.TextChanged, TextBox2.TextChanged
        If Not String.IsNullOrWhiteSpace(TextBox1.Text) AndAlso Not String.IsNullOrWhiteSpace(TextBox2.Text) Then
            Button1.Enabled = True
        End If
    End Sub

    Private Sub CheckBox1_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox1.CheckedChanged
        If CheckBox1.Checked Then
            My.Settings.MainUser = True
        End If
    End Sub
End Class