Public Class Form1
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If TextBox1.Text = "" Then
            MsgBox("You must enter a username.", MsgBoxStyle.OkOnly, "Error")
        ElseIf TextBox2.Text = "" Then
            MsgBox("You must enter a password.", MsgBoxStyle.OkOnly, "Error")
        ElseIf TextBox3.Text = "" Then
            MsgBox("You must enter an email address.", MsgBoxStyle.OkOnly, "Error")
        ElseIf MaskedTextBox1.Text = "" Then
            MsgBox("You must enter a phone number.", MsgBoxStyle.OkOnly, "Error")
        Else
            My.Forms.StableMe1.Show()
            Me.Close()
        End If
    End Sub
End Class