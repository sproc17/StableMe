Public Class Form1

    Private Access As New DBControl

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
            AddUser()
            My.Forms.StableMe1.Show()
            Me.Close()
        End If
    End Sub

    Private Sub TextBoxCheck(sender As Object, e As EventArgs) Handles TextBox1.TextChanged, TextBox2.TextChanged
        If Not String.IsNullOrWhiteSpace(TextBox1.Text) AndAlso Not String.IsNullOrWhiteSpace(TextBox2.Text) AndAlso Not String.IsNullOrWhiteSpace(TextBox3.Text) AndAlso Not String.IsNullOrWhiteSpace(MaskedTextBox1.Text) Then
            Button1.Enabled = True
        End If
    End Sub

    Private Sub AddUser()
        Access.AddParams("@user", TextBox1.Text)
        Access.AddParams("@pass", TextBox2.Text)
        Access.AddParams("@email", TextBox3.Text)
        Access.AddParams("@phone", MaskedTextBox1.Text)

        Access.ExecQuery("INSERT INTO LoginDB (username, [password], email, phone) " & _
                         "VALUES (@user, @pass, @email, @phone); ")

        If Not String.IsNullOrEmpty(Access.exception) Then
            MsgBox(Access.exception)
        End If
    End Sub
End Class