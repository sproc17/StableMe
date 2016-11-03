Public Class Form0

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        My.Forms.Form1.Show()
        Me.Close()
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If TextBox1.Text = My.Settings.User And TextBox2.Text = My.Settings.Pass Then
            My.Forms.StableMe1.Show()
            Me.Close()
        Else
            MsgBox("Username or Password is incorrect", MsgBoxStyle.Information, "Error")
            TextBox1.Clear()
            TextBox2.Clear()
        End If
    End Sub

    Private Sub TextBoxCheck(sender As Object, e As EventArgs) Handles TextBox1.TextChanged, TextBox2.TextChanged
        If Not String.IsNullOrWhiteSpace(TextBox1.Text) AndAlso Not String.IsNullOrWhiteSpace(TextBox2.Text) Then
            Button1.Enabled = True
        End If
    End Sub

    Private Sub Form0_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        If My.Settings.Security = True Then
            My.Forms.Form7.Show()
            Me.Close()
        End If
    End Sub
End Class