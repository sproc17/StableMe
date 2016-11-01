Public Class Form0

    Private Access As New DBControl
    Dim uname As String = ""
    Dim pword As String
    Dim username As String = ""
    Dim pass As String

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        My.Forms.Form1.Show()
        Me.Close()
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
            If TextBox1.Text = "" Or TextBox2.Text = "" Then
                MsgBox("You must enter all information.", MsgBoxStyle.OkOnly, "Error")
            Else
                uname = TextBox1.Text
                pword = TextBox2.Text
            Access.ExecQuery("SELECT Password FROM LoginDB WHERE name= '" & uname & "';")
            Try
                pass = Access.PassFinder()
            Catch ex As Exception
                MsgBox("Username does not exist.")
            End Try
            If (pword = pass) Then
                My.Forms.StableMe1.Show()
                Me.Close()
            Else
                MsgBox("login Failed")
                TextBox1.Clear()
                TextBox2.Clear()
            End If
        End If
    End Sub

    Private Sub TextBoxCheck(sender As Object, e As EventArgs) Handles TextBox1.TextChanged, TextBox2.TextChanged
        If Not String.IsNullOrWhiteSpace(TextBox1.Text) AndAlso Not String.IsNullOrWhiteSpace(TextBox2.Text) Then
            Button1.Enabled = True
        End If
    End Sub
End Class