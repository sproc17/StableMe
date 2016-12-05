Public Class Form8

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If My.Settings.MainUser = True Then
            MaskedTextBox1.Text = My.Settings.Code
        End If
    End Sub
End Class