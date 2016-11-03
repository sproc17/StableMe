Public Class Form8

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If My.Settings.MainUser = True Then
            MaskedTextBox1.Text = My.Settings.Security
        End If
    End Sub

    Private Sub MaskedTextBox1_MaskChanged(sender As Object, e As MaskInputRejectedEventArgs) Handles MaskedTextBox1.MaskInputRejected
        Button1.Enabled() = True
    End Sub
End Class