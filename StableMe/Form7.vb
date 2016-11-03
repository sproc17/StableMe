Public Class Form7

    Private Sub LinkLabel1_LinkClicked(sender As Object, e As LinkLabelLinkClickedEventArgs) Handles LinkLabel1.LinkClicked
        My.Forms.Form1.Show()
        Me.Close()
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If MaskedTextBox1.Text = My.Settings.Security Then
            My.Forms.StableMe1.Show()
            Me.Close()
        Else
            MsgBox("Incorrect Pin", MsgBoxStyle.Information, "Error")
        End If
    End Sub

    Private Sub LinkLabel2_LinkClicked(sender As Object, e As LinkLabelLinkClickedEventArgs) Handles LinkLabel2.LinkClicked
        My.Forms.Form1.Show()
        Me.Close()
    End Sub

    Private Sub MaskedTextBox1_MaskChanged(sender As Object, e As EventArgs) Handles MaskedTextBox1.MaskChanged
        If MaskedTextBox1.MaskCompleted Then
            Button1.Enabled() = True
        End If
    End Sub
End Class