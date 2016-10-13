Public Class Form4

    Private Access As New DBControl

    Private Sub Label2_Click(sender As Object, e As EventArgs) Handles Label2.Click
        If RichTextBox1.Text = "" Then
            MsgBox("Your page is empty. Log journal post anyway?", MsgBoxStyle.OkCancel, "Empty Journal Page")
        Else
            My.Forms.StableMe1.Show()
            Me.Close()
        End If
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If RichTextBox1.Text = "" Then
            MsgBox("Are you sure you want to leave?", MsgBoxStyle.OkCancel, "StableMe")
        Else
            AddJournal()
            My.Forms.StableMe1.Show()
            Me.Close()
        End If
    End Sub

    Private Sub AddJournal()
        Access.AddParams("@journal", RichTextBox1.Text)

        Access.ExecQuery("INSERT INTO journaldb (journal) VALUE @journal")
    End Sub
End Class