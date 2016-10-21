Public Class Form5
    Private access As New DBControl

    Private Sub Label1_Click(sender As Object, e As EventArgs) Handles Label1.Click
        My.Forms.StableMe1.Show()
        Me.Close()
    End Sub

    Private Sub Form5_Shown(sender As Object, e As EventArgs) Handles Me.Shown
        access.ExecQuery("SELECT journal FROM journalDB ORDER BY id DESC")
        If Not String.IsNullOrEmpty(access.exception) Then
            MsgBox(access.exception) : Exit Sub
        End If
        DataGridView1.DataSource = access.DBDT

    End Sub
End Class