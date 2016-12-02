Public Class Form6
    Private Access As New DBControl
    Private Alg As New TrendAlg

    Private Sub Label1_Click(sender As Object, e As EventArgs) Handles Label1.Click
        My.Forms.StableMe1.Show()
        Me.Close()
    End Sub

    Private Sub Form6_Shown(sender As Object, e As EventArgs) Handles MyBase.Load
        Dim sit = Alg.sitMode
        Dim em = Alg.emMode

        Access.ExecQuery("SELECT EmPlusSitMemo FROM TrendsDB WHERE EmotionID1= " & em & "AND SitID1 = " & sit & ";")
        If Not String.IsNullOrEmpty(Access.exception) Then
            MsgBox(Access.exception) : Exit Sub
        End If
        DataGridView1.DataSource = Access.DBDT
    End Sub
End Class