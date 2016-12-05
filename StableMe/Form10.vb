Public Class Form10

    Private Sub Label1_Click(sender As Object, e As EventArgs) Handles Label1.Click
        My.Forms.StableMe1.Show()
        Me.Close()
    End Sub

    Private Sub LinkLabel1_LinkClicked(sender As Object, e As LinkLabelLinkClickedEventArgs) Handles LinkLabel1.LinkClicked
        My.Forms.Form0.Show()
        Me.Close()
    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedIndexChanged
        If ComboBox1.SelectedIndex = 0 Then
            My.Settings.Security = False
        Else
            My.Settings.Security = True
            My.Forms.Form8.Show()
        End If
    End Sub

    Private Sub Form10_Load(sender As Object, e As EventArgs) Handles Me.Load
        ComboBox1.SelectedIndex = 0
        ComboBox2.SelectedIndex = 0
    End Sub

    Private Sub ComboBox2_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox2.SelectedIndexChanged
        If ComboBox2.SelectedItem = "Green" Then
            Me.BackColor = Color.PaleGreen
            Form0.BackColor = Color.PaleGreen
            Form1.BackColor = Color.PaleGreen
            StableMe1.BackColor = Color.PaleGreen
            Form3.BackColor = Color.PaleGreen
            Form4.BackColor = Color.PaleGreen
            Form5.BackColor = Color.PaleGreen
            Form6.BackColor = Color.PaleGreen
            Form7.BackColor = Color.PaleGreen
            Form8.BackColor = Color.PaleGreen
        ElseIf ComboBox2.SelectedItem = "Blue" Then
            Me.BackColor = Color.PaleTurquoise
            Form0.BackColor = Color.PaleTurquoise
            Form1.BackColor = Color.PaleTurquoise
            StableMe1.BackColor = Color.PaleTurquoise
            Form3.BackColor = Color.PaleTurquoise
            Form4.BackColor = Color.PaleTurquoise
            Form5.BackColor = Color.PaleTurquoise
            Form6.BackColor = Color.PaleTurquoise
            Form7.BackColor = Color.PaleTurquoise
            Form8.BackColor = Color.PaleTurquoise
        ElseIf ComboBox2.SelectedItem = "Yellow" Then
            Me.BackColor = Color.LemonChiffon
            Form0.BackColor = Color.LemonChiffon
            Form1.BackColor = Color.LemonChiffon
            StableMe1.BackColor = Color.LemonChiffon
            Form3.BackColor = Color.LemonChiffon
            Form4.BackColor = Color.LemonChiffon
            Form5.BackColor = Color.LemonChiffon
            Form6.BackColor = Color.LemonChiffon
            Form7.BackColor = Color.LemonChiffon
            Form8.BackColor = Color.LemonChiffon
        ElseIf ComboBox2.SelectedItem = "Pink" Then
            Me.BackColor = Color.Pink
            Form0.BackColor = Color.Pink
            Form1.BackColor = Color.Pink
            StableMe1.BackColor = Color.Pink
            Form3.BackColor = Color.Pink
            Form4.BackColor = Color.Pink
            Form5.BackColor = Color.Pink
            Form6.BackColor = Color.Pink
            Form7.BackColor = Color.Pink
            Form8.BackColor = Color.Pink
        Else
            Me.BackColor = Color.Thistle
            Form0.BackColor = Color.Thistle
            Form1.BackColor = Color.Thistle
            StableMe1.BackColor = Color.Thistle
            Form3.BackColor = Color.Thistle
            Form4.BackColor = Color.Thistle
            Form5.BackColor = Color.Thistle
            Form6.BackColor = Color.Thistle
            Form7.BackColor = Color.Thistle
            Form8.BackColor = Color.Thistle
        End If
    End Sub
End Class