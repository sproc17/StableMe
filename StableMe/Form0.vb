Imports System.Data.SqlClient
Imports System.Data.OleDb

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
        Dim exists As String = Nothing
        Dim DBCon As New OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=StableMe.accdb;")
        Using cmd As New OleDbCommand("SELECT 'X' FROM [LoginDB] WHERE [Username] = @un AND [Passphrase] = @pw", DBCon)
            With cmd
                Access.AddParams("@un", TextBox1.Text)
                Access.AddParams("@pw", TextBox2.Text)
                Try
                    exists = .ExecuteScalar.ToString
                Catch ex As Exception
                    MsgBox(ex.ToString)
                End Try
            End With
        End Using
        DBCon.Close()

        If exists IsNot Nothing AndAlso exists.Equals("X") Then
            My.Forms.StableMe1.Show()
            Me.Close()
        Else
            MsgBox("Login Failed")
            TextBox1.Clear()
            TextBox2.Clear()
        End If
    End Sub

    Private Sub TextBoxCheck(sender As Object, e As EventArgs) Handles TextBox1.TextChanged, TextBox2.TextChanged
        If Not String.IsNullOrWhiteSpace(TextBox1.Text) AndAlso Not String.IsNullOrWhiteSpace(TextBox2.Text) Then
            Button1.Enabled = True
        End If
    End Sub
End Class