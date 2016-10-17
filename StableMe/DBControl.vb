Imports System.Data.OleDb
Public Class DBControl
    Private DBCon As New OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=StableMe.accdb;")
    Private DBCmd As OleDbCommand

    Public DBDA As OleDbDataAdapter
    Public DBDT As DataTable
    Public params As New List(Of OleDbParameter)
    Public recordCt As Integer
    Public exception As String

    Public Sub ExecQuery(query As String)
        recordCt = 0
        exception = ""

        Try
            DBCon.Open()
            DBCmd = New OleDbCommand(query, DBCon)
            For Each p As OleDbParameter In params
                DBCmd.Parameters.Add(p)
            Next
            params.Clear()
            DBDT = New DataTable
            DBDA = New OleDbDataAdapter(DBCmd)
            recordCt = DBDA.Fill(DBDT)
        Catch ex As Exception
            exception = ex.Message
        End Try

        If DBCon.State = ConnectionState.Open Then
            DBCon.Close()
        End If
    End Sub

    Public Sub AddParams(name As String, value As Object)
        Dim newParam As New OleDbParameter(name, value)
        params.Add(newParam)
    End Sub
End Class
