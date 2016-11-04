Imports System.Data.OleDb
Public Class TrendAlg
    Private Access As New DBControl
    Dim emCt As Integer = 0
    Dim sitCt As Integer = 0
    Dim happyCt As Integer = 0
    Dim sadCt As Integer = 0
    Dim emptyCt As Integer = 0
    Dim frusCt As Integer = 0
    Dim contCt As Integer = 0
    Dim EmoCt As Integer = 0
    Dim anxCt As Integer = 0
    Dim motCt As Integer = 0
    Dim enerCt As Integer = 0
    Dim drnCt As Integer = 0
    Dim confCt As Integer = 0
    Dim strsCt As Integer = 0
    Dim otherECt As Integer = 0
    Dim angCt As Integer = 0
    Dim otherSCt As Integer = 0
    Dim famCt As Integer = 0
    Dim unsCt As Integer = 0
    Dim frndCt As Integer = 0
    Dim illCt As Integer = 0
    Dim schlCt As Integer = 0
    Dim wrkCt As Integer = 0
    Dim lvLfCt As Integer = 0
    Dim nullCt As Integer = 0
    Public exception As String



    Public Function SetEmCounter()
        Try
            Dim Con As OleDbConnection
            Dim StrCon As String = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=StableMe.accdb;"
            Con = New OleDbConnection(StrCon)
            Con.Open()
            Dim CommEm As New OleDbCommand("SELECT EmID1, EmID2, EmID3, EmID4, EmID5 FROM EmotionLogDB", Con)
            CommEm.CommandType = CommandType.Text
            Dim Dr As OleDbDataReader
            Dr = CommEm.ExecuteScalar
            If Dr.HasRows Then
                While Dr.Read
                    If Dr.Equals(1) Then
                        happyCt = happyCt + 1
                    ElseIf Dr.Equals(2) Then
                        sadCt = sadCt + 1
                    ElseIf Dr.Equals(3) Then
                        emptyCt = emptyCt + 1
                    ElseIf Dr.Equals(4) Then
                        frusCt = frusCt + 1
                    ElseIf Dr.Equals(5) Then
                        contCt = contCt + 1
                    ElseIf Dr.Equals(6) Then
                        EmoCt = EmoCt + 1
                    ElseIf Dr.Equals(7) Then
                        anxCt = anxCt + 1
                    ElseIf Dr.Equals(8) Then
                        motCt = motCt + 1
                    ElseIf Dr.Equals(9) Then
                        enerCt = enerCt + 1
                    ElseIf Dr.Equals(10) Then
                        drnCt = drnCt + 1
                    ElseIf Dr.Equals(11) Then
                        confCt = confCt + 1
                    ElseIf Dr.Equals(12) Then
                        strsCt = strsCt + 1
                    ElseIf Dr.Equals(13) Then
                        otherECt = otherECt + 1
                    ElseIf Dr.Equals(14) Then
                        angCt = angCt + 1
                    Else
                        Continue While
                    End If
                End While
            End If

            Dim CommSit As New OleDbCommand("SELECT SitID1, SitID2, SitID3 FROM EmotionLogDB", Con)
            CommSit.CommandType = CommandType.Text
            Dr = CommSit.ExecuteScalar
            If Dr.HasRows Then
                While Dr.Read
                    If Dr.Equals(1) Then
                        otherSCt = otherSCt + 1
                    ElseIf Dr.Equals(2) Then
                        famCt = famCt + 1
                    ElseIf Dr.Equals(3) Then
                        unsCt = unsCt + 1
                    ElseIf Dr.Equals(4) Then
                        frndCt = frndCt + 1
                    ElseIf Dr.Equals(5) Then
                        illCt = illCt + 1
                    ElseIf Dr.Equals(6) Then
                        schlCt = schlCt + 1
                    ElseIf Dr.Equals(7) Then
                        wrkCt = wrkCt + 1
                    ElseIf Dr.Equals(8) Then
                        lvLfCt = lvLfCt + 1
                    Else
                        Continue While
                    End If
                End While
            End If
            Con.Close()
        Catch ex As Exception
            exception = ex.Message
        End Try
        Dim countEm = New Integer() {happyCt, sadCt, emptyCt, frusCt, contCt, EmoCt, anxCt, motCt, enerCt, drnCt, strsCt, otherECt, angCt}
        Dim contSt = New Integer() {otherSCt, famCt, unsCt, frndCt, illCt, schlCt, wrkCt, lvLfCt}
        Dim tempEm As Integer = 0
        Dim modeEm As Integer = countEm[0]
        For i As Integer = 0 To countEm.Length() - 1
            tempEm = countEm[i]
            Dim tempCt As Integer = 0
            For j As Integer = 1 To countEm.Length()
                If tempEm = countEm[j] Then
                    tempCt = tempCt + 1
                End If
                If tempEm > emCt Then
                    modeEm = tempEm
                    emCt = tempCt
                End If
                Return modeEm
            Next
        Next
    End Function

    'looks through db
    'finds mode of past 7 entries for emotions
    'finds mode of past 7 entries for situations
    'correlates mode to correct trend combo in db
    'returns EmPlusSitMemo value for the set
    'SELECT EmPlusSitMemo FROM TrendsDB WHERE EmotionID1= " & emCt & "AND SitID1 = " & sitCt & ";"

End Class
