Imports System.Data.OleDb
Imports System.Collections.Generic

Public Class TrendAlg
    Private Access As New DBControl
    Dim hECt As Integer
    Dim sECt As Integer
    Dim eECt As Integer
    Dim fECt As Integer
    Dim cECt As Integer
    Dim eEmCt As Integer
    Dim aECt As Integer
    Dim mECt As Integer
    Dim enECt As Integer
    Dim dECt As Integer
    Dim coECt As Integer
    Dim stECt As Integer
    Dim oECt As Integer
    Dim anECt As Integer
    Dim oSCt As Integer
    Dim fSCt As Integer
    Dim uSCt As Integer
    Dim frSCt As Integer
    Dim iSCt As Integer
    Dim sSCt As Integer
    Dim wSCt As Integer
    Dim lSCt As Integer
    Dim eArray(10) As Integer
    Dim sArray(10) As Integer



    Function emMode()
        Dim emNum As Integer

        Dim eStringArray = Access.returnQuery("SELECT TOP 10 Emotion1 FROM EmotionLogDB")
        eArray = Array.ConvertAll(eStringArray, New Converter(Of String, Integer)(AddressOf StringToInteger))
        
        Dim i As Integer
        For i = 0 To i = 9
            If eArray(i) = 1 Then
                hECt = hECt + 1
                i += 1
            ElseIf eArray(i) = 2 Then
                sECt = sECt + 1
                i += 1
            ElseIf eArray(i) = 3 Then
                eECt = eECt + 1
                i += 1
            ElseIf eArray(i) = 4 Then
                fECt = fECt + 1
                i += 1
            ElseIf eArray(i) = 5 Then
                cECt = cECt + 1
                i += 1
            ElseIf eArray(i) = 6 Then
                eEmCt = eEmCt + 1
                i += 1
            ElseIf eArray(i) = 7 Then
                aECt = aECt + 1
                i += 1
            ElseIf eArray(i) = 8 Then
                mECt = mECt + 1
                i += 1
            ElseIf eArray(i) = 9 Then
                enECt = enECt + 1
                i += 1
            ElseIf eArray(i) = 10 Then
                dECt = dECt + 1
                i += 1
            ElseIf eArray(i) = 11 Then
                coECt = coECt + 1
                i += 1
            ElseIf eArray(i) = 12 Then
                stECt = stECt + 1
                i += 1
            ElseIf eArray(i) = 13 Then
                oECt = oECt + 1
                i += 1
            Else
                anECt = anECt + 1
                i += 1
            End If
        Next
        
        Dim k As Integer = 1
        emNum = eArray(0)
        For k = 1 To k = 9
            If emNum < eArray(k) Then
                emNum = eArray(k)
                k += 1
            Else
                k += 1
            End If
        Next
        
        Return emNum
        '
    End Function

    Function sitMode()
        Dim sitNum As Integer

        Dim sStringArray = Access.returnQuery("SELECT TOP 10 Situation1 FROM EmotionLogDB")
        sArray = Array.ConvertAll(sStringArray, New Converter(Of String, Integer)(AddressOf StringToInteger))

        Dim j As Integer = 0
        For j = 0 To j = 9
            If sArray(j) = 1 Then
                oSCt = oSCt + 1
                j += 1
            ElseIf sArray(j) = 2 Then
                fSCt = fSCt + 1
                j += 1
            ElseIf sArray(j) = 3 Then
                uSCt = uSCt + 1
                j += 1
            ElseIf sArray(j) = 4 Then
                frSCt = frSCt + 1
                j += 1
            ElseIf sArray(j) = 5 Then
                iSCt = iSCt + 1
                j += 1
            ElseIf sArray(j) = 6 Then
                sSCt = sSCt + 1
                j += 1
            ElseIf sArray(j) = 7 Then
                wSCt = wSCt + 1
                j += 1
            Else
                lSCt = lSCt + 1
                j += 1
            End If
        Next

        Dim l As Integer = 1
        sitNum = sArray(0)
        For l = 1 To l = 9
            If sitNum < sArray(l) Then
                sitNum = sArray(l)
                l += 1
            Else
                l += 1
            End If
        Next
        Return sitNum
    End Function

    Public Shared Function StringToInteger(ByVal x As String) As Integer
        Dim output As Integer = 0
        Integer.TryParse(x, output)
        Return output
    End Function
End Class
