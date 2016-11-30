Imports System.Data.OleDb
Public Class TrendAlg
    Private Acces As New DBControl
    Dim ids(10, 10) As Integer

    Public Class TrendAlg
        'looks through db
        'finds mode of past 7 entries for emotions
        'finds mode of past 7 entries for situations
        'correlates mode to correct trend combo in db
        'returns EmPlusSitMemo value for the set
        'SELECT EmPlusSitMemo FROM TrendsDB WHERE EmotionID1= " & emCt & "AND SitID1 = " & sitCt & ";"

    End Class
End Class
