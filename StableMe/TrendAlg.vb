Public Class TrendAlg
    Private Access As New DBControl
    Dim emCt As Integer = 0
    Dim sitCt As Integer = 0

    'looks through db
    'finds mode of past 7 entries for emotions
    'finds mode of past 7 entries for situations
    'correlates mode to correct trend combo in db
    'returns EmPlusSitMemo value for the set
    'SELECT EmPlusSitMemo FROM TrendsDB WHERE EmotionID1= " & emCt & "AND SitID1 = " & sitCt & ";"

End Class
