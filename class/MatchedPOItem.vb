Public Class MatchedPOItem
    Public Property Key As POMatchKey
    Public Property Draft As GroupedDraftPO
    Public Property Actual As GroupedActualPO
    Public Property MatchStatus As String ' (เช่น "Matched", "DraftOnly", "ActualOnly")
End Class