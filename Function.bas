Attribute VB_Name = "Function"
Function GetNumber(Optional ByVal Max, Optional ByVal Advice, Optional ByVal Str)
    With AmountForm
        If IsMissing(Max) Then Max = 2 ^ 31 - 1
        If IsMissing(Advice) Then Advice = Max
        If Advice > Max Then Advice = Max
        If IsMissing(Str) Then Str = App.Title
        .Max = Max
        .Advice = Advice
        .Caption = Str
        .Show 1
        GetNumber = .Value
    End With
End Function
Function M(Express, Optional Max, Optional Min)
    If IsMissing(Max) Then Max = Express
    If IsMissing(Min) Then Min = Express
    If Express > Max Then
        M = Max
    ElseIf Express < Min Then
        M = Min
    Else
        M = Express
    End If
End Function
