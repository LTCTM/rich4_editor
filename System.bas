Attribute VB_Name = "System"
Public OpenName As String, Players() As Player, Actor As Player
Public Names(0 To 11) As String, CardNames(0 To 30) As String
Public PMoney As Long
Sub Main()
    '��ȡ�̶�ֵ
    Open App.Path & "\NameList.txt" For Input As #1
    For i = LBound(Names) To UBound(Names)
        Line Input #1, Names(i)
    Next i
    Close #1
    Open App.Path & "\CardList.txt" For Input As #1
    For i = LBound(CardNames) To UBound(CardNames)
        Line Input #1, CardNames(i)
    Next i
    Close #1
    '����������
    With MainForm
        With .CommonDialog1
            .CancelError = True
            .Filter = "DAT�ļ�(*.dat)|*.dat"
        End With
    '��Ƭ
        For i = .CardCombo.LBound To .CardCombo.UBound
        With .CardCombo(i)
            For j = LBound(CardNames) To UBound(CardNames)
                .AddItem CardNames(j)
            Next j
        End With
        Next i
        .Show
    End With
End Sub

