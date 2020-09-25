VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Player"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public Index As Byte, Name As String
Public Control As Byte
Public Cash As Long, Deposit As Long, Loan As Long, Point As Integer
Public Transport As Byte, CCTime As Byte, Vanish As Byte, Prison As Byte, Hospital As Byte, Sleep As Byte, Dream As Byte, Stay As Byte, Slow As Byte, FriendTime As Byte, FriendPlayer As Byte
Dim Items(0 To 12) As Byte, Cards(0 To 14) As Byte
Private Sub Class_Initialize()
    Index = UBound(Players)
    Dim Transform As Byte
    '名字
    Get #1, 36 + Index * 104, Transform
    Name = Names(Transform)
    '道具
    For i = LBound(Items) To UBound(Items)
        Get #1, 1681 + Index * 15 + i, Items(i)
    Next i
    '卡片
    For i = LBound(Cards) To UBound(Cards)
        Get #1, 1621 + Index * 15 + i, Cards(i)
    Next i
    '现金
    For i = 0 To 3
        Get #1, 45 + Index * 104 + i, Transform
        Cash = Cash + Transform * 256 ^ i
    Next i
    '存款
    For i = 0 To 3
        Get #1, 49 + Index * 104 + i, Transform
        Deposit = Deposit + Transform * 256 ^ i
    Next i
    '贷款
    For i = 0 To 3
        Get #1, 53 + Index * 104 + i, Transform
        Loan = Loan + Transform * 256 ^ i
    Next i
    '点券
    For i = 0 To 1
        Get #1, 65 + Index * 104 + i, Transform
        Point = Point + Transform * 256 ^ i
    Next i
    '状态
    Get #1, 34 + Index * 104, Transport
    Get #1, 68 + Index * 104, Vanish
    Get #1, 69 + Index * 104, Prison
    Get #1, 70 + Index * 104, Hospital
    Get #1, 71 + Index * 104, Sleep
    Get #1, 72 + Index * 104, Dream
    Get #1, 73 + Index * 104, Stay
    Get #1, 74 + Index * 104, Slow
    Get #1, 78 + Index * 104, FriendTime
    Get #1, 82 + Index * 104, FriendPlayer
    Get #1, 38 + Index * 104, Control
    If Transport > 2 Then CCTime = IIf(Transport = 3, 64, (Transport - 3) / 4) Else CCTime = 0
End Sub
Sub ChangeItems(Index As Integer)
    Items(Index) = GetNumber(255, 9, "请输入道具数量")
    MainForm.IValues(Index).Caption = Items(Index)
End Sub
Sub ChangeCards(Index As Integer)
    Cards(Index) = MainForm.CardCombo(Index).ListIndex
End Sub
Sub UpdateForm()
    With MainForm
        '资金
        .CashValue.Caption = Format(Cash, "#,##0")
        .DepositValue.Caption = Format(Deposit, "#,##0")
        .LoanValue.Caption = Format(Loan, "#,##0")
        .PointValue.Caption = Point
        '道具
        For i = LBound(Items) To UBound(Items)
            .IValues(i).Caption = Items(i)
        Next i
        '卡片
        For i = LBound(Cards) To UBound(Cards)
            .CardCombo(i).ListIndex = Cards(i)
        Next i
        '状态
        .TCombo.ListIndex = IIf(Transport <= 2, Transport, 3)
        .VanishValue.Caption = Vanish
        .PrisonValue.Caption = IIf(Prison, 0, Prison)
        .HospitalValue.Caption = IIf(Hospital = 255, 0, Hospital)
        .SleepValue.Caption = Sleep
        .DreamValue.Caption = Dream
        .StayValue.Caption = Stay
        .SlowValue.Caption = Slow
        .FriendTimeValue.Caption = FriendTime
        .ControlCombo.ListIndex = Control - 1
        .ControlCombo.Enabled = IIf(.ControlCombo.ListIndex = -1, False, True)
        .BTLButton.Enabled = IIf(Control, False, True)
        Select Case Control
        Case Is = 0
            .ControlLabel = "死亡"
        Case Is = 1
            .ControlLabel = "人类玩家"
        Case Is = 2
            .ControlLabel = "电脑玩家"
        End Select
        '窗体
        With .FriendCombo
            .Clear
            .AddItem "(无)"
            .ItemData(0) = -1
            For i = LBound(Players) To UBound(Players)
                If Not i = Index Then
                    .AddItem Players(i).Name
                    .ItemData(.ListCount - 1) = Players(i).Index
                End If
            Next i
            For i = 0 To .ListCount - 1
                If .ItemData(i) = FriendPlayer - 1 Then .ListIndex = i
            Next i
        End With
        If Prison Then
            .PrisonValue.Enabled = True
            .HospitalValue.Enabled = False
            .VanishValue.Enabled = False
        ElseIf Hospital Then
            .PrisonValue.Enabled = False
            .HospitalValue.Enabled = True
            .VanishValue.Enabled = False
        Else
            .PrisonValue.Enabled = False
            .HospitalValue.Enabled = False
            .VanishValue.Enabled = True
        End If
    End With
End Sub
Sub BackToLife()
    Dim Transform As Byte, TMoney As Long
    For i = 0 To 3
        Get #1, 9867 + i, Transform
        TMoney = TMoney + Transform * 256 ^ i
    Next i
    Cash = TMoney / 2
    Deposit = TMoney / 2
    Items(0) = 1
    Items(1) = 1
    Items(2) = 1
    Items(3) = 1
    Items(7) = 1
    Items(8) = 1
    Get #1, 9851, Transport
    Control = 1
    UpdateForm
End Sub
Sub Save()
    '资金
    Put #1, 45 + Index * 104, Cash
    Put #1, 49 + Index * 104, Deposit
    Put #1, 53 + Index * 104, Loan
    Put #1, 65 + Index * 104, Point
    '道具
    For i = LBound(Items) To UBound(Items)
        Put #1, 1681 + Index * 15 + i, Items(i)
    Next i
    '卡片
    For i = LBound(Cards) To UBound(Cards)
        Put #1, 1621 + Index * 15 + i, Cards(i)
    Next i
    '状态
    Put #1, 34 + Index * 104, Transport
    Put #1, 68 + Index * 104, Vanish
    Put #1, 69 + Index * 104, Prison
    Put #1, 70 + Index * 104, Hospital
    Put #1, 71 + Index * 104, Sleep
    Put #1, 72 + Index * 104, Dream
    Put #1, 73 + Index * 104, Stay
    Put #1, 74 + Index * 104, Slow
    Put #1, 78 + Index * 104, FriendTime
    Put #1, 82 + Index * 104, FriendPlayer
    Put #1, 38 + Index * 104, Control
End Sub
