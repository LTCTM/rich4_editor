VERSION 5.00
Begin VB.Form AmountForm 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   4320
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   2910
   BeginProperty Font 
      Name            =   "Times New Roman"
      Size            =   14.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4320
   ScaleWidth      =   2910
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton AdviceValue 
      Caption         =   "Advice"
      Height          =   435
      Left            =   120
      TabIndex        =   15
      Top             =   1320
      Width           =   1215
   End
   Begin VB.CommandButton Maximum 
      Caption         =   "Max"
      Height          =   435
      Left            =   1560
      TabIndex        =   13
      Top             =   1320
      Width           =   1215
   End
   Begin VB.CommandButton Confirm 
      Caption         =   "确定"
      Height          =   1035
      Left            =   2280
      TabIndex        =   12
      Top             =   3120
      Width           =   495
   End
   Begin VB.CommandButton Clear 
      Caption         =   "C"
      Height          =   435
      Left            =   2280
      TabIndex        =   11
      Top             =   2520
      Width           =   495
   End
   Begin VB.CommandButton Back 
      Caption         =   "←"
      Height          =   435
      Left            =   2280
      TabIndex        =   10
      Top             =   1920
      Width           =   495
   End
   Begin VB.CommandButton Numbers 
      Caption         =   "9"
      Height          =   435
      Index           =   9
      Left            =   1560
      TabIndex        =   9
      Top             =   3120
      Width           =   495
   End
   Begin VB.CommandButton Numbers 
      Caption         =   "8"
      Height          =   435
      Index           =   8
      Left            =   840
      TabIndex        =   8
      Top             =   3120
      Width           =   495
   End
   Begin VB.CommandButton Numbers 
      Caption         =   "7"
      Height          =   435
      Index           =   7
      Left            =   120
      TabIndex        =   7
      Top             =   3120
      Width           =   495
   End
   Begin VB.CommandButton Numbers 
      Caption         =   "6"
      Height          =   435
      Index           =   6
      Left            =   1560
      TabIndex        =   6
      Top             =   2520
      Width           =   495
   End
   Begin VB.CommandButton Numbers 
      Caption         =   "5"
      Height          =   435
      Index           =   5
      Left            =   840
      TabIndex        =   5
      Top             =   2520
      Width           =   495
   End
   Begin VB.CommandButton Numbers 
      Caption         =   "4"
      Height          =   435
      Index           =   4
      Left            =   120
      TabIndex        =   4
      Top             =   2520
      Width           =   495
   End
   Begin VB.CommandButton Numbers 
      Caption         =   "3"
      Height          =   435
      Index           =   3
      Left            =   1560
      TabIndex        =   3
      Top             =   1920
      Width           =   495
   End
   Begin VB.CommandButton Numbers 
      Caption         =   "2"
      Height          =   435
      Index           =   2
      Left            =   840
      TabIndex        =   2
      Top             =   1920
      Width           =   495
   End
   Begin VB.CommandButton Numbers 
      Caption         =   "1"
      Height          =   435
      Index           =   1
      Left            =   120
      TabIndex        =   1
      Top             =   1920
      Width           =   495
   End
   Begin VB.CommandButton Numbers 
      Caption         =   "0"
      Height          =   435
      Index           =   0
      Left            =   840
      TabIndex        =   0
      Top             =   3720
      Width           =   495
   End
   Begin VB.Label Result 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Height          =   435
      Left            =   120
      TabIndex        =   14
      Top             =   120
      Width           =   2655
   End
   Begin VB.Shape Bottom 
      BorderWidth     =   2
      Height          =   375
      Left            =   120
      Top             =   720
      Width           =   2655
   End
   Begin VB.Shape ColorBar 
      BorderStyle     =   0  'Transparent
      BorderWidth     =   2
      FillColor       =   &H0000FF00&
      FillStyle       =   0  'Solid
      Height          =   375
      Left            =   120
      Top             =   720
      Width           =   2655
   End
End
Attribute VB_Name = "AmountForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Max, Advice, Value

Private Sub Form_Load()
    Numbers_Click 0
End Sub
Private Sub Result_Change()
    Dim Number As Double
    With Result
        If Len(.Caption) = 0 Then .Caption = 0
        Number = .Caption
        Number = IIf(Number > Max, Max, Number)
        .Caption = Number
        Value = Number
    End With
    With ColorBar
        If Max = 0 Then
            .Width = 0
        ElseIf Value = Max Then
            .Width = Bottom.Width
            .FillColor = RGB(255, 0, 0)
        ElseIf Value > Advice Then
            .Width = Bottom.Width * (Value / Max)
            .FillColor = RGB(255, 255, 0)
        ElseIf Advice = 0 Then
            .Width = IIf(Value = 0, 0, Bottom.Width * (Value / Max))
            .FillColor = RGB(255, 255, 0)
        Else
            .Width = Bottom.Width * (Value / Advice)
            .FillColor = RGB(0, 255, 0)
        End If
    End With
End Sub
Private Sub Maximum_Click()
    Result = Max
End Sub
Private Sub AdviceValue_Click()
    Result = Advice
End Sub
Private Sub Numbers_Click(Index As Integer)
    Form_KeyPress Asc(Index)
End Sub
Private Sub Back_Click()
    Form_KeyPress 8
End Sub
Private Sub Clear_Click()
    Result = 0
End Sub
Private Sub Form_KeyPress(KeyAscii As Integer)
    With Result
        If KeyAscii = 8 Then
            .Caption = Left(.Caption, Len(.Caption) - 1)
        ElseIf KeyAscii >= 48 And KeyAscii <= 57 Then
            .Caption = .Caption & Chr(KeyAscii)
        End If
    End With
End Sub
Private Sub Confirm_Click()
    Unload Me
End Sub
