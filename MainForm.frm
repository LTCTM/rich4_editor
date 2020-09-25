VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "ComDlg32.OCX"
Begin VB.Form MainForm 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "大富翁修改器"
   ClientHeight    =   7515
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   12165
   BeginProperty Font 
      Name            =   "Times New Roman"
      Size            =   14.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7515
   ScaleWidth      =   12165
   Begin VB.Frame SystemFrame 
      Caption         =   "系统"
      Enabled         =   0   'False
      Height          =   6015
      Left            =   8040
      TabIndex        =   9
      Top             =   1320
      Width           =   3975
      Begin MSComDlg.CommonDialog CommonDialog1 
         Left            =   600
         Top             =   3240
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.Label Label6 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "启动资金："
         Height          =   435
         Index           =   1
         Left            =   120
         TabIndex        =   11
         ToolTipText     =   "影响物价指数"
         Top             =   360
         Width           =   1575
      End
      Begin VB.Label PMoneyValue 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   435
         Left            =   1920
         TabIndex        =   10
         Top             =   360
         Width           =   1815
      End
   End
   Begin VB.Frame PlayerFrame 
      Caption         =   "角色"
      Enabled         =   0   'False
      Height          =   6015
      Left            =   120
      TabIndex        =   5
      Top             =   1320
      Width           =   7695
      Begin VB.CommandButton BTLButton 
         Caption         =   "复活"
         Height          =   435
         Left            =   5520
         TabIndex        =   12
         Top             =   360
         Width           =   855
      End
      Begin VB.ComboBox PlayerCombo 
         Height          =   435
         ItemData        =   "MainForm.frx":0000
         Left            =   1920
         List            =   "MainForm.frx":0002
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   360
         Width           =   1695
      End
      Begin TabDlg.SSTab SSTab1 
         Height          =   4455
         Left            =   120
         TabIndex        =   13
         Top             =   840
         Width           =   7335
         _ExtentX        =   12938
         _ExtentY        =   7858
         _Version        =   393216
         Style           =   1
         Tabs            =   4
         TabsPerRow      =   4
         TabHeight       =   661
         TabCaption(0)   =   "资金"
         TabPicture(0)   =   "MainForm.frx":0004
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "Label3(3)"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).Control(1)=   "PointValue"
         Tab(0).Control(1).Enabled=   0   'False
         Tab(0).Control(2)=   "Label3(2)"
         Tab(0).Control(2).Enabled=   0   'False
         Tab(0).Control(3)=   "LoanValue"
         Tab(0).Control(3).Enabled=   0   'False
         Tab(0).Control(4)=   "Label3(1)"
         Tab(0).Control(4).Enabled=   0   'False
         Tab(0).Control(5)=   "DepositValue"
         Tab(0).Control(5).Enabled=   0   'False
         Tab(0).Control(6)=   "Label3(0)"
         Tab(0).Control(6).Enabled=   0   'False
         Tab(0).Control(7)=   "CashValue"
         Tab(0).Control(7).Enabled=   0   'False
         Tab(0).Control(8)=   "LoanButton(1)"
         Tab(0).Control(8).Enabled=   0   'False
         Tab(0).Control(9)=   "LoanButton(0)"
         Tab(0).Control(9).Enabled=   0   'False
         Tab(0).Control(10)=   "LoanButton(2)"
         Tab(0).Control(10).Enabled=   0   'False
         Tab(0).Control(11)=   "DepositButton(1)"
         Tab(0).Control(11).Enabled=   0   'False
         Tab(0).Control(12)=   "DepositButton(0)"
         Tab(0).Control(12).Enabled=   0   'False
         Tab(0).Control(13)=   "DepositButton(2)"
         Tab(0).Control(13).Enabled=   0   'False
         Tab(0).Control(14)=   "CashButton(1)"
         Tab(0).Control(14).Enabled=   0   'False
         Tab(0).Control(15)=   "CashButton(0)"
         Tab(0).Control(15).Enabled=   0   'False
         Tab(0).Control(16)=   "CashButton(2)"
         Tab(0).Control(16).Enabled=   0   'False
         Tab(0).ControlCount=   17
         TabCaption(1)   =   "道具"
         TabPicture(1)   =   "MainForm.frx":0020
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "Label2(0)"
         Tab(1).Control(0).Enabled=   0   'False
         Tab(1).Control(1)=   "Label2(1)"
         Tab(1).Control(1).Enabled=   0   'False
         Tab(1).Control(2)=   "Label2(2)"
         Tab(1).Control(2).Enabled=   0   'False
         Tab(1).Control(3)=   "Label2(3)"
         Tab(1).Control(3).Enabled=   0   'False
         Tab(1).Control(4)=   "Label2(4)"
         Tab(1).Control(4).Enabled=   0   'False
         Tab(1).Control(5)=   "Label2(5)"
         Tab(1).Control(5).Enabled=   0   'False
         Tab(1).Control(6)=   "Label2(6)"
         Tab(1).Control(6).Enabled=   0   'False
         Tab(1).Control(7)=   "Label2(7)"
         Tab(1).Control(7).Enabled=   0   'False
         Tab(1).Control(8)=   "Label2(12)"
         Tab(1).Control(8).Enabled=   0   'False
         Tab(1).Control(9)=   "Label2(11)"
         Tab(1).Control(9).Enabled=   0   'False
         Tab(1).Control(10)=   "Label2(10)"
         Tab(1).Control(10).Enabled=   0   'False
         Tab(1).Control(11)=   "Label2(9)"
         Tab(1).Control(11).Enabled=   0   'False
         Tab(1).Control(12)=   "Label2(8)"
         Tab(1).Control(12).Enabled=   0   'False
         Tab(1).Control(13)=   "IValues(0)"
         Tab(1).Control(13).Enabled=   0   'False
         Tab(1).Control(14)=   "IValues(1)"
         Tab(1).Control(14).Enabled=   0   'False
         Tab(1).Control(15)=   "IValues(2)"
         Tab(1).Control(15).Enabled=   0   'False
         Tab(1).Control(16)=   "IValues(3)"
         Tab(1).Control(16).Enabled=   0   'False
         Tab(1).Control(17)=   "IValues(4)"
         Tab(1).Control(17).Enabled=   0   'False
         Tab(1).Control(18)=   "IValues(5)"
         Tab(1).Control(18).Enabled=   0   'False
         Tab(1).Control(19)=   "IValues(6)"
         Tab(1).Control(19).Enabled=   0   'False
         Tab(1).Control(20)=   "IValues(7)"
         Tab(1).Control(20).Enabled=   0   'False
         Tab(1).Control(21)=   "IValues(8)"
         Tab(1).Control(21).Enabled=   0   'False
         Tab(1).Control(22)=   "IValues(9)"
         Tab(1).Control(22).Enabled=   0   'False
         Tab(1).Control(23)=   "IValues(10)"
         Tab(1).Control(23).Enabled=   0   'False
         Tab(1).Control(24)=   "IValues(11)"
         Tab(1).Control(24).Enabled=   0   'False
         Tab(1).Control(25)=   "IValues(12)"
         Tab(1).Control(25).Enabled=   0   'False
         Tab(1).ControlCount=   26
         TabCaption(2)   =   "卡片"
         TabPicture(2)   =   "MainForm.frx":003C
         Tab(2).ControlEnabled=   0   'False
         Tab(2).Control(0)=   "CardCombo(14)"
         Tab(2).Control(0).Enabled=   0   'False
         Tab(2).Control(1)=   "CardCombo(13)"
         Tab(2).Control(1).Enabled=   0   'False
         Tab(2).Control(2)=   "CardCombo(12)"
         Tab(2).Control(2).Enabled=   0   'False
         Tab(2).Control(3)=   "CardCombo(11)"
         Tab(2).Control(3).Enabled=   0   'False
         Tab(2).Control(4)=   "CardCombo(10)"
         Tab(2).Control(4).Enabled=   0   'False
         Tab(2).Control(5)=   "CardCombo(9)"
         Tab(2).Control(5).Enabled=   0   'False
         Tab(2).Control(6)=   "CardCombo(8)"
         Tab(2).Control(6).Enabled=   0   'False
         Tab(2).Control(7)=   "CardCombo(7)"
         Tab(2).Control(7).Enabled=   0   'False
         Tab(2).Control(8)=   "CardCombo(6)"
         Tab(2).Control(8).Enabled=   0   'False
         Tab(2).Control(9)=   "CardCombo(5)"
         Tab(2).Control(9).Enabled=   0   'False
         Tab(2).Control(10)=   "CardCombo(4)"
         Tab(2).Control(10).Enabled=   0   'False
         Tab(2).Control(11)=   "CardCombo(3)"
         Tab(2).Control(11).Enabled=   0   'False
         Tab(2).Control(12)=   "CardCombo(2)"
         Tab(2).Control(12).Enabled=   0   'False
         Tab(2).Control(13)=   "CardCombo(1)"
         Tab(2).Control(13).Enabled=   0   'False
         Tab(2).Control(14)=   "CardCombo(0)"
         Tab(2).Control(14).Enabled=   0   'False
         Tab(2).ControlCount=   15
         TabCaption(3)   =   "状态"
         TabPicture(3)   =   "MainForm.frx":0058
         Tab(3).ControlEnabled=   0   'False
         Tab(3).Control(0)=   "Label4(10)"
         Tab(3).Control(0).Enabled=   0   'False
         Tab(3).Control(1)=   "Label4(9)"
         Tab(3).Control(1).Enabled=   0   'False
         Tab(3).Control(2)=   "FriendTimeValue"
         Tab(3).Control(2).Enabled=   0   'False
         Tab(3).Control(3)=   "Label4(8)"
         Tab(3).Control(3).Enabled=   0   'False
         Tab(3).Control(4)=   "SleepValue"
         Tab(3).Control(4).Enabled=   0   'False
         Tab(3).Control(5)=   "Label4(7)"
         Tab(3).Control(5).Enabled=   0   'False
         Tab(3).Control(6)=   "SlowValue"
         Tab(3).Control(6).Enabled=   0   'False
         Tab(3).Control(7)=   "Label4(6)"
         Tab(3).Control(7).Enabled=   0   'False
         Tab(3).Control(8)=   "StayValue"
         Tab(3).Control(8).Enabled=   0   'False
         Tab(3).Control(9)=   "Label4(5)"
         Tab(3).Control(9).Enabled=   0   'False
         Tab(3).Control(10)=   "DreamValue"
         Tab(3).Control(10).Enabled=   0   'False
         Tab(3).Control(11)=   "Label4(4)"
         Tab(3).Control(11).Enabled=   0   'False
         Tab(3).Control(12)=   "HospitalValue"
         Tab(3).Control(12).Enabled=   0   'False
         Tab(3).Control(13)=   "Label4(3)"
         Tab(3).Control(13).Enabled=   0   'False
         Tab(3).Control(14)=   "PrisonValue"
         Tab(3).Control(14).Enabled=   0   'False
         Tab(3).Control(15)=   "Label4(2)"
         Tab(3).Control(15).Enabled=   0   'False
         Tab(3).Control(16)=   "VanishValue"
         Tab(3).Control(16).Enabled=   0   'False
         Tab(3).Control(17)=   "Label4(1)"
         Tab(3).Control(17).Enabled=   0   'False
         Tab(3).Control(18)=   "Label4(0)"
         Tab(3).Control(18).Enabled=   0   'False
         Tab(3).Control(19)=   "Label4(11)"
         Tab(3).Control(19).Enabled=   0   'False
         Tab(3).Control(20)=   "TValue"
         Tab(3).Control(20).Enabled=   0   'False
         Tab(3).Control(21)=   "ControlCombo"
         Tab(3).Control(21).Enabled=   0   'False
         Tab(3).Control(22)=   "FriendCombo"
         Tab(3).Control(22).Enabled=   0   'False
         Tab(3).Control(23)=   "TCombo"
         Tab(3).Control(23).Enabled=   0   'False
         Tab(3).ControlCount=   24
         Begin VB.ComboBox TCombo 
            Height          =   435
            ItemData        =   "MainForm.frx":0074
            Left            =   -69360
            List            =   "MainForm.frx":0084
            Style           =   2  'Dropdown List
            TabIndex        =   40
            Top             =   480
            Width           =   1575
         End
         Begin VB.ComboBox FriendCombo 
            Height          =   435
            ItemData        =   "MainForm.frx":00A2
            Left            =   -69360
            List            =   "MainForm.frx":00A4
            Style           =   2  'Dropdown List
            TabIndex        =   39
            Top             =   1440
            Width           =   1575
         End
         Begin VB.ComboBox CardCombo 
            Height          =   435
            Index           =   0
            ItemData        =   "MainForm.frx":00A6
            Left            =   -74880
            List            =   "MainForm.frx":00A8
            Style           =   2  'Dropdown List
            TabIndex        =   38
            Top             =   480
            Width           =   1335
         End
         Begin VB.ComboBox CardCombo 
            Height          =   435
            Index           =   1
            ItemData        =   "MainForm.frx":00AA
            Left            =   -73440
            List            =   "MainForm.frx":00AC
            Style           =   2  'Dropdown List
            TabIndex        =   37
            Top             =   480
            Width           =   1335
         End
         Begin VB.ComboBox CardCombo 
            Height          =   435
            Index           =   2
            ItemData        =   "MainForm.frx":00AE
            Left            =   -72000
            List            =   "MainForm.frx":00B0
            Style           =   2  'Dropdown List
            TabIndex        =   36
            Top             =   480
            Width           =   1335
         End
         Begin VB.ComboBox CardCombo 
            Height          =   435
            Index           =   3
            ItemData        =   "MainForm.frx":00B2
            Left            =   -70560
            List            =   "MainForm.frx":00B4
            Style           =   2  'Dropdown List
            TabIndex        =   35
            Top             =   480
            Width           =   1335
         End
         Begin VB.ComboBox CardCombo 
            Height          =   435
            Index           =   4
            ItemData        =   "MainForm.frx":00B6
            Left            =   -69120
            List            =   "MainForm.frx":00B8
            Style           =   2  'Dropdown List
            TabIndex        =   34
            Top             =   480
            Width           =   1335
         End
         Begin VB.ComboBox CardCombo 
            Height          =   435
            Index           =   5
            ItemData        =   "MainForm.frx":00BA
            Left            =   -74880
            List            =   "MainForm.frx":00BC
            Style           =   2  'Dropdown List
            TabIndex        =   33
            Top             =   960
            Width           =   1335
         End
         Begin VB.ComboBox CardCombo 
            Height          =   435
            Index           =   6
            ItemData        =   "MainForm.frx":00BE
            Left            =   -73440
            List            =   "MainForm.frx":00C0
            Style           =   2  'Dropdown List
            TabIndex        =   32
            Top             =   960
            Width           =   1335
         End
         Begin VB.ComboBox CardCombo 
            Height          =   435
            Index           =   7
            ItemData        =   "MainForm.frx":00C2
            Left            =   -72000
            List            =   "MainForm.frx":00C4
            Style           =   2  'Dropdown List
            TabIndex        =   31
            Top             =   960
            Width           =   1335
         End
         Begin VB.ComboBox CardCombo 
            Height          =   435
            Index           =   8
            ItemData        =   "MainForm.frx":00C6
            Left            =   -70560
            List            =   "MainForm.frx":00C8
            Style           =   2  'Dropdown List
            TabIndex        =   30
            Top             =   960
            Width           =   1335
         End
         Begin VB.ComboBox CardCombo 
            Height          =   435
            Index           =   9
            ItemData        =   "MainForm.frx":00CA
            Left            =   -69120
            List            =   "MainForm.frx":00CC
            Style           =   2  'Dropdown List
            TabIndex        =   29
            Top             =   960
            Width           =   1335
         End
         Begin VB.ComboBox CardCombo 
            Height          =   435
            Index           =   10
            ItemData        =   "MainForm.frx":00CE
            Left            =   -74880
            List            =   "MainForm.frx":00D0
            Style           =   2  'Dropdown List
            TabIndex        =   28
            Top             =   1440
            Width           =   1335
         End
         Begin VB.ComboBox CardCombo 
            Height          =   435
            Index           =   11
            ItemData        =   "MainForm.frx":00D2
            Left            =   -73440
            List            =   "MainForm.frx":00D4
            Style           =   2  'Dropdown List
            TabIndex        =   27
            Top             =   1440
            Width           =   1335
         End
         Begin VB.ComboBox CardCombo 
            Height          =   435
            Index           =   12
            ItemData        =   "MainForm.frx":00D6
            Left            =   -72000
            List            =   "MainForm.frx":00D8
            Style           =   2  'Dropdown List
            TabIndex        =   26
            Top             =   1440
            Width           =   1335
         End
         Begin VB.ComboBox CardCombo 
            Height          =   435
            Index           =   13
            ItemData        =   "MainForm.frx":00DA
            Left            =   -70560
            List            =   "MainForm.frx":00DC
            Style           =   2  'Dropdown List
            TabIndex        =   25
            Top             =   1440
            Width           =   1335
         End
         Begin VB.ComboBox CardCombo 
            Height          =   435
            Index           =   14
            ItemData        =   "MainForm.frx":00DE
            Left            =   -69120
            List            =   "MainForm.frx":00E0
            Style           =   2  'Dropdown List
            TabIndex        =   24
            Top             =   1440
            Width           =   1335
         End
         Begin VB.CommandButton CashButton 
            Caption         =   "+1000万"
            Height          =   435
            Index           =   2
            Left            =   6000
            TabIndex        =   23
            Top             =   480
            Width           =   1215
         End
         Begin VB.CommandButton CashButton 
            Caption         =   "+10万"
            Height          =   435
            Index           =   0
            Left            =   3360
            TabIndex        =   22
            Top             =   480
            Width           =   1095
         End
         Begin VB.CommandButton CashButton 
            Caption         =   "+100万"
            Height          =   435
            Index           =   1
            Left            =   4680
            TabIndex        =   21
            Top             =   480
            Width           =   1095
         End
         Begin VB.CommandButton DepositButton 
            Caption         =   "+1000万"
            Height          =   435
            Index           =   2
            Left            =   6000
            TabIndex        =   20
            Top             =   960
            Width           =   1215
         End
         Begin VB.CommandButton DepositButton 
            Caption         =   "+10万"
            Height          =   435
            Index           =   0
            Left            =   3360
            TabIndex        =   19
            Top             =   960
            Width           =   1095
         End
         Begin VB.CommandButton DepositButton 
            Caption         =   "+100万"
            Height          =   435
            Index           =   1
            Left            =   4680
            TabIndex        =   18
            Top             =   960
            Width           =   1095
         End
         Begin VB.CommandButton LoanButton 
            Caption         =   "+1000万"
            Height          =   435
            Index           =   2
            Left            =   6000
            TabIndex        =   17
            Top             =   1440
            Width           =   1215
         End
         Begin VB.CommandButton LoanButton 
            Caption         =   "+10万"
            Height          =   435
            Index           =   0
            Left            =   3360
            TabIndex        =   16
            Top             =   1440
            Width           =   1095
         End
         Begin VB.CommandButton LoanButton 
            Caption         =   "+100万"
            Height          =   435
            Index           =   1
            Left            =   4680
            TabIndex        =   15
            Top             =   1440
            Width           =   1095
         End
         Begin VB.ComboBox ControlCombo 
            Height          =   435
            ItemData        =   "MainForm.frx":00E2
            Left            =   -73080
            List            =   "MainForm.frx":00EC
            Style           =   2  'Dropdown List
            TabIndex        =   14
            Top             =   480
            Width           =   1575
         End
         Begin VB.Label TValue 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   1  'Fixed Single
            Height          =   435
            Left            =   -69360
            TabIndex        =   95
            Top             =   960
            Width           =   1575
         End
         Begin VB.Label Label4 
            BorderStyle     =   1  'Fixed Single
            Caption         =   "工程车时间"
            Height          =   435
            Index           =   11
            Left            =   -71160
            TabIndex        =   94
            Top             =   960
            Width           =   1575
         End
         Begin VB.Label Label2 
            BorderStyle     =   1  'Fixed Single
            Caption         =   "机器娃娃："
            Height          =   435
            Index           =   0
            Left            =   -74880
            TabIndex        =   93
            Top             =   480
            Width           =   1575
         End
         Begin VB.Label Label2 
            BorderStyle     =   1  'Fixed Single
            Caption         =   "路障："
            Height          =   435
            Index           =   1
            Left            =   -74880
            TabIndex        =   92
            Top             =   975
            Width           =   1575
         End
         Begin VB.Label Label2 
            BorderStyle     =   1  'Fixed Single
            Caption         =   "地雷："
            Height          =   435
            Index           =   2
            Left            =   -74880
            TabIndex        =   91
            Top             =   1455
            Width           =   1575
         End
         Begin VB.Label Label2 
            BorderStyle     =   1  'Fixed Single
            Caption         =   "定时炸弹："
            Height          =   435
            Index           =   3
            Left            =   -74880
            TabIndex        =   90
            Top             =   1935
            Width           =   1575
         End
         Begin VB.Label Label2 
            BorderStyle     =   1  'Fixed Single
            Caption         =   "机车："
            Height          =   435
            Index           =   4
            Left            =   -74880
            TabIndex        =   89
            Top             =   2415
            Width           =   1575
         End
         Begin VB.Label Label2 
            BorderStyle     =   1  'Fixed Single
            Caption         =   "汽车："
            Height          =   435
            Index           =   5
            Left            =   -74880
            TabIndex        =   88
            Top             =   2895
            Width           =   1575
         End
         Begin VB.Label Label2 
            BorderStyle     =   1  'Fixed Single
            Caption         =   "飞弹："
            Height          =   435
            Index           =   6
            Left            =   -74880
            TabIndex        =   87
            Top             =   3375
            Width           =   1575
         End
         Begin VB.Label Label2 
            BorderStyle     =   1  'Fixed Single
            Caption         =   "遥控骰子："
            Height          =   435
            Index           =   7
            Left            =   -74880
            TabIndex        =   86
            Top             =   3855
            Width           =   1575
         End
         Begin VB.Label Label2 
            BorderStyle     =   1  'Fixed Single
            Caption         =   "核子飞弹："
            Height          =   435
            Index           =   12
            Left            =   -70200
            TabIndex        =   85
            Top             =   2415
            Width           =   1575
         End
         Begin VB.Label Label2 
            BorderStyle     =   1  'Fixed Single
            Caption         =   "工程车："
            Height          =   435
            Index           =   11
            Left            =   -70200
            TabIndex        =   84
            Top             =   1935
            Width           =   1575
         End
         Begin VB.Label Label2 
            BorderStyle     =   1  'Fixed Single
            Caption         =   "传送机："
            Height          =   435
            Index           =   10
            Left            =   -70200
            TabIndex        =   83
            Top             =   1455
            Width           =   1575
         End
         Begin VB.Label Label2 
            BorderStyle     =   1  'Fixed Single
            Caption         =   "时光机："
            Height          =   435
            Index           =   9
            Left            =   -70200
            TabIndex        =   82
            Top             =   975
            Width           =   1575
         End
         Begin VB.Label Label2 
            BorderStyle     =   1  'Fixed Single
            Caption         =   "机器工人："
            Height          =   435
            Index           =   8
            Left            =   -70200
            TabIndex        =   81
            Top             =   495
            Width           =   1575
         End
         Begin VB.Label IValues 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   1  'Fixed Single
            Height          =   435
            Index           =   0
            Left            =   -73080
            TabIndex        =   80
            Top             =   495
            Width           =   615
         End
         Begin VB.Label IValues 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   1  'Fixed Single
            Height          =   435
            Index           =   1
            Left            =   -73080
            TabIndex        =   79
            Top             =   975
            Width           =   615
         End
         Begin VB.Label IValues 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   1  'Fixed Single
            Height          =   435
            Index           =   2
            Left            =   -73080
            TabIndex        =   78
            Top             =   1455
            Width           =   615
         End
         Begin VB.Label IValues 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   1  'Fixed Single
            Height          =   435
            Index           =   3
            Left            =   -73080
            TabIndex        =   77
            Top             =   1935
            Width           =   615
         End
         Begin VB.Label IValues 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   1  'Fixed Single
            Height          =   435
            Index           =   4
            Left            =   -73080
            TabIndex        =   76
            Top             =   2415
            Width           =   615
         End
         Begin VB.Label IValues 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   1  'Fixed Single
            Height          =   435
            Index           =   5
            Left            =   -73080
            TabIndex        =   75
            Top             =   2895
            Width           =   615
         End
         Begin VB.Label IValues 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   1  'Fixed Single
            Height          =   435
            Index           =   6
            Left            =   -73080
            TabIndex        =   74
            Top             =   3375
            Width           =   615
         End
         Begin VB.Label IValues 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   1  'Fixed Single
            Height          =   435
            Index           =   7
            Left            =   -73080
            TabIndex        =   73
            Top             =   3855
            Width           =   615
         End
         Begin VB.Label IValues 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   1  'Fixed Single
            Height          =   435
            Index           =   8
            Left            =   -68400
            TabIndex        =   72
            Top             =   495
            Width           =   615
         End
         Begin VB.Label IValues 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   1  'Fixed Single
            Height          =   435
            Index           =   9
            Left            =   -68400
            TabIndex        =   71
            Top             =   975
            Width           =   615
         End
         Begin VB.Label IValues 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   1  'Fixed Single
            Height          =   435
            Index           =   10
            Left            =   -68400
            TabIndex        =   70
            Top             =   1455
            Width           =   615
         End
         Begin VB.Label IValues 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   1  'Fixed Single
            Height          =   435
            Index           =   11
            Left            =   -68400
            TabIndex        =   69
            Top             =   1935
            Width           =   615
         End
         Begin VB.Label IValues 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   1  'Fixed Single
            Height          =   435
            Index           =   12
            Left            =   -68400
            TabIndex        =   68
            Top             =   2415
            Width           =   615
         End
         Begin VB.Label CashValue 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   1  'Fixed Single
            Height          =   435
            Left            =   1320
            TabIndex        =   67
            Top             =   495
            Width           =   1815
         End
         Begin VB.Label Label3 
            BorderStyle     =   1  'Fixed Single
            Caption         =   "现金："
            Height          =   435
            Index           =   0
            Left            =   120
            TabIndex        =   66
            Top             =   495
            Width           =   975
         End
         Begin VB.Label DepositValue 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   1  'Fixed Single
            Height          =   435
            Left            =   1320
            TabIndex        =   65
            Top             =   975
            Width           =   1815
         End
         Begin VB.Label Label3 
            BorderStyle     =   1  'Fixed Single
            Caption         =   "存款："
            Height          =   435
            Index           =   1
            Left            =   120
            TabIndex        =   64
            Top             =   975
            Width           =   975
         End
         Begin VB.Label LoanValue 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   1  'Fixed Single
            Height          =   435
            Left            =   1320
            TabIndex        =   63
            Top             =   1440
            Width           =   1815
         End
         Begin VB.Label Label3 
            BorderStyle     =   1  'Fixed Single
            Caption         =   "贷款："
            Height          =   435
            Index           =   2
            Left            =   120
            TabIndex        =   62
            Top             =   1455
            Width           =   975
         End
         Begin VB.Label PointValue 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   1  'Fixed Single
            Height          =   435
            Left            =   1320
            TabIndex        =   61
            Top             =   1935
            Width           =   1815
         End
         Begin VB.Label Label3 
            BorderStyle     =   1  'Fixed Single
            Caption         =   "点券："
            Height          =   435
            Index           =   3
            Left            =   120
            TabIndex        =   60
            Top             =   1935
            Width           =   975
         End
         Begin VB.Label Label4 
            BorderStyle     =   1  'Fixed Single
            Caption         =   "交通工具："
            Height          =   435
            Index           =   0
            Left            =   -71160
            TabIndex        =   59
            Top             =   480
            Width           =   1575
         End
         Begin VB.Label Label4 
            BorderStyle     =   1  'Fixed Single
            Caption         =   "消失："
            Height          =   435
            Index           =   1
            Left            =   -74880
            TabIndex        =   58
            Top             =   960
            Width           =   1575
         End
         Begin VB.Label VanishValue 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   1  'Fixed Single
            Height          =   435
            Left            =   -73080
            TabIndex        =   57
            Top             =   960
            Width           =   1575
         End
         Begin VB.Label Label4 
            BorderStyle     =   1  'Fixed Single
            Caption         =   "坐牢："
            Height          =   435
            Index           =   2
            Left            =   -74880
            TabIndex        =   56
            Top             =   1440
            Width           =   1575
         End
         Begin VB.Label PrisonValue 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   1  'Fixed Single
            Height          =   435
            Left            =   -73080
            TabIndex        =   55
            Top             =   1440
            Width           =   1575
         End
         Begin VB.Label Label4 
            BorderStyle     =   1  'Fixed Single
            Caption         =   "住院："
            Height          =   435
            Index           =   3
            Left            =   -74880
            TabIndex        =   54
            Top             =   1920
            Width           =   1575
         End
         Begin VB.Label HospitalValue 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   1  'Fixed Single
            Height          =   435
            Left            =   -73080
            TabIndex        =   53
            Top             =   1920
            Width           =   1575
         End
         Begin VB.Label Label4 
            BorderStyle     =   1  'Fixed Single
            Caption         =   "梦游："
            Height          =   435
            Index           =   4
            Left            =   -74880
            TabIndex        =   52
            Top             =   2880
            Width           =   1575
         End
         Begin VB.Label DreamValue 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   1  'Fixed Single
            Height          =   435
            Left            =   -73080
            TabIndex        =   51
            Top             =   2880
            Width           =   1575
         End
         Begin VB.Label Label4 
            BorderStyle     =   1  'Fixed Single
            Caption         =   "停留："
            Height          =   435
            Index           =   5
            Left            =   -74880
            TabIndex        =   50
            Top             =   3360
            Width           =   1575
         End
         Begin VB.Label StayValue 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   1  'Fixed Single
            Height          =   435
            Left            =   -73080
            TabIndex        =   49
            Top             =   3360
            Width           =   1575
         End
         Begin VB.Label Label4 
            BorderStyle     =   1  'Fixed Single
            Caption         =   "乌龟："
            Height          =   435
            Index           =   6
            Left            =   -74880
            TabIndex        =   48
            Top             =   3840
            Width           =   1575
         End
         Begin VB.Label SlowValue 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   1  'Fixed Single
            Height          =   435
            Left            =   -73080
            TabIndex        =   47
            Top             =   3840
            Width           =   1575
         End
         Begin VB.Label Label4 
            BorderStyle     =   1  'Fixed Single
            Caption         =   "冬眠："
            Height          =   435
            Index           =   7
            Left            =   -74880
            TabIndex        =   46
            Top             =   2400
            Width           =   1575
         End
         Begin VB.Label SleepValue 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   1  'Fixed Single
            Height          =   435
            Left            =   -73080
            TabIndex        =   45
            Top             =   2400
            Width           =   1575
         End
         Begin VB.Label Label4 
            BorderStyle     =   1  'Fixed Single
            Caption         =   "同盟时间："
            Height          =   435
            Index           =   8
            Left            =   -71160
            TabIndex        =   44
            Top             =   1920
            Width           =   1575
         End
         Begin VB.Label FriendTimeValue 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   1  'Fixed Single
            Height          =   435
            Left            =   -69360
            TabIndex        =   43
            Top             =   1920
            Width           =   1575
         End
         Begin VB.Label Label4 
            BorderStyle     =   1  'Fixed Single
            Caption         =   "同盟人物："
            Height          =   435
            Index           =   9
            Left            =   -71160
            TabIndex        =   42
            Top             =   1440
            Width           =   1575
         End
         Begin VB.Label Label4 
            BorderStyle     =   1  'Fixed Single
            Caption         =   "控制权："
            Height          =   435
            Index           =   10
            Left            =   -74880
            TabIndex        =   41
            Top             =   480
            Width           =   1575
         End
      End
      Begin VB.Label Label1 
         Caption         =   "当前角色："
         Height          =   435
         Left            =   120
         TabIndex        =   8
         Top             =   360
         Width           =   1560
      End
      Begin VB.Label ControlLabel 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   435
         Left            =   3840
         TabIndex        =   7
         Top             =   360
         Width           =   1455
      End
   End
   Begin VB.Frame FileFrame 
      Caption         =   "选择文件"
      Height          =   975
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   11895
      Begin VB.CheckBox Check1 
         Caption         =   "自动备份"
         Height          =   435
         Left            =   10200
         TabIndex        =   4
         Top             =   360
         Value           =   1  'Checked
         Width           =   1575
      End
      Begin VB.CommandButton SaveButton 
         Caption         =   "保存"
         Enabled         =   0   'False
         Height          =   435
         Left            =   9120
         TabIndex        =   2
         Top             =   360
         Width           =   855
      End
      Begin VB.CommandButton OpenButton 
         Caption         =   "打开"
         Height          =   435
         Left            =   8040
         TabIndex        =   1
         Top             =   360
         Width           =   855
      End
      Begin VB.Label OpenFileName 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   435
         Left            =   120
         TabIndex        =   3
         Top             =   360
         Width           =   7695
      End
   End
End
Attribute VB_Name = "MainForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub BTLButton_Click()
    msg1 = MsgBox("确定要复活此角色？", vbExclamation + vbYesNo)
    If msg1 = vbYes Then Actor.BackToLife
End Sub
Private Sub CardCombo_Click(Index As Integer)
    Actor.ChangeCards (Index)
End Sub
Private Sub CashButton_Click(Index As Integer)
    Actor.Cash = M(Actor.Cash + 10 ^ (5 + Index), 2 ^ 31 - 1)
    CashValue.Caption = Format(Actor.Cash, "#,##0")
End Sub
Private Sub DepositButton_Click(Index As Integer)
    Actor.Deposit = M(Actor.Deposit + 10 ^ (5 + Index), 2 ^ 31 - 1)
    DepositValue.Caption = Format(Actor.Deposit, "#,##0")
End Sub
Private Sub LoanButton_Click(Index As Integer)
    Actor.Loan = M(Actor.Loan + 10 ^ (5 + Index), 2 ^ 31 - 1)
    LoanValue.Caption = Format(Actor.Loan, "#,##0")
End Sub
Private Sub CashValue_DblClick()
    Actor.Cash = GetNumber(, , "请输入现金")
    CashValue.Caption = Format(Actor.Cash, "#,##0")
End Sub
Private Sub DepositValue_DblClick()
    Actor.Deposit = GetNumber(, , "请输入存款")
    DepositValue.Caption = Format(Actor.Deposit, "#,##0")
End Sub
Private Sub LoanValue_DblClick()
    Actor.Loan = GetNumber(, , "请输入贷款")
    LoanValue.Caption = Format(Actor.Loan, "#,##0")
End Sub

Private Sub PMoneyValue_DblClick()
    PMoney = GetNumber(, , "请输入数值")
    PMoneyValue.Caption = Format(PMoney, "#,##0")
End Sub
Private Sub PointValue_DblClick()
    Actor.Point = GetNumber(32767, , "请输入点券")
    PointValue.Caption = Actor.Point
End Sub
Private Sub OpenButton_Click()
    On Error GoTo aa
    '选择文档
    With CommonDialog1
        .ShowOpen
        OpenFileName.Caption = .FileName
        OpenName = .FileName
    End With
    Close #1
    '备份
    If Check1.Value = Checked Then FileCopy OpenName, Left(OpenName, Len(OpenName) - 4) & "-备份.DAT"
    '打开文档
    Dim PlayerCount As Byte, Transform As Byte
    Open OpenName For Binary As #1
    PMoney = 0
    For i = 0 To 3
        Get #1, 9867 + i, Transform
        PMoney = PMoney + Transform * 256 ^ i
    Next i
    Get #1, 13, PlayerCount
    For i = 0 To PlayerCount - 1
        ReDim Preserve Players(i)
        Set Players(i) = New Player
    Next i
    '窗体
    PlayerFrame.Enabled = True
    SystemFrame.Enabled = True
    SaveButton.Enabled = True
    With PlayerCombo
        .Clear
        For i = LBound(Players) To UBound(Players)
            .AddItem Players(i).Name
        Next i
        .ListIndex = 0
    End With
    PMoneyValue.Caption = Format(PMoney, "#,###,###,##0")
aa:
End Sub
Private Sub PlayerCombo_Click()
    Set Actor = Players(PlayerCombo.ListIndex)
    Actor.UpdateForm
End Sub
Private Sub SaveButton_Click()
    For i = LBound(Players) To UBound(Players)
        Players(i).Save
    Next i
    Put #1, 9867, PMoney
End Sub
Private Sub IValues_DblClick(Index As Integer)
    Actor.ChangeItems (Index)
End Sub
Private Sub VanishValue_DblClick()
    Actor.Vanish = GetNumber(127, 3, "请输入天数")
    VanishValue.Caption = Actor.Vanish
End Sub
Private Sub PrisonValue_DblClick()
    Actor.Prison = GetNumber(127, 9, "请输入天数")
    PrisonValue.Caption = Actor.Prison
    If Actor.Prison = 0 Then Actor.Prison = 255
End Sub
Private Sub HospitalValue_DblClick()
    Actor.Hospital = GetNumber(127, 5, "请输入天数")
    HospitalValue.Caption = Actor.Hospital
    If Actor.Hospital = 0 Then Actor.Hospital = 255
End Sub
Private Sub SleepValue_DblClick()
    Actor.Sleep = GetNumber(127, 5, "请输入天数")
    SleepValue.Caption = Actor.Sleep
End Sub
Private Sub DreamValue_DblClick()
    Actor.Dream = GetNumber(127, 5, "请输入天数")
    DreamValue.Caption = Actor.Dream
End Sub
Private Sub StayValue_DblClick()
    Actor.Stay = GetNumber(127, 1, "请输入天数")
    StayValue.Caption = Actor.Stay
End Sub
Private Sub SlowValue_DblClick()
    Actor.Slow = GetNumber(127, 3, "请输入天数")
    SlowValue.Caption = Actor.Slow
End Sub
Private Sub FriendTimeValue_DblClick()
    Actor.FriendTime = GetNumber(127, 5, "请输入天数")
    FriendTimeValue.Caption = Actor.FriendTime
End Sub
Private Sub FriendCombo_Click()
    If FriendCombo.ListIndex = 0 Then
        FriendTimeValue.Enabled = False
        Actor.FriendTime = 0
        FriendTimeValue.Caption = 0
    Else
        FriendTimeValue.Enabled = True
    End If
    Actor.FriendPlayer = FriendCombo.ItemData(FriendCombo.ListIndex) + 1
End Sub
Private Sub TCombo_Click()
    If TCombo.ListIndex <= 2 Then
        Actor.Transport = TCombo.ListIndex
        TValue.Caption = 0
        TValue.Enabled = False
    Else
        If Actor.CCTime = 0 Then    '之前没开工程车，也没调过工程车
            Actor.CCTime = 7
            Actor.Transport = 31
        End If
        TValue.Caption = Actor.CCTime
        TValue.Enabled = True
    End If
End Sub
Private Sub TValue_DblClick()
    Actor.CCTime = GetNumber(64, 7, "请输入天数")
    Actor.CCTime = IIf(Actor.CCTime = 0, 1, Actor.CCTime)
    Actor.Transport = IIf(Actor.CCTime = 64, 3, Actor.CCTime * 4 + 3)
    TValue.Caption = Actor.CCTime
End Sub
Private Sub ControlCombo_Click()
    Actor.Control = ControlCombo.ListIndex + 1
    MainForm.ControlLabel.Caption = ControlCombo.Text
End Sub
