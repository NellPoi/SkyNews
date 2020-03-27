VERSION 5.00
Begin VB.Form index 
   Appearance      =   0  'Flat
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "index"
   ClientHeight    =   7995
   ClientLeft      =   7770
   ClientTop       =   4860
   ClientWidth     =   14025
   FillColor       =   &H000000FF&
   FillStyle       =   0  'Solid
   BeginProperty Font 
      Name            =   "微软雅黑"
      Size            =   9
      Charset         =   134
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H8000000D&
   Icon            =   "index.frx":0000
   LinkTopic       =   "index"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   533
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   935
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox p2_bg 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      FillColor       =   &H80000000&
      ForeColor       =   &H80000000&
      Height          =   5295
      Left            =   9120
      ScaleHeight     =   351
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   311
      TabIndex        =   12
      Top             =   2280
      Width           =   4695
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "说明"
         BeginProperty Font 
            Name            =   "微软雅黑"
            Size            =   10.5
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000011&
         Height          =   285
         Left            =   360
         TabIndex        =   13
         Top             =   240
         Width           =   420
      End
   End
   Begin VB.PictureBox p1_bg 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   2415
      Left            =   840
      ScaleHeight     =   159
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   495
      TabIndex        =   4
      Top             =   2280
      Width           =   7455
      Begin VB.CheckBox check2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "29"
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   9
         Left            =   4680
         TabIndex        =   20
         Top             =   1080
         Width           =   2895
      End
      Begin VB.CheckBox check2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "28"
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   8
         Left            =   4680
         TabIndex        =   19
         Top             =   720
         Width           =   2895
      End
      Begin VB.CheckBox check2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "27"
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   7
         Left            =   4080
         TabIndex        =   18
         Top             =   1800
         Width           =   2895
      End
      Begin VB.CheckBox check2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "26"
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   6
         Left            =   4080
         TabIndex        =   17
         Top             =   1440
         Width           =   2895
      End
      Begin VB.CheckBox check2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "25"
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   5
         Left            =   4080
         TabIndex        =   16
         Top             =   1080
         Width           =   2895
      End
      Begin VB.CheckBox check2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "24"
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   4
         Left            =   4080
         TabIndex        =   15
         Top             =   720
         Width           =   2895
      End
      Begin VB.CheckBox check1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "----"
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   2
         Left            =   600
         TabIndex        =   14
         Top             =   1440
         Width           =   2415
      End
      Begin VB.CheckBox check2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "23"
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   3
         Left            =   3480
         TabIndex        =   10
         Top             =   1800
         Width           =   2895
      End
      Begin VB.CheckBox check2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "22"
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   2
         Left            =   3480
         TabIndex        =   9
         Top             =   1440
         Width           =   2895
      End
      Begin VB.CheckBox check2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "21"
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   1
         Left            =   3480
         TabIndex        =   8
         Top             =   1080
         Width           =   2895
      End
      Begin VB.CheckBox check2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "20"
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   0
         Left            =   3480
         TabIndex        =   7
         Top             =   720
         Width           =   2895
      End
      Begin VB.CheckBox check1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "----"
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   1
         Left            =   600
         TabIndex        =   6
         Top             =   1080
         Width           =   2415
      End
      Begin VB.CheckBox check1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "----"
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   0
         Left            =   600
         TabIndex        =   5
         Top             =   720
         Width           =   2415
      End
      Begin VB.Line Line3 
         BorderColor     =   &H80000000&
         X1              =   0
         X2              =   496
         Y1              =   160
         Y2              =   160
      End
      Begin VB.Label title_Type 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "国服"
         BeginProperty Font 
            Name            =   "微软雅黑"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   600
         TabIndex        =   11
         Top             =   240
         Width           =   480
      End
   End
   Begin VB.CheckBox devMode 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "调用数组型Web控件"
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   5400
      TabIndex        =   2
      Top             =   840
      Width           =   2055
   End
   Begin VB.Timer checkUI 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   120
      Top             =   720
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000000&
      X1              =   592
      X2              =   592
      Y1              =   128
      Y2              =   504
   End
   Begin VB.Line line_Type2 
      BorderColor     =   &H8000000D&
      BorderWidth     =   2
      X1              =   112
      X2              =   160
      Y1              =   128
      Y2              =   128
   End
   Begin VB.Line line_Type1 
      BorderColor     =   &H8000000D&
      BorderWidth     =   2
      X1              =   64
      X2              =   96
      Y1              =   128
      Y2              =   128
   End
   Begin VB.Label div2_type2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "国际服"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   1680
      TabIndex        =   3
      Top             =   1560
      Width           =   720
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000000&
      X1              =   0
      X2              =   592
      Y1              =   128
      Y2              =   128
   End
   Begin VB.Label div2_type1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "国服"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   960
      TabIndex        =   1
      Top             =   1560
      Width           =   480
   End
   Begin VB.Label div2_title 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Web Browser Frames"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   21.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   570
      Left            =   840
      TabIndex        =   0
      Top             =   600
      Width           =   4365
   End
End
Attribute VB_Name = "index"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Click()
    Dim tt, ttt As String, ts As Variant
    For ts = 0 To check2.Count - 1 Step 1
    ttt = check2(ts).Left
    tt = tt & check2(ts).Caption & "————" & ttt & "————" & check2(ts).Visible & "————" & check2(ts).Value & Chr(13)
    Next ts
    MsgBox "发现新大陆了？这只是一个调试窗口而已" & Chr(13) & tt & Chr(13) & "CONSOLE===>" & check2.Count
End Sub

Private Sub Form_Load()
    checkUI.Enabled = True
    Dim i, n As Integer
        For i = 0 To check1.Count - 1 Step 1
            Rem 负责加载 国服 数列
            check1(i).BackColor = RGB(251, 251, 251)
            check1(i).Caption = LoadResString(100 + i)
        Next i
    line_Type2.Visible = False
        For n = 0 To check2.Count - 1 Step 1
            Rem 负责加载 国际服 数列
            check2(n).BackColor = RGB(251, 251, 251)
            check2(n).Caption = LoadResString(200 + n)
            check2(n).Visible = False
        Next n
    p1_bg.BorderStyle = 0
    p2_bg.BorderStyle = 0
    p1_bg.BackColor = RGB(251, 251, 251)
End Sub

Private Sub checkUI_Timer()
    Dim p2_paddingVal As Variant
    p2_paddingVal = 15
    Rem 此参数决定p2_bg模块内边距大小
    devMode.Top = Me.Height - 300
    Me.WindowState = 2
    Line1.X2 = ScaleWidth
    Line2.X1 = p2_bg.Left - p2_paddingVal
    Line2.X2 = p2_bg.Left - p2_paddingVal
    Line2.Y2 = Me.Height
    Line3.X2 = p1_bg.Width
    p1_bg.Width = Line2.X1 - (Line2.X1 / 100 * 15)
    '使p1_bg宽度能动态的调整，而不是僵硬的变化，注意，拉伸范围过大会导致参数错误崩溃
    p2_bg.Top = Line1.Y1 + p2_paddingVal
    p2_bg.Left = ScaleWidth - 300
    p2_bg.Height = ScaleHeight - p2_bg.Top - p2_paddingVal
    p2_bg.Width = ScaleWidth - p2_bg.Left - p2_paddingVal
End Sub

Private Sub div2_type1_Click()
    title_Type = "国服"
    Dim i, n As Integer
    line_Type1.Visible = True
    line_Type2.Visible = False
    For i = 0 To check1.Count - 1 Step 1
        check1(i).Visible = True
    Next i
    For n = 0 To check2.Count - 1 Step 1
        check2(n).Visible = False
    Next n
End Sub

Private Sub div2_type2_Click()
    title_Type = "国际服"
    Dim i, n, row As Integer
    line_Type1.Visible = False
    line_Type2.Visible = True
    For i = 0 To check1.Count - 1 Step 1
        check1(i).Visible = False
        check1(row).Left = 40
    Next i
    Dim col As Variant
    If check2.Count / 4 <> 0 Then
        col = check2.Count + 1
    Else
        col = check2.Count
    End If
    For n = 0 To col Step 1
        For row = 0 To 3 Step 1
            If n = 0 Then
            check2(row).Left = 40
            check2(row).Visible = True
            ElseIf n > 0 Then
            If row + n * 3 + 1 > check2.Count - 1 Then
                Exit For
            End If
            check2(row + n * 3 + 1).Left = 130 * (n + 1)
            check2(row + n * 3 + 1).Visible = True
            End If
        Next row
    Next n
    'For n = 0 To check2.Count / 4 Step 1
    '   If n Mod 4 = 0 Then
    '        Dim nn As Variant
    '        For nn = n To n + 3 Step 1
    '        check2(nn).Caption = LoadResString(200 + n)
    '        Next nn
    '    End If
    '    check2(n).Visible = True
    '    check2(n).Left = 40
    'Next n
End Sub
Private Sub devMode_Click()
    If devMode.Value = 1 Then
        Dim t
        t = _
        "这将调用多个Web网页组件来加载网页，请确保你的网络顺畅" & Chr(13) & _
        "同时部分网页需要开启VPN来加载，否则无法正常显示页面" & Chr(13) & _
        "是否继续？"
            If MsgBox(t, vbOKCancel, "开发者模式") = vbOK Then
            Else
                devMode.Value = 0
        End If
    Else
    End If
End Sub
