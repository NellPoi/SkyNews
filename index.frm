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
         Caption         =   "Helpshift"
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
         Caption         =   "官方网站"
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
         Caption         =   "推特Twitter"
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
         Caption         =   "Discord"
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
         Caption         =   "官方微博"
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   1
         Left            =   600
         TabIndex        =   6
         Top             =   720
         Width           =   2895
      End
      Begin VB.CheckBox check1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "官方网站"
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   0
         Left            =   600
         TabIndex        =   5
         Top             =   1080
         Width           =   2895
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
      Top             =   120
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000000&
      X1              =   592
      X2              =   592
      Y1              =   128
      Y2              =   504
   End
   Begin VB.Line Line_Type2 
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

Private Sub Command1_Click()
Dim a As Integer
a = InputBox("width" & p1_bg.Width & Chr(13) & index.Width, "植入参数")
p1_bg.Width = Val(a)
checkUI.Enabled = False
End Sub

Private Sub Form_Load()
    checkUI.Enabled = True
    Dim i, n As Integer
        For i = 0 To check1.Count - 1 Step 1
            check1(i).BackColor = RGB(251, 251, 251)
        Next i
    Line_Type2.Visible = False
        For n = 0 To check2.Count - 1 Step 1
            check2(n).BackColor = RGB(251, 251, 251)
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
    p1_bg.Width = Line2.X1 - (Line2.X1 / 100 * 15) '使p1_bg宽度能动态的调整，而不是僵硬的变化
    p2_bg.Top = Line1.Y1 + p2_paddingVal
    p2_bg.Left = ScaleWidth - 300
    p2_bg.Height = ScaleHeight - p2_bg.Top - p2_paddingVal
    p2_bg.Width = ScaleWidth - p2_bg.Left - p2_paddingVal
    'index.Width = mainForm.mainWidth
    'index.Height = mainForm.mainHeight
    'Line1.X2 = mainForm.mainWidth
    'devMode.Left = 200
    'Line2.X1 = p2_bg.Left - 300
    'Line2.X2 = p2_bg.Left - 300
    'Line2.Y2 = Me.Height
    'Line3.X2 = p1_bg.Width
    'p1_bg.Width = Line2.X1 / 100 * 85
    'p2_bg.Left = Me.Width - 300
    'p2_bg.Height = Me.Height - p1_bg.Top - 300
    'p2_bg.Width = Me.Width / 100 * 30
End Sub

Private Sub div2_type1_Click()
    title_Type = "国服"
    Dim i, n As Integer
    line_Type1.Visible = True
    Line_Type2.Visible = False
    For i = 0 To check1.Count - 1 Step 1
        check1(i).Visible = True
    Next i
    For n = 0 To check2.Count - 1 Step 1
        check2(n).Visible = False
    Next n
End Sub

Private Sub div2_type2_Click()
    title_Type = "国际服"
    Dim i, n As Integer
    line_Type1.Visible = False
    Line_Type2.Visible = True
    For i = 0 To check1.Count - 1 Step 1
        check1(i).Visible = False
    Next i
    For n = 0 To check2.Count - 1 Step 1
        check2(n).Visible = True
        check2(n).Left = 600
    Next n
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


