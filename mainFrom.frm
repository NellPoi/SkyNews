VERSION 5.00
Begin VB.MDIForm mainForm 
   Appearance      =   0  'Flat
   BackColor       =   &H00FFFFFF&
   Caption         =   "SkyNews"
   ClientHeight    =   10650
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   17310
   Icon            =   "mainFrom.frx":0000
   LinkTopic       =   "mainForm"
   StartUpPosition =   2  '屏幕中心
   Begin VB.PictureBox div1 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   10575
      Left            =   0
      ScaleHeight     =   703
      ScaleMode       =   0  'User
      ScaleWidth      =   1152
      TabIndex        =   0
      Top             =   0
      Width           =   17310
      Begin VB.PictureBox div1_btn_bg 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         BorderStyle     =   0  'None
         FillColor       =   &H80000004&
         FillStyle       =   3  'Vertical Line
         BeginProperty Font 
            Name            =   "微软雅黑"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   495
         Left            =   1560
         ScaleHeight     =   33
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   97
         TabIndex        =   4
         Top             =   3960
         Width           =   1455
         Begin VB.Label div1_btn_t 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "准备就绪"
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
            Left            =   240
            TabIndex        =   5
            Top             =   120
            Width           =   960
         End
      End
      Begin VB.Timer pushConfig 
         Enabled         =   0   'False
         Interval        =   1
         Left            =   14880
         Top             =   360
      End
      Begin VB.Label developer 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "S"
         BeginProperty Font 
            Name            =   "Wingdings 3"
            Size            =   42
            Charset         =   2
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   600
         Left            =   16080
         TabIndex        =   6
         Top             =   9720
         Width           =   915
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "一个up主自用来搬运外网资源的App"
         BeginProperty Font 
            Name            =   "微软雅黑"
            Size            =   15
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000017&
         Height          =   405
         Left            =   1560
         TabIndex        =   3
         Top             =   2520
         Width           =   4875
      End
      Begin VB.Label div1_info1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "##:##:##"
         BeginProperty Font 
            Name            =   "微软雅黑"
            Size            =   15
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   11400
         TabIndex        =   2
         Top             =   1800
         Width           =   1320
      End
      Begin VB.Label div1_title 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Sky:The Children Of The Light"
         BeginProperty Font 
            Name            =   "微软雅黑 Light"
            Size            =   36
            Charset         =   134
            Weight          =   290
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   930
         Left            =   1440
         TabIndex        =   1
         Top             =   1320
         Width           =   9660
      End
   End
End
Attribute VB_Name = "mainForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public mainWidth, mainHeight As Variant
Rem 传递子窗体尺寸定位参数
Private Sub MDIForm_Load()
    div1.BorderStyle = 0
    div1_info1.Caption = "Ver" & App.Major & "." & App.Minor & "." & App.Revision
    div1_btn_t.Left = (div1_btn_bg.Width - div1_btn_t.Width) / 2
    div1_btn_t.Top = (div1_btn_bg.Height - div1_btn_t.Height) / 2
    mainWidth = ScaleWidth
    mainHeight = ScaleHeight
    pushConfig.Enabled = True
End Sub
Private Sub div1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    div1_btn_bg.BorderStyle = 0
End Sub

Private Sub div1_btn_t_Click()
    index.Show
    div1.Visible = False
End Sub

Private Sub div1_btn_bg_Click()
    index.Show
    div1.Visible = False
End Sub

Private Sub div1_btn_bg_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    div1_btn_bg.BackColor = &H80000010
    div1_btn_bg.BorderStyle = 0
End Sub
Private Sub div1_btn_bg_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    div1_btn_bg.BackColor = &H80000000
    div1_btn_bg.BorderStyle = 0
End Sub

Private Sub div1_btn_bg_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    div1_btn_bg.BorderStyle = 1
End Sub

Private Sub MDIForm_Resize()
    'mainWidth = ScaleWidth
    'mainHeight = ScaleHeight
    developer.Left = Me.Width - developer.Width - 300
    developer.Top = Me.Height - developer.Height - 300
    div1.Height = Me.Height
End Sub

Private Sub pushConfig_Timer()
    mainWidth = ScaleWidth
    mainHeight = ScaleHeight
End Sub
