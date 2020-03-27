VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.MDIForm mainForm 
   Appearance      =   0  'Flat
   BackColor       =   &H00FFFFFF&
   Caption         =   "SkyNews"
   ClientHeight    =   11610
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   18975
   Icon            =   "mainFrom.frx":0000
   LinkTopic       =   "mainForm"
   StartUpPosition =   2  '屏幕中心
   Begin VB.PictureBox div1 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "Microsoft YaHei UI"
         Size            =   15.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   11295
      Left            =   0
      ScaleHeight     =   751
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   1263
      TabIndex        =   0
      Top             =   0
      Width           =   18975
      Begin VB.TextBox Text2 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Microsoft YaHei UI"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C0C0&
         Height          =   375
         Left            =   1560
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   8
         Text            =   "mainFrom.frx":25CA
         Top             =   8640
         Width           =   5895
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Microsoft YaHei UI"
            Size            =   18
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000011&
         Height          =   1455
         Left            =   1560
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   7
         Text            =   "mainFrom.frx":25D4
         Top             =   6840
         Width           =   5775
      End
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
         ScaleWidth      =   105
         TabIndex        =   4
         Top             =   3360
         Width           =   1575
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
      Begin VB.Label developerEntrance 
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
         Left            =   12960
         TabIndex        =   6
         ToolTipText     =   "What's this?"
         Top             =   1560
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
         Left            =   11280
         TabIndex        =   2
         ToolTipText     =   "Release Version"
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
Private Sub MDIForm_Load()
    div1.BorderStyle = 0
    div1_info1.Caption = "Ver" & App.Major & "." & App.Minor & "." & App.Revision
    div1_btn_t.Left = (div1_btn_bg.Width - div1_btn_t.Width) / 2
    div1_btn_t.Top = (div1_btn_bg.Height - div1_btn_t.Height) / 2
    div1.Height = Me.Height
    Text1.ForeColor = RGB(102, 102, 102)
    Text2.ForeColor = RGB(153, 153, 153)
    Text2.Text = "轻触'" & div1_btn_t.Caption & "'以继续"
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
    developerEntrance.Left = div1.ScaleWidth - 80
    developerEntrance.Top = div1.ScaleHeight - 90
    div1.Height = Me.Height
    If Me.Width < 14000 Or Me.Height < 9000 Then
        Me.Width = 14000
        Me.Height = 10000
        MsgBox "再小就和丁丁一样小了！", vbCritical
    End If
End Sub
Private Sub developerEntrance_Click()
    developer.Show
    div1.Visible = False
End Sub

Private Sub developerEntrance_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    developerEntrance.ForeColor = &H8000000D
End Sub

Private Sub developerEntrance_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    developerEntrance.ForeColor = vbBlack
End Sub

