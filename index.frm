VERSION 5.00
Begin VB.Form index 
   Appearance      =   0  'Flat
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "index"
   ClientHeight    =   7995
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   14025
   FillColor       =   &H000000FF&
   FillStyle       =   0  'Solid
   BeginProperty Font 
      Name            =   "΢���ź�"
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
   ScaleHeight     =   7995
   ScaleWidth      =   14025
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox p2_bg 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   5295
      Left            =   9120
      ScaleHeight     =   351
      ScaleMode       =   0  'User
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
      ScaleHeight     =   2385
      ScaleMode       =   0  'User
      ScaleWidth      =   7425
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
         Caption         =   "�ٷ���վ"
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
         Caption         =   "����Twitter"
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
         Caption         =   "�ٷ�΢��"
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
         Caption         =   "�ٷ���վ"
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
         X2              =   7440
         Y1              =   2400
         Y2              =   2400
      End
      Begin VB.Label title_Type 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "����"
         BeginProperty Font 
            Name            =   "΢���ź�"
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
      Caption         =   "����������Web�ؼ�"
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   5400
      TabIndex        =   2
      Top             =   840
      Width           =   2055
   End
   Begin VB.Timer checkUI 
      Interval        =   10
      Left            =   120
      Top             =   120
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000000&
      X1              =   8880
      X2              =   8880
      Y1              =   1920
      Y2              =   7560
   End
   Begin VB.Line Line_Type2 
      BorderColor     =   &H8000000D&
      BorderWidth     =   2
      X1              =   1680
      X2              =   2400
      Y1              =   1920
      Y2              =   1920
   End
   Begin VB.Line line_Type1 
      BorderColor     =   &H8000000D&
      BorderWidth     =   2
      X1              =   960
      X2              =   1440
      Y1              =   1920
      Y2              =   1920
   End
   Begin VB.Label div2_type2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "���ʷ�"
      BeginProperty Font 
         Name            =   "΢���ź�"
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
      X2              =   8880
      Y1              =   1920
      Y2              =   1920
   End
   Begin VB.Label div2_type1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "����"
      BeginProperty Font 
         Name            =   "΢���ź�"
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
         Name            =   "΢���ź�"
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
Private Sub form_load()
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
    p1_bg.BackColor = RGB(251, 251, 251)
End Sub

Private Sub checkUI_Timer()
    index.Width = mainForm.mainWidth
    index.Height = mainForm.mainHeight
    Line1.X2 = mainForm.mainWidth
    devMode.Top = Me.Height - 300
    devMode.Left = 200
    Line2.X1 = p2_bg.Left - 300
    Line2.X2 = p2_bg.Left - 300
    Line2.Y2 = Me.Height
    Line3.X2 = p1_bg.Width
    p1_bg.Width = Line2.X1 / 100 * 85
    p2_bg.Left = Me.Width - 300
    p2_bg.Height = Me.Height - p1_bg.Top - 300
    p2_bg.Width = Me.Width / 100 * 30
End Sub

Private Sub div2_type1_Click()
    title_Type = "����"
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
    title_Type = "���ʷ�"
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
        "�⽫���ö��Web��ҳ�����������ҳ����ȷ���������˳��" & Chr(13) & _
        "ͬʱ������ҳ��Ҫ����VPN�����أ������޷�������ʾҳ��" & Chr(13) & _
        "�Ƿ������"
            If MsgBox(t, vbOKCancel, "������ģʽ") = vbOK Then
            Else
                devMode.Value = 0
        End If
    Else
    End If
End Sub


