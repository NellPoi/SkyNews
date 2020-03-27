VERSION 5.00
Begin VB.Form developer 
   Appearance      =   0  'Flat
   BackColor       =   &H008080FF&
   BorderStyle     =   0  'None
   ClientHeight    =   7920
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   13830
   BeginProperty Font 
      Name            =   "Î¢ÈíÑÅºÚ"
      Size            =   9
      Charset         =   134
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "developer"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   528
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   922
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox header_bg 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   975
      Left            =   0
      ScaleHeight     =   945
      ScaleWidth      =   13785
      TabIndex        =   0
      Top             =   0
      Width           =   13815
   End
End
Attribute VB_Name = "developer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Click()
    header_bg.BorderStyle = 0
End Sub

Private Sub Form_Load()
    header_bg.Width = ScaleWidth
End Sub
