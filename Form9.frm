VERSION 5.00
Begin VB.Form Form9 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Teknet Media Manager"
   ClientHeight    =   30
   ClientLeft      =   1950
   ClientTop       =   600
   ClientWidth     =   8265
   FillStyle       =   0  'Solid
   Icon            =   "Form9.frx":0000
   LinkTopic       =   "Form9"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   Picture         =   "Form9.frx":0442
   ScaleHeight     =   30
   ScaleWidth      =   8265
   WindowState     =   1  'Minimized
   Begin VB.Label lblSettings 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   180
      TabIndex        =   2
      Top             =   720
      Width           =   690
   End
   Begin VB.Label lblOpen 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   135
      TabIndex        =   1
      Top             =   450
      Width           =   735
   End
   Begin VB.Label lblPower 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H80000005&
      Height          =   195
      Left            =   180
      TabIndex        =   0
      Top             =   180
      Width           =   690
   End
End
Attribute VB_Name = "Form9"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Dim Responce
If MovieId Then
    Form2.MediaPlayer.Pause
    Responce = MsgBox("Are you sure ?", vbQuestion + vbYesNo, "Teknet Media Manager")
    If Responce = 6 Then End
    Form2.MediaPlayer.Play
Else
    Responce = MsgBox("Are you sure ?", vbQuestion + vbYesNo, "Teknet Media Manager")
    If Responce = 6 Then End
End If
Cancel = True
End Sub

Private Sub Form_Resize()
Form3.Frame2.Refresh
Form3.Show
End Sub

Private Sub lblPower_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblPower.BorderStyle = 1
End Sub
