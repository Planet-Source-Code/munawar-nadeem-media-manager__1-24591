VERSION 5.00
Begin VB.Form Form5 
   BackColor       =   &H00008000&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Teknet Media Manager (Media File(s) List)"
   ClientHeight    =   4425
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3990
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form5"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4425
   ScaleWidth      =   3990
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command2 
      Caption         =   "&Hide"
      Height          =   375
      Left            =   240
      TabIndex        =   2
      Top             =   3960
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Ok"
      Height          =   375
      Left            =   2640
      TabIndex        =   1
      Top             =   3960
      Width           =   975
   End
   Begin VB.ListBox lstFoundFiles 
      BackColor       =   &H00008000&
      ForeColor       =   &H0000FFFF&
      Height          =   3765
      ItemData        =   "Form5.frx":0000
      Left            =   150
      List            =   "Form5.frx":0002
      TabIndex        =   0
      Top             =   120
      Width           =   3735
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00008000&
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   1290
      TabIndex        =   3
      Top             =   4020
      Width           =   1215
   End
End
Attribute VB_Name = "Form5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
lstFoundFiles_DblClick
End Sub

Private Sub Command2_Click()
Form5.Hide
End Sub


Private Sub lstFoundFiles_DblClick()
Form5ID = True
Form5.Hide
ItemsInFile = Form5.lstFoundFiles.ListCount - 1
GetFileHeader
CurrentIndexOne = 0
CurrentIndex = CurrentIndexOne
If ItemsInFile > 0 Then zeroItemsInFile = False
Form5.lstFoundFiles.ListIndex = CurrentIndex
Form2.MediaPlayer.FileName = Form5.lstFoundFiles.List(CurrentIndexOne)
Form2.Timer1.Enabled = True
End Sub


