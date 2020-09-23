VERSION 5.00
Begin VB.Form Form8 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00008000&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   1380
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   4695
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form8"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1380
   ScaleWidth      =   4695
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      BackColor       =   &H00008000&
      Caption         =   "&Ok"
      Height          =   330
      Left            =   1785
      MaskColor       =   &H0000C000&
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   945
      Width           =   960
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   105
      Picture         =   "Form8.frx":0000
      Top             =   210
      Width           =   480
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00008000&
      Caption         =   "Sorry....!"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   330
      Index           =   0
      Left            =   735
      TabIndex        =   1
      Top             =   210
      Width           =   3480
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00008000&
      Caption         =   "Invalid / Unknown File Format "
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   225
      Index           =   1
      Left            =   735
      TabIndex        =   0
      Top             =   525
      Width           =   3480
   End
End
Attribute VB_Name = "Form8"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Unload Form8
End Sub

Private Sub SysInfo1_ConfigChangeCancelled()

End Sub

Private Sub Form_Unload(Cancel As Integer)
Form3.lblDrive.Caption = ""
End Sub
