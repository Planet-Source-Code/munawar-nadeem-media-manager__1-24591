VERSION 5.00
Begin VB.Form Form6 
   BackColor       =   &H00008000&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   975
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   4875
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form6"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   975
   ScaleWidth      =   4875
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text7 
      Height          =   285
      Left            =   105
      TabIndex        =   8
      Text            =   "Text7"
      Top             =   3360
      Width           =   1800
   End
   Begin VB.TextBox Text6 
      Height          =   345
      Left            =   105
      TabIndex        =   7
      Text            =   "Text6"
      Top             =   3045
      Width           =   1800
   End
   Begin VB.TextBox Text5 
      Height          =   285
      Left            =   105
      TabIndex        =   4
      Text            =   "Text5"
      Top             =   2730
      Width           =   1800
   End
   Begin VB.TextBox Text4 
      Height          =   330
      Left            =   105
      TabIndex        =   3
      Text            =   "Text4"
      Top             =   2310
      Width           =   1800
   End
   Begin VB.TextBox Text3 
      Height          =   330
      Left            =   105
      TabIndex        =   2
      Text            =   "Text3"
      Top             =   1995
      Width           =   1800
   End
   Begin VB.TextBox Text2 
      Height          =   330
      Left            =   105
      TabIndex        =   1
      Text            =   "Text2"
      Top             =   1680
      Width           =   1800
   End
   Begin VB.TextBox Text1 
      Height          =   330
      Left            =   105
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   1365
      Width           =   1800
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   105
      Picture         =   "Form6.frx":0000
      Top             =   210
      Width           =   480
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00008000&
      Caption         =   "Your CD-ROM Is Being Initialized"
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
      Left            =   630
      TabIndex        =   6
      Top             =   525
      Width           =   3990
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00008000&
      Caption         =   "Please Wait...!"
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
      Left            =   630
      TabIndex        =   5
      Top             =   210
      Width           =   4005
   End
End
Attribute VB_Name = "Form6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Refresh
End Sub
