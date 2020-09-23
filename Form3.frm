VERSION 5.00
Begin VB.Form Form3 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H00008080&
   BorderStyle     =   0  'None
   ClientHeight    =   1050
   ClientLeft      =   450
   ClientTop       =   0
   ClientWidth     =   8730
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   HelpContextID   =   1
   Icon            =   "Form3.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form3"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Form3.frx":0442
   ScaleHeight     =   1050
   ScaleWidth      =   8730
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmd0 
      BackColor       =   &H00000080&
      Height          =   285
      Left            =   6555
      Picture         =   "Form3.frx":062A
      Style           =   1  'Graphical
      TabIndex        =   7
      ToolTipText     =   "Startup Media Selection"
      Top             =   510
      Width           =   330
   End
   Begin VB.CommandButton cmdMovieForward 
      BackColor       =   &H00000080&
      Height          =   285
      Left            =   3510
      MaskColor       =   &H00000080&
      Picture         =   "Form3.frx":0768
      Style           =   1  'Graphical
      TabIndex        =   33
      ToolTipText     =   "Forward"
      Top             =   2205
      Width           =   375
   End
   Begin VB.CommandButton cmdMinimize 
      BackColor       =   &H00000080&
      Height          =   285
      Left            =   5460
      Picture         =   "Form3.frx":0902
      Style           =   1  'Graphical
      TabIndex        =   31
      ToolTipText     =   "Minimize"
      Top             =   225
      Width           =   405
   End
   Begin VB.CommandButton cmdRec 
      BackColor       =   &H00000080&
      Height          =   285
      Left            =   1800
      Picture         =   "Form3.frx":0A55
      Style           =   1  'Graphical
      TabIndex        =   27
      ToolTipText     =   "Record"
      Top             =   2205
      Width           =   420
   End
   Begin VB.CommandButton cmdStop 
      BackColor       =   &H00000080&
      Height          =   285
      Left            =   4065
      Picture         =   "Form3.frx":0BFB
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Stop"
      Top             =   225
      Width           =   435
   End
   Begin VB.CommandButton cmdMovieBack 
      BackColor       =   &H00000080&
      Height          =   285
      Left            =   3330
      Picture         =   "Form3.frx":0DCF
      Style           =   1  'Graphical
      TabIndex        =   34
      ToolTipText     =   "Reverse"
      Top             =   2205
      Width           =   375
   End
   Begin VB.CommandButton cmdset 
      BackColor       =   &H00000080&
      Height          =   285
      Left            =   2880
      Picture         =   "Form3.frx":0F5F
      Style           =   1  'Graphical
      TabIndex        =   32
      ToolTipText     =   "System Settings"
      Top             =   2160
      Width           =   555
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   0
      Top             =   2730
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C000C0&
      BorderStyle     =   0  'None
      Height          =   720
      HelpContextID   =   1
      Left            =   180
      TabIndex        =   0
      Top             =   150
      Width           =   8355
      Begin VB.CommandButton cmd1 
         BackColor       =   &H00000080&
         Height          =   285
         Left            =   5685
         Picture         =   "Form3.frx":119F
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "1st Media Selection"
         Top             =   75
         Width           =   345
      End
      Begin VB.CommandButton cmd2 
         BackColor       =   &H00000080&
         Height          =   285
         Left            =   6030
         Picture         =   "Form3.frx":12DB
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "2nd Media Selection"
         Top             =   75
         Width           =   345
      End
      Begin VB.CommandButton cmd3 
         BackColor       =   &H00000080&
         Height          =   285
         Left            =   6375
         Picture         =   "Form3.frx":140C
         Style           =   1  'Graphical
         TabIndex        =   10
         ToolTipText     =   "3rd Media Selection"
         Top             =   75
         Width           =   330
      End
      Begin VB.CommandButton cmd9 
         BackColor       =   &H00000080&
         Height          =   285
         Left            =   6030
         Picture         =   "Form3.frx":154E
         Style           =   1  'Graphical
         TabIndex        =   16
         ToolTipText     =   "9th Media Selection"
         Top             =   360
         Width           =   345
      End
      Begin VB.CommandButton cmdGo 
         BackColor       =   &H00000080&
         Height          =   285
         Left            =   7125
         Picture         =   "Form3.frx":167B
         Style           =   1  'Graphical
         TabIndex        =   19
         ToolTipText     =   "Jump to specified index"
         Top             =   360
         Width           =   465
      End
      Begin VB.CommandButton cmdP 
         BackColor       =   &H00000080&
         Height          =   285
         Left            =   7590
         Picture         =   "Form3.frx":182F
         Style           =   1  'Graphical
         TabIndex        =   18
         ToolTipText     =   "Increase Specific Selection"
         Top             =   360
         Width           =   435
      End
      Begin VB.CommandButton cmd7 
         BackColor       =   &H00000080&
         Height          =   285
         Left            =   7695
         Picture         =   "Form3.frx":1955
         Style           =   1  'Graphical
         TabIndex        =   14
         ToolTipText     =   "7th Media Selection"
         Top             =   75
         Width           =   330
      End
      Begin VB.CommandButton cmd6 
         BackColor       =   &H00000080&
         Height          =   285
         Left            =   7365
         Picture         =   "Form3.frx":1A8A
         Style           =   1  'Graphical
         TabIndex        =   13
         ToolTipText     =   "6th Media Selection"
         Top             =   75
         Width           =   330
      End
      Begin VB.CommandButton cmd5 
         BackColor       =   &H00000080&
         Height          =   285
         Left            =   7035
         Picture         =   "Form3.frx":1BBB
         Style           =   1  'Graphical
         TabIndex        =   12
         ToolTipText     =   "5th Media Selection"
         Top             =   75
         Width           =   330
      End
      Begin VB.CommandButton cmd4 
         BackColor       =   &H00000080&
         Height          =   285
         Left            =   6705
         Picture         =   "Form3.frx":1CEE
         Style           =   1  'Graphical
         TabIndex        =   11
         ToolTipText     =   "4th Media Selection"
         Top             =   75
         Width           =   330
      End
      Begin VB.CommandButton Command2 
         Appearance      =   0  'Flat
         BackColor       =   &H00000080&
         Height          =   285
         Left            =   90
         MaskColor       =   &H00000080&
         Picture         =   "Form3.frx":1E2F
         Style           =   1  'Graphical
         TabIndex        =   1
         ToolTipText     =   "Power"
         Top             =   90
         Width           =   600
      End
      Begin VB.CommandButton cmdopenFile 
         BackColor       =   &H00000080&
         Height          =   285
         Left            =   90
         Picture         =   "Form3.frx":2054
         Style           =   1  'Graphical
         TabIndex        =   30
         ToolTipText     =   "Open Media File(s)"
         Top             =   360
         Width           =   600
      End
      Begin VB.CommandButton cmdPlay 
         BackColor       =   &H00000080&
         Height          =   285
         Index           =   0
         Left            =   4320
         Picture         =   "Form3.frx":224E
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Play"
         Top             =   75
         Width           =   555
      End
      Begin VB.CommandButton cmdSize 
         BackColor       =   &H00000080&
         Height          =   285
         Left            =   4860
         Picture         =   "Form3.frx":23F5
         Style           =   1  'Graphical
         TabIndex        =   29
         ToolTipText     =   "Full Screen mod"
         Top             =   75
         Width           =   420
      End
      Begin VB.CommandButton cmdFastBack 
         BackColor       =   &H00000080&
         Height          =   285
         Left            =   3885
         Picture         =   "Form3.frx":25D5
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Quick Reverse"
         Top             =   360
         Width           =   435
      End
      Begin VB.CommandButton cmdPause 
         BackColor       =   &H00004080&
         Height          =   285
         Left            =   4320
         MaskColor       =   &H00000080&
         Picture         =   "Form3.frx":27A7
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Pause"
         Top             =   360
         Width           =   555
      End
      Begin VB.CommandButton cmdFastNext 
         BackColor       =   &H00000080&
         Height          =   285
         Left            =   4860
         Picture         =   "Form3.frx":29AC
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Quick Forward"
         Top             =   360
         Width           =   420
      End
      Begin VB.CommandButton Command5 
         BackColor       =   &H00000080&
         Height          =   285
         Index           =   8
         Left            =   5265
         Picture         =   "Form3.frx":2B63
         Style           =   1  'Graphical
         TabIndex        =   28
         ToolTipText     =   "DVD / VCD"
         Top             =   360
         Width           =   420
      End
      Begin VB.CommandButton cmdN 
         BackColor       =   &H00000080&
         Height          =   285
         Left            =   6705
         Picture         =   "Form3.frx":2D28
         Style           =   1  'Graphical
         TabIndex        =   17
         ToolTipText     =   "Decrease Specific Selection"
         Top             =   360
         Width           =   420
      End
      Begin VB.Timer Timer2 
         Interval        =   100
         Left            =   315
         Top             =   2625
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Height          =   540
         Left            =   720
         TabIndex        =   20
         Top             =   90
         Width           =   3135
         Begin VB.Label lblDrive 
            Alignment       =   2  'Center
            BackColor       =   &H00000000&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Left            =   45
            TabIndex        =   26
            Top             =   45
            Width           =   900
         End
         Begin VB.Label lblTime 
            Alignment       =   2  'Center
            BackColor       =   &H00000000&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FFFF&
            Height          =   225
            Left            =   45
            TabIndex        =   25
            Top             =   270
            Width           =   900
         End
         Begin VB.Label lblPosition 
            Alignment       =   2  'Center
            BackColor       =   &H00000000&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FFFF&
            Height          =   225
            Left            =   990
            TabIndex        =   24
            Top             =   270
            Width           =   1065
         End
         Begin VB.Label lblDuration 
            Alignment       =   2  'Center
            BackColor       =   &H00000000&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FFFF&
            Height          =   240
            Left            =   990
            TabIndex        =   23
            Top             =   45
            Width           =   1065
         End
         Begin VB.Label lblIndex 
            Alignment       =   2  'Center
            BackColor       =   &H00000000&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0080FF80&
            Height          =   225
            Left            =   2070
            TabIndex        =   22
            Top             =   45
            Width           =   960
         End
         Begin VB.Label lblStatus 
            Alignment       =   2  'Center
            BackColor       =   &H00000000&
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
            Left            =   2070
            TabIndex        =   21
            Top             =   270
            Width           =   960
         End
      End
      Begin VB.CommandButton cmd8 
         BackColor       =   &H00000080&
         Height          =   285
         Left            =   5670
         Picture         =   "Form3.frx":2E54
         Style           =   1  'Graphical
         TabIndex        =   15
         ToolTipText     =   "8th Media Selection"
         Top             =   360
         Width           =   360
      End
      Begin VB.Label lblGo 
         Alignment       =   2  'Center
         BackColor       =   &H00000000&
         ForeColor       =   &H00FFFFFF&
         Height          =   570
         Left            =   8055
         TabIndex        =   36
         ToolTipText     =   "Currently Selected Media Index"
         Top             =   75
         Width           =   210
      End
   End
   Begin VB.Label lblTask 
      Alignment       =   2  'Center
      BackColor       =   &H00008000&
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   165
      Left            =   540
      TabIndex        =   35
      Top             =   2250
      Width           =   3150
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmd0_Click()
        If ItemsInFile >= "0" Then
            Form3.lblGo.Caption = "0"
            cmdGo_Click
        End If
End Sub
Private Sub cmd1_Click()
        If ItemsInFile >= "1" Then
            Form3.lblGo.Caption = "1"
            cmdGo_Click
        End If
End Sub
Private Sub cmd2_Click()
        If ItemsInFile >= "2" Then
            Form3.lblGo.Caption = "2"
            cmdGo_Click
        End If
End Sub
Private Sub cmd3_Click()
        If ItemsInFile >= "3" Then
            Form3.lblGo.Caption = "3"
            cmdGo_Click
        End If
End Sub
Private Sub cmd4_Click()
        If ItemsInFile >= "4" Then
            Form3.lblGo.Caption = "4"
            cmdGo_Click
        End If
End Sub
Private Sub cmd5_Click()
        
        If ItemsInFile >= "5" Then
            Form3.lblGo.Caption = "5"
            cmdGo_Click
        End If
End Sub
Private Sub cmd6_Click()
        If ItemsInFile >= "6" Then
            Form3.lblGo.Caption = "6"
            cmdGo_Click
        End If
End Sub
Private Sub cmd7_Click()
        If ItemsInFile >= "7" Then
            Form3.lblGo.Caption = "7"
            cmdGo_Click
        End If

End Sub
Private Sub cmd8_Click()
        If ItemsInFile >= "8" Then
            Form3.lblGo.Caption = "8"
            cmdGo_Click
        End If

End Sub
Private Sub cmd9_Click()
        If ItemsInFile >= "9" Then
            Form3.lblGo.Caption = "9"
            cmdGo_Click
        End If
End Sub

Private Sub cmdFastBack_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If ButtonIndex = True Then Form3.lblStatus.Caption = "<<<"
End Sub

Private Sub cmdFastBack_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If ButtonIndex = True Then Form3.lblStatus.Caption = ">"
End Sub


Private Sub cmdFastNext_Click()
        On Error GoTo ErrH
If ButtonIndex = True Then
        With Form2.MediaPlayer
            .CurrentPosition = .CurrentPosition + 5
        End With
End If
ErrH:
    On Error GoTo 0
End Sub
Private Sub cmdFastBack_Click()
        On Error GoTo ErrH
If ButtonIndex = True Then
        With Form2.MediaPlayer
            .CurrentPosition = .CurrentPosition - 5
        End With
End If
ErrH:
    On Error GoTo 0
End Sub

Private Sub cmdFastNext_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If ButtonIndex = True Then Form3.lblStatus.Caption = ">>>"
End Sub

Private Sub cmdFastNext_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If ButtonIndex = True Then Form3.lblStatus.Caption = ">"
End Sub


Private Sub cmdGo_Click()
Dim MyValue
If ButtonIndex = True Then
If CurrentIndex <> Form3.lblGo.Caption Then
If zeroItemsInFile = False Then
        MyValue = Form3.lblGo.Caption
        Form5.lstFoundFiles.ListIndex = MyValue
        CurrentIndexOne = Trim(Form3.lblGo.Caption)
        '--------------------
        Form2.MediaPlayer.FileName = Form5.lstFoundFiles.List(MyValue)
        '--------------------
End If
End If
End If
End Sub
Private Sub cmdMinimize_Click()
        Form3.Hide
End Sub
Private Sub cmdMovieBack_Click()
        On Error GoTo ErrH
If ButtonIndex = True Then
        With Form2.MediaPlayer
            .CurrentPosition = .CurrentPosition - 1
        End With
End If
ErrH:
    On Error GoTo 0
End Sub

Private Sub cmdMovieBack_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If ButtonIndex = True Then Form3.lblStatus.Caption = "<<"
End Sub


Private Sub cmdMovieBack_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If ButtonIndex = True Then Form3.lblStatus.Caption = ">"
End Sub


Private Sub cmdMovieForward_Click()
        On Error GoTo ErrH
If ButtonIndex = True Then
        With Form2.MediaPlayer
            .CurrentPosition = .CurrentPosition + 1
        End With
End If
ErrH:
    On Error GoTo 0
End Sub

Private Sub cmdMovieForward_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If ButtonIndex = True Then Form3.lblStatus.Caption = ">>"
End Sub

Private Sub cmdMovieForward_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If ButtonIndex = True Then Form3.lblStatus.Caption = ">"
End Sub

Private Sub cmdN_Click()
If ButtonIndex = True Then
        If Form3.lblGo.Caption <> 0 Then
            Form3.lblGo.Caption = Form3.lblGo.Caption - 1
        End If
End If
End Sub
Private Sub cmdopenFile_Click()
        On Error GoTo ErrH
If ButtonIndex = True Then
        ButtonIndex = False
        Form2.MediaPlayer.Stop
        Form3.lblStatus.Caption = "Stop"
        TimeValue = False
        Form2.Hide
        Form3.lblGo.Caption = ""
        Form3.lblStatus.Caption = ""
        Form3.lblDrive.Caption = ""
        Form3.lblDuration.Caption = ""
        Form3.lblPosition.Caption = ""
        Form3.lblIndex.Caption = ""
End If
Form4.Show
Form4.cmdFind.Enabled = False
Exit Sub
ErrH:
    On Error GoTo 0
End Sub
Private Sub cmdP_Click()
If ButtonIndex = True Then
        If Form3.lblGo.Caption < ItemsInFile Then
            Form3.lblGo.Caption = Form3.lblGo.Caption + 1
        End If
End If
End Sub
Private Sub cmdPause_Click()
        On Error GoTo ErrH
If ButtonIndex = True Then
        Form2.MediaPlayer.Pause
        Form3.lblStatus.Caption = "Pause"
End If
ErrH:
    On Error GoTo 0
End Sub
Private Sub cmdPlay_Click(Index As Integer)
        On Error GoTo ErrH
If ButtonIndex = True Then
        Form2.MediaPlayer.Play
        Form3.lblStatus.Caption = "Play"
End If
ErrH:
    On Error GoTo 0
End Sub

Private Sub cmdSize_Click()
If ButtonIndex = True Then
Form2.MediaPlayer.DisplaySize = mpFullScreen
End If
End Sub
Private Sub cmdStop_Click()
        On Error GoTo ErrH
If ButtonIndex = True Then
        ButtonIndex = False
        Form2.MediaPlayer.Stop
        Form3.lblStatus.Caption = "Stop"
        TimeValue = False
        Form2.Hide
        Form3.lblGo.Caption = ""
        Form3.lblStatus.Caption = ""
        Form3.lblDrive.Caption = ""
        Form3.lblDuration.Caption = ""
        Form3.lblPosition.Caption = ""
        Form3.lblIndex.Caption = ""
        MovieId = False
        MovieFlag = False
End If
ErrH:
    On Error GoTo 0
End Sub

Private Sub cmdStyle_Click()
End Sub


Private Sub Command2_Click()
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
End Sub
Private Sub Command5_Click(Index As Integer)
If Not MovieFlag Then
    Unload Form1
    Unload Form2
    Unload Form4
    Unload Form5
    Unload Form6
    Unload Form8
    Startup
    Form2.Timer1.Enabled = True
End If
End Sub

Private Sub Form_Activate()
Form9.WindowState = 1
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo ErrH
    If KeyCode = vbKeyRight Then
        Form2.MediaPlayer.CurrentPosition = Form2.MediaPlayer.CurrentPosition + 1
    End If
    If KeyCode = vbKeyLeft Then
        Form2.MediaPlayer.CurrentPosition = Form2.MediaPlayer.CurrentPosition - 1
    End If
    If KeyCode = vbKeyF10 Then Form2.MediaPlayer.DisplaySize = mpFullScreen
    If KeyCode = vbKeyF2 Then Form7.Show
Exit Sub
ErrH:
    If Err = 380 Then
        On Error GoTo 0
        Exit Sub
    End If
End Sub

Private Sub Form_Load()
Form9.Show
Refresh
Frame2.Refresh
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
End
End Sub

Private Sub Form_Resize()
If Form3.WindowState = 0 Then
    Form3.Caption = ""
    Left = 1670
    Top = 20
    Height = 1050
    Width = 8730
End If
If Form3.WindowState = 1 Then
    Form3.Caption = "Teknet Media Manager"
End If
End Sub


Private Sub Timer1_Timer()
    Form3.lblTime.Caption = Format(Time, "hh:mm:ss")
End Sub



Private Sub Timer2_Timer()
DoEvents
End Sub
