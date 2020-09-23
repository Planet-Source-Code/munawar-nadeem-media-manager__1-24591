VERSION 5.00
Object = "{22D6F304-B0F6-11D0-94AB-0080C74C7E95}#1.0#0"; "MSDXM.OCX"
Begin VB.Form Form2 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00C00000&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   7500
   ClientLeft      =   390
   ClientTop       =   975
   ClientWidth     =   11160
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   FillStyle       =   0  'Solid
   Icon            =   "Form2.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Palette         =   "Form2.frx":0442
   ScaleHeight     =   7587.062
   ScaleMode       =   0  'User
   ScaleWidth      =   11378.82
   ShowInTaskbar   =   0   'False
   Begin VB.HScrollBar HScroll1 
      Enabled         =   0   'False
      Height          =   195
      Left            =   120
      TabIndex        =   9
      Top             =   7200
      Width           =   10965
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   180
      Top             =   5760
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00C00000&
      Caption         =   "Evaluation Version"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   240
      Left            =   3480
      TabIndex        =   8
      Top             =   2760
      Width           =   4605
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Teknet"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   48
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   915
      Left            =   3360
      TabIndex        =   7
      Top             =   3000
      Width           =   4635
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FF00FF&
      BorderWidth     =   2
      X1              =   3425.882
      X2              =   8075.294
      Y1              =   2549.253
      Y2              =   2549.253
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "All rights Reserved"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   4
      Left            =   3360
      TabIndex        =   6
      Top             =   4560
      Width           =   4815
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Copyright 1991-2001 by Teknet Technologies, Developed By Munawar Nadeem"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   540
      Index           =   3
      Left            =   3360
      TabIndex        =   5
      Top             =   3960
      Width           =   4815
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00C00000&
      BackStyle       =   0  'Transparent
      Caption         =   "Ver 1.0"
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   9.75
         Charset         =   0
         Weight          =   300
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   1
      Left            =   7080
      TabIndex        =   4
      Top             =   2160
      Width           =   810
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00C00000&
      BackStyle       =   0  'Transparent
      Caption         =   "Teknet Media Manager"
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   11.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   0
      Left            =   3240
      TabIndex        =   3
      Top             =   2160
      Width           =   3330
   End
   Begin VB.Label lblPanel 
      BackColor       =   &H00C00000&
      Height          =   2865
      Left            =   3120
      TabIndex        =   2
      Top             =   2040
      Width           =   5160
   End
   Begin VB.Label lblVideoTime 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "                     "
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
      Height          =   195
      Left            =   240
      TabIndex        =   1
      Top             =   240
      Width           =   1515
   End
   Begin MediaPlayerCtl.MediaPlayer MediaPlayer 
      Height          =   7080
      Left            =   90
      TabIndex        =   0
      ToolTipText     =   "Teknet Cinema"
      Top             =   90
      Width           =   10965
      AudioStream     =   -1
      AutoSize        =   0   'False
      AutoStart       =   -1  'True
      AnimationAtStart=   -1  'True
      AllowScan       =   -1  'True
      AllowChangeDisplaySize=   -1  'True
      AutoRewind      =   -1  'True
      Balance         =   0
      BaseURL         =   ""
      BufferingTime   =   5
      CaptioningID    =   ""
      ClickToPlay     =   0   'False
      CursorType      =   0
      CurrentPosition =   0
      CurrentMarker   =   0
      DefaultFrame    =   ""
      DisplayBackColor=   0
      DisplayForeColor=   16777215
      DisplayMode     =   0
      DisplaySize     =   4
      Enabled         =   -1  'True
      EnableContextMenu=   0   'False
      EnablePositionControls=   0   'False
      EnableFullScreenControls=   -1  'True
      EnableTracker   =   0   'False
      Filename        =   ""
      InvokeURLs      =   0   'False
      Language        =   -1
      Mute            =   0   'False
      PlayCount       =   1
      PreviewMode     =   0   'False
      Rate            =   1
      SAMILang        =   ""
      SAMIStyle       =   ""
      SAMIFileName    =   ""
      SelectionStart  =   -1
      SelectionEnd    =   -1
      SendOpenStateChangeEvents=   -1  'True
      SendWarningEvents=   -1  'True
      SendErrorEvents =   -1  'True
      SendKeyboardEvents=   -1  'True
      SendMouseClickEvents=   0   'False
      SendMouseMoveEvents=   0   'False
      SendPlayStateChangeEvents=   -1  'True
      ShowCaptioning  =   0   'False
      ShowControls    =   0   'False
      ShowAudioControls=   0   'False
      ShowDisplay     =   0   'False
      ShowGotoBar     =   0   'False
      ShowPositionControls=   0   'False
      ShowStatusBar   =   0   'False
      ShowTracker     =   0   'False
      TransparentAtStart=   -1  'True
      VideoBorderWidth=   1
      VideoBorderColor=   8388608
      VideoBorder3D   =   0   'False
      Volume          =   0
      WindowlessVideo =   0   'False
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Private Sub Form_Activate()
Form2.WindowState = 0
Form2.Caption = ""
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo ErrH
    If KeyCode = vbKeyRight Then
        Form2.MediaPlayer.CurrentPosition = Form2.MediaPlayer.CurrentPosition + 1
    End If
    If KeyCode = vbKeyLeft Then
        Form2.MediaPlayer.CurrentPosition = Form2.MediaPlayer.CurrentPosition - 1
    End If
    If KeyCode = vbKeyF2 Then Form7.Show
    If KeyCode = vbKeyF10 Then Form2.MediaPlayer.DisplaySize = mpFullScreen
Exit Sub
ErrH:
    If Err = 380 Then
        On Error GoTo 0
        Exit Sub
    End If
End Sub

Private Sub Form_Load()
    '----------------------------------------------
    ItemsInFile = Form5.lstFoundFiles.ListCount - 1
    If ItemsInFile > 0 Then
        zeroItemsInFile = False
    Else
        zeroItemsInFile = True
    End If
    '----------------------------------------------
End Sub

Private Sub Form_Unload(Cancel As Integer)
MovieFlag = False
Form5.lstFoundFiles.Clear
Form5ID = False
MovieId = False
End Sub

Private Sub HScroll1_Change()
DoEvents
Form2.MediaPlayer.CurrentPosition = HScroll1.Value
End Sub

Private Sub HScroll1_Scroll()
DoEvents
End Sub

Private Sub MediaPlayer_EndOfStream(ByVal result As Long)
Dim RunIndex As Integer
DoEvents
'---------------------------------------
MovieFlag = False
MovieId = False
HScroll1.Enabled = False
'---------------------------------------
ButtonIndex = False
If zeroItemsInFile = False Then
    If CurrentIndex < ItemsInFile Then
        CurrentIndexOne = CurrentIndex + 1
        Form5.lstFoundFiles.ListIndex = CurrentIndexOne
    '--------------------
        MediaPlayer.FileName = Form5.lstFoundFiles.List(CurrentIndexOne)
    '--------------------
        Form4.Label1.Caption = CurrentIndexOne & " - " & ItemsInFile
        GetFileHeader
    End If
End If
End Sub


Private Sub MediaPlayer_Error()
DoEvents
Form3.lblGo.Caption = ""
Form3.lblStatus.Caption = ""
Form3.lblDrive.Caption = "Error"
Form3.lblDuration.Caption = ""
Form3.lblPosition.Caption = ""
Form3.lblIndex.Caption = ""
'----------------------------------
Unload Form2
Form8.Show
End Sub


Private Sub MediaPlayer_KeyDown(KeyCode As Integer, ShiftState As Integer)
On Error GoTo ErrH
    If KeyCode = vbKeyRight Then
        Form2.MediaPlayer.CurrentPosition = Form2.MediaPlayer.CurrentPosition + 1
    End If
    If KeyCode = vbKeyLeft Then
        Form2.MediaPlayer.CurrentPosition = Form2.MediaPlayer.CurrentPosition - 1
    End If
    If KeyCode = vbKeyF2 Then Form7.Show
    If KeyCode = vbKeyF10 Then Form2.MediaPlayer.DisplaySize = mpFullScreen
Exit Sub
ErrH:
    If Err = 380 Then
        On Error GoTo 0
        Exit Sub
    End If
End Sub

Private Sub MediaPlayer_NewStream()
DoEvents
'===========================
MovieFlag = True
MovieId = True
'---------------------------
TimeValue = True
Form3.lblStatus.Caption = ">"
Form3.lblGo.Caption = CurrentIndexOne
'--------------------------------
HScroll1.Enabled = True
HScroll1.Value = 0
HScroll1.Max = MediaPlayer.Duration
'--------------------------------
CurrentDuration = MediaPlayer.Duration
CurrentFile = UCase$(MediaPlayer.FileName)
CurrentIndex = CurrentIndexOne
'--------------------------------
If FindString = "AVI" Then AviOpen = True
'==========================
ButtonIndex = True
'--------------------------
End Sub

Private Sub MediaPlayer_PlayStateChange(ByVal OldState As Long, ByVal NewState As Long)
If NewState Then
    Unload Form6
 '------------------------------------
 If FindString = "DAT" Or FindString = "AVI" Then
    Form2.Show
    Form2.Refresh
End If
End If
End Sub


Private Sub Timer1_Timer()
'-----------------------
DoEvents
'====================================
If TimeValue Then
    Form3.lblIndex.Caption = CurrentIndexOne & " - " & ItemsInFile
    Form3.lblDuration.Caption = Format(Form2.MediaPlayer.Duration, "00:00.00")
    Form3.lblPosition.Caption = Format(Form2.MediaPlayer.CurrentPosition, "00:00.00")
End If
End Sub
