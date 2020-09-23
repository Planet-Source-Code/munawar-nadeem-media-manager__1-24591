VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form4 
   BackColor       =   &H00008000&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Teknet Media Manager (Open File)"
   ClientHeight    =   3780
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5670
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form4"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3780
   ScaleWidth      =   5670
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   4995
      Top             =   3420
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdSearch1 
      Caption         =   "Command1"
      Height          =   255
      Left            =   1440
      TabIndex        =   11
      Top             =   4200
      Width           =   1575
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "&Help"
      Enabled         =   0   'False
      Height          =   375
      Left            =   4440
      TabIndex        =   10
      Top             =   2760
      Width           =   1095
   End
   Begin VB.CheckBox chkAuto 
      BackColor       =   &H00008000&
      Caption         =   "Select All Files"
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   2160
      TabIndex        =   9
      Top             =   3360
      Value           =   1  'Checked
      Width           =   1335
   End
   Begin VB.CommandButton cmdAdvance 
      Caption         =   "&Search"
      Enabled         =   0   'False
      Height          =   375
      Left            =   4440
      TabIndex        =   8
      Top             =   2160
      Width           =   1095
   End
   Begin VB.CommandButton cmdFind 
      Caption         =   "&View Files"
      Height          =   375
      Left            =   4440
      TabIndex        =   7
      Top             =   1560
      Width           =   1095
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   4440
      TabIndex        =   6
      Top             =   960
      Width           =   1095
   End
   Begin VB.CommandButton cmdOpen 
      Caption         =   "&Open"
      Enabled         =   0   'False
      Height          =   375
      Left            =   4440
      TabIndex        =   5
      Top             =   360
      Width           =   1095
   End
   Begin VB.ComboBox ComboPtr 
      BackColor       =   &H00008000&
      ForeColor       =   &H0000FFFF&
      Height          =   315
      Left            =   120
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   3360
      Width           =   1935
   End
   Begin VB.TextBox txtSearchSpec 
      BackColor       =   &H00008000&
      ForeColor       =   &H0000FFFF&
      Height          =   285
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   1935
   End
   Begin VB.FileListBox FilList 
      BackColor       =   &H00008000&
      ForeColor       =   &H0000FFFF&
      Height          =   2820
      Left            =   120
      TabIndex        =   2
      Top             =   480
      Width           =   1950
   End
   Begin VB.DriveListBox DrvList 
      BackColor       =   &H00008000&
      ForeColor       =   &H0000FFFF&
      Height          =   315
      Left            =   2160
      TabIndex        =   1
      Top             =   120
      Width           =   1935
   End
   Begin VB.DirListBox DirList 
      BackColor       =   &H00008000&
      ForeColor       =   &H0080FFFF&
      Height          =   2790
      Left            =   2160
      TabIndex        =   0
      Top             =   480
      Width           =   1935
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00008000&
      ForeColor       =   &H0000FFFF&
      Height          =   195
      Left            =   3390
      TabIndex        =   12
      Top             =   3360
      Width           =   75
   End
   Begin VB.Line Line3 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      Index           =   2
      X1              =   4320
      X2              =   5520
      Y1              =   840
      Y2              =   840
   End
   Begin VB.Line Line3 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      Index           =   1
      X1              =   4320
      X2              =   5520
      Y1              =   2640
      Y2              =   2640
   End
   Begin VB.Line Line3 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      Index           =   0
      X1              =   4320
      X2              =   5520
      Y1              =   2040
      Y2              =   2040
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000009&
      BorderWidth     =   2
      X1              =   4320
      X2              =   5520
      Y1              =   1440
      Y2              =   1440
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      Index           =   0
      X1              =   4320
      X2              =   4320
      Y1              =   360
      Y2              =   3120
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Sub PatternIndex()
    '---------------------------------
    ComboPtr.AddItem "All Media Formats"
    ComboPtr.AddItem "Video CD"
    ComboPtr.AddItem "Audio CD"
    ComboPtr.AddItem "Video For Windows"
    ComboPtr.AddItem "Active Movie Files"
    ComboPtr.AddItem "Sound"
    ComboPtr.AddItem "MIDI Sequencer"
    ComboPtr.AddItem "All Files"
    ComboPtr.ListIndex = 0
    Select Case ComboPtr.List(ComboPtr.ListIndex)
        Case "All Media Formats"
            Form4.FilList.Pattern = "*.dat;*.avi;*.mpg;*.mpeg"
            txtSearchSpec.Text = "*.dat;*.avi;*.mpg;*.mpeg"
        Case "Video CD"
            Form4.FilList.Pattern = "*.dat"
            txtSearchSpec.Text = "*.dat"
        Case "Audio CD"
            Form4.FilList.Pattern = "*.mp3"
            txtSearchSpec.Text = "*.mp3"
        Case "Sound"
            Form4.FilList.Pattern = "*.wav"
            txtSearchSpec.Text = "*.wav"
        Case "Video For Windows"
            Form4.FilList.Pattern = "*.avi"
            txtSearchSpec.Text = "*.avi"
        Case "Active Movie Files"
            Form4.FilList.Pattern = "*.asf;*.asx;*.ivf;*.lsf;*.lsx"
            txtSearchSpec.Text = "*.asf;*.asx;*.ivf;*.lsf;*.lsx"
        Case "MIDI Sequencer"
            Form4.FilList.Pattern = "*.mid;*.mid;*.rmi"
            txtSearchSpec.Text = "*.mid;*.mid;*.rmi"
        Case "All Files"
            Form4.FilList.Pattern = "*.*"
            txtSearchSpec.Text = "*.*"
    End Select
End Sub

Public Sub cmdAdvance_Click()
Form4.Hide
Form5.Hide
Form3.Refresh
'-----------------------
Form6.Show
Form6.Label1(1).Caption = "Searching For Selected Media Files"
Form6.Refresh
'-----------------------
Form5.lstFoundFiles.Clear
Form4.FilList.Refresh
Call cmdSearch1_Click
Load Form5
If Form5.lstFoundFiles.ListCount - 1 >= 0 Then
    Unload Form6
    Form5.Show
    Label1.Caption = Form5.lstFoundFiles.ListCount - 1
    Form5.Label1.Caption = Form5.lstFoundFiles.ListCount - 1
    Form4.cmdFind.Enabled = True
Else
    Form4.Show
End If
End Sub

Private Sub cmdCancel_Click()
Form4.Hide
End Sub

Private Sub cmdFind_Click()
Form5.Show
End Sub

Private Sub cmdOpen_Click()
Dim CountD
On Error GoTo OpenHandler
Form3.Frame2.Refresh
If chkAuto.Value = Checked Then
If Form4.FilList.ListCount - 1 = 0 Then GoTo SingleMod
    Form5ID = True
    Form3.lblTask.Caption = "Multiple File Mode"
    Form5.lstFoundFiles.Clear
    Form4.Hide
    '--- Form4 list is not empty then ---
    If FilList.ListCount - 1 <> "-1" Then
        CurrentDrive = Form4.DrvList.Drive
        CurrentPath = Form4.FilList.Path
        For CountD = 0 To Form4.FilList.ListCount - 1
             Form5.lstFoundFiles.AddItem CurrentPath + "\" + Form4.FilList.List(CountD)
        Next CountD
        Form5.lstFoundFiles.ListIndex = 0
        CurrentIndex = 0
        GetFileHeader
        '---------------------------------
        ItemsInFile = Form5.lstFoundFiles.ListCount - 1
        '----------------------------------------------
        If ItemsInFile > 0 Then
            zeroItemsInFile = False
        Else
            zeroItemsInFile = True
        End If
    '----------------------------------------------

            '------------------------------
            TotalIndex = ItemsInFile
            CurrentIndex = 0
            CurrentIndexOne = CurrentIndex
            '----------------------------
        Select Case FindString
            Case "DAT"
                CurrentFile = Form5.lstFoundFiles.List(CurrentIndex)
                Form2.MediaPlayer.FileName = CurrentFile
            Case "AVI"
                CurrentFile = Form5.lstFoundFiles.List(CurrentIndex)
                Form2.MediaPlayer.FileName = CurrentFile
            
            Case "MPG"
                CurrentFile = Form5.lstFoundFiles.List(CurrentIndex)
                Form2.MediaPlayer.FileName = CurrentFile

            Case Else
                CurrentFile = Form5.lstFoundFiles.List(CurrentIndex)
                Form2.MediaPlayer.FileName = CurrentFile
                Form2.Hide
        End Select
    End If
End If
SingleMod:
If chkAuto.Value = Unchecked Then
If FilList.ListIndex <> "-1" Then
    Form4.Hide
    Form5.lstFoundFiles.Clear
    CurrentDrive = Form4.DrvList.Drive
    CurrentPath = Form4.FilList.Path
    Form5.lstFoundFiles.AddItem CurrentPath + "\" + Form4.FilList.List(Form4.FilList.ListIndex)
    Form5ID = True
    Form3.lblTask.Caption = "Single File Mode"
    CurrentIndex = 0
    GetFileHeader
    ItemsInFile = Form5.lstFoundFiles.ListCount - 1
    TotalIndex = ItemsInFile
    CurrentIndex = 0
    CurrentIndexOne = CurrentIndex
    CurrentFile = Form5.lstFoundFiles.List(CurrentIndex)
'====================================
        Select Case FindString
            Case "DAT"
                CurrentFile = Form5.lstFoundFiles.List(CurrentIndex)
                Form2.MediaPlayer.FileName = CurrentFile
            Case "AVI"
                CurrentFile = Form5.lstFoundFiles.List(CurrentIndex)
                Form2.MediaPlayer.FileName = CurrentFile
            
            Case "MPG"
                CurrentFile = Form5.lstFoundFiles.List(CurrentIndex)
                Form2.MediaPlayer.FileName = CurrentFile
            
            Case Else
                CurrentFile = Form5.lstFoundFiles.List(CurrentIndex)
                Form2.MediaPlayer.FileName = CurrentFile
                Form2.Hide
        End Select
'=====================================
End If
End If
Exit Sub
'-------------------------------
OpenHandler:
    If Err = -2147417848 Then
    Form8.Show
    Form8.Refresh
    On Error GoTo 0
    End If
End Sub

Public Sub cmdSearch1_Click()
Dim FirstPath As String, DirCount As Integer, NumFiles As Integer
Dim Result As Integer
    If DirList.Path <> DirList.List(DirList.ListIndex) Then
        DirList.Path = DirList.List(DirList.ListIndex)
        Exit Sub         ' Exit so user can take a look before searching.
    End If

    FilList.Pattern = txtSearchSpec.Text
    FirstPath = DirList.Path
    DirCount = DirList.ListCount

    NumFiles = 0                       ' Reset found files indicator.
    Result = DirDiver(FirstPath, DirCount, "")
    FilList.Path = DirList.Path

End Sub

Private Sub ComboPtr_Click()
    Select Case ComboPtr.List(ComboPtr.ListIndex)
        Case "Video CD"
            Form4.FilList.Pattern = "*.dat"
            txtSearchSpec.Text = "*.dat"
        Case "Movie mpeg"
            Form4.FilList.Pattern = "*.mpeg;*.mpg;*.m1v;*.mp2"
            txtSearchSpec.Text = "*.mpeg;*.mpg;*.m1v;*.mp2"
        Case "Audio CD"
            Form4.FilList.Pattern = "*.mp3"
            txtSearchSpec.Text = "*.mp3"
        Case "Sound"
            Form4.FilList.Pattern = "*.wav;*.snd;*.av;*.aif;*.aifc;*.aiff;*.wma"
            txtSearchSpec.Text = "*.wav;*.snd;*.av;*.aif;*.aifc;*.aiff;*.wma"
        Case "Video For Windows"
            Form4.FilList.Pattern = "*.avi;*.asf;*.wmv"
            txtSearchSpec.Text = "*.avi;*.asf;*.wmv"
        Case "Windows Media Files"
            Form4.FilList.Pattern = "*.asf;*.wm;*.wma;*.wmv"
            txtSearchSpec.Text = "*.asf;*.wm;*.wma;*.wmv"
        Case "Active Movie Files"
            Form4.FilList.Pattern = "*.asf;*.asx;*.ivf;*.lsf;*.lsx"
            txtSearchSpec.Text = "*.asf;*.asx;*.ivf;*.lsf;*.lsx"
        Case "MIDI Sequencer"
            Form4.FilList.Pattern = "*.mid;*.mid;*.rmi"
            txtSearchSpec.Text = "*.mid;*.mid;*.rmi"
        Case "All Files"
            Form4.FilList.Pattern = "*.*"
            txtSearchSpec.Text = "*.*"
    End Select
End Sub


Private Sub DirList_Change()
    FilList.Path = DirList.Path
End Sub

Private Sub dirList_LostFocus()
    DirList.Path = DirList.List(DirList.ListIndex)
End Sub

Private Sub drvList_Change()
    On Error GoTo DriveHandler
    DirList.Path = DrvList.Drive
    Exit Sub

DriveHandler:
    DrvList.Drive = DirList.Path
    Exit Sub
End Sub

Private Sub FilList_DblClick()
cmdOpen_Click
End Sub


Private Sub FilList_PathChange()
If Form4.FilList.ListCount - 1 >= 0 Then Form4.cmdOpen.Enabled = True
End Sub

Private Sub FilList_PatternChange()
If Form4.FilList.ListCount - 1 >= 0 Then Form4.cmdOpen.Enabled = True
End Sub


Private Sub Form_Load()
    TimeValue = False
    Call PatternIndex
    '---------------------------------
    If IsCDReady Then DrvList.Drive = CDROMDRIVE
    Form4.cmdFind.Enabled = False
    '---------------------------------
    Form3.lblGo.Caption = ""
    Form3.lblStatus.Caption = ""
    Form3.lblDuration.Caption = ""
    Form3.lblPosition.Caption = ""
    Form3.lblIndex.Caption = ""
    MovieId = False
End Sub

Private Sub ResetSearch()
    ' Reinitialize before starting a new search.
    lstFoundFiles.Clear
    SearchFlag = False ' Flag indicating search in progress.
    DirList.Path = CurDir: DrvList.Drive = DirList.Path ' Reset the path.
End Sub


Private Sub Form_LostFocus()
Form3.Refresh
Form3.Frame2.Refresh
Form4.Refresh
End Sub

Private Sub txtSearchSpec_Change()
   'Update file list box if user changes pattern.
    FilList.Pattern = txtSearchSpec.Text
End Sub


