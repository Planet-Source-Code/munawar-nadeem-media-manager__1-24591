Attribute VB_Name = "Module1"
Option Explicit
'--------------------------------
Global TotalIndex As Integer
Global CurrentIndex As Integer
Global CurrentIndexOne As Integer
Global ButtonIndex As Boolean
'-------------------------------
Global SearchFlag As Boolean
Global OpenWindow As Boolean
Global TimeValue As Boolean
'-------------------------------
Global ItemsInFile As Integer
Global zeroItemsInFile As Boolean
Global AviOpen As Boolean
'---------------------------
Global MovieFlag As Boolean
Global CDROMDRIVE As String
'---------------------------

Global IsCDPresent As Boolean
Global IsCDReady As Boolean
Global IsCDAudio As Boolean
Global IsCDVideo As Boolean
Global IsCDDVD As Boolean
'------------------------

Global GetNum As Integer
'------------------------

Global CurrentDrive As String
Global CurrentPath As String
Global CurrentFile As String
Global CurrentStatus As String
'---------------------------

Global DirectOpen As Boolean
Global Form5ID As Boolean
'---------------------------

Global FindString As String
Global CurrentDuration As Long
'-----------------------------
Global DriveLabel As String
'-----------------------------------------------

Private Declare Function GetDriveType Lib "kernel32" _
    Alias "GetDriveTypeA" (ByVal nDrive As String) As Long
    
Private Declare Function GetLogicalDriveStrings _
    Lib "kernel32" Alias "GetLogicalDriveStringsA" _
    (ByVal nBufferLength As Long, _
    ByVal lpBuffer As String) As Long
Public Function FindDVD()
Dim CountDVD As Integer
    If IsCDDVD Then
            OpenWindow = False
            Form3.lblDrive = "Video CD"
        '-----------------------------------
        With Form4
                .DrvList.Drive = CurrentDrive
                .DirList.Path = CurrentPath
                .FilList.Path = CurrentPath
        End With
        '------------------------------------
            For CountDVD = 0 To Form4.FilList.ListCount - 1
                Form5.lstFoundFiles.AddItem Form4.FilList.List(CountDVD)
            Next CountDVD
            '------------------------------
            ItemsInFile = Form5.lstFoundFiles.ListCount - 1
            '------------------------------
            TotalIndex = ItemsInFile
            CurrentIndex = 0
            CurrentIndexOne = CurrentIndex
            '----------------------------
            CurrentFile = CurrentPath + "\" + Form5.lstFoundFiles.List(CurrentIndex)
            Form2.MediaPlayer.FileName = CurrentFile
            Form2.Show
    End If
End Function

Public Function GetFileHeader()
    If Form5.lstFoundFiles.ListCount - 1 <> "-1" Then
        FindString = UCase$(Mid$(Form5.lstFoundFiles.List(CurrentIndex), Len(Form5.lstFoundFiles.List(CurrentIndex)) - 2, 3))
        Select Case FindString
            Case "MP3"
                Form3.lblDrive.Caption = "Audio MP3"
                Form3.lblDrive.Caption = "Audio MP3"
                IsCDAudio = True
            Case "M3U"
                Form3.lblDrive.Caption = "Audio M3U"
                Form3.lblDrive.Caption = "Audio M3U"
                IsCDAudio = True
            Case "WAV"
                Form3.lblDrive.Caption = "Sound Wav"
                Form3.lblDrive.Caption = "Sound Wav"
                IsCDAudio = True
            Case "DAT"
                Form3.lblDrive.Caption = "Video CD"
                Form3.lblDrive.Caption = "Video CD"
                IsCDVideo = True
            Case "AVI"
                Form3.lblDrive.Caption = "Video AVI"
                Form3.lblDrive.Caption = "Video AVI"
                IsCDVideo = True
            Case Else
                Form3.lblDrive.Caption = "Unknown"
                Form3.lblDrive.Caption = "Unknown"
                'IsCDVideo = True
        End Select
    End If
End Function

Public Function Avi_FileSearch()
Dim FirstPath As String, DirCount As Integer, NumFiles As Integer
Dim Result As Integer
    If Form4.DirList.Path <> Form4.DirList.List(Form4.DirList.ListIndex) Then
        Form4.DirList.Path = Form4.DirList.List(Form4.DirList.ListIndex)
        Exit Function    ' Exit so user can take a look before searching.
    End If

    Form4.FilList.Pattern = "*.avi" 'txtSearchSpec.Text
    FirstPath = Form4.DirList.Path
    DirCount = Form4.DirList.ListCount

    NumFiles = 0                       ' Reset found files indicator.
    Result = DirDiver(FirstPath, DirCount, "")
    Form4.FilList.Path = Form4.DirList.Path
End Function

Function AviStartupValues()
    '-------------------------------
If IsCDReady Then
Avi_FileSearch
'---------------------------
Form5ID = True
ItemsInFile = Form5.lstFoundFiles.ListCount - 1
CurrentIndexOne = 0
CurrentIndex = CurrentIndexOne
If ItemsInFile > 0 Then zeroItemsInFile = False
GetFileHeader
Form5.lstFoundFiles.ListIndex = CurrentIndex
Form2.MediaPlayer.FileName = Form5.lstFoundFiles.List(CurrentIndexOne)
Form2.Timer1.Enabled = True
    Unload Form6
    Form3.Show
End If
End Function


Function SetStartupValues()
    '-------------------------------
If IsCDReady Then
    Form4.DrvList.Drive = CurrentDrive
    Form4.DirList.Path = Form4.DrvList.Drive
    CurrentPath = Form4.DirList.Path
    Form4.txtSearchSpec.Text = "*.dat;*.avi;*.mp3;*.wav"
    Form4.cmdSearch1_Click
'---------------------------
        Form5.lstFoundFiles.ListIndex = 0
        GetFileHeader
        '---------------------------------
            ItemsInFile = Form5.lstFoundFiles.ListCount - 1
            '------------------------------
            TotalIndex = ItemsInFile
            CurrentIndex = 0
            CurrentIndexOne = CurrentIndex
            '----------------------------
        Select Case FindString
            Case "DAT"
                CurrentFile = CurrentPath + "\" + Form5.lstFoundFiles.List(CurrentIndex)
                Form2.MediaPlayer.FileName = CurrentFile
            Case Else
                CurrentFile = CurrentPath + "\" + Form5.lstFoundFiles.List(CurrentIndex)
                Form2.MediaPlayer.FileName = CurrentFile
        End Select
    End If
End Function

Public Function TestCDdrive()
    On Error GoTo Handler
    If IsCDPresent Then
        Open CDROMDRIVE + "mpegav\avseq01.dat" For Input As #1
        Close #1
        '-----------------
        IsCDVideo = True
        IsCDDVD = True
        IsCDReady = True
        '---------------------
        Form6.Show
        Form6.Refresh
        '-----------------
        OpenWindow = False
        '-----------------
        CurrentPath = CDROMDRIVE + "mpegav"
        '----------------------------------
        On Error GoTo 0
        Exit Function
    End If
IsCDDVD = False
IsCDVideo = False
IsCDReady = False
On Error GoTo 0
Close #1
Exit Function

Handler:
    DoEvents
    Select Case Err
        Case 71
            DoEvents
            '--- Disk Not Ready -------
            IsCDReady = False
            CurrentIndex = 0
            OpenWindow = True
        Case 53
            DoEvents
            '--- File Not Found -------
            IsCDReady = True
            IsCDDVD = False
        Case 76
            DoEvents
            '--- Path Not Found -------
            IsCDReady = True
            IsCDDVD = False
        Case -2147024865
            DoEvents
            IsCDReady = False
            IsCDDVD = False
            CurrentIndex = 0
            OpenWindow = True
            Exit Function
        Case Else
            DoEvents
            End
            MsgBox Error$ & "--" & Err
    End Select
On Error GoTo 0
End Function
Public Function DirDiver(NewPath As String, DirCount As Integer, BackUp As String) As Integer
Static FirstErr As Integer
Dim DirsToPeek As Integer, AbandonSearch As Integer, ind As Integer
Dim OldPath As String, ThePath As String, entry As String
Dim retval As Integer
    SearchFlag = True           ' Set flag so the user can interrupt.
    DirDiver = False            ' Set to True if there is an error.
    retval = DoEvents()         ' Check for events (for instance, if the user chooses Cancel).
    If SearchFlag = False Then
        DirDiver = True
        Exit Function
    End If
    On Local Error GoTo DirDriverHandler
    DirsToPeek = Form4.DirList.ListCount                  ' How many directories below this?
    Do While DirsToPeek > 0 And SearchFlag = True
        OldPath = Form4.DirList.Path                      ' Save old path for next recursion.
        Form4.DirList.Path = NewPath
        If Form4.DirList.ListCount > 0 Then
            ' Get to the node bottom.
            Form4.DirList.Path = Form4.DirList.List(DirsToPeek - 1)
            AbandonSearch = DirDiver((Form4.DirList.Path), DirCount%, OldPath)
        End If
        ' Go up one level in directories.
        DirsToPeek = DirsToPeek - 1
        If AbandonSearch = True Then Exit Function
    Loop
    ' Call function to enumerate files.
    If Form4.FilList.ListCount Then
        If Len(Form4.DirList.Path) <= 3 Then             ' Check for 2 bytes/character
            ThePath = Form4.DirList.Path                  ' If at root level, leave as is...
        Else
            ThePath = Form4.DirList.Path + "\"            ' Otherwise put "\" before the filename.
        End If
        For ind = 0 To Form4.FilList.ListCount - 1        ' Add conforming files in this directory to the list box.
            entry = ThePath + Form4.FilList.List(ind)
            Form5.lstFoundFiles.AddItem entry
        Next ind
    End If
    If BackUp <> "" Then        ' If there is a superior directory, move it.
        Form4.DirList.Path = BackUp
    End If
    Exit Function
DirDriverHandler:
    DoEvents
    If Err = 7 Then             ' If Out of Memory error occurs, assume the list box just got full.
        DirDiver = True         ' Create Msg and set return value AbandonSearch.
        MsgBox "You've filled the list box. Abandoning search..."
        Exit Function           ' Note that the exit procedure resets Err to 0.
    Else
                                ' Otherwise display error message and quit.
        If IsCDReady Then
            If Err = 68 Then
                Form6.Show
                Form6.Refresh
                Resume
            End If
        End If
        
        MsgBox Error
        
        End
    End If
End Function

Public Function strfnFindCD() As String
    Dim strDrives As String     'store list of drives
    Dim strThisDrive As String  'one drive from strDrives
    Dim intCounter As Integer
    Dim lngLenStr As Long       'length of returned string
    Dim lngDriveType As Long    'gets drive type as follows
    Dim Responce
    'Drive Types:
    '0: drive type cannot be determined
    '1: specified drive doesn't exist
    '2: removeable-disk can be removed, eg floppy or zip
    '3: fixed-disk cannot be removed, eg hard disk
    '4: remote-eg remote network drive
    '5: CD-ROM drive
    '6: RAM disk
    
    strDrives = Space$(255)
    
    lngLenStr = GetLogicalDriveStrings(255, strDrives)
    'strDrives has the names of all the root directories.
    'Each entry takes four characters-three for the name plus
    'a null character. String ends with a second null.
    'Example: A:\[Null]C:\[Null]D:\[Null][Null]
    
    For intCounter = 1 To lngLenStr Step 4
    'Count by fours to get the letter of each drive
        strThisDrive = Mid$(strDrives, intCounter, 3)
        lngDriveType = GetDriveType(strThisDrive)
        If lngDriveType = 5 Then 'It's a CDROM
            strfnFindCD = strThisDrive
            CDROMDRIVE = strfnFindCD
            IsCDPresent = True
            CurrentDrive = CDROMDRIVE
            Exit Function
        End If
    Next intCounter
    
    IsCDPresent = False
    strfnFindCD = "None" 'System doesn't have a CDROM
    Responce = MsgBox("No Valid CD-ROM Drive Found", vbExclamation + vbOKOnly, "Teknet Media Manager")
    End
End Function
Public Function DirExists(ByVal sDirName As String) As Boolean
    On Error Resume Next
    DirExists = (GetAttr(sDirName) And vbDirectory) = vbDirectory
    On Error GoTo 0
End Function
Public Function FileExists(ByVal sPathName As String) As Boolean
    On Error Resume Next
    FileExists = (GetAttr(sPathName) And vbNormal) = vbNormal
    On Error GoTo 0
End Function

Public Function Startup()
'======================================================
    strfnFindCD
    If IsCDPresent Then TestCDdrive
    If IsCDReady Then FindDVD
    Unload Form6
    Form3.Show
'======================================================
End Function

