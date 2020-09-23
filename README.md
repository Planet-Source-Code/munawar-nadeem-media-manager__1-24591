<div align="center">

## Media Manager


</div>

### Description

A complete Media Manager Project, that can play almost all types of media formats like, DVD, MP3, AVI, from CD-ROM or Hard Drive. Built in function "strfnFindCD" provides automatic detection of your CD-ROM drive. Auto play VCDs or DVDs. Select multiple media files and play without any interruption.
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |2001-06-30 16:19:06
**By**             |[Munawar Nadeem](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/munawar-nadeem.md)
**Level**          |Advanced
**User Rating**    |5.0 (20 globes from 4 users)
**Compatibility**  |VB 6\.0
**Category**       |[Complete Applications](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/complete-applications__1-27.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[Media Mana219536302001\.zip](https://github.com/Planet-Source-Code/munawar-nadeem-media-manager__1-24591/archive/master.zip)

### API Declarations

```
Private Declare Function GetDriveType Lib "kernel32" _
 Alias "GetDriveTypeA" (ByVal nDrive As String) As Long
Private Declare Function GetLogicalDriveStrings _
 Lib "kernel32" Alias "GetLogicalDriveStringsA" _
 (ByVal nBufferLength As Long, _
 ByVal lpBuffer As String) As Long
```





