<div align="center">

## A Simple Directory Search Module


</div>

### Description

I tried to make this as light as possible just add it to your app and call SearchDirs. It takes two arguments the path to search in and the file or directory to find. It puts alot of Data on the Call Stack however it's very very fast. 3 API's, 5 Constants, 3 User Types, no activex or OLE runtime objects.

Thanks and feel free to email if you have any questions
 
### More Info
 
Boolean


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Chris King](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/chris-king.md)
**Level**          |Advanced
**User Rating**    |4.3 (26 globes from 6 users)
**Compatibility**  |VB 4\.0 \(32\-bit\), VB 5\.0, VB 6\.0
**Category**       |[Windows API Call/ Explanation](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/windows-api-call-explanation__1-39.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/chris-king-a-simple-directory-search-module__1-14249/archive/master.zip)

### API Declarations

```
Private Declare Function FindFirstFile Lib "kernel32" Alias "FindFirstFileA" _
 (ByVal lpFileName As String, lpFindFileData As WIN32_FIND_DATA) As Long
 Private Declare Function FindClose Lib "kernel32" (ByVal hFindFile As Long) As Long
 Private Declare Function FindNextFile Lib "kernel32" Alias "FindNextFileA" _
 (ByVal hFindFile As Long, lpFindFileData As WIN32_FIND_DATA) As Long
 Private Const MaxLFNPath = 260
 Private Const INVALID_HANDLE_VALUE = -1
 Private Const VBKEYDOT = 46
 Private Const VBBACKSLASH = "\"
 Private Const VBALLFILES = "*.*"
 Private Type FILETIME
 dwLowDateTime As Long
 dwHighDateTime As Long
 End Type
 Private Type WIN32_FIND_DATA
 dwFileAttributes As Long
 ftCreationTime As FILETIME
 ftLastAccessTime As FILETIME
 ftLastWriteTime As FILETIME
 nFileSizeHigh As Long
 nFileSizeLow As Long
 dwReserved0 As Long
 dwReserved1 As Long
 cFileName As String * MaxLFNPath
 cShortFileName As String * 14
 End Type
 Private WFD As WIN32_FIND_DATA
```


### Source Code

```
'Chris King 01/08/2000 c_king@mtv.com
 Option Explicit
Public Function SearchDirs(Curpath$, strFName$)
 Dim strProg$
 Dim dirs%
 Dim dirbuf$()
 Dim hItem&
 Dim i%
 Dim rtn As Boolean
 If Curpath$ = "" Then Exit Function
 If strFName$ = "" Then Exit Function
 If Right(strFName$, 1) = VBBACKSLASH Then
 strFName = Left(strFName, InStr(1, strFName, VBBACKSLASH, vbTextCompare) - 1)
 End If
 If Right(Curpath$, 1) <> VBBACKSLASH Then
 Curpath$ = Curpath$ & VBBACKSLASH
 End If
 hItem& = FindFirstFile(Curpath$ & VBALLFILES, WFD)
 If hItem& <> INVALID_HANDLE_VALUE Then
 Do
 If (WFD.dwFileAttributes And vbDirectory) Then
 If Asc(WFD.cFileName) <> VBKEYDOT Then
 If (dirs% Mod 10) = 0 Then ReDim Preserve dirbuf$(dirs% + 10)
 dirs% = dirs% + 1
 dirbuf$(dirs%) = Left$(WFD.cFileName, InStr(WFD.cFileName, vbNullChar) - 1)
 End If
 End If
 strProg$ = Left(WFD.cFileName, InStr(WFD.cFileName, vbNullChar) - 1)
 If UCase(strProg$) = UCase(strFName$) Then
 SearchDirs = True
 Exit Function
 Else
 SearchDirs = False
 End If
 DoEvents
 Loop While FindNextFile(hItem&, WFD)
 Call FindClose(hItem&)
 End If
 For i% = 1 To dirs%
 rtn = SearchDirs(Curpath$ & dirbuf$(i%) & VBBACKSLASH, strFName$)
 SearchDirs = rtn
 If rtn Then Exit Function
 Next i%
End Function
```

