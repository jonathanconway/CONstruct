Attribute VB_Name = "ModFile"
' =============================================================================
' Module Name:      ModFile
' Module Type:      Code Module
' Description:      API wrapper for file/folder manipulation
' Author(s):        VB Web, Edited by Jonathan A. Conway
' -----------------------------------------------------------------------------
' Log:
'
' 04 08 23 :
'   - ModFile integrated into CONstruct, modified by JAC
' 00 03 18 :
'   - Contents of ModFile written
' =============================================================================


Option Explicit
Public Const MAX_PATH = 260

Private Const ERROR_NO_MORE_FILES = 18&
Private Const FILE_ATTRIBUTE_DIRECTORY = &H10

Private Const FILE_ATTRIBUTE_NORMAL = &H80
Private Const FILE_ATTRIBUTE_HIDDEN = &H2
Private Const FILE_ATTRIBUTE_SYSTEM = &H4
Private Const FILE_ATTRIBUTE_TEMPORARY = &H100

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
    cFileName As String * MAX_PATH
    cAlternate As String * 14
End Type

Private Declare Function FindFirstFile Lib "kernel32" Alias "FindFirstFileA" _
    (ByVal lpFileName As String, lpFindFileData As WIN32_FIND_DATA) As Long
Private Declare Function FindClose Lib "kernel32" (ByVal hFindFile As Long) As Long

Private Declare Function FindNextFile Lib "kernel32" Alias "FindNextFileA" _
    (ByVal hFindFile As Long, lpFindFileData As WIN32_FIND_DATA) As Long

Private Function StripTerminator(ByVal strString As String) As String
    Dim intZeroPos As Integer
    intZeroPos = InStr(strString, Chr$(0))
    If intZeroPos > 0 Then
        StripTerminator = Left$(strString, intZeroPos - 1)
    Else
        StripTerminator = strString
    End If
End Function

Public Function IsFileExistant(Optional ByVal sFile As String = "", _
        Optional ByVal sFolder As String = "") As Boolean

    Dim lpFindFileData As WIN32_FIND_DATA
    Dim lFileHdl  As Long
    Dim sTemp As String
    Dim sTemp2 As String
    Dim lRet As Long
    Dim iLastIndex  As Integer
    Dim strPath As String
    Dim sStartDir As String
    
    'On Error Resume Next
    '// both params are empty
    If sFile = "" And sFolder = "" Then Exit Function
    '// both are full, empty folder param
    If sFile <> "" And sFolder <> "" Then sFolder = ""
    If sFolder <> "" Then
        '// set start directory
        sStartDir = sFolder
    Else
        '// extract start directory from file path
        sStartDir = Left$(sFile, InStrRev(sFile, "\"))
        '// just get filename
        sFile = Right$(sFile, Len(sFile) - InStrRev(sFile, "\"))
    End If
    '// add trailing \ to start directory if required
    If Right$(sStartDir, 1) <> "\" Then sStartDir = sStartDir & "\"
    
    sStartDir = sStartDir & "*.*"
    
    '// get a file handle
    lFileHdl = FindFirstFile(sStartDir, lpFindFileData)
    
    If lFileHdl <> -1 Then
        If sFolder <> "" Then
            '// folder exists
            IsFileExistant = True
        Else
            Do Until lRet = ERROR_NO_MORE_FILES
                strPath = Left$(sStartDir, Len(sStartDir) - 4) & "\"
                '// if it is a file
                If (lpFindFileData.dwFileAttributes And FILE_ATTRIBUTE_NORMAL) = vbNormal Then
                    sTemp = StrConv(StripTerminator(lpFindFileData.cFileName), vbProperCase)
                    '// remove LCase$ if you want the search to be case sensitive (unlikely!)
                    If LCase$(sTemp) = LCase$(sFile) Then
                        IsFileExistant = True '// file found
                        Exit Do
                    End If
                End If
                '// based on the file handle iterate through all files and dirs
                lRet = FindNextFile(lFileHdl, lpFindFileData)
                If lRet = 0 Then Exit Do
            Loop
        End If
    End If
    '// close the file handle
    lRet = FindClose(lFileHdl)
End Function


