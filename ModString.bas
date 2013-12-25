Attribute VB_Name = "ModString"
' =============================================================================
' Module Name:      ModString
' Description:      Common (and not so common) string handling and parsing
'                   subs and functions.
' Author(s):        Jonathan A. Conway
' -----------------------------------------------------------------------------
' Log:
'                   04 07 05 : Updated InStr32() with a version that fixes
'                   some pretty major bugs.
'                   04 05 02 : Created GetFilter(), ShiftArrayToZero() and
'                   IsWordTime() procedures.
' =============================================================================


Option Explicit


' Public Methods
' ==============

Public Function TrimWhiteSpace(ByRef Source As String)

    Dim sSrc As String
    Dim bIsDone As Boolean
    Dim sChar As String
    
    sSrc = Source
    sSrc = Trim$(sSrc)
    
    bIsDone = False
    
    Do Until bIsDone
        sChar = Left$(sSrc, 1)
        If sChar = vbTab _
            Or sChar = Chr(13) _
            Or sChar = Chr(10) Then
            
            sSrc = Mid$(sSrc, 2, Len(sSrc) - 1)
        Else
            bIsDone = True
        End If
    Loop
    
    bIsDone = False
    
    Do Until bIsDone
        sChar = Right$(sSrc, 1)
        If sChar = vbTab _
            Or sChar = Chr(13) _
            Or sChar = Chr(10) Then
            
            sSrc = Mid$(sSrc, 1, Len(sSrc) - 1)
        Else
            bIsDone = True
        End If
    Loop
    
    TrimWhiteSpace = sSrc

End Function


Public Function GetEndWithoutLineBreaks(ByRef Source As String) As Long

    ' Returns a long integer representing the index position at which
    ' the specified string should end. This value can be used to save
    ' a string to disk without writing in unnecessary line-breaks.

    GetEndWithoutLineBreaks = -1
    
    Dim lPos As Long
    Dim bChr As Byte
    
    lPos = Len(Source)
    
    Do While True
    
        bChr = Asc(Mid$(Source, lPos, 1))
    
        If bChr = 13 Or bChr = 10 Then
            lPos = lPos - 1
        Else
            GetEndWithoutLineBreaks = lPos
            Exit Function
        End If
        
    Loop

End Function


Public Function GetOccurrences(ByRef Source As String, ByVal MatchString As String, ByVal Compare As VbCompareMethod) As Long

    ' Returns the total number of occurrances of one string within
    ' another.
    
    ' Works happily with strings over 32, 767 bytes in size.

    Dim lPos As Long
    Dim lInStr As Long
    Dim lCount As Long
    Dim bIsDone As Boolean
    
    bIsDone = False
    lPos = 1
    
    Do Until bIsDone
        lInStr = InStr32(lPos, Source, MatchString, Compare)
        If lInStr = 0 Then
            bIsDone = True
        Else
            lCount = lCount + 1
            lPos = lInStr + Len(MatchString)
        End If
    Loop

    GetOccurrences = lCount
    
End Function


Public Function SplitBySpace(ByVal Source As String) As String()

    ' Splits a string into an array of strings by the first blank
    ' space.
    
    ' For example...
    
    ' This string...
    ' "definequote 533           HELLO, WORLD!!"
    
    ' Would by split like this using the regular Split$() function...
    ' "definequote"
    ' "533"
    ' "HELLO,"
    ' "WORLD!!"
    
    ' But using SplitBySpace$(), we get...
    ' "definequote"
    ' " 533"
    ' "           HELLO,"
    ' " WORLD!!"
    
    Dim sReturn() As String
    Dim sBuffer As String
    Dim iLen As Integer
    Dim c As Byte
    Dim i As Integer
    
    ReDim sReturn(0) As String
    
    iLen = Len(Source)
    
    For i = 1 To iLen
        c = Asc%(Mid$(Source, i, 1))
        If c <> 32 Then
            ' Append character to the buffer
            sBuffer = sBuffer & Chr$(c)
            
            If i = iLen Then
                ' Add buffer contents to array
                ReDim Preserve sReturn(LBound(sReturn) To UBound(sReturn) + 1)
                sReturn(UBound(sReturn)) = sBuffer
            End If
        Else
            If Len(Trim$(sBuffer)) > 0 Then
                ' Add buffer contents to array
                ReDim Preserve sReturn(LBound(sReturn) To UBound(sReturn) + 1)
                sReturn(UBound(sReturn)) = sBuffer
                
                ' Reset buffer
                sBuffer = Chr$(32)
            Else
                ' Append blank space to the buffer
                sBuffer = sBuffer & Chr(32)
            End If
        End If
    Next

    ShiftArrayToZero sReturn
    SplitBySpace = sReturn

End Function

Public Function GetLineCount(ByVal Source As String) As Long

    Dim s As String
    Dim i As Integer
    Dim b As Byte
    Dim iCount As Integer
    
    If Len(Trim(Source)) = 0 Then
        iCount = 0
    Else
        iCount = 1
        s = Source
        s = Replace(s, Chr(13) & Chr(10), Chr(13))
        
        For i = 1 To Len(s)
            b = Asc(Mid$(s, i, 1))
            If b = 13 Then
                iCount = iCount + 1
            End If
        Next
    End If
    
    GetLineCount = iCount

End Function

Public Function GetFirstLine(ByRef Source As String) As String
    Dim i As Integer
    i = InStr(1, Source, vbNewLine)
    If i = 0 Then
        GetFirstLine = Source
    Else
        GetFirstLine = Mid$(Source, 1, i)
    End If
End Function

Public Function GetLastLine(ByRef Source As String) As String
    Dim i As Long
    i = InStrRev(Source, Chr(13))
    If i = 0 Then
        GetLastLine = Source
    Else
        GetLastLine = Mid$(Source, i, (Len(Source) - i) + 1)
    End If
End Function


Public Function InStr32(ByVal Start As Long, ByRef Source As String, ByVal FindWhat As String, ByVal CompareMethod As VbCompareMethod) As Long
    
    ' Replacement function for InStr() that works with strings longer
    ' that 32,767 bytes. Return value and arguments are Long Integers.
    
    Dim sChunks()   As String       ' 32767-byte-long chunks of text
    Dim iLen        As Long         ' Length of total length of source
    Dim i           As Long         ' Loop counter
    Dim iResult     As Long         ' Result
    Dim sSource     As String       ' Source code
    
    If iLen > 32767 Then
        ' Calculate length (taking into account starting position)
        ' This is LengthOfSource - Start - 1
        ' The -1 is to convert the one-based length to a zero-based position
        iLen = Len(Source) - (Start - 1)
        sSource = Mid$(Source, Start, iLen)
        
        ReDim sChunks(0 To CInt(iLen / 32767))
        For i = 0 To UBound(sChunks)
            sChunks(i) = Mid$(sSource, (32767 * i) + 1, 32767)
        Next
        
        i = 0
        Do While i <= UBound(sChunks)
            iResult = InStr(1, sChunks(i), FindWhat, CompareMethod)
            If iResult > 0 Then
                iResult = iResult + (32767 * i)
                Exit Do
            End If
            i = i + 1
        Loop
        
        If iResult > 0 Then
            iResult = iResult + Start
        End If
    Else
        iResult = InStr(Start, Source, FindWhat, CompareMethod)
    End If

    InStr32 = iResult

End Function

Public Sub GetWordList(ByRef Source As String, ByRef outWordList() As String)

    Dim s As String
    s = Source

    s = Replace(s, Chr(10), " ")    ' Replace line feeds
    s = Replace(s, Chr(13), " ")    '  and carriage returns
    s = Replace(s, vbTab, " ")      '   and tabs

    outWordList = SplitBySpace(s)   ' Split string

End Sub

Public Sub InStrMulti(ByVal Start As Long, ByRef Source As String, ByRef SearchFor() As String, ByRef outPosition As Long, ByRef outItemFound As Long)

    ' Searches a string and returns the first item found (and its position)
    ' out of several search items.

    Dim i               As Integer
    Dim lInStr          As Long
    
    Dim lCurrentPos     As Long
    Dim lCurrentItem    As Long
    
    Dim bMatchFound     As Boolean
    
    bMatchFound = False
    
    ' Record first item
    'lCurrentPos = InStr32(Start, Source, SearchFor(LBound(SearchFor)), vbTextCompare)
    
    For i = LBound(SearchFor) To UBound(SearchFor)
        ' Get location of item
        lInStr = InStr32(Start, Source, SearchFor(i), vbTextCompare)
        If lInStr > 0 Then
            If bMatchFound = False Or lInStr < lCurrentPos Then
                lCurrentPos = lInStr
                lCurrentItem = i
                bMatchFound = True
            End If
        End If
    Next

    ' Return results (may be null/minus values)
    outPosition = lCurrentPos
    outItemFound = lCurrentItem
    
End Sub

Public Sub MoveOutOfWord(ByRef Source As String, ByRef Position As Long)

    ' This function will check if a position within a string is inside a
    ' word (i.e. any character other than a space, tab or line break) and
    ' if it is, it moves it backward until it is 'outside' the word.
    
    ' Input:    Source string
    ' Input:    Begin position within string
    
    ' Output:   New position within string (that is outside a word)

    Dim sB As String, sA As String
    Dim lPos As Long
        
    lPos = Position
    
    Do While lPos > 0
        sB = Mid$(Source, Position, 1)
        sA = Mid$(Source, Position - 1, 1)
        If (sB = " " Or sB = vbTab Or sB = Chr(10) Or sB = Chr(13)) Or _
            (sA = " " Or sA = vbTab Or sA = Chr(10) Or sA = Chr(13)) Then
            Position = lPos
            Exit Do
        End If
        lPos = lPos - 1
    Loop

End Sub


Public Function IsLineWhitespace(ByVal Source As String) As Boolean

    ' Returns a Boolean indicating whether or not the specified string
    ' contains nothing but whitespace (i.e. spaces, tabs, line breaks)

    Dim sText As String
    
    sText = Source
    sText = Replace$(sText, " ", "")
    sText = Replace$(sText, Chr(9), "")
    sText = Replace$(sText, Chr(10), "")
    sText = Replace$(sText, Chr(13), "")

    IsLineWhitespace = (Len(Trim$(sText)) = 0)

End Function


' Private Methods
' ===============

Private Sub ShiftArrayToZero(ByRef SourceArray() As String)
    ' Shifts every item in SourceArray back by one and removes the last item
    Dim i As Integer
    For i = LBound(SourceArray) To (UBound(SourceArray) - 1)
        SourceArray(i) = SourceArray(i + 1)
    Next
    ReDim Preserve SourceArray(LBound(SourceArray) To UBound(SourceArray) - 1)
End Sub


Public Function IsWordTime(ByVal Word As String) As Boolean
    ' Is the input string in valid Time format?
    Dim sWord As String
    sWord = Trim$(Word)
    IsWordTime = False
    If InStr(1, sWord, ":") <> 0 Then
        If IsNumeric(Replace$(sWord, ":", "")) Then
            IsWordTime = True
        End If
    End If
End Function

