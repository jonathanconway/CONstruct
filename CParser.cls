VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CParser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' =============================================================================
' Module Name:      CParser
' Module Type:      Class
' Description:      Multi-purpose generic parser; takes a source string and a
'                   series of token defenitions as input and returns a
'                   collection of matching blocks as output.
'
' Dependencies:     - Function InStr32()
'                   - Class CBlock
'                   // - Class CBlocks
'                   - Class CToken
'
' Author(s):        Jonathan A. Conway
' -----------------------------------------------------------------------------
' Log:
'
' 04 08 11 :
'   - Added some error handling code to the Parse() method as a kind of
'     "quick-fix" solution to a bug. The bug happens when a beginning token
'     is found right at the end of the file, with no matching end token. After
'     the fix, the parser simply ignores that structure, which is OK for now.
'
' 04 07 05 :
'   - Added code for assigning generated code blocks a "structure" tag to
'     identify what kind of structure each one describes.
'   - Fixed bug that occurs when an end token isn't found -- the program
'     should now deal with this problem by setting the end marker to the end
'     of the source text.
'     NOTE: The goal is to make CONstruct handle badly formed script code
'     withot crashing or exhibiting wierd side-effects.
'
' 04 06 19 :
'   - Commenced writing CToken from scratch
' =============================================================================



Option Explicit



Private WithEvents m_oBlocks As CObjectCollection
Attribute m_oBlocks.VB_VarHelpID = -1

Private m_oDefinition As CDefinition

Private m_bIsLoading As Boolean
Private m_bIsParsing As Boolean


Public Event ParseProgress(ByVal CurrentPosition As Long, _
                           ByVal TotalLength As Long)
Public Event ParseComplete()

Public Event BlockAdded(ByRef Block As CBlock)
Public Event BlockDeleted(ByVal BlockIndex As Integer)



Public Property Get IsLoading() As Boolean
    IsLoading = m_bIsLoading
End Property


Public Property Get IsParsing() As Boolean
    IsParsing = m_bIsParsing
End Property


Private Function GetStretchBefore(ByRef Source As String, ByVal Position As Long) As String

    If Position > 32767 Then
        GetStretchBefore = Mid$(Source, Position - 32767, 32767)
    Else
        GetStretchBefore = Mid$(Source, 1, Position)
    End If

End Function


Public Function GetObjectAtPos(ByRef Source As String, ByVal Position As Long) As CBlock
    
    Dim sBefore As String
    Dim sAfter As String
    Dim oBeginToken As CToken
    Dim oEndToken As CToken
    Dim lBeginPos As Long
    Dim lEndPos As Long
    Dim bNoEndToken As Boolean
    Dim lTotalLength As Long

    Dim oBlock As CBlock
    Dim oChosenBlock As CBlock
    Dim sText As String
    Dim sBeginSignature As String

    m_bIsLoading = True

    lTotalLength = Len(Source)
    
    If Position > lTotalLength Or Position < 1 Then
        m_bIsLoading = False
        Exit Function
    End If
    
    ' 1. Get stretch of up to 32767 characters before current pos
    sBefore = GetStretchBefore(Source, Position)
    sAfter = Mid$(Source, Position + 1, 32767)
    
    ' 2. From that stretch, locate nearest begin token
    LocateAnyBeginTokenReverse sBefore, oBeginToken, lBeginPos
    
    If (lBeginPos > 0) And (Not oBeginToken Is Nothing) Then
    
        ' 3. From the nearest begin token, locate end token
        LocateTokenOfStructure _
                        sBefore & sAfter, _
                        oBeginToken.Structure, _
                        lBeginPos + Len(oBeginToken.Text), _
                        0, _
                        tktEnder, _
                        oEndToken, _
                        lEndPos

        ' Calculate end position of the block
        If Not oEndToken Is Nothing Then
            ' End token found -- well formed code
            ' End marker = InStr of token + Length of token - Begin pos
            lEndPos = ((lEndPos + Len(oEndToken.Text)) - lBeginPos)
            bNoEndToken = False
        Else
            ' No end token found -- badly formed code!!
            ' End marker = end of text
            lEndPos = (lTotalLength - lBeginPos) + 1
            bNoEndToken = True
        End If
        
        ' 4. Check signature against signatures of other objects. If no match
        '    is found:
        '       a. Create new block
        '       b. Insert the code of the new object in it
        '       c. Refresh the list of current objects for deletes
        '       d. Return the new object
        
        'If (lBeginPos + lEndPos) >= Len(sBefore) Then
        If Position < (lBeginPos + lEndPos) Then
        
            sText = Mid$(sBefore & sAfter, lBeginPos, lEndPos)
            sBeginSignature = GetFirstLine(sText)
            
            For Each oBlock In m_oBlocks
                If (InStr(1, sBeginSignature, oBlock.BeginSignature) <> 0) _
                    Or (InStr(1, oBlock.BeginSignature, sBeginSignature) <> 0) Then
                
                'If sBeginSignature = oBlock.BeginSignature Then
                    Set oChosenBlock = oBlock
                    Exit For
                End If
            Next
            
            If oChosenBlock Is Nothing Then
            
                ' Add a new block for this structure
                Set oChosenBlock = New CBlock
                With oChosenBlock
                    
                    ' Set block properties
                    .Structure = oBeginToken.Structure
                    .Text = sText
                    
                    ' Add "noendmarker" error
                    If bNoEndToken Then
                        .AddError eBlockErrors.be_NoEndMarker
                    End If
            
                    m_oBlocks.Add oChosenBlock
                
                End With
            
                ' Clean up invalid blocks
                CleanUpBlocks Source
            
            End If
            
            Set GetObjectAtPos = oChosenBlock
        
        End If
    
    End If
    
    DoEvents
    
    m_bIsLoading = False
    
End Function


Public Sub CleanUpBlocks(ByRef Source As String)

    ' Does some basic "housekeeping" for blocks, removing blocks whose
    ' signature can no longer be found in code, as well as updating
    ' the text of blocks where necessary.
    
    ' If you're integrating CParser into your own app, it's a good
    ' idea to have this method called once every few seconds.
    
    Dim oBlock As CBlock
    Dim oBeginToken As CToken
    Dim lBeginPos As Long
    Dim oEndToken As CToken
    Dim lEndPos As Long
    Dim bNoEndToken As Boolean
    Dim lTotalLength As Long
    
    bNoEndToken = False
    lTotalLength = Len(Source)
    
    For Each oBlock In m_oBlocks
        lBeginPos = InStr32(1, Source, oBlock.BeginSignature, vbTextCompare)
        
        If lBeginPos = 0 Then
            ' Delete block if not found...
            m_oBlocks.Delete oBlock.Index
        Else
            ' Else, update block source
            
            ' Get the token used...
            LocateTokenOfStructure Source, oBlock.Structure, lBeginPos, 0, tktBeginner, oBeginToken, lBeginPos
            
            ' Make sure signature begins on new line
            If oBeginToken Is Nothing Then
                ' Delete block if not found...
                m_oBlocks.Delete oBlock.Index
                Exit Sub
            End If
                
            If oBeginToken.BeginsOnNewLine Then
                If Not IsPositionOnNewLine(Source, lBeginPos) Then
                    ' Delete block if not found...
                    m_oBlocks.Delete oBlock.Index
                    Exit Sub
                End If
            End If
                
            ' Locate end token for Block=>Structure=>Token
            LocateTokenOfStructure _
                            Source, _
                            oBlock.Structure, _
                            lBeginPos, _
                            0, _
                            tktEnder, _
                            oEndToken, _
                            lEndPos
    
            ' Calculate end position of the block
            If Not oEndToken Is Nothing And lEndPos > -1 Then
                ' End token found -- well formed code
                ' End marker = InStr of token + Length of token - Begin pos
                lEndPos = ((lEndPos + Len(oEndToken.Text)) - lBeginPos)
                bNoEndToken = False
            Else
                ' No end token found -- badly formed code!!
                ' End marker = end of text
                lEndPos = (lTotalLength - lBeginPos) + 1
                bNoEndToken = True
            End If
            
            oBlock.Text = Mid$(Source, lBeginPos, lEndPos)
        End If
    
    Next

End Sub


Public Sub Parse(ByRef Source As String, Optional ByVal IncrementalParse As Boolean)

    On Error GoTo ProcedureError

    Dim bDone As Boolean
    Dim bNoEndToken As Boolean
    
    Dim lInStr As Long
    Dim lTotalLength As Long
    
    Dim oBeginToken As CToken
    Dim oEndToken As CToken
    
    Dim lBeginPos As Long
    Dim lEndPos As Long
    Dim lCursor As Long
    
    Dim oBlock As CBlock
    
    Dim sSource As String
    
    m_bIsLoading = True         ' Set IsLoading flag
    m_bIsParsing = True         ' Set IsParsing flag
    lCursor = 1
    lTotalLength = Len(Source)
    bDone = False
    
    ' Clear blocks prior to adding items, IF SPECIFIED!!
    If IsMissing(IncrementalParse) Then
        m_oBlocks.Clear
    Else
        If IncrementalParse = False Then
            m_oBlocks.Clear
        End If
    End If
    
    sSource = vbNewLine & Source
    
    Do While Not bDone
        
        RaiseEvent ParseProgress(lCursor, lTotalLength)
        
        ' Try to find a beginner token
        LocateAnyBeginToken sSource, lCursor, 0, oBeginToken, lBeginPos
    
        If (lBeginPos > 0) And (Not oBeginToken Is Nothing) Then
            
            ' Beginner token located!
            ' Now try to find an ending token
            LocateTokenOfStructure _
                            sSource, _
                            oBeginToken.Structure, _
                            lBeginPos + Len(oBeginToken.Text), _
                            0, _
                            tktEnder, _
                            oEndToken, _
                            lEndPos
        
            If Not oEndToken Is Nothing Then
                ' End token found -- well formed code
                ' End marker = InStr of token + Length of token - Begin pos
                lEndPos = ((lEndPos + Len(oEndToken.Text)) - lBeginPos)
                    
                bNoEndToken = False
            Else
                ' No end token found -- badly formed code!!
                ' End marker = end of text
                lEndPos = (lTotalLength - lBeginPos) + 1
                bNoEndToken = True
            End If
            
            ' Add a new block for this structure
            Set oBlock = New CBlock
            With oBlock
                ' Set block properties
                .Structure = oBeginToken.Structure
                .Text = Mid$(sSource, lBeginPos, lEndPos)
                
                ' Add "noendmarker" error
                If bNoEndToken Then
                    .AddError eBlockErrors.be_NoEndMarker
                End If
            End With
            m_oBlocks.Add oBlock
            
            RaiseEvent BlockAdded(oBlock)
        
            lCursor = lBeginPos + lEndPos
        
        Else
            bDone = True
        End If
        
        DoEvents
    
    Loop

    m_bIsLoading = False        ' Reset IsLoading flag
    m_bIsParsing = False        ' Reset IsParsing flag
    RaiseEvent ParseComplete    ' Signal that parsing is completed

    Exit Sub
    
ProcedureError:
    If err.Number = 5 Then
        
    Else
        MsgBox err.Number & vbNewLine & err.Description
    End If
    m_bIsLoading = False        ' Reset IsLoading flag

End Sub


Public Function ParseKeyWords(ByRef Source As String) As CObjectCollection

    'On Error GoTo ProcedureError

    Dim bDone As Boolean
    
    Dim lInStr As Long
    Dim lTotalLength As Long
    
    Dim lBeginPos As Long
    Dim lCursor As Long
    
    Dim oKeyWord As CKeyWord
    Dim oKeyWords As CObjectCollection
    
    
    Dim sSource As String
    
    
    m_bIsLoading = True         ' Set IsLoading flag
    m_bIsParsing = True         ' Set IsParsing flag
    lCursor = 1
    lTotalLength = Len(Source)
    bDone = False
    Set oKeyWords = New CObjectCollection
    
    sSource = vbNewLine & Source
    
    Do While Not bDone
        
        ' Look for any keyword within the source
        LocateAnyKeyword sSource, lCursor, 0, oKeyWord, lBeginPos
    
        If (lBeginPos > 0) And (Not oKeyWord Is Nothing) Then
            
            ' Keyword located
            
            ' Add to the collection
            
            oKeyWords.Add oKeyWord
            
            lCursor = lBeginPos + Len(oKeyWord.Text)
        
        Else
            bDone = True
        End If
        
        DoEvents
    
    Loop

    m_bIsLoading = False        ' Reset IsLoading flag
    m_bIsParsing = False        ' Reset IsParsing flag

    Set ParseKeyWords = oKeyWords
    
    Exit Function
    
ProcedureError:
    If err.Number = 5 Then
        
    Else
        MsgBox err.Number & vbNewLine & err.Description
    End If
    m_bIsLoading = False        ' Reset IsLoading flag

End Function









Private Function IsPositionOnNewLine(ByRef Source As String, ByVal Position As Long) As Boolean

    Dim bIsDone As Boolean
    Dim bResult As Boolean
    Dim lPos As Long
    
    bResult = False
    bIsDone = False
    lPos = IIf(Position > 1, Position - 1, Position)
    
    Do Until bIsDone
        Select Case Asc(Mid$(Source, lPos, 1))
            Case 32     ' Space; ignore
            Case 9      ' Tab; ignore
            Case 10     ' Newline; ignore
            
            Case 13     ' Carriage Return
                bResult = True
                bIsDone = True
            
            Case Else   ' Something else -- position is *not* on new line!
                bResult = False
                bIsDone = True
        End Select
        lPos = lPos - 1
    Loop
    
    IsPositionOnNewLine = bResult
    
End Function


Private Sub LocateAnyBeginToken(ByRef Source As String, _
                                ByVal BeginPosition As Long, _
                                ByVal EndPosition As Long, _
                                ByRef outToken As CToken, _
                                ByRef outTokenPostion As Long)

    Dim oToken As CToken
    Dim lPos As Long
    Dim lInStr As Long
    Dim oItem As CToken
    Dim bMatchFound As Boolean
    
    bMatchFound = False
    
    For Each oToken In m_oDefinition.Tokens
    
        With oToken
            ' Make sure token is a beginner token
            If .TokenType = tktBeginner Then
                ' Search for token text
                lInStr = InStr32(BeginPosition, Source, .Text, IIf(.IsCaseSensitive, vbBinaryCompare, vbTextCompare))
                If lInStr > 0 Then
                    If .BeginsOnNewLine Then
                        If IsPositionOnNewLine(Source, lInStr) Then
                            ' If this is the first time -OR- this match is earliest...
                            If bMatchFound = False Or lInStr < lPos Then
                                lPos = lInStr
                                Set oItem = oToken
                                bMatchFound = True
                            End If
                        End If
                    End If
                End If
            End If
        End With

    Next

    If Not oItem Is Nothing Then
        Set outToken = oItem
        outTokenPostion = IIf(lPos = 0, -1, lPos)
    Else
        outTokenPostion = -1
    End If

End Sub


Private Sub LocateTokenOfStructure(ByRef Source As String, _
                                   ByRef Structure As CStructure, _
                                   ByVal BeginPosition As Long, _
                                   ByVal EndPosition As Long, _
                                   ByVal TokenType As eTokenType, _
                                   ByRef outToken As CToken, _
                                   ByRef outPosition As Long)

    Dim lPos As Long
    Dim lInStr As Long
    Dim oItem As CToken
    Dim oToken As CToken
    Dim bMatchFound As Boolean
    
    bMatchFound = False
                            
    For Each oToken In m_oDefinition.Tokens
        With oToken
            ' Make sure token is a beginner token
            If (.TokenType = TokenType) And (.Structure.ToString() = Structure.ToString()) Then
                ' Search for token text
                lInStr = InStr32(BeginPosition, Source, .Text, IIf(.IsCaseSensitive, vbBinaryCompare, vbTextCompare))
                If lInStr > 0 Then
                    ' If this is the first time -OR- this match is earliest...
                    If bMatchFound = False Or lInStr < lPos Then
                        lPos = lInStr
                        Set oItem = oToken
                        bMatchFound = True
                    End If
                End If
            End If
        End With
    Next

    If Not oItem Is Nothing Then
        Set outToken = oItem
        outPosition = IIf(lPos = 0, -1, lPos)
    Else
        outPosition = -1
    End If

End Sub



Private Sub LocateAnyBeginTokenReverse(ByRef Source As String, _
                                       ByRef outToken As CToken, _
                                       ByRef outTokenPostion As Long)

    Dim oToken As CToken
    Dim lPos As Long
    Dim lInStr As Long
    Dim oItem As CToken
    Dim bMatchFound As Boolean
    
    lPos = Len(Source)
    bMatchFound = False
    
    For Each oToken In m_oDefinition.Tokens
    
        With oToken
            ' Make sure token is a beginner token
            If .TokenType = tktBeginner Then
                ' Search for token text
                lInStr = InStrRev(Source, .Text, , IIf(.IsCaseSensitive, vbBinaryCompare, vbTextCompare))
                If lInStr > 0 Then
                    
                    
                    
                    If .BeginsOnNewLine Then
                        If IsPositionOnNewLine(Source, lInStr) Then
                            ' If this is the first time -OR- this match is earliest...
                            If bMatchFound = False Or lInStr > lPos Then
                                lPos = lInStr
                                Set oItem = oToken
                                bMatchFound = True
                            End If
                        End If
                    End If
                    
                    
                    '' If this is the first time -OR- this match is earliest...
                    'If bMatchFound = False Or lInStr > lPos Then
                    '    lPos = lInStr
                    '    Set oItem = oToken
                    '    bMatchFound = True
                    'End If
                
                End If
            End If
        End With

    Next

    If Not oItem Is Nothing Then
        Set outToken = oItem
        outTokenPostion = IIf(lPos = 0, -1, lPos)
    Else
        outTokenPostion = -1
    End If

End Sub



Private Sub LocateAnyKeyword(ByRef Source As String, _
                                ByVal BeginPosition As Long, _
                                ByVal EndPosition As Long, _
                                ByRef outKeyWord As CKeyWord, _
                                ByRef outKeyWordPosition As Long)

    Dim oPrim As CPrimitive
    Dim lPos As Long
    Dim lInStr As Long
    Dim oItem As CPrimitive
    Dim bMatchFound As Boolean
    
    bMatchFound = False
    
    For Each oPrim In m_oDefinition.Primitives
    
        With oPrim
            ' Search for token text
            lInStr = InStr32(BeginPosition, Source, .PrimitiveName, vbTextCompare)
            If lInStr > 0 Then
                ' If this is the first time -OR- this match is earliest...
                If bMatchFound = False Or lInStr < lPos Then
                    lPos = lInStr
                    Set oItem = oPrim
                    bMatchFound = True
                End If
            End If
        End With

    Next

    If Not oItem Is Nothing Then
        Dim oKeyWord As CKeyWord
        Set oKeyWord = New CKeyWord
        With oKeyWord
            .Text = Mid$(Source, lPos, Len(oItem.PrimitiveName))
            .Primitive = oItem
            .BeginPosition = lPos
        End With
        Set outKeyWord = oKeyWord
        outKeyWordPosition = IIf(lPos = 0, -1, lPos)
    Else
        outKeyWordPosition = -1
    End If

End Sub





Public Property Get Definition() As CDefinition
    Set Definition = m_oDefinition
End Property

Public Property Let Definition(ByRef NewValue As CDefinition)
    Set m_oDefinition = NewValue
End Property

Public Property Get Blocks() As CObjectCollection
    Set Blocks = m_oBlocks
End Property



Private Sub m_oBlocks_ItemAdded(Item As Object)
    If Not m_bIsParsing Then
        RaiseEvent BlockAdded(Item)
    End If
End Sub

Private Sub m_oBlocks_ItemDeleted(ByVal Index As Integer)
    RaiseEvent BlockDeleted(Index)
End Sub


Private Sub Class_Initialize()
    Set m_oBlocks = New CObjectCollection
    m_oBlocks.Parent = Me
    m_bIsLoading = False        ' Not loading by default
End Sub

