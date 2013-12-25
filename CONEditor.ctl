VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.UserControl CONEditor 
   BackStyle       =   0  'Transparent
   ClientHeight    =   4095
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6195
   KeyPreview      =   -1  'True
   PaletteMode     =   4  'None
   ScaleHeight     =   273
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   413
   Begin VB.Timer tmrFireSelChange 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   6240
      Top             =   4020
   End
   Begin RichTextLib.RichTextBox rtb 
      Height          =   4095
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   6195
      _ExtentX        =   10927
      _ExtentY        =   7223
      _Version        =   393217
      BorderStyle     =   0
      ScrollBars      =   3
      DisableNoScroll =   -1  'True
      Appearance      =   0
      RightMargin     =   10000
      TextRTF         =   $"CONEditor.ctx":0000
   End
   Begin VB.PictureBox picCorner 
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      Height          =   270
      Left            =   6000
      ScaleHeight     =   270
      ScaleWidth      =   270
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   3840
      Visible         =   0   'False
      Width           =   270
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "&Edit"
      Visible         =   0   'False
      Begin VB.Menu mnuEditUndo 
         Caption         =   "&Undo"
      End
      Begin VB.Menu mnuEditBar1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditCut 
         Caption         =   "Cu&t"
      End
      Begin VB.Menu mnuEditCopy 
         Caption         =   "&Copy"
      End
      Begin VB.Menu mnuEditPaste 
         Caption         =   "&Paste"
      End
      Begin VB.Menu mnuEditDelete 
         Caption         =   "&Delete"
      End
      Begin VB.Menu mnuEditBar2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditSelectAll 
         Caption         =   "Select &All"
      End
      Begin VB.Menu mnuEditBar3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditProperties 
         Caption         =   "&Edit"
      End
      Begin VB.Menu mnuEditPrimitive 
         Caption         =   "&Edit"
      End
   End
End
Attribute VB_Name = "CONEditor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' =============================================================================
' Module Name:      CONEditor
' Module Type:      Custom Control
' Description:      Text-editing control tailored to editing CON files. Based
'                   off Rich Text Box control.
' Author(s):        Jonathan A. Conway
' -----------------------------------------------------------------------------
' Log:
'
' 04 07 27 :
'   - Added SelIndent and SelOutdent methods
'   - Made TAB and SHIFT+TAB work when multiple lines are selected. Yay!!
'
' 04 07 26 :
'   - Made some preliminary efforts to get indentation working properly
'
' 04 07 24 :
'   - Added SelLine, SelLineText properties
'   - Added SelectLine, GetCharFromLine, GetLineFromChar methods
'   - These properties/methods were necessary to implement bookmarks
'      (see CONBookmarks, FrmMain modules)
'
' 04 07 18 :
'   - Fixed line-break-appendage problem (for now). See documentation for a
'     summary of the new functionality. Changes were made to the Save()
'     procedure and the GetEndWithoutLineBreaks() procedure was created.
'
' 04 03 24 :
'   - Created CONEditor
' =============================================================================

Option Explicit


' Enumerations
' ============

Public Enum eTextCase
    [TC_UPPERCASE] = 0
    [TC_LOWERCASE] = 1
End Enum


' Private Variables
' -----------------

Private m_sFilename As String           ' Current filename
Private m_bIsDirty As Boolean           ' Dirty flag
Private m_iUntitledID As Long

Private WithEvents m_oParser As CParser ' Reference to Parser object
Attribute m_oParser.VB_VarHelpID = -1

' Cache of Find details
Private m_sFindWhat As String           ' Search string
Private m_bFindWholeWord As Boolean     ' Whole word
Private m_bFindMatchCase As Boolean     ' Match case

Private m_bIsBusy As Boolean

Private m_bIndentOnKeyUp As Boolean
Private m_sSelTextCache As String

Private m_oBlock As CBlock


' Public Events
' -------------

Public Event FileLoadBegin()
Public Event FileLoadProgress()
Public Event FileLoadEnd()

Public Event DirtyStateChanged()
Public Event FilenameChanged()

Public Event SelChange(ByVal SelLine As Long, ByVal SelColumn As Integer)
Public Event SelBlockChange(ByRef Block As CBlock)
Public Event Change()

Public Event ParseProgress(ByVal CurrentPosition As Long, ByVal TotalLength As Long)

Public Event ObjectEdit(ByRef Block As CBlock)
Public Event KeyWordEdit(ByRef Primitive As CPrimitive, ByVal Text As String)



' Public Properties
' -----------------


Public Property Let SelCase(ByVal NewValue As eTextCase)

    Dim lSelStart As Long
    Dim lSelLength As Long
    
    With rtb
        lSelStart = .SelStart
        lSelLength = .SelLength
        
        Select Case NewValue
            Case eTextCase.TC_LOWERCASE
                .SelText = LCase(.SelText)
            Case eTextCase.TC_UPPERCASE
                .SelText = UCase(.SelText)
        End Select
        
        .SelStart = lSelStart
        .SelLength = lSelLength
    End With

End Property


Public Property Get UntitledID() As Long
    UntitledID = m_iUntitledID
End Property

Public Property Let UntitledID(ByVal NewValue As Long)
    m_iUntitledID = NewValue
End Property


Public Property Get Caption() As String
    If Trim$(m_sFilename) <> "" Then
        Caption = GetFilenameFromPath(m_sFilename)
    Else
        Caption = "[Untitled " & m_iUntitledID & "]"
    End If
End Property


Public Property Get Filename() As String
    Filename = m_sFilename
End Property


Public Property Get IsDirty() As Boolean
    IsDirty = m_bIsDirty
End Property

Public Property Let IsDirty(ByVal NewValue As Boolean)
    m_bIsDirty = NewValue
    RaiseEvent DirtyStateChanged
End Property


Public Property Get Enabled() As Boolean
    Enabled = rtb.Enabled
End Property

Public Property Let Enabled(ByVal NewValue As Boolean)
    rtb.Enabled = NewValue
End Property


Public Property Get SelText() As String
    ' [CONTENTS CHANGED]
    SelText = rtb.SelText
End Property

Public Property Let SelText(ByVal NewValue As String)
    ' [CONTENTS CHANGED]
    rtb.SelText = NewValue
End Property


Public Property Get SelStart() As Long
Attribute SelStart.VB_MemberFlags = "400"
    SelStart = rtb.SelStart
End Property

Public Property Let SelStart(ByVal NewValue As Long)
    rtb.SelStart = NewValue
End Property


Public Property Get SelLength() As Long
Attribute SelLength.VB_MemberFlags = "400"
    SelLength = rtb.SelLength
End Property

Public Property Let SelLength(ByVal NewValue As Long)
    rtb.SelLength = NewValue
End Property


Public Property Get Text() As String
    ' [CONTENTS CHANGED]
    Text = rtb.Text
End Property

Public Property Let Text(ByVal NewValue As String)
    ' [CONTENTS CHANGED]
    rtb.Text = NewValue
End Property


Public Property Get UndoType() As ERECUndoTypeConstants
Attribute UndoType.VB_MemberFlags = "400"
    UndoType = SendMessageLong(rtb.hWnd, EM_GETUNDONAME, 0, 0)
End Property


Public Property Get RedoType() As ERECUndoTypeConstants
Attribute RedoType.VB_MemberFlags = "400"
    RedoType = SendMessageLong(rtb.hWnd, EM_GETREDONAME, 0, 0)
End Property


Public Property Get CanPaste() As Boolean
Attribute CanPaste.VB_MemberFlags = "400"
    CanPaste = SendMessageLong(rtb.hWnd, EM_CANPASTE, 0, 0)
End Property


Public Property Get CanCopy() As Boolean
    If rtb.SelLength < 0 Then
        CanCopy = True
    End If
End Property


Public Property Get CanUndo() As Boolean
Attribute CanUndo.VB_MemberFlags = "400"
    CanUndo = True
    'CanUndo = SendMessageLong(rtb.hWnd, EM_CANUNDO, 0, 0)
End Property


Public Property Get CanRedo() As Boolean
Attribute CanRedo.VB_MemberFlags = "400"
    CanRedo = True
    'CanRedo = SendMessageLong(rtb.hWnd, EM_CANREDO, 0, 0)
End Property


Public Property Get Parser() As CParser
    Set Parser = m_oParser
End Property


Public Property Get FindWhat() As String
    FindWhat = m_sFindWhat
End Property

Public Property Let FindWhat(ByVal NewValue As String)
    m_sFindWhat = NewValue
End Property


Public Property Get FindMatchCase() As Boolean
    FindMatchCase = m_bFindMatchCase
End Property

Public Property Let FindMatchCase(ByVal NewValue As Boolean)
    m_bFindMatchCase = NewValue
End Property


Public Property Get FindWholeWord() As Boolean
    FindWholeWord = m_bFindWholeWord
End Property

Public Property Let FindWholeWord(ByVal NewValue As Boolean)
    m_bFindWholeWord = NewValue
End Property


Public Property Get SelLine() As Long
    SelLine = GetCurrentLine(rtb)
End Property

Public Property Let SelLine(ByVal NewValue As Long)
    ' *** TODO ***
    ' Find out a way of selecting a particular line in rtb control!!
    'rtb.SelStart = GetCharFromLine(rtb.GetLineFromChar(NewValue + 1) + 1) + 1
End Property


Public Property Let Locked(ByVal NewValue As Boolean)
    rtb.Locked = NewValue
End Property

Public Property Get SelLineText() As String

    ' Get contents of current line
    Dim sText As String
    Dim lBegin As Long, lEnd As Long
    
    lBegin = GetCharFromLine(rtb.GetLineFromChar(rtb.SelStart + 1))
    lEnd = InStr32(lBegin + 1, (rtb.Text & Chr(13)), Chr(13), vbBinaryCompare)
    
    sText = Mid$(rtb.Text & Chr(13), IIf(lBegin = 0, 1, lBegin), lEnd - lBegin)
    
    Do Until Mid$(sText, 1, 1) <> Chr(13) And Mid$(sText, 1, 1) <> Chr(10)
        sText = Mid$(sText, 2, Len(sText) - 1)
    Loop
    
    SelLineText = sText

End Property


Public Property Get SelColumn() As Long
    SelColumn = GetCurrentColumn(rtb)
End Property



' Public Methods
' --------------

Public Function FilterByObjects(ByRef Structs As CList) As String

    Dim oStruct As CListItem
    Dim oBlock As CBlock
    Dim sText As String
    Dim sHeader As String
    Dim lCount As Long
    Dim lTotal As Long
    
    sText = vbNewLine
    sHeader = "/* Filter Results for """ & m_sFilename & """" & vbNewLine
    
    For Each oStruct In Structs
        lCount = 0
        
        For Each oBlock In m_oParser.Blocks
            If oBlock.Structure.StructureName = oStruct.Label Then
                sText = sText & oBlock.Text & vbNewLine & vbNewLine
                lCount = lCount + 1
            End If
        Next oBlock
    
        sHeader = sHeader & "/* " & oStruct.Label & ": " & lCount & vbNewLine
        lTotal = lTotal + lCount
    Next oStruct

    sHeader = sHeader & "/* " & vbNewLine & "/* Total: " & lTotal & vbNewLine & "*/" & vbNewLine & vbNewLine
    sText = sHeader & sText

    FilterByObjects = sText

End Function


Public Sub LoadFile(ByVal Filename As String)
    
    Dim sText As String
    Dim lHnd As Long
    
    lHnd = FreeFile()
    
    ' [CONTENTS CHANGED]
    rtb.Enabled = False
    rtb.Locked = True
    
    Open Filename For Input As #lHnd
    sText = Input(LOF(lHnd), #lHnd)     ' Read file contents
    Close #lHnd
    
    rtb.Text = sText                    ' Load contents into rich text control
    
    RaiseEvent FileLoadBegin
    
    m_oParser.Parse sText               ' Send contents to parser
    
    ' Set variables, flags, etc.
    m_sFilename = Filename
    IsDirty = False
    
    rtb.Enabled = True
    rtb.Locked = False
    
    RaiseEvent FileLoadEnd
    
End Sub


Public Sub Save()

    If Trim$(m_sFilename) <> "" And IsDirty Then
        
        Dim lHnd As Long
        Dim sText As String
        
        lHnd = FreeFile()
        
        ' Try to remove trailing line breaks
        sText = rtb.Text
        sText = Mid$(sText, 1, GetEndWithoutLineBreaks(sText))
        
        Open Filename For Output As #lHnd
        Print #lHnd, sText          ' Write file contents
        Close #lHnd
        IsDirty = False             ' Reset dirty flag
    
    End If

End Sub


Public Sub SaveAs(ByVal NewFilename As String)

    m_sFilename = NewFilename
    RaiseEvent FilenameChanged
    Save

End Sub


Public Sub SaveCopyAs(ByVal NewFilename As String)
    
    If Trim$(m_sFilename) <> "" Then
        Save
        FileCopy m_sFilename, NewFilename
    End If

End Sub


Public Sub Undo()
    SendMessageLong rtb.hWnd, EM_UNDO, 0, 0
End Sub


Public Sub Redo()
    SendMessageLong rtb.hWnd, EM_REDO, 0, 0
End Sub


Public Sub Cut()
    SendMessageLong rtb.hWnd, WM_CUT, 0, 0
End Sub


Public Sub Copy()
    SendMessageLong rtb.hWnd, WM_COPY, 0, 0
End Sub


Public Sub Paste()
    SendMessageLong rtb.hWnd, WM_PASTE, 0, 0
    DoEvents
    SetFont Settings.ReadSetting("EditorFontName"), Settings.ReadSetting("EditorFontSize")
End Sub


Public Sub Clear()
    rtb.SelText = Empty
End Sub


Public Sub SelClear()
    rtb.SelText = ""
End Sub


Public Sub SelCut()
    SelCopy
    SelClear
End Sub


Public Sub SelCopy()
    Clipboard.SetText rtb.SelText, vbCFText
End Sub


Public Sub SelPaste()
    rtb.SelText = Clipboard.GetText(vbCFText)
End Sub


Public Sub SelectAll()
    rtb.SelStart = 0
    rtb.SelLength = Len(rtb.Text) + 2
End Sub


Public Function IsSearchWordSelected() As Boolean

    Dim bReturn As Boolean
    
    bReturn = False

    If m_bFindMatchCase Then
        If rtb.SelText = m_sFindWhat Then bReturn = True
    Else
        If LCase(rtb.SelText) = LCase(m_sFindWhat) Then bReturn = True
    End If

    IsSearchWordSelected = bReturn

End Function


Public Function FindNext() As Boolean
    
    Dim iSelStart As Long, bDone As Boolean
    
    iSelStart = rtb.SelStart
    If iSelStart = Len(rtb.Text) Then iSelStart = 0
    
    If IsSearchWordSelected() Then iSelStart = iSelStart + rtb.SelLength
    
    bDone = (rtb.Find(m_sFindWhat, iSelStart, Len(rtb.Text), _
            IIf(m_bFindWholeWord, rtfWholeWord, 0) _
            Or IIf(m_bFindMatchCase, rtfMatchCase, 0))) = -1
    
    FindNext = bDone
    
End Function


Public Function ReplaceAll(ByVal ReplaceText As String) As Long

    Dim iCount As Long

    rtb.SelStart = 0

    Do While FindNext() = False
        rtb.SelText = ReplaceText
        iCount = iCount + 1
    Loop

    ReplaceAll = iCount

End Function


Public Sub GetRangeFromBlock(ByRef Block As CBlock, ByRef outBeginPos As Long, ByRef outEndPos As Long)
    GetRangeFromSignature Block.BeginSignature, Block.EndSignature, True, outBeginPos, outEndPos
End Sub

Public Sub GetRangeFromSignature(ByVal inBeginSignature As String, ByVal inEndSignature As String, inEndSignatureInBody As Boolean, ByRef outBeginPos As Long, ByRef outEndPos As Long)
    
    ' Find beginning signature in code
    outBeginPos = InStr32(1, rtb.Text, inBeginSignature, vbTextCompare) - 1
    If outBeginPos = -1 Then
        outEndPos = -1
        Exit Sub
    End If
    
    ' Find ending signature in code
    outEndPos = InStr32(outBeginPos + 1, rtb.Text, inEndSignature, vbTextCompare) - 1
    If outEndPos = -1 Then outEndPos = Len(rtb.Text)
    outEndPos = outEndPos

    ' Take into account whether signature is part of body
    If Not inEndSignatureInBody Then _
        outEndPos = outEndPos '+ Len(inEndSignature)

End Sub


Public Sub UpdateBlock(ByRef OldBlock As CBlock, ByRef NewBlock As CBlock)

    ' Update block text within rtb
    
    Dim iBegin As Long, iEnd As Long

    GetRangeFromBlock OldBlock, iBegin, iEnd

    If iBegin < 0 Then Exit Sub
    With rtb
        .SelStart = iBegin
        .SelLength = Len(OldBlock.Text)
        '.SelText = NewBlock.Text
        InsertText NewBlock.Text, True
    End With

End Sub

Public Sub GotoBlock(ByRef Block As CBlock)

    ' Select object within rtb
    
    m_bIsBusy = True
    
    Dim iBegin As Long, iEnd As Long

    GetRangeFromBlock Block, iBegin, iEnd

    If iBegin < 0 Then Exit Sub
    rtb.SelStart = iBegin
    rtb.SelLength = Len(Block.BeginSignature)

    rtb.SetFocus

    m_bIsBusy = False

End Sub

Public Sub InsertText(ByVal Source As String, ByVal IsSelected As Boolean)
    With rtb
        Dim lSelStart As Long
        
        ' Insert the specified text
        .SelText = Source
        lSelStart = .SelStart - Len(Source)
    
        ' Re-parse the newly inserted code if necessary
        If Settings.ReadSetting("Code_DynamicParsing") = "yes" Then
            m_oParser.CleanUpBlocks .Text
            
            ' Parse the insert code incrementally
            m_oParser.Parse Source, True
            
            If Settings.ReadSetting("Code_AutoFormat") = "yes" Then
                FormatCode  ' Format & color the code
            End If
        End If
    
        ' Highlight text if specified to do so
        If IsSelected Then
            .SelStart = lSelStart
            .SelLength = Len(Source)
        End If
    End With
End Sub

Public Sub AppendText(ByVal Source As String, ByVal IsSelected As Boolean)
    rtb.SelStart = IIf(Len(rtb.Text) = 0, 0, Len(rtb.Text) - 1)
    InsertText Source, IsSelected
    
    
    'With rtb
    '    .SelStart = Len(rtb.Text)
    '    DoEvents
    '    .SelText = Source
    '    DoEvents
    '    If IsSelected Then
    '        .SelStart = .SelStart - Len(Source)
    '        .SelLength = Len(Source)
    '    End If
    'End With
End Sub

Public Sub SetFont(ByVal FontName As String, ByVal FontSize As String)
    Dim bOldDirty As Boolean
    bOldDirty = m_bIsDirty
    If Len(Trim$(FontName)) <> 0 Then rtb.Font.Name = FontName
    If Len(Trim$(FontSize)) <> 0 Then rtb.Font.Size = FontSize
    rtb.Font.Bold = False
    rtb.Font.Italic = False
    rtb.Font.Underline = False
    rtb.Font.Strikethrough = False
    
    Dim lSelStart As Long
    Dim lSelLength As Long
    
    lSelStart = rtb.SelStart
    lSelLength = rtb.SelLength
    
    rtb.SelStart = 0
    rtb.SelLength = Len(rtb.Text)
    rtb.SelColor = vbBlack
    
    rtb.SelStart = lSelStart
    rtb.SelLength = lSelLength
    
    'm_bIsDirty = bOldDirty
    IsDirty = bOldDirty
End Sub

Public Function GetCharFromLine(ByVal LineNumber As Long) As Long

    Dim iInStr As Long
    Dim iCount As Long
    
    If LineNumber = -1 Then Exit Function
    
    iCount = 0
    iInStr = 0
    
    Do Until iCount = LineNumber
        iInStr = InStr32(iInStr + 1, rtb.Text, Chr(13), vbBinaryCompare)
        iCount = iCount + 1
    Loop
    
    GetCharFromLine = iInStr

End Function


Public Sub SelectLine(ByVal LineNumber As Long)

    Dim lBegin As Long, lEnd As Long
    
    If LineNumber > 1 Then
        lBegin = GetCharFromLine(LineNumber - 1)
    Else
        lBegin = 0
    End If
    
    lEnd = InStr32(lBegin + 1, rtb.Text, Chr(13), vbBinaryCompare) '- lBegin
    If lEnd <= 0 Then
        lEnd = Len(rtb.Text)
    End If
    lEnd = lEnd - lBegin
    
    rtb.SelStart = IIf(LineNumber = 1, 0, lBegin + 1)
    rtb.SelLength = lEnd

End Sub

Public Function GetLineFromChar(ByVal Char As Long) As Long
    GetLineFromChar = rtb.GetLineFromChar(Char)
End Function

Public Sub SelOutdent()
    
    ' Take into account the possibility of this method being called from
    ' outside. If the IndentOnKeyUp flag is False then the method has been
    ' called externally, so the SelTextCache must be taken care of.
    If Not m_bIndentOnKeyUp Then
        m_sSelTextCache = rtb.SelText
    End If
    
    ' Remove all trailing line breaks, etc.
    Do While Right$(m_sSelTextCache, 1) = Chr(13) Or _
             Right$(m_sSelTextCache, 1) = Chr(10)
        m_sSelTextCache = Mid$(m_sSelTextCache, 1, Len(m_sSelTextCache) - 1)
    Loop
    
    ' Delete tabs...
    m_sSelTextCache = Replace$(m_sSelTextCache, Chr(10) & vbTab, Chr(10))
    If Left$(m_sSelTextCache, 1) = vbTab Then _
        m_sSelTextCache = Right$(m_sSelTextCache, Len(m_sSelTextCache) - 1)
    
    ' Insert updated text
    With rtb
        .Visible = False
        
        If m_bIndentOnKeyUp Then
            ' Only do this if called internally; this means that the
            ' original text was deleted so it must be restored at the right
            ' position.
            .SelStart = IIf(rtb.SelStart > 0, rtb.SelStart - 1, 0)
            .SelLength = 2
        End If
        
        .SelText = m_sSelTextCache
        
        ' Restore original selection
        .SelStart = rtb.SelStart - Len(m_sSelTextCache)
        .SelLength = Len(m_sSelTextCache)
        
        .Visible = True
        .SetFocus
    End With

End Sub

Public Sub SelIndent()

    ' Take into account the possibility of this method being called from
    ' outside. If the IndentOnKeyUp flag is False then the method has been
    ' called externally, so the SelTextCache must be taken care of.
    If Not m_bIndentOnKeyUp Then
        m_sSelTextCache = rtb.SelText
    End If

    ' Remove all trailing line breaks, etc.
    Do While Right$(m_sSelTextCache, 1) = Chr(13) Or _
             Right$(m_sSelTextCache, 1) = Chr(10)
        m_sSelTextCache = Mid$(m_sSelTextCache, 1, Len(m_sSelTextCache) - 1)
    Loop
    
    ' Insert tabs...
    m_sSelTextCache = vbTab & Replace$(m_sSelTextCache, Chr(10), Chr(10) & vbTab)
    
    ' Insert updated text
    With rtb
        .Enabled = False
        
        If m_bIndentOnKeyUp Then
            ' Only do this if called internally; this means that the
            ' original text was deleted so it must be restored at the right
            ' position.
            .SelStart = rtb.SelStart - 1
            .SelLength = 2
        End If
        
        .SelText = m_sSelTextCache
        
        ' Restore original selection
        .SelStart = rtb.SelStart - Len(m_sSelTextCache)
        .SelLength = Len(m_sSelTextCache)
        
        .Enabled = True
        .SetFocus
    End With

End Sub

Public Sub SelComment()

    Dim sText As String
    
    sText = rtb.SelText
    
    If GetLineCount(sText) > 1 Then
        
        Dim bIsDone As Boolean
        Dim sChr As String
        Dim sBuffer As String
        
        bIsDone = False
        Do Until bIsDone
            sChr = Right$(Trim$(sText), 1)
            If sChr = Chr(13) Or sChr = Chr(10) Then
                sText = Left$(sText, Len(sText) - 1)
                sBuffer = sBuffer & sChr
            Else
                bIsDone = True
            End If
        Loop
        
        sText = "//" & Replace$(sText, Chr(10), Chr(10) & "//") & sBuffer
    
        rtb.SelText = sText
        rtb.SelStart = rtb.SelStart - Len(sText)
        rtb.SelLength = Len(sText)
    
    Else
    
        Dim lSelStart As Long
        Dim lSelLength As Long
    
        lSelStart = rtb.SelStart
        lSelLength = rtb.SelLength
        sText = rtb.Text
        
        Dim lPos As Long
        If rtb.GetLineFromChar(rtb.SelStart + 1) = 0 Then
            lPos = 0
        Else
            lPos = GetCharFromLine(rtb.GetLineFromChar(rtb.SelStart + 1)) + 1
        End If
        lPos = lPos + 1
        sText = Mid$(sText, 1, lPos - 1) & "//" & Mid$(sText, lPos, Len(sText) - (lPos - 1))
        
        rtb.Text = sText
        rtb.SelStart = lSelStart + 2
        rtb.SelLength = lSelLength
    
    End If

End Sub

Public Sub SelUncomment()

    Dim sText As String

    sText = rtb.SelText
    
    If GetLineCount(sText) > 1 Then
        
        Dim bIsDone As Boolean
        Dim sChr As String
        Dim sBuffer As String
        
        bIsDone = False
        Do Until bIsDone
            sChr = Right$(Trim$(sText), 1)
            If sChr = Chr(13) Or sChr = Chr(10) Then
                sText = Left$(sText, Len(sText) - 1)
                sBuffer = sBuffer & sChr
            Else
                bIsDone = True
            End If
        Loop
        
        sText = Replace$(sText, Chr(10) & "//", Chr(10)) & sBuffer
        If Left$(sText, 2) = "//" Then
            sText = Mid$(sText, 3, Len(sText) - 2)
        End If
    
    Else
    
    End If

    rtb.SelText = sText
    rtb.SelStart = rtb.SelStart - Len(sText)
    rtb.SelLength = Len(sText)

End Sub




' Private Methods
' ---------------

Public Function ClearIndents(ByRef Source As String) As String

    Dim sText As String
    
    sText = Source
    
    Do While InStr32(1, sText, Chr(13) & vbTab, vbBinaryCompare) <> 0
        sText = Replace$(sText, Chr(13) & vbTab, Chr(13), , , vbBinaryCompare)
    Loop
    Do While InStr32(1, sText, Chr(10) & vbTab, vbBinaryCompare) <> 0
        sText = Replace$(sText, Chr(10) & vbTab, Chr(10), , , vbBinaryCompare)
    Loop
    
    
    ClearIndents = sText

End Function


Public Sub DeleteIndents()

    Dim sText As String
    
    sText = rtb.Text
    
    sText = ClearIndents(sText)         ' Remove all indentation beforehand
    
    rtb.Text = sText                    ' Return text to richtextbox control

End Sub

Public Sub BulkIndent()

    Dim sText As String
    
    sText = rtb.Text
    
    sText = ClearIndents(sText)         ' Remove all indentation beforehand
    sText = BulkIndentSection(sText)    ' Do a bulk auto-indent of text
    
    rtb.Text = sText                    ' Return text to richtextbox control

End Sub


Public Function BulkIndentSection(ByVal Source As String) As String
    
    ' Perform a bulk indentation on the selected text.
    ' (This procedure is recursive)
    
    ' For example, given this source text:
    
    ' struct MyStructure
    ' {
    ' some code
    ' }
    
    ' ... The code is indented like so:
    
    ' struct MyStructure
    ' {
    '   some code
    ' }
    
    Dim lBegin As Long
    Dim lEnd As Long
    Dim bIsDone As Boolean
    Dim lPos As Long
    Dim sSource As String
    Dim sValue As String
    
    Dim sOpener As String
    Dim sCloser As String
    
    sSource = Source
    
    sOpener = "{"
    sCloser = "}"
    
    lPos = 1
    bIsDone = False
    
    Do Until bIsDone
        lBegin = InStr32(lPos, sSource, sOpener, vbBinaryCompare)
        If lBegin > 0 Then
            lEnd = InStr32(lBegin, sSource, sCloser, vbBinaryCompare)
            If lEnd > 0 Then
                sValue = Trim$(Mid$(sSource, lBegin + 3, (lEnd - lBegin) - 3))
                sValue = Indent(sValue)
                sValue = BulkIndentSection(sValue)
                sValue = sOpener & vbNewLine & sValue & vbNewLine & sCloser
                sSource = Mid$(sSource, 1, lBegin - 1) & sValue & Mid$(sSource, lEnd + 1, (Len(sSource) - lEnd) + 1)
                lPos = lEnd
            
            Else
                lPos = lBegin + 1
            End If
        Else
            bIsDone = True
        End If
    Loop
    
    BulkIndentSection = sSource

End Function

Private Function TrimLineBreaks(ByRef Source As String, ByVal DirectionLeft As Boolean) As String

    ' Remove all trailing line breaks, etc.

    Dim sSource As String
    
    sSource = Source
    
    If DirectionLeft Then
        Do While Left$(sSource, 1) = Chr(13) Or _
                 Left$(sSource, 1) = Chr(10)
            sSource = Mid$(sSource, 2, Len(sSource) - 1)
        Loop
    Else
        Do While Right$(sSource, 1) = Chr(13) Or _
                 Right$(sSource, 1) = Chr(10)
            sSource = Mid$(sSource, 1, Len(sSource) - 1)
        Loop
    End If
    
    TrimLineBreaks = sSource

End Function

Private Function Indent(ByRef Source As String) As String

    Dim sSource As String
    sSource = Source
    
    sSource = TrimLineBreaks(sSource, False)
        
    ' Insert tabs...
    sSource = vbTab & Replace$(sSource, Chr(10), Chr(10) & vbTab)
    
    Indent = sSource

End Function

Private Sub AutoIndent()

    ' Carries out AutoIndentation, assuming that it is called just after the
    ' user pressed [ENTER] (ch. 13) on their keyboard.
    
    ' 1. Sustains indentation:
    '    If user presses [ENTER] on a line that is indented, the next line is
    '    indented the same amount as the first.
    
    ' 2. AutoIndent on block opening:
    '    If the user presses [ENTER] directly after having entered a block
    '    opener (in this case, the "{" symbol, the indent of the subsequent
    '    line is increased by one.

    On Error GoTo ProcedureError

    Dim lLastLinePos As Long
    Dim iTabCount As Integer
    Dim sLastLine As String
    Dim i As Integer
    
    ' Get contents of previous line
    lLastLinePos = GetCharFromLine(rtb.GetLineFromChar(rtb.SelStart + 1) - 1)
    If lLastLinePos = 0 Then
        lLastLinePos = 1
        sLastLine = Mid$(rtb.Text, lLastLinePos, rtb.SelStart - 2)
    Else
        lLastLinePos = lLastLinePos + 2
        sLastLine = Mid$(rtb.Text, lLastLinePos, rtb.SelStart - 3)
    End If
    
    ' Enumerate total number of indentation tabs
    For i = 1 To Len(sLastLine)
        If Mid$(sLastLine, i, 1) = vbTab Then
            iTabCount = iTabCount + 1
        Else
            Exit For
        End If
    Next
    
    ' Was the previous character a block opener ("{")?
    If Asc(Mid$(rtb.Text, IIf((rtb.SelStart - 2) <= 0, 1, rtb.SelStart - 2), 1)) = 123 Then
        iTabCount = iTabCount + 1       ' Increase indent by one
    End If

    ' Repeat tabs on next line for each indentation level
    rtb.SelText = String(iTabCount, vbTab)

    Exit Sub
    
ProcedureError:
    If err.Number = 5 Or err.Number = 6 Then
    Else
        MsgBox err.Number & vbNewLine & err.Description
    End If

End Sub


Private Sub AutoKeyWordColor()

    If Settings.ReadSetting("Code_DynamicParsing") <> "yes" Then Exit Sub
    If Settings.ReadSetting("Code_AutoFormat") <> "yes" Then Exit Sub
    If Settings.ReadSetting("AutoFormat_EnableKeywords") <> "yes" Then Exit Sub

    ' 1. Is the selected word a KeyWord?
    ' 2. If yes, get absolute position of that keyword
    ' 3. Format keyword
    
    Dim lBeforePos As Long
    Dim sBeforeText As String
    
    lBeforePos = rtb.SelStart - 32755
    If lBeforePos < 0 Then lBeforePos = 0
    sBeforeText = Mid$(rtb.Text, lBeforePos + 1, ((rtb.SelStart + 1) - lBeforePos) - 1)
    
    Dim oPrim As CPrimitive
    For Each oPrim In m_oParser.Definition.Primitives
    
        If UCase(Right$(sBeforeText, Len(oPrim.PrimitiveName))) _
                = UCase(oPrim.PrimitiveName) Then
            
            ' Match found
            With rtb
                Dim lColor As Long
                Dim sFontName As String
                Dim iFontSize As Integer
                Dim bBold As Boolean
                Dim bItalic As Boolean
                Dim bUnderline As Boolean
                Dim bStrikethrough As Boolean
                
                ' Save current font
                lColor = .SelColor
                sFontName = .SelFontName
                iFontSize = .SelFontSize
                bBold = .SelBold
                bItalic = .SelItalic
                bUnderline = .SelUnderline
                bStrikethrough = .SelStrikeThru
                
                .SelStart = .SelStart - Len(oPrim.PrimitiveName())
                .SelLength = Len(oPrim.PrimitiveName())
                
                ' Set all formatting attributes for this selection
                .SelColor = FixLong(GetTagAttribute(Settings.ReadSetting("AutoFormat_[Keywords]"), "fontcolor"))
                .SelFontName = GetTagAttribute(Settings.ReadSetting("AutoFormat_[Keywords]"), "fontname")
                .SelFontSize = FixInteger(GetTagAttribute(Settings.ReadSetting("AutoFormat_[Keywords]"), "fontsize"))
                .SelBold = (GetTagAttribute(Settings.ReadSetting("AutoFormat_[Keywords]"), "fontbold") = "yes")
                .SelItalic = (GetTagAttribute(Settings.ReadSetting("AutoFormat_[Keywords]"), "fontitalic") = "yes")
                .SelUnderline = (GetTagAttribute(Settings.ReadSetting("AutoFormat_[Keywords]"), "fontunderline") = "yes")
                .SelStrikeThru = (GetTagAttribute(Settings.ReadSetting("AutoFormat_[Keywords]"), "fontstrikethru") = "yes")
            
                .SelStart = .SelStart + Len(oPrim.PrimitiveName)
                
                ' Restore previous font
                .SelColor = lColor
                .SelFontName = sFontName
                .SelFontSize = iFontSize
                .SelBold = bBold
                .SelItalic = bItalic
                .SelUnderline = bUnderline
                .SelStrikeThru = bStrikethrough
                
            End With
            
        End If
    
    Next
    
End Sub

Private Function GetLineText(ByVal LineNumber As Long) As String

    ' Get contents at specified line number
    
    Dim sText As String
    Dim lBegin As Long, lEnd As Long
    
    lBegin = GetCharFromLine(rtb.GetLineFromChar(GetCharFromLine(LineNumber)))
    lEnd = InStr32(lBegin + 1, (rtb.Text & Chr(13)), Chr(13), vbBinaryCompare)
    
    sText = Mid$(rtb.Text & Chr(13), IIf(lBegin = 0, 1, lBegin), lEnd - lBegin)
    
    Do Until Mid$(sText, 1, 1) <> Chr(13) And Mid$(sText, 1, 1) <> Chr(10)
        sText = Mid$(sText, 2, Len(sText) - 1)
    Loop
    
    GetLineText = sText

End Function

Private Sub AutoOutdent()

    ' Carries out AutoOutdentation, assuming that it is called just after the
    ' user pressed the "}" key (closing block) on their keyboard.
    
    ' On pressing of the "}" key, if:
    '   1. There are no other characters on the line apart from whitespace
    '   2. There is at least *one* tab directly before the "}" character
    '
    ' ... then ...
    '
    ' Outdent the block by one and place the cursor at the end of the line

    On Error GoTo ProcedureError

    Dim lLastLinePos As Long
    Dim iTabCount As Integer
    Dim sLastLine As String
    Dim i As Integer
    
    '' Get contents of previous line
    'lLastLinePos = GetCharFromLine(rtb.GetLineFromChar(rtb.SelStart + 1)) '- 1)
    'If lLastLinePos = 0 Then
    '    lLastLinePos = 1
    '    sLastLine = Mid$(rtb.Text, lLastLinePos, rtb.SelStart - 2)
    'Else
    '    lLastLinePos = lLastLinePos + 2
    '    sLastLine = Mid$(rtb.Text, lLastLinePos, rtb.SelStart - 3)
    'End If
    
    sLastLine = GetLineText(SelLine)
    
    'SelLineText
    
    'MsgBox sLastLine
    'Stop
    
    If IsLineWhitespace(Replace(sLastLine, "}", "")) Then
    
        ' Enumerate total number of indentation tabs
        For i = 1 To Len(sLastLine)
            If Mid$(sLastLine, i, 1) = vbTab Then
                iTabCount = iTabCount + 1
            Else
                Exit For
            End If
        Next
        
        If iTabCount > 0 Then
            Dim sText As String
            Dim lSelStart As Long
            
            lSelStart = rtb.SelStart
            
            sText = rtb.Text
            Mid$(sText, rtb.SelStart - 1, 2) = "} " '& Chr(13)
            rtb.Text = sText
            
            rtb.SelStart = lSelStart
        End If
        
    End If

    Exit Sub
    
ProcedureError:
    If err.Number = 5 Then
    Else
        MsgBox err.Number & vbNewLine & err.Description
    End If

End Sub

Public Sub ClearColoring()
    
    ' Clear all formatting by setting everything to a single font
    Me.SetFont Settings.ReadSetting("EditorFontName"), Settings.ReadSetting("EditorFontSize")

End Sub

Public Sub FormatCode()
    
    BulkColorBlocks
    
    If Settings.ReadSetting("AutoFormat_EnableKeywords") = "yes" Then
        BulkColorKeywords
    End If

End Sub

Private Sub ColorKeywords(ByRef Block As CBlock)

    m_bIsBusy = True
    
    ' Get position of block within source
    Dim lBlockPos As Long, iEnd As Long
    GetRangeFromBlock Block, lBlockPos, iEnd
    
    If lBlockPos = -1 Then
        m_bIsBusy = False
        Exit Sub
    End If
    
    ' Process & format keywords
    Dim oKeyWords As CObjectCollection
    Dim oKeyWord As CKeyWord

    Set oKeyWords = m_oParser.ParseKeyWords(Block.Text)
    
    For Each oKeyWord In oKeyWords
        With rtb
            .SelStart = (lBlockPos + (oKeyWord.BeginPosition - 1)) - 2
            .SelLength = Len(oKeyWord.Text)
            
            ' Set all formatting attributes for this selection
            .SelColor = FixLong(GetTagAttribute(Settings.ReadSetting("AutoFormat_[Keywords]"), "fontcolor"))
            .SelFontName = GetTagAttribute(Settings.ReadSetting("AutoFormat_[Keywords]"), "fontname")
            .SelFontSize = FixInteger(GetTagAttribute(Settings.ReadSetting("AutoFormat_[Keywords]"), "fontsize"))
            .SelBold = (GetTagAttribute(Settings.ReadSetting("AutoFormat_[Keywords]"), "fontbold") = "yes")
            .SelItalic = (GetTagAttribute(Settings.ReadSetting("AutoFormat_[Keywords]"), "fontitalic") = "yes")
            .SelUnderline = (GetTagAttribute(Settings.ReadSetting("AutoFormat_[Keywords]"), "fontunderline") = "yes")
            .SelStrikeThru = (GetTagAttribute(Settings.ReadSetting("AutoFormat_[Keywords]"), "fontstrikethru") = "yes")
        End With
    Next
    
    m_bIsBusy = False

End Sub

Private Sub BulkColorKeywords()

    ' Method:
    
    ' For each block:
    '   1. Get block's starting position
    '   2. Find tokens
    '       a. Select token
    '       b. Format selection

    Dim oBlock As CBlock
    For Each oBlock In m_oParser.Blocks
        ColorKeywords oBlock
    Next

End Sub

Private Sub BulkColorBlocks()

    Dim oBlock As CBlock
    For Each oBlock In m_oParser.Blocks
        ColorBlock oBlock
    Next

End Sub

Private Sub ColorBlock(ByRef Block As CBlock)
    
    Dim lSelStart As Long
    Dim lSelLength As Long
    Dim lInStr As Long
    
    lSelStart = rtb.SelStart
    lSelLength = rtb.SelLength
    
    lInStr = InStr32(1, rtb.Text, Block.BeginSignature, vbBinaryCompare)
    If lInStr > 0 Then
        With rtb
            .SelStart = lInStr - 1
            .SelLength = Len(Block.Text)
            
            ' Set all formatting attributes for this selection
            .SelColor = FixLong(GetTagAttribute(Settings.ReadSetting("AutoFormat_" & Block.Structure.StructureName), "fontcolor"))
            .SelFontName = GetTagAttribute(Settings.ReadSetting("AutoFormat_" & Block.Structure.StructureName), "fontname")
            .SelFontSize = FixInteger(GetTagAttribute(Settings.ReadSetting("AutoFormat_" & Block.Structure.StructureName), "fontsize"))
            .SelBold = (GetTagAttribute(Settings.ReadSetting("AutoFormat_" & Block.Structure.StructureName), "fontbold") = "yes")
            .SelItalic = (GetTagAttribute(Settings.ReadSetting("AutoFormat_" & Block.Structure.StructureName), "fontitalic") = "yes")
            .SelUnderline = (GetTagAttribute(Settings.ReadSetting("AutoFormat_" & Block.Structure.StructureName), "fontunderline") = "yes")
            .SelStrikeThru = (GetTagAttribute(Settings.ReadSetting("AutoFormat_" & Block.Structure.StructureName), "fontstrikethru") = "yes")
        End With
    End If

    rtb.SelStart = lSelStart
    rtb.SelLength = lSelLength

End Sub



' Event Handlers
' --------------

Private Sub m_oParser_BlockAdded(ByRef Block As CBlock)
    If Not m_oParser.IsParsing Then
        If Not m_oParser.IsLoading Then
            AppendText Block.Text, True
            'If Settings.ReadSetting("Code_AutoFormat") = "yes" Then
            '    ColorBlock Block
            'End If
        End If
    End If
End Sub

Private Sub m_oParser_ParseComplete()
    If Settings.ReadSetting("Code_AutoFormat") = "yes" Then
        FormatCode
    End If
End Sub

Private Sub m_oParser_ParseProgress(ByVal CurrentPosition As Long, ByVal TotalLength As Long)
    RaiseEvent ParseProgress(CurrentPosition, TotalLength)
End Sub

Private Sub mnuEditCopy_Click()
    Me.Copy
End Sub

Private Sub mnuEditCut_Click()
    Me.Cut
End Sub

Private Sub mnuEditDelete_Click()
    Me.Clear
End Sub

Private Sub mnuEditPaste_Click()
    Me.Paste
End Sub

Private Sub mnuEditPrimitive_Click()
    SelectLine SelLine
    RaiseEvent KeyWordEdit(GetSelPrimitive(), TrimWhiteSpace(SelLineText))
End Sub

Private Sub mnuEditProperties_Click()
    RaiseEvent ObjectEdit(m_oBlock)
End Sub

Private Sub mnuEditSelectAll_Click()
    Me.SelectAll
End Sub

Private Sub picCorner_GotFocus()
    rtb.SetFocus
End Sub


Private Sub rtb_Change()
    If rtb.Enabled Then
        If Not rtb.Locked Then
            If Not m_bIsBusy Then
                If Not IsDirty Then IsDirty = True
                RaiseEvent Change
            End If
        End If
    End If
End Sub


Public Sub rtb_KeyDown(KeyCode As Integer, Shift As Integer)
    
    If KeyCode = 9 Then
        If GetLineCount(rtb.SelText) > 1 Then
            m_bIndentOnKeyUp = True
            m_sSelTextCache = rtb.SelText
        End If
    End If

End Sub

Private Sub rtb_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = 9 Then         ' [TAB] key pressed
        ' Do multi-line indent
        If m_bIndentOnKeyUp Then
            If Shift = 0 Then
                SelIndent
            Else
                SelOutdent
            End If
            m_bIndentOnKeyUp = False
        End If
    
    ElseIf KeyCode = 13 Then    ' [ENTER] key pressed
        ' Auto-indentation
        DoEvents
        If Settings.ReadSetting("Code_AutoIndent") = "yes" Then
            If Len(rtb.Text) > 2 Then
                AutoIndent
            End If
        End If
    
    ElseIf KeyCode = 221 Then   ' "}" key pressed
        ' Auto-outdentation
        DoEvents
        If Settings.ReadSetting("Code_AutoIndent") = "yes" Then
            If Len(rtb.Text) > 2 Then
                AutoOutdent
            End If
        End If
    End If

    ' Take care of keyword colouring
    AutoKeyWordColor

End Sub

Private Sub UpdateEditBlockMenu()

    Set m_oBlock = m_oParser.GetObjectAtPos(rtb.Text, rtb.SelStart + 1)
    mnuEditProperties.Visible = Not (m_oBlock Is Nothing)
    If mnuEditProperties.Visible Then mnuEditBar3.Visible = True
    If Not (m_oBlock Is Nothing) Then
        mnuEditProperties.Caption = "&Edit Block " & m_oBlock.ToString() & "..."
    End If

End Sub

Private Function IsTextWhiteSpace(ByRef Source As String) As Boolean

    Dim sText As String
    sText = Source
    
    sText = Trim$(sText)
    sText = Replace$(sText, vbTab, " ")
    sText = Replace$(sText, Chr(10), " ")
    sText = Replace$(sText, Chr(13), " ")
    
    IsTextWhiteSpace = (Len(sText) = 0)

End Function

Private Function GetSelPrimitive() As CPrimitive

    Dim sLineText As String
    Dim oKeyWord As CKeyWord
    Dim oKeyWords As CObjectCollection
    
    sLineText = SelLineText

    Set oKeyWords = m_oParser.ParseKeyWords(sLineText)
    
    
    ' Remove non-stand-alone primitives, as these can't be edited
    ' with the Primitives dialog. (They use the Structure dialog instead.)
    
    ' Also remove primitives with no parameters
    
    For Each oKeyWord In oKeyWords
        If (Not oKeyWord.Primitive.IsStandAlone) _
            Or (oKeyWord.Primitive.Parameters.Count = 0) Then
            oKeyWords.Delete oKeyWord.Index
        End If
    Next

    If oKeyWords.Count > 0 Then
        Set GetSelPrimitive = oKeyWords(1).Primitive
    End If

End Function

Private Sub UpdateEditPrimitiveMenu()

    Dim oPrim As CPrimitive
    Set oPrim = GetSelPrimitive()
    
    mnuEditPrimitive.Visible = False
    
    If Not (oPrim Is Nothing) Then
        ' Enable menu
        mnuEditPrimitive.Caption = "Edit P&rimitive " & oPrim.PrimitiveName & "..."
        mnuEditPrimitive.Visible = True
        mnuEditBar3.Visible = True
    End If
    
End Sub


Private Sub rtb_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)

    If Button = vbRightButton Then
        
        DoEvents
        
        mnuEditBar3.Visible = False
    
        UpdateEditBlockMenu
        UpdateEditPrimitiveMenu
        
        PopupMenu mnuEdit
        
    End If

End Sub

Private Sub rtb_SelChange()
    
    If m_bIsBusy = False Then
        
        RaiseEvent SelChange(SelLine, SelColumn)
        
        If Settings.ReadSetting("Code_DynamicParsing") = "yes" Then
        
            Set m_oBlock = m_oParser.GetObjectAtPos(rtb.Text, rtb.SelStart + 1)
            
            RaiseEvent SelBlockChange(m_oBlock)
            
            If rtb.SelLength = 0 Then
                If Settings.ReadSetting("Code_AutoFormat") = "yes" Then
                    If Not (m_oBlock Is Nothing) Then
                        With rtb
                            ' Set all formatting attributes for this selection
                            .SelColor = FixLong(GetTagAttribute(Settings.ReadSetting("AutoFormat_" & m_oBlock.Structure.StructureName), "fontcolor"))
                            .SelFontName = GetTagAttribute(Settings.ReadSetting("AutoFormat_" & m_oBlock.Structure.StructureName), "fontname")
                            .SelFontSize = FixInteger(GetTagAttribute(Settings.ReadSetting("AutoFormat_" & m_oBlock.Structure.StructureName), "fontsize"))
                            .SelBold = (GetTagAttribute(Settings.ReadSetting("AutoFormat_" & m_oBlock.Structure.StructureName), "fontbold") = "yes")
                            .SelItalic = (GetTagAttribute(Settings.ReadSetting("AutoFormat_" & m_oBlock.Structure.StructureName), "fontitalic") = "yes")
                            .SelUnderline = (GetTagAttribute(Settings.ReadSetting("AutoFormat_" & m_oBlock.Structure.StructureName), "fontunderline") = "yes")
                            .SelStrikeThru = (GetTagAttribute(Settings.ReadSetting("AutoFormat_" & m_oBlock.Structure.StructureName), "fontstrikethru") = "yes")
                        End With
                    Else
                        rtb.SelColor = vbBlack
                        With rtb
                            ' Remove all formatting attributes for this selection
                            .SelColor = vbBlack
                            .SelFontName = Settings.ReadSetting("EditorFontName")
                            .SelFontSize = Settings.ReadSetting("EditorFontSize")
                            .SelBold = False
                            .SelItalic = False
                            .SelUnderline = False
                            .SelStrikeThru = False
                        End With
                    
                    End If
                End If
            End If
        
        End If

    End If
End Sub


Private Sub UserControl_GotFocus()
    On Error GoTo ProcedureError
    
    rtb.SetFocus
    
    Exit Sub
ProcedureError:
    If err.Number = 5 Then
        Debug.Print err.Number & vbNewLine & err.Description
    Else
        MsgBox err.Number & vbNewLine & err.Description
    End If
End Sub

Private Sub UserControl_Resize()
    
    On Error Resume Next
    
    Dim iW As Integer, iH As Integer
    
    iW = UserControl.ScaleWidth
    iH = UserControl.ScaleHeight
    
    rtb.Width = iW - rtb.Left
    rtb.Height = iH
    
    picCorner.Left = iW - 17
    picCorner.Top = iH - 17

End Sub


Private Sub UserControl_Initialize()
    
    ' For revealing multiple-undo/redo functionality in RichTextBox control
    Dim lStyle As Long
    lStyle = TM_RICHTEXT Or TM_MULTILEVELUNDO Or TM_MULTICODEPAGE
    SendMessageLong rtb.hWnd, EM_SETTEXTMODE, lStyle, 0
    
    ' Set undo limit to 100, which didn't work!! AARGHHH!!!
    'SendMessageLong rtb.hWnd, EM_SETUNDOLIMIT, 100, 0
    
    ' Set font/size values
    'SetFont Settings.ReadSetting("EditorFontName"), Settings.ReadSetting("EditorFontSize")

    m_bIsBusy = True
    
    rtb.SelIndent = 10

    ' Initialize a few flags...
    m_bIsBusy = False
    m_bIndentOnKeyUp = False
    
    Set m_oParser = New CParser

End Sub

Private Sub UserControl_Terminate()
    ' Garbage collection
    Set m_oParser = Nothing
End Sub

