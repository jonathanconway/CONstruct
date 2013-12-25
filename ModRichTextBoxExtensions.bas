Attribute VB_Name = "ModRichTextBoxExtensions"
' =============================================================================
' Module Name:      ModRichTextBoxExtensions
' Module Type:      Code Module
' Description:      Give richtextbox control Undo/Redo functionality using
'                   the SendMessage() API call.
' Author(s):        [Original author unknown]
'                   Modified by Jonathan A. Conway
'                   ElementK Journals -- for current line/column stuff
' -----------------------------------------------------------------------------
' Log:
'                   ?? ?? ?? : Got the code somewhere on the web and modified
'                              it to suit my purposes
'
' =============================================================================


Option Explicit


'// View Types
Public Enum ERECViewModes
    ercDefault = 0
    ercWordWrap = 1
    ercWYSIWYG = 2
End Enum
'// Undo Types
Public Enum ERECUndoTypeConstants
    ercUID_UNKNOWN = 0
    ercUID_TYPING = 1
    ercUID_DELETE = 2
    ercUID_DRAGDROP = 3
    ercUID_CUT = 4
    ercUID_PASTE = 5
End Enum
'// Text Modes
Public Enum TextMode
    TM_PLAINTEXT = 1
    TM_RICHTEXT = 2 ' /* default behavior */
    TM_SINGLELEVELUNDO = 4
    TM_MULTILEVELUNDO = 8 ' /* default behavior */
    TM_SINGLECODEPAGE = 16
    TM_MULTICODEPAGE = 32 ' /* default behavior */
End Enum

Public Const WM_COPY = &H301
Public Const WM_CUT = &H300
Public Const WM_PASTE = &H302

Public Const WM_USER = &H400
Public Const EM_SETTEXTMODE = (WM_USER + 89)
Public Const EM_UNDO = &HC7
Public Const EM_REDO = (WM_USER + 84)
Public Const EM_CANPASTE = (WM_USER + 50)
Public Const EM_CANUNDO = &HC6&
Public Const EM_CANREDO = (WM_USER + 85)
Public Const EM_GETUNDONAME = (WM_USER + 86)
Public Const EM_GETREDONAME = (WM_USER + 87)
Public Const EM_SETUNDOLIMIT = (WM_USER + 82)


Public Declare Function SendMessageLong Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Any) As Long


Private Declare Function SendMessageByNum Lib "user32" _
    Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, _
    ByVal wParam As Long, ByVal lParam As Long) As Long


Private Const EM_LINEFROMCHAR = &HC9
Private Const EM_LINEINDEX = &HBB

Public Function GetCurrentLine(TxtBox As Object) As Long
With TxtBox
GetCurrentLine = SendMessageByNum(.hWnd, EM_LINEFROMCHAR, _
CLng(.SelStart), 0&) + 1
End With
End Function
Public Function GetCurrentColumn(TxtBox As Object) As Long
With TxtBox
GetCurrentColumn = .SelStart - SendMessageByNum(.hWnd, _
EM_LINEINDEX, -1&, 0&) + 1
End With
End Function




'// Returns the undo/redo type
Public Function TranslateUndoType(ByVal eType As ERECUndoTypeConstants) As String
    Select Case eType
        Case ercUID_UNKNOWN
            TranslateUndoType = "Last Action"
        Case ercUID_TYPING
            TranslateUndoType = "Typing"
        Case ercUID_PASTE
            TranslateUndoType = "Paste"
        Case ercUID_DRAGDROP
            TranslateUndoType = "Drag Drop"
        Case ercUID_DELETE
            TranslateUndoType = "Delete"
        Case ercUID_CUT
            TranslateUndoType = "Cut"
    End Select
End Function
