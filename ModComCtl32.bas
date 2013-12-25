Attribute VB_Name = "ModComCtl32"
' =============================================================================
' Module Name:      ModComCtl32
' Module Type:      Code Module
' Description:      Code for implementing CommonControls32 (for Windows XP
'                   visual styles support)
' Author(s):        Jonathan A. Conway
' -----------------------------------------------------------------------------
' Log:
'
' 04 07 16 :
'   - Combined all code related to visual styles into ModComCtl32
'
' ?? ?? ?? :
'   - Created ModComCtl32
' =============================================================================


Option Explicit


' API Declares
' ============

Public Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" _
            (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, _
            ByVal lpsz2 As String) As Long
Private Declare Function InitCommonControlsEx Lib "comctl32.dll" _
   (iccex As tagInitCommonControlsEx) As Boolean
Private Const ICC_USEREX_CLASSES = &H200

' Constants
' =========

' Toolbar constants
Public Const WM_USER = &H400
Public Const TBSTYLE_FLAT As Long = &H800
Public Const TB_SETSTYLE = WM_USER + 56
Public Const TB_GETSTYLE = WM_USER + 57
Public Const TB_SETIMAGELIST = WM_USER + 48
Public Const TB_SETHOTIMAGELIST = WM_USER + 52
Public Const TB_SETDISABLEDIMAGELIST = WM_USER + 54


' Type Definitions
' ================

' For initializing common controls
Private Type tagInitCommonControlsEx
   lngSize As Long
   lngICC As Long
End Type


' Public Methods
' ==============

Public Function InitCommonControlsVB() As Boolean
   On Error Resume Next
   Dim iccex As tagInitCommonControlsEx
   ' Ensure CC available:
   With iccex
       .lngSize = LenB(iccex)
       .lngICC = ICC_USEREX_CLASSES
   End With
   InitCommonControlsEx iccex
   InitCommonControlsVB = (err.Number = 0)
   On Error GoTo 0
End Function

