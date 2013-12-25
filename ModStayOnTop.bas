Attribute VB_Name = "ModStayOnTop"
' =============================================================================
' Module Name:      ModStayOnTop
' Module Type:      Code Module
' Description:      Make any form a top-most window or a normal window
' Author(s):        Microsoft Knowledge Base (modified by Jonathan A. Conway)
' -----------------------------------------------------------------------------
' Log:
'                   ?? ?? ?? : Got the code from a Microsoft KB article and
'                              modified it to suit my purposes
' =============================================================================

Option Explicit

Private Const SWP_NOMOVE = 2
Private Const SWP_NOSIZE = 1
Private Const HWND_TOPMOST = -1
Private Const HWND_NOTOPMOST = -2
Private Const XFLAGS = SWP_NOMOVE Or SWP_NOSIZE

Private Declare Function SetWindowPos Lib "user32" _
      (ByVal hWnd As Long, _
      ByVal hWndInsertAfter As Long, _
      ByVal x As Long, _
      ByVal y As Long, _
      ByVal cx As Long, _
      ByVal cy As Long, _
      ByVal wFlags As Long) As Long


Private Function SetTopMostWindow(hWnd As Long, Topmost As Boolean) As Long
    ' Function sets a window as always on top, or turns this off
      
    ' hwnd - handle the the window to affect
    ' Topmost - do you want it always on top or not
     
   On Error GoTo ErrHandler
     
   If Topmost = True Then 'Make the window topmost
      SetTopMostWindow = SetWindowPos(hWnd, HWND_TOPMOST, 0, 0, 0, 0, XFLAGS)
   Else
      SetTopMostWindow = SetWindowPos(hWnd, HWND_NOTOPMOST, 0, 0, 0, 0, XFLAGS)
      'SetTopMostWindow = SetWindowPos(hWnd, HWND_TOP, 0, 0, 0, 0, XFLAGS)
      SetTopMostWindow = False
   End If
      
    Exit Function
ErrHandler:
    Select Case err.Number
        Case Else
            err.Raise err.Number, err.Source & "+modAPIStuff/SetTopMostWindow", err.Description
    End Select
End Function


Public Sub StayOnTop(ByRef Target As Form, ByVal IsTopMost As Boolean)
    SetTopMostWindow Target.hWnd, IsTopMost
End Sub

