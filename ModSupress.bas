Attribute VB_Name = "ModSupress"
Option Explicit
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Copyright ©1996-2004 VBnet, Randy Birch, All Rights Reserved.
' Some pages may also contain other copyrights by the author.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Distribution: You can freely use this code in your own
'               applications, but you may not reproduce
'               or publish this code on any web site,
'               online service, or distribute as source
'               on any media without express permission.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Declare Function CallWindowProc Lib "user32" _
    Alias "CallWindowProcA" _
   (ByVal lpPrevWndFunc As Long, _
    ByVal hwnd As Long, _
    ByVal Msg As Long, _
    ByVal wParam As Long, _
    ByVal lParam As Long) As Long

Private Declare Function SetWindowLong Lib "user32" _
    Alias "SetWindowLongA" _
   (ByVal hwnd As Long, _
    ByVal nIndex As Long, _
    ByVal dwNewLong As Long) As Long

Private Const GWL_WNDPROC As Long = (-4)
Private Const WM_CONTEXTMENU As Long = &H7B
Public defWndProc As Long


Public Sub Hook(hwnd As Long)

   If defWndProc = 0 Then
    
      'defWndProc = SetWindowLong(hwnd, _
                                 GWL_WNDPROC, _
                                 AddressOf WindowProc)
   
      defWndProc = SetWindowLong(hwnd, _
                                 WM_PASTE, _
                                 AddressOf WindowProc)
   
   End If
                                 
End Sub


Public Sub UnHook(hwnd As Long)

    If defWndProc > 0 Then
    
      'Call SetWindowLong(hwnd, GWL_WNDPROC, defWndProc)
      Call SetWindowLong(hwnd, WM_PASTE, defWndProc)
      
      defWndProc = 0
      
   End If
   
    
End Sub


Public Function WindowProc(ByVal hwnd As Long, _
                           ByVal uMsg As Long, _
                           ByVal wParam As Long, _
                           ByVal lParam As Long) As Long

'   Select Case uMsg
'      Case WM_CONTEXTMENU
'
'        'this executes when the window is hooked
'         Form1.PopupMenu Form1.mnuPopup
'         WindowProc = 1
'
'      Case Else
'
'         WindowProc = CallWindowProc(defWndProc, _
'                                     hwnd, _
'                                     uMsg, _
'                                     wParam, _
'                                     lParam)
'   End Select
    
End Function


