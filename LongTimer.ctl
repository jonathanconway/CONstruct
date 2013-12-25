VERSION 5.00
Begin VB.UserControl LongTimer 
   CanGetFocus     =   0   'False
   ClientHeight    =   420
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   420
   InvisibleAtRuntime=   -1  'True
   ScaleHeight     =   420
   ScaleWidth      =   420
   Begin VB.Timer tmr 
      Interval        =   1000
      Left            =   840
      Top             =   1365
   End
   Begin VB.Image Image1 
      Height          =   420
      Left            =   0
      Picture         =   "LongTimer.ctx":0000
      Top             =   0
      Width           =   420
   End
End
Attribute VB_Name = "LongTimer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit


Private m_lHours As Long
Private m_lMinutes As Long
Private m_lSeconds As Long
Private m_lMilliseconds As Long



Public Event Tick()



Public Sub Reset()
    m_lMilliseconds = 0
    m_lSeconds = 0
    m_lMinutes = 0
    m_lHours = 0
End Sub


Public Property Get Enabled() As Boolean
    Enabled = tmr.Enabled
End Property

Public Property Let Enabled(ByVal NewValue As Boolean)
    tmr.Enabled = NewValue
End Property



Public Property Get CurrentMilliSecond() As Long
    CurrentMilliSecond = m_lMilliseconds
End Property

Public Property Get CurrentSecond() As Long
    CurrentSecond = m_lSeconds
End Property

Public Property Get CurrentMinute() As Long
    CurrentMinute = m_lMinutes
End Property

Public Property Get CurrentHour() As Long
    CurrentHour = m_lHours
End Property




Public Property Get TotalMilliSeconds() As Long
    TotalMilliSeconds = m_lMilliseconds _
                        + (m_lSeconds * 1000) _
                        + ((m_lMinutes * 60) * 1000) _
                        + (((m_lHours * 60) * 60) * 1000)
End Property

Public Property Get TotalSeconds() As Long
    TotalSeconds = m_lSeconds _
                    + (m_lMinutes * 60) _
                    + ((m_lHours * 60) * 60)
End Property

Public Property Get TotalMinutes() As Long
    TotalMinutes = m_lMinutes _
                    + ((m_lHours) * 60)
End Property

Public Property Get TotalHours() As Long
    TotalHours = m_lHours
End Property




Private Sub tmr_Timer()
    
    m_lSeconds = m_lSeconds + 1
   
    If m_lSeconds > 60 Then
        m_lSeconds = 0
        m_lMinutes = m_lMinutes + 1
    End If

    If m_lMinutes > 60 Then
        m_lMinutes = 0
        m_lHours = m_lHours + 1
    End If

    RaiseEvent Tick

End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    tmr.Enabled = StringToBoolean(PropBag.ReadProperty("Enabled", "True"))
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    PropBag.WriteProperty "Enabled", StringToBoolean(tmr.Enabled), "True"
End Sub


Private Function StringToBoolean(ByVal Source As String) As Boolean
    StringToBoolean = IIf(Source = "True", True, False)
End Function

Private Function BooleanToString(ByVal Source As Boolean) As String
    BooleanToString = IIf(Source = True, "True", "False")
End Function

Private Sub UserControl_Resize()
    
    UserControl.Height = 420
    UserControl.Width = 420

End Sub



