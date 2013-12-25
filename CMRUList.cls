VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CMRUList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit


Private m_sItems() As String

Private m_iMaxItems As Integer


' 0     -   one
' 1     -   two
' 2     -   three
' 3     -   four


Public Property Get MaxItems() As Integer
    MaxItems = m_iMaxItems
End Property

Public Property Let MaxItems(ByVal NewValue As Integer)
    m_iMaxItems = NewValue
    If NewValue > 0 Then
        ReDim Preserve m_sItems(0 To m_iMaxItems - 1)
    End If
End Property


Public Property Get Items() As String()
    Items = m_sItems
End Property




Public Sub AddMRUItem(ByVal Filename As String)

    ' Adds an item at the front
    
    If UBound(m_sItems) = 0 Then
        If m_sItems(0) = "" Then
            m_sItems(0) = Filename
            Exit Sub
        End If
    End If
    
    If UBound(m_sItems) < (m_iMaxItems - 1) Then
        ReDim Preserve m_sItems(LBound(m_sItems) To UBound(m_sItems) + 1)
    End If

    Dim i As Integer
    i = UBound(m_sItems)
    Do While i > LBound(m_sItems)
        m_sItems(i) = m_sItems(i - 1)
        i = i - 1
    Loop

    m_sItems(LBound(m_sItems)) = Filename

End Sub


Public Function ToString() As String
    ToString = Join$(m_sItems, vbTab)
End Function


Public Sub LoadFromString(ByRef Source As String)
    m_sItems = Split(Source, vbTab)
End Sub





Private Sub Class_Initialize()
    ReDim m_sItems(0)
End Sub

