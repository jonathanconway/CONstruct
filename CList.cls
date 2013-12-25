VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit


Private m_oItems As Collection


Public Function HasItem(ByVal Source As String) As Boolean

    HasItem = False
    
    Dim oItem As CListItem
    For Each oItem In m_oItems
        If oItem.Label = Source Then
            HasItem = True
            Exit Function
        End If
    Next

End Function

Public Sub ProcessString(ByVal Source As String)

    Dim sValues() As String
    Dim i As Integer
    Dim bIsDone As Boolean
    Dim oItem As CListItem
    
    Clear
    
    If Len(Source) = 0 Then Exit Sub
    
    sValues = Split(Source, vbTab, , vbBinaryCompare)
    i = LBound(sValues)
    
    Do While Not bIsDone
    
        Set oItem = New CListItem
        
        oItem.Label = sValues(i)
        If i + 1 < UBound(sValues) Then
            oItem.Value = sValues(i + 1)
        End If
        m_oItems.Add oItem
        
        i = i + 2
        If i >= UBound(sValues) Then bIsDone = True
    
    Loop

End Sub


Public Function GenerateString() As String

    Dim oItem As CListItem
    Dim sString As String
    
    For Each oItem In m_oItems
        sString = sString & oItem.Label & vbTab & oItem.Value & vbTab
    Next

    GenerateString = sString

End Function

Private Function GetIndexOfLabel(ByVal Label As String) As Integer

    GetIndexOfLabel = -1

    Dim i As Integer
    For i = 1 To m_oItems.Count
        If m_oItems(i).Label = Label Then
            GetIndexOfLabel = i
            Exit Function
        End If
    Next

End Function

Private Function GetIndexOfValue(ByVal Value As String) As Integer
    
    GetIndexOfValue = -1

    Dim i As Integer
    For i = 1 To m_oItems.Count
        If m_oItems(i).Value = Value Then
            GetIndexOfValue = i
            Exit Function
        End If
    Next

End Function

Public Sub AddItem(ByVal Label As String, ByVal Value As String)

    Dim oNewObj As CListItem
    Static lIndex As Long
    
    Set oNewObj = New CListItem
    With oNewObj
        .Label = Label
        .Value = Value
    End With
    
    m_oItems.Add oNewObj

End Sub

Public Property Get Count() As Integer
    Count = m_oItems.Count
End Property

Public Sub Delete(ByVal Label As String)
    Dim i As Integer
    For i = 1 To m_oItems.Count
        If m_oItems(i).Label = Label Then
            m_oItems.Remove i
            Exit Sub
        End If
    Next
End Sub

Public Sub Clear()
    Set m_oItems = Nothing
    Set m_oItems = New Collection
End Sub

Public Function GetItemAtIndex(ByVal Index As Integer) As CListItem
    If Index > 0 And Index <= m_oItems.Count Then
        Set GetItemAtIndex = m_oItems.Item(Index)
    End If
End Function

Public Function FindValue(ByVal Value As String) As CListItem
    Dim iIndex As Integer
    iIndex = GetIndexOfValue(Value)
    If iIndex > -1 Then
        Set FindValue = m_oItems(iIndex)
    End If
End Function


Public Function Move(ByVal Label As String, ByVal DirectionUp As Boolean) As Boolean

    Move = False
    
    If m_oItems.Count <= 1 Then Exit Function
    
    Dim iIndex As Integer
    Dim iOtherIndex As Integer
    Dim sLabel As String
    Dim sValue As String
    
    iIndex = GetIndexOfLabel(Label)
    
    ' Do error checks
    If iIndex = -1 Then Exit Function
    If DirectionUp And iIndex = 1 Then Exit Function
    If Not DirectionUp And iIndex = m_oItems.Count Then Exit Function
    
    ' Swap values
    With m_oItems(iIndex)
        sLabel = .Label
        sValue = .Value
        iOtherIndex = IIf(DirectionUp, iIndex - 1, iIndex + 1)
        .Label = m_oItems(iOtherIndex).Label
        .Value = m_oItems(iOtherIndex).Value
    End With
    With m_oItems(iOtherIndex)
        .Label = sLabel
        .Value = sValue
    End With

    Move = True

End Function

Public Function Item(ByVal Label As String) As CListItem
Attribute Item.VB_UserMemId = 0
    Dim i As Integer
    For i = 1 To m_oItems.Count
        If m_oItems(i).Label = Label Then
            Set Item = m_oItems(i)
            Exit Function
        End If
    Next
End Function

' NewEnum must return the IUnknown interface of a
' collection's enumerator.
Public Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
    Set NewEnum = m_oItems.[_NewEnum]
End Function

Private Sub Class_Initialize()
    Set m_oItems = New Collection
End Sub

Private Sub Class_Terminate()
    Set m_oItems = Nothing
End Sub




