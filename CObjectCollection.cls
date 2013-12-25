VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CObjectCollection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' =============================================================================
' Module Name:      CObjectCollection
' Module Type:      Class Module
' Description:      Multi-purpose, generic, re-useable collection class
'                   containing objects of type Object
' Author(s):        Jonathan A. Conway
' -----------------------------------------------------------------------------
' Log:
'
' 04 07 26 :
'   - Created CObjectCollection class
' =============================================================================


Option Explicit


Public Event ItemAdded(ByRef Item As Object)
Public Event ItemAccessed(ByRef Item As Object)
Public Event ItemDeleted(ByVal Index As Integer)


Private m_oItems As Collection

Private m_oParent As Object





Public Function Add(ByRef NewObject As Object) As Long

    On Error GoTo ProcedureError

    Dim oNewObj As Object
    Static lIndex As Long
    
    Set oNewObj = NewObject
    lIndex = lIndex + 1
    
    oNewObj.Index = lIndex
    
    If Not (m_oParent Is Nothing) Then
        oNewObj.Parent = m_oParent
    End If
    
    m_oItems.Add oNewObj, Chr(97) & lIndex
    Add = lIndex
    
    RaiseEvent ItemAdded(NewObject)
    
    Exit Function

ProcedureError:
    If err.Number = 438 Then
        ' Ignore error 438, some objects don't have Index property
        Resume Next
    Else
        MsgBox err.Number & vbNewLine & err.Description
    End If

End Function


Public Function Count() As Long
    Count = m_oItems.Count
End Function


Public Sub Delete(ByVal Index As Long)
    m_oItems.Remove Chr(97) & Index
    RaiseEvent ItemDeleted(Index)
End Sub

Public Sub Clear()
    Set m_oItems = Nothing
    Set m_oItems = New Collection
End Sub

Public Function Item(ByVal Index As Long) As Object
Attribute Item.VB_UserMemId = 0
    Set Item = m_oItems.Item(Chr(97) & Index)
    RaiseEvent ItemAccessed(Item)
End Function

' NewEnum must return the IUnknown interface of a
' collection's enumerator.
Public Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
    Set NewEnum = m_oItems.[_NewEnum]
End Function


Public Function FindItem(ByVal ItemString As String) As Object
    On Error GoTo ProcedureError
    
    Dim oItem As Object
    For Each oItem In m_oItems
        If oItem.ToString() = ItemString Then
            Set FindItem = oItem
            Exit Function
        End If
    Next
    Exit Function
    
ProcedureError:
    If err.Number = 438 Then
        ' Ignore error 438; some items won't have a ToString method
        Resume Next
    Else
        MsgBox err.Number & vbNewLine & err.Description
    End If
End Function


Private Sub Class_Initialize()
    Set m_oItems = New Collection
End Sub

Private Sub Class_Terminate()
    Set m_oItems = Nothing
End Sub


Public Property Let Parent(ByRef NewValue As Object)
    Set m_oParent = NewValue
End Property
