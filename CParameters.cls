VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CParameters"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' =============================================================================
' Module Name:      CParameters
' Module Type:      Class Module
' Description:      Collection class containing parameters for a primitive.
'                   (See CParameter)
' Author(s):        Jonathan A. Conway
' -----------------------------------------------------------------------------
' Log:
'
' 04 07 14 :
'   - Incorporated CParameters into the CONstruct main project
' 04 07 12 :
'   - Made a bug fix that involved the addition of an Index property to the
'     CParameter class.
' 04 07 11 (circa) :
'   - Created CParameters
' =============================================================================



Option Explicit


Private m_oItems As Collection

Public Function Add(ByRef NewObject As CParameter) As Long

    Dim oNewObj As CParameter
    Static lIndex As Long
    
    Set oNewObj = NewObject
    
    lIndex = lIndex + 1
    oNewObj.Index = lIndex
    m_oItems.Add oNewObj, Chr(97) & lIndex
    Add = lIndex

End Function


Public Function Count() As Long
    Count = m_oItems.Count
End Function


Public Sub Delete(ByVal Index As Long)
    m_oItems.Remove Chr(97) & Index
End Sub

Public Sub Clear()
    Set m_oItems = Nothing
    Set m_oItems = New Collection
End Sub

Public Function Item(ByVal Index As Long) As CParameter
Attribute Item.VB_UserMemId = 0
    Set Item = m_oItems.Item(Chr(97) & Index)
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



