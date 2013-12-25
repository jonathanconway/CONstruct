VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CParameter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' =============================================================================
' Module Name:      CParameter
' Module Type:      Class Module
' Description:      Encapsulates elements of a parameter item belonging to a
'                   primitive.
' Author(s):        Jonathan A. Conway
' -----------------------------------------------------------------------------
' Log:
'
' 04 07 16 :
'   - Modified code for setting value with a "choice" datatype to conform to
'     the new method of storing choice items
' 04 07 14 :
'   - Incorporated CParameter into the CONstruct main project
' 04 07 12 :
'   - Added an Index property so as to differentiate each primitive by a
'     unique index number
' 04 07 11 (circa) :
'   - Created CParameter
' =============================================================================



Option Explicit


' Private Variables
' -----------------

Private m_sParameterName As String
Private m_sDescription As String
Private m_iDataType As Integer
Private m_iBuilder As Integer
Private m_bIsOptional As Boolean
Private m_sData As String
Private m_bHasSpaces As Boolean     ' * New

Private m_lIndex As Long
Private m_sValue As String
Private m_oList As CList


' Properties
' ----------

Public Property Get ParameterName() As String
    ParameterName = m_sParameterName
End Property

Public Property Let ParameterName(ByVal NewValue As String)
    m_sParameterName = NewValue
End Property

Public Property Get Description() As String
    Description = m_sDescription
End Property

Public Property Let Description(ByVal NewValue As String)
    m_sDescription = NewValue
End Property

Public Property Get DataType() As Integer
    DataType = m_iDataType
End Property

Public Property Let DataType(ByVal NewValue As Integer)
    m_iDataType = NewValue
End Property

Public Property Get Builder() As Integer
    Builder = m_iBuilder
End Property

Public Property Let Builder(ByVal NewValue As Integer)
    m_iBuilder = NewValue
End Property

Public Property Get IsOptional() As Boolean
    IsOptional = m_bIsOptional
End Property

Public Property Let IsOptional(ByVal NewValue As Boolean)
    m_bIsOptional = NewValue
End Property

Public Property Get Data() As String
    'If m_iDataType = eParameterTypes.ptChoice Then
    '    m_sData = m_oList.GenerateString()
    'End If
    Data = m_sData
End Property

Public Property Let Data(ByVal NewValue As String)
    m_sData = NewValue
    If m_iDataType = eParameterTypes.ptChoice Then
        Set m_oList = New CList
        m_oList.ProcessString m_sData
    End If
End Property

Public Property Get Index() As Long
    Index = m_lIndex
End Property

Public Property Let Index(ByVal NewValue As Long)
    m_lIndex = NewValue
End Property

Public Property Get List() As CList
    Set List = m_oList
End Property

Public Property Get HasSpaces() As Boolean
    HasSpaces = m_bHasSpaces
End Property

Public Property Let HasSpaces(ByVal NewValue As Boolean)
    m_bHasSpaces = NewValue
End Property



' Methods
' -------

Public Function ToString() As String
    Dim str As String
    str = m_sParameterName
    If m_bIsOptional Then str = "[" & str & "]"
    ToString = str
End Function



Private Sub Class_Initialize()
    Set m_oList = New CList
End Sub



Public Property Get Value() As String
    If Not m_bIsOptional Then
        If Len(Trim$(m_sValue)) = 0 Then
            Select Case m_iDataType
                Case eParameterTypes.ptString
                    If Len(Trim$(m_sData)) = 0 Then
                        m_sValue = "[empty]"
                    Else
                        m_sValue = m_sData
                    End If
                Case eParameterTypes.ptNumber
                    m_sValue = GetTagAttribute(m_sData, "min")
                    If Len(m_sValue) = 0 Then m_sValue = 0
                Case eParameterTypes.ptYesNo
                    m_sValue = GetTagAttribute(m_sData, "no")
                    If Len(m_sValue) = 0 Then m_sValue = "no"
                Case eParameterTypes.ptChoice
                    Dim sChoices() As String
                    If Len(m_sData) > 0 Then
                        sChoices = Split(m_sData, vbTab, , vbBinaryCompare)
                        If UBound(sChoices) > 0 Then
                            m_sValue = sChoices(1)
                        Else
                            m_sValue = sChoices(0)
                        End If
                    Else
                        m_sValue = "[empty]"
                    End If
            End Select
        End If
    End If
    Value = m_sValue
End Property

Public Property Let Value(ByVal NewValue As String)
    m_sValue = NewValue
End Property

