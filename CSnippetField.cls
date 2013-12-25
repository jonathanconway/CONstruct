VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CSnippetField"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit


Private m_sFieldName As String
Private m_iDataType As Integer
Private m_lIndex As Long


Public Property Get FieldName() As String
    FieldName = m_sFieldName
End Property

Public Property Let FieldName(ByVal NewValue As String)
    m_sFieldName = NewValue
End Property



Public Property Get DataType() As Integer
    DataType = m_iDataType
End Property

Public Property Let DataType(ByVal NewValue As Integer)
    m_iDataType = NewValue
End Property


Public Property Get Index() As Long
    Index = m_lIndex
End Property

Public Property Let Index(ByVal NewValue As Long)
    m_lIndex = NewValue
End Property



Public Function ToString() As String
    ToString = m_sFieldName
End Function

