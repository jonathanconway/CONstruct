VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CListItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit


Private m_sValue As String
Private m_sLabel As String



Public Property Get Label() As String
    Label = m_sLabel
End Property

Public Property Let Label(ByVal NewValue As String)
    m_sLabel = NewValue
End Property

Public Property Get Value() As String
    Value = m_sValue
End Property

Public Property Let Value(ByVal NewValue As String)
    m_sValue = NewValue
End Property



Public Function ToString() As String
    ToString = m_sValue
End Function
