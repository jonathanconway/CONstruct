VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CKeyWord"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit


Private m_sText As String               ' Current text (code)
Private m_iIndex As Integer             ' Index of the keyword

Private m_oPrimitive As CPrimitive      ' Holds current primitive

Private m_iBeginPosition As Integer     ' Position of the keyword

'Private m_oParent As CParser            ' Object reference to parent



Public Function ToString() As String
    ToString = m_sText
End Function


'Public Property Let Parent(ByRef NewValue As CParser)
'    Set m_oParent = NewValue
'End Property


Public Property Get Index() As Integer
    Index = m_iIndex
End Property

Public Property Let Index(ByVal NewValue As Integer)
    m_iIndex = NewValue
End Property



Public Property Get Primitive() As CPrimitive
    If Not m_oPrimitive Is Nothing Then
        Set Primitive = m_oPrimitive
    End If
End Property

Public Property Let Primitive(ByRef NewValue As CPrimitive)
    Set m_oPrimitive = Nothing
    Set m_oPrimitive = NewValue
End Property



Public Property Get Text() As String
    Text = m_sText
End Property

Public Property Let Text(ByVal NewValue As String)
    m_sText = NewValue
End Property


Public Property Get BeginPosition() As Integer
    BeginPosition = m_iBeginPosition
End Property

Public Property Let BeginPosition(ByVal NewValue As Integer)
    m_iBeginPosition = NewValue
End Property


