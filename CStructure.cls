VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CStructure"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' - BeginTokens[]
' - EndTokens[]
' - Format
' - IntrinsicPrimitive
' - IsCodeContainer

' - DoIndent()

' INTERNAL VARIABLES:
' - IndentationMarkers[] --> PUT IN CDefinition INSTEAD!!!



Private m_sStructureName As String
Private m_sDescription As String
Private m_sDukeVersion As String
Private m_lHelpContextID As Long
Private m_iImageID As Integer
Private m_iIndex As Integer
Private m_oPrimitive As CPrimitive


Public Property Get Index() As Integer
    Index = m_iIndex
End Property

Public Property Let Index(ByVal NewValue As Integer)
    m_iIndex = NewValue
End Property


Public Property Get ImageID() As Integer
    ImageID = m_iImageID
End Property

Public Property Let ImageID(ByVal NewValue As Integer)
    m_iImageID = NewValue
End Property


Public Property Get Primitive() As CPrimitive
    Set Primitive = m_oPrimitive
End Property

Public Property Let Primitive(ByRef NewValue As CPrimitive)
    Set m_oPrimitive = NewValue
End Property



Public Property Get StructureName() As String
    StructureName = m_sStructureName
End Property

Public Property Let StructureName(ByVal NewValue As String)
    m_sStructureName = NewValue
End Property


Public Property Get Description() As String
    Description = m_sDescription
End Property

Public Property Let Description(ByVal NewValue As String)
    m_sDescription = NewValue
End Property

Public Property Get DukeVersion() As String
    DukeVersion = m_sDukeVersion
End Property

Public Property Let DukeVersion(ByVal NewValue As String)
    m_sDukeVersion = NewValue
End Property

Public Property Get HelpContextID() As Long
    HelpContextID = m_lHelpContextID
End Property

Public Property Let HelpContextID(ByVal NewValue As Long)
    m_lHelpContextID = NewValue
End Property


Public Function ToString() As String
    ToString = m_sStructureName
End Function
