VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CToken"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' =============================================================================
' Module Name:      CToken
' Module Type:      Class
' Description:      Encapsulates data of a single token defenition
' Author(s):        Jonathan A. Conway
' -----------------------------------------------------------------------------
' Log:
'
' 04 08 06 :
'   - Changed the datatype of the Structure property to a CStructure object
'
' 04 06 19 :
'   - Commenced writing CToken from scratch
' =============================================================================



Option Explicit



' Enumerations
' ============

Public Enum eTokenType
    [tktBeginner] = 0       ' Begins a structure (example: "if ...")
    [tktEnder] = 1          ' Ends a structure (example: "... endif")
    [tktStandAlone] = 2     ' Stands by itself (example: "return")
    [tktDelimiter] = 3      ' Seperates two or more items (e.g. ", ")
End Enum



' Private Variables
' =================

Private m_sText As String
Private m_oType As eTokenType
Private m_bIsCaseSensitive As Boolean
Private m_bBeginsOnNewLine As Boolean
Private m_bIsInsideStructure As Boolean
Private m_iIndex As Integer

Private m_oStructure As CStructure


' Public Properties
' =================

Public Property Get Index() As Integer
    Index = m_iIndex
End Property

Public Property Let Index(ByVal NewValue As Integer)
    m_iIndex = NewValue
End Property

Public Property Get Text() As String
    Text = m_sText
End Property

Public Property Let Text(ByVal NewValue As String)
    m_sText = NewValue
End Property

Public Property Get TokenType() As eTokenType
    TokenType = m_oType
End Property

Public Property Let TokenType(ByVal NewValue As eTokenType)
    m_oType = NewValue
End Property

Public Property Get IsCaseSensitive() As Boolean
    IsCaseSensitive = m_bIsCaseSensitive
End Property

Public Property Let IsCaseSensitive(ByVal NewValue As Boolean)
    m_bIsCaseSensitive = NewValue
End Property

Public Property Get BeginsOnNewLine() As Boolean
    BeginsOnNewLine = m_bBeginsOnNewLine
End Property

Public Property Let BeginsOnNewLine(ByVal NewValue As Boolean)
    m_bBeginsOnNewLine = NewValue
End Property

Public Property Get IsInsideStructure() As Boolean
    IsInsideStructure = m_bIsInsideStructure
End Property

Public Property Let IsInsideStructure(ByVal NewValue As Boolean)
    m_bIsInsideStructure = NewValue
End Property

Public Property Get Structure() As CStructure
    Set Structure = m_oStructure
End Property

Public Property Let Structure(ByRef NewValue As CStructure)
    Set m_oStructure = NewValue
End Property


Public Function ToString() As String
    ToString = m_sText
End Function

