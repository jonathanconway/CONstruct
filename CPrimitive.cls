VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CPrimitive"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' =============================================================================
' Module Name:      CPrimitive
' Module Type:      Class Module
' Description:      Encapsulates elements of a primitive in a primitives
'                   definition. Contains a collection of parameters belonging
'                   to the primitive.
'                   (see CParameter)
' Author(s):        Jonathan A. Conway
' -----------------------------------------------------------------------------
' Log:
'
' 04 07 14 :
'   - Incorporated CPrimitive into the CONstruct main project
'   - Added a Category property to store the category the primitive goes under
'   - Added a Syntax property, which generates a string describing the syntax
' 04 07 12 :
'   - Added an Index property so as to differentiate each primitive by a
'     unique index number
' 04 07 11 (circa) :
'   - Created CDefinition
' =============================================================================



Option Explicit


' Private Variables
' -----------------

Private m_sPrimitiveName As String
Private m_sCategory As String
Private m_sDescription As String
Private m_sDukeVersion As String
Private m_lIndex As Long
Private m_lHelpContextID As Long
Private m_oParameters As CObjectCollection
Private m_bIsStandAlone As Boolean


' Properties
' ----------

Public Property Get PrimitiveName() As String
    PrimitiveName = m_sPrimitiveName
End Property

Public Property Let PrimitiveName(ByVal NewValue As String)
    m_sPrimitiveName = NewValue
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

Public Property Get Category() As String
    Category = m_sCategory
End Property

Public Property Let Category(ByVal NewValue As String)
    m_sCategory = NewValue
End Property

Public Property Get Index() As Long
    Index = m_lIndex
End Property

Public Property Let Index(ByVal NewValue As Long)
    m_lIndex = NewValue
End Property

Public Property Get Parameters() As CObjectCollection
    Set Parameters = m_oParameters
End Property

Public Property Get IsStandAlone() As Boolean
    IsStandAlone = m_bIsStandAlone
End Property

Public Property Let IsStandAlone(ByVal NewValue As Boolean)
    m_bIsStandAlone = NewValue
End Property

Public Property Get Syntax() As String

    Dim sSyntax As String
    Dim oParam As CParameter
    
    sSyntax = m_sPrimitiveName & " "
    
    For Each oParam In m_oParameters
        sSyntax = sSyntax & oParam.ToString() & " "
    Next

    Syntax = sSyntax

End Property


Private Sub Class_Initialize()
    Set m_oParameters = New CObjectCollection
End Sub

Private Sub Class_Terminate()
    Set m_oParameters = Nothing
End Sub


Public Function ToString() As String
    ToString = m_sPrimitiveName
End Function

