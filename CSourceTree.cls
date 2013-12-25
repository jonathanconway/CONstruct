VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CSourceTree"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit


' Private Variables
' -----------------

Private m_iIndex As Integer         ' Index within collection

Private WithEvents m_oObjects As CObjects
Attribute m_oObjects.VB_VarHelpID = -1

Private WithEvents m_oParser As CParser
Attribute m_oParser.VB_VarHelpID = -1


' Public Events
' -------------

Public Event ObjectAdded(ByRef Object As Object)
Public Event ObjectDeleted(ByVal ObjectID As Integer)
Public Event ObjectUpdated(ByRef Object As Object)
Public Event ObjectsCleared()
Public Event ParseProgress(ByVal CurrentPosition As Long, ByVal TotalLength As Long)


' Public Properties
' -----------------

Public Property Get Objects() As CObjects
    Set Objects = m_oObjects
End Property


Public Property Get Parser() As CParser
    Set Parser = m_oParser
End Property

Public Property Let Parser(ByRef NewValue As CParser)
    Set m_oParser = NewValue
End Property


Public Property Get Index() As Integer
    Index = m_iIndex
End Property

Public Property Let Index(ByVal NewValue As Integer)
    m_iIndex = NewValue
End Property



' Public Methods
' --------------

Public Sub ClearObjects()

    Set m_oObjects = Nothing
    Set m_oObjects = New CObjects
    RaiseEvent ObjectsCleared

End Sub




Public Sub SetCode(ByVal Source As String)

    ' This is where Source is converted into structures that are created
    
    ' 1. Source => Parser
    ' 2. Parser => List of blocks
    ' 2. List of blocks => Object

    Dim oBlocks As CObjectCollection
    Dim oBlock As CBlock
    
    Dim oObject As Object
    
    Dim bSkip As Boolean
    
    ' Clear out objects collection
    ClearObjects
    
    ' Parse the code
    m_oParser.Parse Source
    
    Set oBlocks = m_oParser.Blocks
    
    For Each oBlock In oBlocks
        bSkip = False
        
        Select Case oBlock.Structure
            ''Case gbtCStyleComment To [gbtC++StyleComment]
            ''    Set oObject = New CComment
            Case [csDefine]
                Set oObject = New CDefine
            Case [csActor]
                Set oObject = New CActor
'            Case [gbtState]
'                Set oObject = New CState
'            Case [gbtAction]
'                Set oObject = New CAction
'            Case [gbtAI]
'                Set oObject = New CAI
            Case Else
                bSkip = True
        End Select

        If Not bSkip Then
            oObject.SetCode oBlock.Text
            m_oObjects.Add oObject
        End If
    Next

    'm_sCode = Source

End Sub



' Event Handlers
' --------------

Private Sub Class_Initialize()
    
    Set m_oObjects = New CObjects

End Sub



Private Sub Class_Terminate()

    ' Garbage Collection
    Set m_oParser = Nothing
    Set m_oObjects = Nothing

End Sub

Private Sub m_oObjects_ObjectAdded(Object As Object)
    RaiseEvent ObjectAdded(Object)
End Sub

Private Sub m_oObjects_ObjectDeleted(ByVal ObjectID As Integer)
    RaiseEvent ObjectDeleted(ObjectID)
End Sub

Private Sub m_oObjects_ObjectUpdated(Object As Object)
    RaiseEvent ObjectUpdated(Object)
End Sub

Private Sub m_oParser_ParseProgress(ByVal CurrentPosition As Long, ByVal TotalLength As Long)
    RaiseEvent ParseProgress(CurrentPosition, TotalLength)
End Sub
