VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CDefinition"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' =============================================================================
' Module Name:      CDefinition
' Module Type:      Class Module
' Description:      Encapsulates elements of a Primitive Definition. Used to
'                   manipulate primitives and primitive parameters in various
'                   helpers in CONstruct.
' Author(s):        Jonathan A. Conway
' -----------------------------------------------------------------------------
' Log:
'
' 04 07 30 :
'   - Added structures and tokens collections to the Definition class and
'     updated loading/saving code to reflect these code.
'   - Modularized loading/saving code to be easier to read/modify
'
' 04 07 27 :
'   - Added a type definition, tStructure so that structures can be stored
'
' 04 07 14 :
'   - Incorporated CDefinition into the CONstruct main project
'
' 04 07 13 :
'   - Fixed several bugs including problem with adding an extra blank
'     primitive object when loading from a file.
'
' 04 07 11 (circa) :
'   - Created CDefinition
' =============================================================================


Option Explicit


' Type Definitions
' ================


Private Type tToken
    tkText As String                ' Text value of the token
    tkTokenType As Integer          ' Token type (see eTokenTypes)
    tkStructure As String           ' String name of related structure
    tkIsInsideStructure As Boolean  ' Is structure inclusive of text?
    tkIsCaseSensitive As Boolean    ' Is token case sensitive?
    tkBeginsOnNewLine As Boolean    ' Must the token begin on a new line?
End Type

Private Type tParameter
    paName As String                ' Name of the parameter (e.g. "Radius")
    paDescription As String         ' Brief description of what parameter does
    paDataType As Integer           ' Data-type of the parameter
    paBuilder As Integer            ' Which builder (if any) to implement
    paIsOptional As Boolean         ' Is the parameter required or optional?
    paData As String                ' Any extra data required goes here
    paHasSpaces As Boolean          ' Can the value contain spaces?
End Type

Private Type tPrimitive
    pmName As String                ' Name of the primitive (e.g. "Move")
    pmCategory As String            ' Category primitive goes under
    pmDescription As String         ' Brief description of what primitive does
    pmDukeVersion As String         ' Lowest Duke3D version supported
    pmHelpContextID As Long         ' Help Context ID
    pmParameters() As tParameter    ' Array of parameters
    pmIsStandAlone As Boolean       ' Is the primitive a single unit?
End Type

Private Type tStructure
    stName As String                ' Name of the structure (e.g. "Actor")
    stDescription As String         ' Description of the structure
    stDukeVersion As String         ' Lowest Duke3D version supported
    stHelpContextID As Long         ' Help context ID for the structure
    stIntrinsicPrimitive As String  ' Name of primitive to use for structure
    stImageID As Integer            ' ID number of representative image
End Type

Private Type tDefinition
    dfAuthor As String              ' Author of the definition
    dfComments As String            ' Any additional information
    dfPrimitives() As tPrimitive    ' Array of primitives
    dfStructures() As tStructure    ' Array of structures
    dfTokens() As tToken            ' Array of tokens
End Type


' Enumerations
' ============

Public Enum eParameterTypes
    [ptString] = 0              ' Alphanumeric string (regular text-box)
    [ptNumber] = 1              ' Number (data must include min/max values)
    [ptYesNo] = 2               ' Yes/No checkbox (data includes result values)
    [ptChoice] = 3              ' Combo-box (data must include options/data)
End Enum


' Private Variables
' =================

Private m_sAuthor As String
Private m_sComments As String

Private m_oPrimitives As CObjectCollection
Private m_oStructures As CObjectCollection
Private m_oTokens As CObjectCollection

Private m_oDef As tDefinition
    


' Public Properties
' =================

Public Property Get Author() As String
    Author = m_sAuthor
End Property

Public Property Let Author(ByVal NewValue As String)
    m_sAuthor = NewValue
End Property

Public Property Get Comments() As String
    Comments = m_sComments
End Property

Public Property Let Comments(ByVal NewValue As String)
    m_sComments = NewValue
End Property

Public Property Get Primitives() As CObjectCollection
    Set Primitives = m_oPrimitives
End Property

Public Property Get Structures() As CObjectCollection
    Set Structures = m_oStructures
End Property

Public Property Get Tokens() As CObjectCollection
    Set Tokens = m_oTokens
End Property



' Public Methods
' ==============

Public Sub SaveBinary(ByVal Filename As String)
    
    ' Saves the all primitives and parameters in the current definition to
    ' a binary file.
    
    Dim lHnd As Long
    
    ' Transfer class instances to UDT instances
    ' Set Definition properties
    m_oDef.dfAuthor = m_sAuthor
    m_oDef.dfComments = m_sComments
    
    SavePrimitives      ' Save primitives
    SaveStructures      ' Save structure
    SaveTokens          ' Save tokens
    
    ' Write UDT instance to the specified filename
    lHnd = FreeFile()
    Open Filename For Binary As #lHnd
    Put #lHnd, , m_oDef
    Close #lHnd
    
End Sub

Public Sub LoadBinary(ByVal Filename As String)

    ' Loads definition (including all primitives and parameters) from a
    ' binary file.
    
    Dim lHnd As Long                ' File handle
    
    ' Clear internal primitives collection prior to loading file
    m_oPrimitives.Clear
    m_oStructures.Clear
    m_oTokens.Clear
    
    ' Load file contents into variables
    lHnd = FreeFile()
    Open Filename For Binary As #lHnd
    Get #lHnd, , m_oDef
    Close #lHnd
    
    ' Get Definition properties
    m_sAuthor = m_oDef.dfAuthor
    m_sComments = m_oDef.dfComments
    
    LoadPrimitives      ' Load all primitives
    LoadStructures      ' Load all structures
    LoadTokens          ' Load all tokens
    
End Sub


' Private Methods
' ===============

' Loading/Saving Methods
' ----------------------

Private Sub SaveTokens()
    
    ' Save all tokens to UDT instance
    
    Dim oToken As CToken
    Dim i As Integer

    If m_oTokens.Count = 0 Then Exit Sub

    ReDim m_oDef.dfTokens(0 To m_oTokens.Count() - 1)
    For Each oToken In m_oTokens
        With m_oDef.dfTokens(i)
            ' Set properties
            .tkBeginsOnNewLine = oToken.BeginsOnNewLine
            .tkIsCaseSensitive = oToken.IsCaseSensitive
            .tkIsInsideStructure = oToken.IsInsideStructure
            .tkStructure = oToken.Structure.StructureName
            .tkText = oToken.Text
            .tkTokenType = oToken.TokenType
        End With
        i = i + 1
    Next oToken

End Sub

Private Sub SaveStructures()

    ' Save all structures to UDT instance
    
    Dim oStruct As CStructure
    Dim i As Integer

    If m_oStructures.Count = 0 Then Exit Sub

    ReDim m_oDef.dfStructures(0 To m_oStructures.Count() - 1)
    For Each oStruct In m_oStructures
        With m_oDef.dfStructures(i)
            ' Set properties
            .stDescription = oStruct.Description
            .stDukeVersion = oStruct.DukeVersion
            .stHelpContextID = oStruct.HelpContextID
            .stImageID = oStruct.ImageID
            .stIntrinsicPrimitive = oStruct.Primitive.PrimitiveName
            .stName = oStruct.StructureName
        End With
        i = i + 1
    Next oStruct

End Sub

Private Sub SavePrimitives()

    ' Save all primitives to UDT instance

    Dim oPrim As CPrimitive
    Dim i As Integer

    If m_oPrimitives.Count = 0 Then Exit Sub

    ReDim m_oDef.dfPrimitives(0 To m_oPrimitives.Count() - 1)
    For Each oPrim In m_oPrimitives
        With m_oDef.dfPrimitives(i)
            ' Set properties
            .pmCategory = oPrim.Category
            .pmDescription = oPrim.Description
            .pmDukeVersion = oPrim.DukeVersion
            .pmHelpContextID = oPrim.HelpContextID
            .pmIsStandAlone = oPrim.IsStandAlone
            .pmName = oPrim.PrimitiveName
            
            ' Add parameters
            SaveParameters oPrim, i
        End With
        i = i + 1
    Next oPrim

End Sub

Private Sub SaveParameters(ByRef Source As CPrimitive, ByVal PrimitiveIndex As Integer)

    ' Save all primitive parameters to UDT instance, under specified primitive
    
    Dim oParam As CParameter
    Dim j As Integer
    
    If Source.Parameters.Count = 0 Then Exit Sub
    
    ' Add Parameters
    With m_oDef.dfPrimitives(PrimitiveIndex)
        ReDim .pmParameters(0 To IIf(Source.Parameters.Count() = 0, 0, Source.Parameters.Count() - 1))
        j = -0
        For Each oParam In Source.Parameters
            With m_oDef.dfPrimitives(PrimitiveIndex).pmParameters(j)
                ' Set parameter properties
                .paBuilder = oParam.Builder
                .paData = oParam.Data
                .paDataType = oParam.DataType
                .paDescription = oParam.Description
                .paHasSpaces = oParam.HasSpaces
                .paIsOptional = oParam.IsOptional
                .paName = oParam.ParameterName
            End With
            j = j + 1
        Next
    End With

End Sub

Private Sub LoadPrimitives()

    On Error GoTo ProcedureError

    Dim i As Integer
    Dim oPrim As CPrimitive

    For i = LBound(m_oDef.dfPrimitives) To UBound(m_oDef.dfPrimitives)
        Set oPrim = New CPrimitive
        With m_oDef.dfPrimitives(i)
            ' Get Properties
            oPrim.Category = .pmCategory
            oPrim.Description = .pmDescription
            oPrim.DukeVersion = .pmDukeVersion
            oPrim.HelpContextID = .pmHelpContextID
            oPrim.IsStandAlone = .pmIsStandAlone
            oPrim.PrimitiveName = .pmName
        End With
        
        LoadParameters oPrim, i
        
        m_oPrimitives.Add oPrim
    Next i

    Exit Sub

ProcedureError:
    If err.Number = 9 Then
    Else
        MsgBox err.Number & vbNewLine & err.Description
    End If

End Sub

Private Sub LoadParameters(ByRef Source As CPrimitive, ByVal PrimitiveIndex As Integer)

    On Error GoTo ProcedureError

    Dim i As Integer
    Dim oParam As CParameter

    With m_oDef.dfPrimitives(PrimitiveIndex)
        ' Get Parameters
        For i = LBound(.pmParameters) To UBound(.pmParameters)
            If .pmParameters(i).paName = "" Then Exit For
            Set oParam = New CParameter
            With m_oDef.dfPrimitives(PrimitiveIndex).pmParameters(i)
                ' Set parameter properties
                oParam.DataType = .paDataType
                oParam.Builder = .paBuilder
                oParam.Description = .paDescription
                oParam.HasSpaces = .paHasSpaces
                oParam.IsOptional = .paIsOptional
                oParam.ParameterName = .paName
                oParam.Data = .paData
            End With
            Source.Parameters.Add oParam
        Next i
    End With

    Exit Sub

ProcedureError:
    If err.Number = 9 Then
    Else
        MsgBox err.Number & vbNewLine & err.Description
    End If

End Sub

Private Sub LoadStructures()

    On Error GoTo ProcedureError

    Dim i As Integer
    Dim oStruct As CStructure

    For i = LBound(m_oDef.dfStructures) To UBound(m_oDef.dfStructures)
        Set oStruct = New CStructure
        With m_oDef.dfStructures(i)
            ' Put properties in new structure object
            oStruct.Description = .stDescription
            oStruct.DukeVersion = .stDukeVersion
            oStruct.HelpContextID = .stHelpContextID
            oStruct.ImageID = .stImageID
            oStruct.StructureName = .stName
            
            ' Populate structure's "Primitive" property if possible
            If Len(Trim$(.stIntrinsicPrimitive)) > 0 Then
                oStruct.Primitive = m_oPrimitives.FindItem(.stIntrinsicPrimitive)
            End If
            'oStruct.IntrinsicPrimitive = .stIntrinsicPrimitive
        End With
        
        m_oStructures.Add oStruct
    Next i
    
    Exit Sub

ProcedureError:
    If err.Number = 9 Then
    Else
        MsgBox err.Number & vbNewLine & err.Description
    End If

End Sub

Private Sub LoadTokens()
    
    On Error GoTo ProcedureError
    
    Dim i As Integer
    Dim oToken As CToken

    For i = LBound(m_oDef.dfTokens) To UBound(m_oDef.dfTokens)
        Set oToken = New CToken
        With m_oDef.dfTokens(i)
            ' Put properties into new token object
            oToken.BeginsOnNewLine = .tkBeginsOnNewLine
            oToken.IsCaseSensitive = .tkIsCaseSensitive
            oToken.IsInsideStructure = .tkIsInsideStructure
            oToken.Text = .tkText
            oToken.TokenType = .tkTokenType
            
            ' Populate token's "Structure" property if possible
            If Len(Trim$(.tkStructure)) > 0 Then
                oToken.Structure = m_oStructures.FindItem(.tkStructure)
            End If
        End With
        
        m_oTokens.Add oToken
    Next i
    
    Exit Sub

ProcedureError:
    If err.Number = 9 Then
    Else
        MsgBox err.Number & vbNewLine & err.Description
    End If

End Sub



' Event Handlers
' --------------

Private Sub Class_Initialize()
    Set m_oPrimitives = New CObjectCollection
    Set m_oStructures = New CObjectCollection
    Set m_oTokens = New CObjectCollection
End Sub

Private Sub Class_Terminate()
    Set m_oPrimitives = Nothing
    Set m_oStructures = Nothing
    Set m_oTokens = Nothing
End Sub

