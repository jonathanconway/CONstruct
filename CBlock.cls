VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CBlock"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit


Public Enum eBlockErrors
    [be_NoErrors] = 0
    [be_NoEndMarker] = 1
    [be_NoBlockBegin] = 2
    [be_NoBlockEnd] = 3
End Enum



Private m_sText As String               ' Current text (code)
Private m_iErrors() As String           ' Array of Errors
Private m_iIndex As Integer             ' Index of the block

Private m_oStructure As CStructure      ' Holds current structure
Private m_oPrimitive As CPrimitive      ' Holds current primitive

Private m_oParent As CParser            ' Object reference to parent

Private m_oValues As CList              ' Current parameter values



Public Sub AddError(ByVal Error As Integer)
    If UBound(Filter(m_iErrors, Error, True, vbBinaryCompare)) = -1 Then
        ' Error not found
        ReDim Preserve m_iErrors(LBound(m_iErrors) To UBound(m_iErrors) + 1)
        m_iErrors(UBound(m_iErrors)) = Error
    End If
End Sub


Public Function ToString() As String
    On Error GoTo ProcedureError
    
    ToString = m_oValues.GetItemAtIndex(1).Value
    
    Exit Function
    
ProcedureError:
    If err.Number = 91 Then
        Resume Next
    Else
        MsgBox err.Number & vbNewLine & err.Description
    End If

End Function



Public Property Let Parent(ByRef NewValue As CParser)
    Set m_oParent = NewValue
End Property


Public Property Get Index() As Integer
    Index = m_iIndex
End Property

Public Property Let Index(ByVal NewValue As Integer)
    m_iIndex = NewValue
End Property


Public Property Get BeginSignature() As String
    BeginSignature = GetFirstLine(m_sText)
End Property

Public Property Get EndSignature() As String
    EndSignature = GetLastLine(m_sText)
End Property


Public Property Get Structure() As CStructure
    If Not m_oStructure Is Nothing Then
        Set Structure = m_oStructure
    End If
End Property

Public Property Let Structure(ByRef NewValue As CStructure)
    Set m_oPrimitive = Nothing
    Set m_oStructure = NewValue
    If Not (m_oStructure Is Nothing) Then
        If Not (m_oStructure.Primitive Is Nothing) Then
            Set m_oValues = ProcessText(GetFirstLine(m_sText))
        End If
    End If
End Property



Public Property Get Primitive() As CPrimitive
    Set Primitive = GetCurrentPrimitive()
End Property

Public Property Let Primitive(ByRef NewValue As CPrimitive)
    
    Set m_oStructure = Nothing
    Set m_oPrimitive = NewValue
    
    If Not (m_oPrimitive Is Nothing) Then
        Set m_oValues = ProcessText(GetFirstLine(m_sText))
    End If
    
End Property


Private Function GetCurrentPrimitive() As CPrimitive

    If Not (m_oStructure Is Nothing) Then
        If Not (m_oStructure.Primitive Is Nothing) Then
            Set GetCurrentPrimitive = m_oStructure.Primitive
        Else
            If Not (m_oPrimitive Is Nothing) Then
                Set GetCurrentPrimitive = m_oPrimitive
            End If
        End If
    Else
        If Not (m_oPrimitive Is Nothing) Then
            Set GetCurrentPrimitive = m_oPrimitive
        End If
    End If

End Function


Public Property Get Values() As CList
    Set Values = m_oValues
End Property

Public Sub SetValue(ByVal Label As String, ByVal Value As String)
    m_oValues(Label).Value = Value

    Dim oPrim As CPrimitive
    
    Set oPrim = GetCurrentPrimitive()

    ' Update code from values
    If Not (oPrim Is Nothing) Then
        
        Dim sFirstLine As String
        Dim lInStr As Long
        
        sFirstLine = ProcessValues(oPrim, m_oValues)
        lInStr = InStr32(1, m_sText, Chr(13), vbBinaryCompare)
        
        If lInStr > 0 Then
            m_sText = Mid$(m_sText, lInStr, (Len(m_sText) - lInStr) + 1)
            m_sText = sFirstLine & m_sText
        Else
            m_sText = sFirstLine
        End If
    
    End If
End Sub


Public Property Get ErrorString() As String
    
    ErrorString = Join(Filter(m_iErrors, "0", False), ", ")
    
End Property


Public Property Get Text() As String
    If Len(Trim$(m_sText)) = 0 Then
        
        If Not (Me.Structure Is Nothing) Then
            If Not (m_oParent Is Nothing) Then
                Dim oToken As CToken
                
                ' Insert beginner token
                For Each oToken In m_oParent.Definition.Tokens
                    If oToken.Structure.StructureName = Me.Structure.StructureName Then
                        If oToken.TokenType = tktBeginner Then
                            'If oToken.BeginsOnNewLine Then
                            '    m_sText = vbNewLine
                            'End If
                            m_sText = m_sText & oToken.Text
                        End If
                    End If
                Next
                
                ' Insert ender token
                For Each oToken In m_oParent.Definition.Tokens
                    If oToken.Structure.StructureName = Me.Structure.StructureName Then
                        If oToken.TokenType = tktEnder Then
                            If oToken.BeginsOnNewLine Then
                                m_sText = m_sText & vbNewLine
                            End If
                            m_sText = m_sText & oToken.Text & vbNewLine
                        End If
                    End If
                Next
            End If
        Else
            m_sText = Me.Primitive.PrimitiveName
        End If
    End If
    Text = m_sText
End Property

Public Property Let Text(ByVal NewValue As String)
    m_sText = NewValue
    
    ReDim m_iErrors(0)
    m_iErrors(0) = "0"
    
    ' User set new text -- convert to values
    If Not (GetCurrentPrimitive() Is Nothing) Then
        Set m_oValues = ProcessText(GetFirstLine(m_sText))
    End If
    
    ' Check for unclosed blocks
    CheckUnclosedBlocks
    
End Property


Private Sub CheckUnclosedBlocks()

    ' Checks to see whether there are any unclosed blocks in the
    ' specified string (i.e. "{" char without accompanying "}" char)
    
    Dim lTotalOpen As Long
    Dim lTotalClose As Long
    
    lTotalOpen = ModString.GetOccurrences(m_sText, "{", vbBinaryCompare)
    lTotalClose = ModString.GetOccurrences(m_sText, "}", vbBinaryCompare)
    
    If lTotalOpen > lTotalClose Then
        AddError eBlockErrors.be_NoBlockEnd
    ElseIf lTotalClose > lTotalOpen Then
        AddError eBlockErrors.be_NoBlockBegin
    End If

End Sub

Private Function ProcessText(ByVal Source As String) As CList

    Dim sValues() As String
    Dim sText As String
    Dim i As Integer
    Dim oList As CList
    Dim oParam As CParameter
    
    Set oList = New CList
    
    If GetCurrentPrimitive() Is Nothing Then Exit Function
    
    sText = Replace$(Source, vbTab, " ")    ' Change tabs to spaces
    sText = Replace$(sText, Chr(13), "")    ' Remove newline characters
    
    sValues = Split(sText, " ")             ' Split values by space
    
    ' Re-size values array to prevent the loop (below) from running off
    ' the upper boundary of the array while looping through parameters
    
    ' NOTE: Array is zero-based, but notice that the UBound is increased
    '       by one. This is to take into account the first token, which
    '       would be the actual name of the structure.
    
    ReDim Preserve sValues(0 To GetCurrentPrimitive().Parameters.Count)
    
    For Each oParam In GetCurrentPrimitive().Parameters
        ' Do a post-increment (this means that the first item is always
        ' ignored. That's because it will be the name of the structure,
        ' and not an actual parameter value.
        i = i + 1
        oList.AddItem oParam.ParameterName, sValues(i)
    Next

    ' Return the generated values list
    Set ProcessText = oList

End Function


Private Function ProcessValues(ByRef Primitive As CPrimitive, ByRef Source As CList) As String

    Dim oParam As CParameter
    Dim sCode As String
    
    For Each oParam In Primitive.Parameters
        oParam.Value = Source(oParam.ParameterName).Value
    Next
    
    sCode = LCase(Primitive.PrimitiveName) & " "
    
    For Each oParam In Primitive.Parameters
        sCode = sCode & oParam.Value & IIf(Len(Trim$(oParam.Value)) = 0, "", " ")
    Next

    ProcessValues = sCode
    
End Function


'Private Function GetDefaultText() As String
'
'    ' Generates default text for the block by appending default begin/
'    ' end tokens into a string.
'
'    Dim sText As String
'    Dim oToken As CToken
'
'    ' Double-check that object references are set
'    If m_oParent Is Nothing Then Exit Function
'    If m_oParent.Definition Is Nothing Then Exit Function
'    If m_oParent.Definition.Tokens Is Nothing Then Exit Function
'
'    ' Insert default begin token (if one exists)
'    For Each oToken In m_oParent.Definition.Tokens
'        If oToken.TokenType = tktBeginner Then
'            If oToken.Structure Is m_oStructure Then
'                If oToken.IsInsideStructure Then
'                    sText = oToken.Text
'                    If oToken.BeginsOnNewLine Then sText = vbNewLine & sText
'                End If
'            End If
'        End If
'    Next oToken
'
'    ' Append default end token (if one exists)
'    For Each oToken In m_oParent.Definition.Tokens
'        If oToken.TokenType = tktEnder Then
'            If oToken.Structure Is m_oStructure Then
'                If oToken.IsInsideStructure Then
'                    If oToken.BeginsOnNewLine Then sText = sText & vbNewLine
'                    sText = sText & oToken.Text
'                End If
'            End If
'        End If
'    Next oToken
'
'    GetDefaultText = sText      ' Return result as string
'
'End Function


Private Sub Class_Initialize()
    Set m_oValues = New CList
    ReDim m_iErrors(0)
    m_iErrors(0) = "0"
End Sub



'
'Public Property Get Errors() As String
'    Errors = Join(m_iErrors, ", ")
'End Property
'

