VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CSnippet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit



Private m_sText As String
Private m_sLabel As String
Private m_sCategory As String
Private m_sFilename As String
Private m_lIndex As Long

Private m_oFields As CObjectCollection



Public Sub LoadFile(ByVal Filename As String)

    Dim lHnd As Long
    Dim sText As String
    
    lHnd = FreeFile()
    
    Open Filename For Input As #lHnd
    sText = Input(LOF(lHnd), #lHnd)    ' Read file contents
    Close #lHnd

    If Len(Trim$(sText)) > 0 Then
        LoadHeader sText
        LoadFields sText
    End If
    
    m_sText = CleanComments(sText)

    m_sFilename = Filename
    
End Sub


Public Sub SaveFile(ByVal Filename As String)

    Dim lHnd As Long
    Dim sText As String
    
    lHnd = FreeFile()
    
    sText = AddComments(m_sText)
    sText = Mid$(sText, 1, GetEndWithoutLineBreaks(sText))
    
    Open Filename For Output As #lHnd
    Print #lHnd, sText          ' Write file contents
    Close #lHnd
    
    m_sFilename = Filename
    
End Sub

Private Function CleanComments(ByVal Source As String) As String

    Dim lHBegin As Long
    Dim lHEnd As Long
    Dim sText As String
    Dim sComment As String
    Dim lFBegin As Long
    Dim lFEnd As Long
    Dim bIsDone As Boolean
    Dim lPos As Long
    
    sText = Source
    
    ' Remove header
    
    lHBegin = InStr32(1, sText, "/** snippet", vbTextCompare)
    If lHBegin > 0 Then
        lHEnd = InStr32(lHBegin, sText, "*/", vbBinaryCompare)
        If lHEnd > 0 Then
            sComment = Trim$(Mid$(sText, lHBegin, (lHEnd - lHBegin) + 4))
            sText = Replace$(sText, sComment, "")
        End If
    End If

    
    ' Remove fields
    
    'lPos = 1
    bIsDone = False
    
    Do Until bIsDone
        lFBegin = InStr32(1, sText, "/** field", vbTextCompare)
        If lFBegin > 0 Then
            lFEnd = InStr32(lFBegin, sText, "*/", vbBinaryCompare)
            If lFEnd > 0 Then
                sComment = Trim$(Mid$(sText, lFBegin, (lFEnd - lFBegin) + 4))
                sText = Replace$(sText, sComment, "")
            Else
                bIsDone = True
            End If
        Else
            bIsDone = True
        End If
    Loop

    CleanComments = sText

End Function

Private Function AddComments(ByVal Source As String) As String

    Dim sText As String
    
    sText = "/** snippet [label]=" & m_sLabel & ", [category]=" & m_sCategory & " */" & vbNewLine
    
    Dim oField As CSnippetField
    For Each oField In m_oFields
        sText = sText & "/** field [fieldname]=" & oField.FieldName & ", [datatype]=" & oField.DataType & " */" & vbNewLine
    Next
    
    sText = sText & Source
    AddComments = sText

End Function


Private Sub LoadHeader(ByRef Source As String)

    Dim lHBegin As Long
    Dim lHEnd As Long
    Dim sHeader As String
    
    lHBegin = InStr32(1, Source, "/** snippet", vbTextCompare)
    If lHBegin > 0 Then
        lHEnd = InStr32(lHBegin, Source, "*/", vbBinaryCompare)
        
        If lHEnd > 0 Then
        
            sHeader = Trim$(Mid$(Source, lHBegin + 11, (lHEnd - lHBegin) - 11))
            
            m_sLabel = ModTag.GetTagAttribute(sHeader, "label")
            m_sCategory = ModTag.GetTagAttribute(sHeader, "category")
        
        End If
    
    End If

End Sub

Private Sub LoadFields(ByRef Source As String)

    Dim lFBegin As Long
    Dim lFEnd As Long
    Dim bIsDone As Boolean
    Dim sField As String
    Dim lPos As Long
    
    Dim oField As CSnippetField
    
    m_oFields.Clear
    
    lPos = 1
    bIsDone = False
    
    Do Until bIsDone
        lFBegin = InStr32(lPos, Source, "/** field", vbTextCompare)
        If lFBegin > 0 Then
            lFEnd = InStr32(lFBegin, Source, "*/", vbBinaryCompare)
            If lFEnd > 0 Then
                sField = Trim$(Mid$(Source, lFBegin + 9, (lFEnd - lFBegin) - 9))

                Set oField = New CSnippetField
                oField.FieldName = ModTag.GetTagAttribute(sField, "fieldname")
                oField.DataType = FixInteger(ModTag.GetTagAttribute(sField, "datatype"))
                m_oFields.Add oField
                
                lPos = lFEnd
            Else
                lPos = lFBegin + 9
            End If
        Else
            bIsDone = True
        End If
    Loop

End Sub



Public Property Get Index() As Long
    Index = m_lIndex
End Property

Public Property Let Index(ByVal NewValue As Long)
    m_lIndex = NewValue
End Property


Public Property Get Fields() As CObjectCollection
    Set Fields = m_oFields
End Property


Public Property Get Filename() As String
    Filename = m_sFilename
End Property



Public Property Get Text() As String
    Text = m_sText
End Property

Public Property Let Text(ByVal NewValue As String)
    m_sText = NewValue
End Property



Public Property Get Label() As String
    Label = m_sLabel
End Property

Public Property Let Label(ByVal NewValue As String)
    m_sLabel = NewValue
End Property



Public Property Get Category() As String
    Category = m_sCategory
End Property

Public Property Let Category(ByVal NewValue As String)
    m_sCategory = NewValue
End Property


Private Sub Class_Initialize()
    Set m_oFields = New CObjectCollection
End Sub

Private Sub Class_Terminate()
    Set m_oFields = Nothing
End Sub


Public Function ToString() As String
    ToString = m_sLabel
End Function
