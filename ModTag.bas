Attribute VB_Name = "ModTag"
' =============================================================================
' Module Name:      ModTag
' Module Type:      Code Module
' Description:      Contains functions for reading/writing tags (such as for
'                   usercontrols) in a standard format. The format is as
'                   follows:
'
'                   [Label]=Value, ...
'
'                   Label is a short ID name to look up the value and Value
'                   is... well... the value
'
' Author(s):        Jonathan A. Conway
' -----------------------------------------------------------------------------
' Log:
'
' 04 07 05 :
'   - Created SetTagAttribute() function & tested/debugged it
' 04 07 04 :
'   - Created the module and the function GetTagAttribute() & tested/debugged
' =============================================================================


Option Explicit



Public Function GetTagAttribute(ByRef Source As String, ByVal Label As String)

    Dim iInstr As Integer
    Dim iNextCommaPos As Integer
    
    ' Get position of label marker
    iInstr = InStr(1, Source, "[" & Label & "]=", vbBinaryCompare)

    If iInstr > 0 Then
        ' Get pos of next delimiter
        iInstr = iInstr + Len(Label) + 3
        iNextCommaPos = InStr(iInstr, Source, ",", vbBinaryCompare)
    
        ' Retrieve and return value, using end-of-string if there's no comma
        If iNextCommaPos > 0 Then
            GetTagAttribute = Mid$(Source, iInstr, iNextCommaPos - iInstr)
        Else
            GetTagAttribute = Mid$(Source, iInstr, (Len(Source) - iInstr) + 1)
        End If
    End If

End Function


Public Function SetTagAttribute(ByRef Source As String, ByVal Label As String, ByVal NewValue As String) As String

    Dim iInstr As Integer
    Dim iBeginPos As Integer
    Dim iEndPos As Integer
    Dim sResult As String
    
    ' Get position of beginning of target tag
    iInstr = InStr(1, Source, "[" & Label & "]=", vbBinaryCompare)
    iBeginPos = iInstr

    If iInstr > 0 Then
        ' Get ending position
        iInstr = iInstr + Len(Label) + 3
        iEndPos = InStr(iInstr, Source, ",", vbBinaryCompare)
        
        ' If not found, specify the end-of-string instead
        If iEndPos <= 0 Then iEndPos = Len(Source) + 1
        
        ' Replace old value with new
        sResult = Replace$(Source, Mid$(Source, iBeginPos, iEndPos - iBeginPos), "[" & Label & "]=" & NewValue, , , vbBinaryCompare)
    Else
        ' Create attribute & append to text
        sResult = Source
        If (Not Right$(Trim$(sResult), 1) = ",") And Len(Trim$(sResult)) > 0 Then
            sResult = sResult & ", "
        End If
        sResult = sResult & "[" & Label & "]=" & NewValue
    End If

    ' Return result
    SetTagAttribute = sResult

End Function
