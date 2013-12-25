VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.UserControl CONPrimitive 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000005&
   ClientHeight    =   4620
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6450
   ControlContainer=   -1  'True
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ScaleHeight     =   308
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   430
   Begin VB.CommandButton cmdParam 
      Caption         =   "..."
      Height          =   300
      Index           =   0
      Left            =   3210
      TabIndex        =   3
      Top             =   795
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.ComboBox cboParam 
      Height          =   315
      Index           =   0
      Left            =   1575
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   1125
      Visible         =   0   'False
      Width           =   1905
   End
   Begin VB.CheckBox chkParam 
      BackColor       =   &H80000005&
      Height          =   300
      Index           =   0
      Left            =   1575
      TabIndex        =   1
      Top             =   525
      Visible         =   0   'False
      Width           =   225
   End
   Begin VB.TextBox txtParam 
      Height          =   300
      Index           =   0
      Left            =   1575
      TabIndex        =   0
      Top             =   795
      Visible         =   0   'False
      Width           =   1695
   End
   Begin RichTextLib.RichTextBox rtbSyntax 
      Height          =   225
      Left            =   1050
      TabIndex        =   5
      Top             =   0
      Width           =   2760
      _ExtentX        =   4868
      _ExtentY        =   397
      _Version        =   393217
      BorderStyle     =   0
      Enabled         =   -1  'True
      HideSelection   =   0   'False
      ReadOnly        =   -1  'True
      Appearance      =   0
      TextRTF         =   $"CONPrimitive.ctx":0000
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label lblSyntax 
      BackStyle       =   0  'Transparent
      Caption         =   "Syntax:"
      Height          =   225
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Width           =   540
   End
   Begin VB.Label lblParam 
      BackStyle       =   0  'Transparent
      Caption         =   "########"
      Height          =   225
      Index           =   0
      Left            =   0
      TabIndex        =   4
      Top             =   825
      Visible         =   0   'False
      Width           =   1485
   End
End
Attribute VB_Name = "CONPrimitive"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' =============================================================================
' Module Name:      CONPrimitive
' Module Type:      User Control
' Description:      Re-useable control that acts as an automated container for
'                   controls used to edit a primitive block. The control
'                   recieves as input a primitive object, as well as input
'                   from the user through child controls. The settings the
'                   user then chooses from the child controls are outputted
'                   from the user control as a Block object of the specified
'                   type of primitive.
' Author(s):        Jonathan A. Conway
' -----------------------------------------------------------------------------
' Log:
'
' 04 08 03 :
'   - Created CONPrimitive and began coding dynamic control instantiation
'     code.
' =============================================================================



Option Explicit


' Public Events
' =============


Public Event StatusChanged(ByVal CurrentStatus As String)


Private m_oBlock As CBlock

Private m_bIsLoading As Boolean

Private m_oFirstControl As Control


Public Property Get Block() As CBlock
    Set Block = m_oBlock
End Property

Public Property Let Block(ByRef NewValue As CBlock)
    Set m_oBlock = NewValue
    LoadBlock
End Property


Public Property Get BackColor() As Long
    BackColor = UserControl.BackColor
End Property

Public Property Let BackColor(ByRef NewValue As Long)
    'UserControl.BackColor = NewValue
    'rtbSyntax.BackColor = NewValue
End Property




Public Sub Refresh()

    LoadBlock

End Sub


Private Function GetParamControlFromIndex(ByVal ParameterIndex As Integer) As Control

    ' Returns a reference to a control (textbox, checkbox or combo) that
    ' matches the specified parameter index.

    Dim cCtl As Control

    For Each cCtl In UserControl.Controls
        If TypeOf cCtl Is TextBox Or _
           TypeOf cCtl Is ComboBox Or _
           TypeOf cCtl Is CheckBox Then

            If cCtl.Tag = CStr(ParameterIndex) Then
                Set GetParamControlFromIndex = cCtl
            End If
        End If
    Next

End Function


Private Sub LoadBlock()
    m_bIsLoading = True
    
    If Not (m_oBlock.Primitive Is Nothing) Then
        LoadPrimitive m_oBlock.Primitive
    End If

    m_bIsLoading = False
End Sub

Private Sub LoadPrimitive(ByRef Source As CPrimitive)

    Dim oParam As CParameter
    Dim iCurrentTop As Integer
    Dim bFirstTime As Boolean       ' Is this the first time round?
    Dim cCtl As Control

    ClearControls                   ' Clear out controls from old primitive

    With rtbSyntax
        .Text = Source.Syntax
        .SelStart = 0
        .SelLength = Len(Source.Syntax)
        .SelBold = False
        .SelItalic = False
        .SelUnderline = False
        .SelColor = vbBlack
    End With

    bFirstTime = True
    iCurrentTop = 25

    For Each oParam In Source.Parameters

        Load lblParam(lblParam.UBound + 1)  ' Load labels
        With lblParam(lblParam.UBound)
            .Left = 0
            .Top = iCurrentTop
            .Caption = oParam.ParameterName & ":"
            .Tag = oParam.Index
            .Visible = True
        End With

        If bFirstTime Then
            ' Load controls for the *first* parameter AND get reference to it
            Set cCtl = LoadControls(oParam, iCurrentTop)
            bFirstTime = False
        Else
            ' Load parameter controls
            LoadControls oParam, iCurrentTop
        End If

        iCurrentTop = iCurrentTop + 25      ' Increment top variable

    Next

    PositionParameterControls   ' Position/size all controls

    Set m_oFirstControl = cCtl  ' Cache reference to first visible control

End Sub

Private Function GetListIndexFromCombo(ByRef Source As ComboBox, ByVal TextValue As String) As Long

    GetListIndexFromCombo = -1
    
    If Source.ListCount > 0 Then
        Dim i As Integer
        For i = 0 To Source.ListCount - 1
            If Source.List(i) = TextValue Then
                GetListIndexFromCombo = i
                Exit Function
            End If
        Next
    End If

End Function

Private Function LoadControls(ByRef Parameter As CParameter, ByVal Top As Integer) As Control
    Dim oCtl As Control
    Dim sValue As String
    
    sValue = m_oBlock.Values(Parameter.ParameterName).Value
    
    Select Case Parameter.DataType
        Case eParameterTypes.ptString To eParameterTypes.ptNumber
            Load txtParam(txtParam.UBound + 1)
            txtParam(txtParam.UBound).Text = sValue
            Set oCtl = txtParam(txtParam.UBound)
        
        Case eParameterTypes.ptYesNo
            Load chkParam(chkParam.UBound + 1)
            chkParam(chkParam.UBound).Value = BooleanToChecked(sValue = GetTagAttribute(Parameter.Data, "yes"))
            Set oCtl = chkParam(chkParam.UBound)
        
        Case eParameterTypes.ptChoice
            Load cboParam(cboParam.UBound + 1)
            Dim oItem As CListItem
            For Each oItem In Parameter.List
                cboParam(cboParam.UBound).AddItem oItem.Label
            Next
            If Parameter.List.Count > 0 Then
                Set oItem = Parameter.List.FindValue(sValue)
                If oItem Is Nothing Then
                    Set oItem = Parameter.List.GetItemAtIndex(1)
                End If
                cboParam(cboParam.UBound).ListIndex = GetListIndexFromCombo(cboParam(cboParam.UBound), oItem.Label)
            End If
            Set oCtl = cboParam(cboParam.UBound)
    End Select

    ' Set generic attributes (apply to all types of controls)
    With oCtl
        .Top = Top
        .Tag = Parameter.Index
        .Visible = True
    End With
    
    ' Add a builder if needed
    If Parameter.Builder > -1 Then
        Load cmdParam(cmdParam.UBound + 1)
        With cmdParam(cmdParam.UBound)
            .Top = Top
            .Tag = Parameter.Index
            .Visible = True
        End With
    End If

    Set LoadControls = oCtl

End Function


Private Sub ClearControls()

    Dim i As Integer

    For i = lblParam.lbound + 1 To lblParam.UBound
        Unload lblParam(i)
    Next
    For i = txtParam.lbound + 1 To txtParam.UBound
        Unload txtParam(i)
    Next
    For i = chkParam.lbound + 1 To chkParam.UBound
        Unload chkParam(i)
    Next
    For i = cboParam.lbound + 1 To cboParam.UBound
        Unload cboParam(i)
    Next
    For i = cmdParam.lbound + 1 To cmdParam.UBound
        Unload cmdParam(i)
    Next

End Sub



Private Sub PositionParameterControls()

    On Error Resume Next

    Dim i As Integer
    
    For i = txtParam.lbound + 1 To txtParam.UBound
        txtParam(i).Width = (UserControl.ScaleWidth - txtParam(i).Left) - 5
    Next
    For i = chkParam.lbound + 1 To chkParam.UBound
        chkParam(i).Width = (UserControl.ScaleWidth - chkParam(i).Left) - 5
    Next
    For i = cboParam.lbound + 1 To cboParam.UBound
        cboParam(i).Width = (UserControl.ScaleWidth - cboParam(i).Left) - 5
    Next
    For i = cmdParam.lbound + 1 To cmdParam.UBound
        cmdParam(i).Left = (UserControl.ScaleWidth - cmdParam(i).Width) - 5
        With GetParamControlFromIndex(cmdParam(i).Tag)
            .Width = (.Width - cmdParam(i).Width) - 5
        End With
    Next

End Sub

Private Sub HighlightLabel(ByVal ParameterIndex As Integer)
    Dim i As Integer
    For i = lblParam.lbound To lblParam.UBound
        'lblParam(i).FontBold = (lblParam(i).Tag = CStr(ParameterIndex))
        With lblParam(i)
            If .Tag = CStr(ParameterIndex) Then
                .ForeColor = vbRed
                .FontUnderline = True
            Else
                .ForeColor = vbWindowText
                .FontUnderline = False
            End If
        End With
    Next

    FormatSyntaxBox ParameterIndex, m_oBlock.Primitive, rtbSyntax
End Sub

Private Sub ShowDescription(ByVal Index As Integer)
    If Index = -1 Then
        RaiseEvent StatusChanged("")
    Else
        With m_oBlock.Primitive.Parameters(Index)
            RaiseEvent StatusChanged(.Description)
        End With
    End If
End Sub

Private Sub FormatSyntaxBox(ByVal SelectedParameterIndex As Long, ByRef Primitive As CPrimitive, ByRef Target As RichTextBox)

    Dim iInStr As Integer

    Dim oParam As CParameter

    Dim iLen As Integer

    Target.Visible = False

    With Target
        ' Format each parameter
        For Each oParam In Primitive.Parameters
            iInStr = InStr(1, Primitive.Syntax, " " & oParam.ToString(), vbBinaryCompare) - 1
            If iInStr <> 0 Then
                .Find " " & oParam.ToString(), , , rtfMatchCase
                iLen = .SelLength
                .SelStart = .SelStart + 1
                .SelLength = iLen - 1
                .SelItalic = True

                '.SelBold = (oParam.Index = SelectedParameterIndex)
                .SelUnderline = (oParam.Index = SelectedParameterIndex)

                .SelColor = IIf(oParam.Index = SelectedParameterIndex, vbRed, vbBlack)
                .SelStart = 0
            End If
        Next

        ' Format primitive name
        iInStr = InStr(1, Primitive.Syntax, Primitive.PrimitiveName, vbBinaryCompare)
        If iInStr <> 0 Then
            .Find Primitive.PrimitiveName, , , rtfMatchCase
            .SelItalic = False
            .SelBold = True
            .SelColor = vbBlack
        End If

        .SelStart = 0
    End With

    Target.Visible = True

End Sub

Private Sub cmdParam_Click(Index As Integer)
    
    With m_oBlock.Primitive.Parameters(cmdParam(Index).Tag)
        
        Load FrmBuilder
        ' Set current data
        FrmBuilder.Data = m_oBlock.Values(.ParameterName).Value
        FrmBuilder.Builder = .Builder
            
        ' Make sure there was no problem with builder
        If Not (FrmBuilder.Error) Then
            
            FrmBuilder.Caption = "Building " & m_oBlock.Primitive.PrimitiveName & "." & .ParameterName
            FrmBuilder.Show vbModal   ' Show form modally
                
            If Not (FrmBuilder.IsCancelled) Then
                ' Grab altered data from form
                m_oBlock.SetValue .ParameterName, FrmBuilder.Data
                
                ' Update controls with value
                Dim sValue As String
                Dim iIndex As Integer
                
                sValue = m_oBlock.Values(.ParameterName).Value
                iIndex = cmdParam(Index).Tag
                
                Select Case .DataType
                    Case eParameterTypes.ptString To eParameterTypes.ptNumber
                        GetParamControlFromIndex(iIndex).Text = sValue
                        'txtParam(Index).Text = sValue
                    Case eParameterTypes.ptYesNo
                        GetParamControlFromIndex(iIndex).Value = IIf(GetTagAttribute(.Data, "yes") Like sValue, vbChecked, vbUnchecked)
                        'chkParam(Index).Value = IIf(GetTagAttribute(.Data, "yes") Like sValue, vbChecked, vbUnchecked)
                    Case eParameterTypes.ptChoice
                        GetParamControlFromIndex(iIndex).Text = sValue
                        'cboParam(Index).Text = sValue
                End Select
            End If
    
        End If
    
        Unload FrmBuilder
    
    End With
    
    
End Sub

Private Sub rtbSyntax_GotFocus()
    On Error GoTo ProcedureError
    
    If Not (m_oFirstControl Is Nothing) Then
        m_oFirstControl.SetFocus    ' Set focus to first visible control
    End If
    Exit Sub
    
ProcedureError:
    If err.Number = 5 Then
    Else
        MsgBox err.Number & vbNewLine & err.Description
    End If
End Sub

Private Sub txtParam_GotFocus(Index As Integer)
    HighlightLabel txtParam(Index).Tag
    ShowDescription txtParam(Index).Tag
End Sub


Public Sub FocusFirstControl()
    If Not (m_oFirstControl Is Nothing) Then
        m_oFirstControl.SetFocus    ' Set focus to first visible control
    End If
End Sub


Private Sub chkParam_GotFocus(Index As Integer)
    HighlightLabel chkParam(Index).Tag
    ShowDescription chkParam(Index).Tag
End Sub

Private Sub cboParam_GotFocus(Index As Integer)
    HighlightLabel cboParam(Index).Tag
    ShowDescription cboParam(Index).Tag
End Sub

Private Sub cmdParam_GotFocus(Index As Integer)
    HighlightLabel cmdParam(Index).Tag
    ShowDescription cmdParam(Index).Tag
End Sub


Private Sub txtParam_Change(Index As Integer)
    If Not m_bIsLoading Then
        With m_oBlock.Primitive.Parameters(txtParam(Index).Tag)
            If .DataType = eParameterTypes.ptNumber Then _
                txtParam(Index) = FixLong(txtParam(Index).Text)
            m_oBlock.SetValue .ParameterName, txtParam(Index).Text
        End With
    End If
End Sub

Private Sub cboParam_Click(Index As Integer)
    'Dim oParam As CParameter
    'oparam.List(
    
    If Not m_bIsLoading Then
        With m_oBlock.Primitive.Parameters(cboParam(Index).Tag)
            m_oBlock.SetValue .ParameterName, .List(cboParam(Index).Text).Value
        End With
    End If
End Sub

Private Sub chkParam_Click(Index As Integer)
    If Not m_bIsLoading Then
        With m_oBlock.Primitive.Parameters(chkParam(Index).Tag)
            m_oBlock.SetValue .ParameterName, _
                GetTagAttribute(.Data, _
                IIf(chkParam(Index).Value = vbChecked, "yes", "no"))
        End With
    End If
End Sub


Private Sub UserControl_GotFocus()
    FocusFirstControl
End Sub

Private Sub UserControl_Initialize()
    On Error Resume Next
    If Settings.ReadSetting("General_CompatibleLook") = "yes" Then
        UserControl.BackColor = vb3DFace
        'lblSyntax.BackColor = vb3DFace
        'lblParam(0).BackColor = vb3DFace
        'rtbSyntax.Appearance = rtfThreeD
        rtbSyntax.BackColor = vb3DFace
        chkParam(0).BackColor = vb3DFace
    End If
End Sub

Private Sub UserControl_Resize()
    On Error Resume Next
    rtbSyntax.Width = UserControl.ScaleWidth - rtbSyntax.Left
End Sub


'Private Sub lblParam_Click(Index As Integer)
'    HighlightLabel lblParam(Index).Tag
'    ShowDescription lblParam(Index).Tag
'
'    With m_oBlock.Primitive.Parameters(lblParam(Index).Tag)
'        Select Case .DataType
'            Case eParameterTypes.ptString To eParameterTypes.ptNumber
'                txtParam(Index).SetFocus
'            Case eParameterTypes.ptYesNo
'                chkParam(Index).SetFocus
'            Case eParameterTypes.ptChoice
'                cboParam(Index).SetFocus
'        End Select
'    End With
'End Sub

