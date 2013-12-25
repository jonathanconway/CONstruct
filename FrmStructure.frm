VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Begin VB.Form FrmStructure 
   Caption         =   "Add Structure"
   ClientHeight    =   5895
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5640
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FrmStructure.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   393
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   376
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin VB.PictureBox picStructure 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   3765
      Left            =   150
      ScaleHeight     =   251
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   356
      TabIndex        =   9
      Top             =   1200
      Width           =   5340
      Begin VB.ComboBox cboObjectType 
         Height          =   315
         Left            =   1050
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Top             =   0
         Width           =   2190
      End
      Begin CONstruct.CONPrimitive cp 
         Height          =   3165
         Left            =   0
         TabIndex        =   12
         Top             =   375
         Width           =   5265
         _ExtentX        =   9287
         _ExtentY        =   5583
      End
      Begin VB.Label lblObjectType 
         BackStyle       =   0  'Transparent
         Caption         =   "Object Type:"
         Height          =   240
         Left            =   0
         TabIndex        =   11
         Top             =   30
         Width           =   1065
      End
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "&Help"
      CausesValidation=   0   'False
      Height          =   375
      Left            =   4320
      TabIndex        =   7
      Top             =   5070
      Width           =   1215
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      CausesValidation=   0   'False
      Height          =   375
      Left            =   3000
      TabIndex        =   6
      Top             =   5070
      Width           =   1215
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Save"
      Height          =   375
      Left            =   1680
      TabIndex        =   5
      Top             =   5070
      Width           =   1215
   End
   Begin ComctlLib.TabStrip tabMain 
      Height          =   4170
      Left            =   75
      TabIndex        =   4
      Top             =   825
      Width           =   5490
      _ExtentX        =   9684
      _ExtentY        =   7355
      _Version        =   327682
      BeginProperty Tabs {0713E432-850A-101B-AFC0-4210102A8DA7} 
         NumTabs         =   2
         BeginProperty Tab1 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "&Structure"
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "&Code"
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
   Begin ComctlLib.StatusBar barStatus 
      Align           =   2  'Align Bottom
      Height          =   330
      Left            =   0
      TabIndex        =   3
      Top             =   5565
      Width           =   5640
      _ExtentX        =   9948
      _ExtentY        =   582
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   1
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            AutoSize        =   1
            Object.Width           =   9446
            Object.Tag             =   ""
         EndProperty
      EndProperty
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
   Begin VB.TextBox txtCode 
      Height          =   3540
      Left            =   150
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   8
      Top             =   1200
      Width           =   5265
   End
   Begin VB.Image imgErrors 
      Height          =   240
      Left            =   3360
      Picture         =   "FrmStructure.frx":014A
      Top             =   465
      Width           =   240
   End
   Begin VB.Label lblErrors 
      BackStyle       =   0  'Transparent
      Caption         =   "Errors Found"
      Height          =   225
      Left            =   3675
      TabIndex        =   13
      Top             =   480
      Width           =   1800
   End
   Begin VB.Image imgIcons 
      Height          =   720
      Index           =   3
      Left            =   1260
      Picture         =   "FrmStructure.frx":06D4
      Top             =   4725
      Visible         =   0   'False
      Width           =   720
   End
   Begin VB.Image imgIcons 
      Height          =   720
      Index           =   2
      Left            =   840
      Picture         =   "FrmStructure.frx":159E
      Top             =   4725
      Visible         =   0   'False
      Width           =   720
   End
   Begin VB.Image imgIcons 
      Height          =   720
      Index           =   1
      Left            =   420
      Picture         =   "FrmStructure.frx":2468
      Top             =   4725
      Visible         =   0   'False
      Width           =   720
   End
   Begin VB.Image imgIcons 
      Height          =   720
      Index           =   0
      Left            =   105
      Picture         =   "FrmStructure.frx":44AA
      Top             =   4725
      Visible         =   0   'False
      Width           =   720
   End
   Begin ComctlLib.ImageList iml 
      Left            =   525
      Top             =   5175
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   48
      ImageHeight     =   48
      MaskColor       =   8421376
      UseMaskColor    =   0   'False
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   1
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmStructure.frx":A8FC
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Label lblType 
      BackStyle       =   0  'Transparent
      Caption         =   "Define"
      Height          =   255
      Left            =   960
      TabIndex        =   2
      Top             =   240
      Width           =   4575
   End
   Begin VB.Label lblInfo 
      BackStyle       =   0  'Transparent
      Caption         =   "Length: 321; Lines: 3"
      Height          =   255
      Left            =   945
      TabIndex        =   1
      Top             =   480
      Width           =   2475
   End
   Begin VB.Label lblName 
      BackStyle       =   0  'Transparent
      Caption         =   "[Empty]"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   960
      TabIndex        =   0
      Top             =   0
      Width           =   4575
   End
   Begin VB.Image imgHead 
      Height          =   720
      Left            =   120
      Top             =   0
      Width           =   720
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H80000010&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   735
      Left            =   0
      Top             =   0
      Width           =   5655
   End
End
Attribute VB_Name = "FrmStructure"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' =============================================================================
' Module Name:      FrmStructure
' Module Type:      User Form
' Description:      Form for creating/editing structures in a script.
' Author(s):        Jonathan A. Conway
' -----------------------------------------------------------------------------
'
'
' =============================================================================



Option Explicit



' Private Variables
' =================


Private m_oBlock As CBlock

Private m_bIsCancelled As Boolean
Private m_oOldBlock As CBlock

Private m_bIsAdding As Boolean
Private m_bIsBusy As Boolean


' Public Properties
' =================


Public Property Let IsAdding(ByVal NewValue As Boolean)
    m_bIsAdding = NewValue
    If NewValue Then
        cp.Block = New CBlock
        cp.Block.Structure = FrmMain.ceCONs(FrmMain.CurrentCON).Parser.Definition.Structures(1)
        Me.Block = cp.Block
    End If
End Property

Public Property Get Block() As CBlock
    Set Block = m_oBlock
End Property

Public Property Let Block(ByVal NewValue As CBlock)
    Set m_oBlock = NewValue
    cp.Block = m_oBlock
    txtCode.Text = m_oBlock.Text
    UpdateBlock
    
    ' Cache old block for use when updating code
    CopyBlock NewValue
End Property

Public Property Get IsCancelled() As Boolean
    IsCancelled = m_bIsCancelled
End Property



' Private Methods
' ===============

Private Sub UpdateBlock()
    m_bIsBusy = True
    
    Me.Caption = m_oBlock.ToString() & " (" & m_oBlock.Structure.StructureName & ")"
    cboObjectType.Text = m_oBlock.Structure.StructureName
    lblName.Caption = m_oBlock.ToString()
    lblType.Caption = m_oBlock.Structure.StructureName()
    lblInfo.Caption = "Length: " & Len(m_oBlock.Text) & _
                      ", Lines: " & ModString.GetLineCount(m_oBlock.Text)
    
    Dim bError As Boolean
    bError = (Len(Trim$(m_oBlock.ErrorString)) > 0)
    imgErrors.Visible = bError
    lblErrors.Visible = bError
    If bError Then
        lblErrors.Caption = "Errors found" '& m_oBlock.ErrorString
    End If
    
    imgHead.Picture = imgIcons(m_oBlock.Structure.ImageID - 1).Picture

    m_bIsBusy = False
End Sub


Private Sub CopyBlock(ByRef Source As CBlock)
    ' Make a physical backup copy of the block.
    
    ' When it is necessary to search source code for the code of the
    ' original block so as to replace it, use this copy.
    
    Set m_oOldBlock = New CBlock
    m_oOldBlock.Text = Source.Text
End Sub


Private Sub LoadTypeCombo()
    ' Populate the structure-type combo box with values
    
    Dim oStruct As CStructure
    
    For Each oStruct In FrmMain.ceCONs(FrmMain.CurrentCON).Parser.Definition.Structures
        cboObjectType.AddItem oStruct.StructureName
    Next
End Sub


Private Sub ShowErrors()

    Load FrmStructureErrors
    With FrmStructureErrors
        Dim sErrors() As String
        Dim iError As Integer
        Dim i As Integer
        Dim sText As String
        
        sErrors = Split(m_oBlock.ErrorString, ",")
        
        For i = LBound(sErrors) To UBound(sErrors)
            iError = FixInteger(sErrors(i))
            If iError > 0 Then
                Select Case iError
                    Case 1  ' No end marker
                        sText = sText & _
                        "Error 1: Missing End Marker" & vbNewLine & _
                        "The ending marker for the structure was not found." & vbNewLine & vbNewLine
                    Case 2  ' No block begin
                        sText = sText & _
                        "Error 2: Missing Block Begin" & vbNewLine & _
                        "A required opening brace ""{"" was not found." & vbNewLine & vbNewLine
                    Case 3  ' No block end
                        sText = sText & _
                        "Error 2: Missing Block End" & vbNewLine & _
                        "A required closing brace ""}"" was not found." & vbNewLine & vbNewLine
                End Select
            End If
        Next
        
        .txtErrors.Text = sText
        
        .Show vbModal
    End With
    Unload FrmStructureErrors

End Sub



' Event Handlers
' ==============


Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdSave_Click()
    m_bIsCancelled = False
    
    'tabMain.Tabs(1).Selected = True
    
    If Not m_bIsAdding Then
        ' Do this only if we're *not* adding a new item
        FrmMain.ceCONs(FrmMain.CurrentCON).UpdateBlock m_oOldBlock, cp.Block
    End If
    
    Me.Hide
End Sub



Private Sub cp_StatusChanged(ByVal CurrentStatus As String)
    barStatus.Panels(1).Text = CurrentStatus
End Sub



Private Sub Form_Load()
    'Set imgHead.Picture = LoadPicture("F:\My Documents\My Source\Construct\Define_Large.ico")
    m_bIsCancelled = True

    LoadTypeCombo
    SetCompatibleColours Me
End Sub

Private Sub Form_Resize()

    On Error GoTo ProcedureError
    
    ' Top header panel & labels
    Shape1.Width = Me.ScaleWidth + 1
    lblName.Width = Me.ScaleWidth
    lblType.Width = Me.ScaleWidth
    lblInfo.Width = Me.ScaleWidth

    ' Tab control
    tabMain.Width = Me.ScaleWidth - 10
    tabMain.Height = Me.ScaleHeight - 115
    
    ' Client area -- primitive structure & code textbox
    With picStructure
        .Width = tabMain.Width - 10
        .Height = tabMain.Height - 31
    End With
    cp.Width = picStructure.ScaleWidth
    cp.Height = picStructure.ScaleHeight - cp.Top
    
    txtCode.Width = picStructure.Width
    txtCode.Height = picStructure.Height
    
    ' Command buttons
    cmdHelp.Top = tabMain.Top + tabMain.Height + 5
    cmdCancel.Top = cmdHelp.Top
    cmdSave.Top = cmdHelp.Top
    cmdHelp.Left = Me.ScaleWidth - 88
    cmdCancel.Left = cmdHelp.Left - 88
    cmdSave.Left = cmdCancel.Left - 88

    Exit Sub

ProcedureError:
    If err.Number = 380 Then
    Else
        MsgBox err.Number & vbNewLine & err.Description
    End If

End Sub


Private Sub imgErrors_Click()
    ShowErrors
End Sub

Private Sub lblErrors_Click()
    ShowErrors
End Sub

Private Sub tabMain_Click()
    Select Case tabMain.SelectedItem.Index
        Case 1      ' Structure tab
            With cp
                .Block.Text = txtCode.Text
                .Refresh
                .SetFocus
                .FocusFirstControl
            End With
            UpdateBlock
            picStructure.ZOrder vbBringToFront
            
        Case 2      ' Code tab
            With txtCode
                .Text = cp.Block.Text
                .SetFocus
                .ZOrder vbBringToFront
            End With
    
    End Select
End Sub

Private Sub tabMain_GotFocus()
    tabMain_Click
End Sub

Private Sub cboObjectType_Click()
    
    If Me.Visible = True Then
        If Not m_bIsBusy Then
            cp.Block.Structure = FrmMain.ceCONs(FrmMain.CurrentCON).Parser.Definition.Structures.FindItem(cboObjectType.Text)
            With cp
                '.Block.Text = txtCode.Text
                .Refresh
                .SetFocus
                .FocusFirstControl
            End With
            UpdateBlock
        End If
    End If

End Sub




'Private Sub cp_StructureChanged(StructureName As String)
'    cp.ChangeStructure FrmMain.ceCONs(FrmMain.CurrentCON).Parser.Definition.Structures.FindItem(StructureName)
'End Sub

'
'Private Sub DumpStuff()
'
'    Dim cctl As Control
'
'    For Each cctl In Me.Controls
'
'        If (TypeOf cctl Is CommandButton) Or (TypeOf cctl Is tabstrip) Then
'
'            Debug.Print cctl.Name & "  ::  L:" & cctl.Left & " T:" & cctl.Top & " W:" & cctl.Width & " H:" & cctl.Height
'
'        End If
'
'    Next
'
'End Sub
