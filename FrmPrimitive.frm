VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Begin VB.Form FrmPrimitive 
   Caption         =   "Insert Primitive"
   ClientHeight    =   4290
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6855
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FrmPrimitive.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   286
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   457
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picTabs 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1950
      Index           =   0
      Left            =   2655
      ScaleHeight     =   130
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   261
      TabIndex        =   12
      Top             =   1470
      Width           =   3915
      Begin CONstruct.CONPrimitive cp 
         Height          =   1905
         Left            =   45
         TabIndex        =   13
         Top             =   45
         Width           =   3840
         _ExtentX        =   7064
         _ExtentY        =   4101
      End
   End
   Begin ComctlLib.TabStrip tabMain 
      Height          =   2325
      Left            =   2625
      TabIndex        =   11
      Top             =   1155
      Width           =   4005
      _ExtentX        =   7064
      _ExtentY        =   4101
      _Version        =   327682
      BeginProperty Tabs {0713E432-850A-101B-AFC0-4210102A8DA7} 
         NumTabs         =   2
         BeginProperty Tab1 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Primitive"
            Key             =   "Primitive"
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Code"
            Key             =   "Code"
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
   Begin ComctlLib.StatusBar barStatus 
      Align           =   2  'Align Bottom
      Height          =   315
      Left            =   0
      TabIndex        =   9
      Top             =   3975
      Width           =   6855
      _ExtentX        =   12091
      _ExtentY        =   556
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   1
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            AutoSize        =   1
            Object.Visible         =   0   'False
            Object.Width           =   12674
            Object.Tag             =   ""
         EndProperty
      EndProperty
   End
   Begin ComctlLib.TreeView tvwPrimitives 
      Height          =   3090
      Left            =   75
      TabIndex        =   8
      Top             =   405
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   5450
      _Version        =   327682
      HideSelection   =   0   'False
      Indentation     =   397
      LabelEdit       =   1
      LineStyle       =   1
      Style           =   7
      ImageList       =   "imlTree"
      Appearance      =   1
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   5250
      TabIndex        =   4
      Top             =   3570
      Width           =   1380
   End
   Begin VB.CommandButton cmdInsert 
      Caption         =   "Insert"
      Default         =   -1  'True
      Height          =   375
      Left            =   3780
      TabIndex        =   3
      Top             =   3570
      Width           =   1380
   End
   Begin VB.PictureBox picTabs 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1950
      Index           =   1
      Left            =   2655
      ScaleHeight     =   130
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   261
      TabIndex        =   5
      Top             =   1470
      Width           =   3915
      Begin VB.TextBox txtCode 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1575
         Left            =   45
         TabIndex        =   6
         Top             =   315
         Width           =   3825
      End
      Begin VB.Label lblCode 
         BackStyle       =   0  'Transparent
         Caption         =   "Code Preview:"
         Height          =   225
         Left            =   45
         TabIndex        =   7
         Top             =   45
         Width           =   1065
      End
   End
   Begin VB.Label lblNotes 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "#####################"
      Height          =   195
      Left            =   3360
      TabIndex        =   10
      Top             =   840
      Width           =   3465
   End
   Begin ComctlLib.ImageList imlTree 
      Left            =   1200
      Top             =   3300
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   16711935
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   3
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmPrimitive.frx":06C2
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmPrimitive.frx":0A14
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmPrimitive.frx":0D66
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Image imgCat 
      Height          =   720
      Left            =   2550
      Picture         =   "FrmPrimitive.frx":10B8
      Top             =   375
      Width           =   720
   End
   Begin VB.Label Label1 
      Caption         =   "Primitives:"
      Height          =   225
      Left            =   105
      TabIndex        =   2
      Top             =   105
      Width           =   1170
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000014&
      X1              =   175
      X2              =   448
      Y1              =   73
      Y2              =   73
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000010&
      X1              =   175
      X2              =   448
      Y1              =   72
      Y2              =   72
   End
   Begin VB.Label lblDescription 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "#####################"
      Height          =   195
      Left            =   3360
      TabIndex        =   1
      Top             =   630
      Width           =   3465
   End
   Begin VB.Label lblPrimitive 
      BackStyle       =   0  'Transparent
      Caption         =   "#################"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   3360
      TabIndex        =   0
      Top             =   420
      Width           =   3435
   End
   Begin VB.Image imgPrim 
      Height          =   720
      Left            =   2550
      Picture         =   "FrmPrimitive.frx":2DB2
      Top             =   375
      Width           =   720
   End
End
Attribute VB_Name = "FrmPrimitive"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' =============================================================================
' Module Name:      FrmPrimitive
' Module Type:      User Form
' Description:      Modal "helper" form for cutting primitive code and
'                   inserting into a file. Dynamically creates controls for
'                   allowing the user to enter in parameter values. Relies
'                   on the Primitives classes:
'                       - CDefinition       - CPrimitives
'                       - CPrimitive        - CParameters
'                       - CParameter
' Author(s):        Jonathan A. Conway
' -----------------------------------------------------------------------------
' Log:
'
' 04 07 16 :
'   - Fixed "default values" bug
'   - Some "eyecandy" changes like adding a folder image
'   - Added description text per primitive.parameter via ShowDescription()
'   - Some other bugfixes that I've lost track of
'
' 04 07 14 :
'   - Major modifications to the visual layout of the form
'   - Coded functions for handling dynamic parameter controls as well as
'     event procedures
'   - Dynamic resizing code
'   - Fixed *heaps* of bugs along the way
'
' 04 05 05 :
'   - Fixed bugs in Find/Replace/Goto
'   - Added method for launching Duke Nukem 3D
'   - Added method for showing the Options form
'   - Misc. mods and fixes.
' =============================================================================

Option Explicit



' Private Enumerations
' ====================



' Private Variables
' =================

Private m_oDef As CDefinition
Private m_iPrimitiveIndex As Integer
Private m_bIsCancelled As Boolean
Private m_sCode As String
Private m_oSelectedPrimitive As CPrimitive

Private m_bIsLoading As Boolean



' Public Properties
' =================


Public Property Get IsCancelled() As Boolean
    IsCancelled = m_bIsCancelled
End Property

Public Property Get Code() As String
    Code = m_sCode
End Property


Public Property Let SelectedPrimitive(ByRef NewValue As CPrimitive)
    Set m_oSelectedPrimitive = NewValue
    
    Me.Show
    ShowPrimitive m_oSelectedPrimitive.Index
End Property



' Public Methods
' ==============



' Private Methods
' ===============


' Loading/Initialization Methods
' ------------------------------


Private Sub RefreshTree()
    
    Dim sCategories As String       ' List of all categories processed so far
    Dim sCategory As String         ' Category label for the current category
    Dim oPrim As CPrimitive         ' "Cookie-cutter" for new primitives
    
    tvwPrimitives.Nodes.Clear       ' Remove all nodes
    
    For Each oPrim In m_oDef.Primitives
        If oPrim.IsStandAlone Then
            ' String together a category label
            sCategory = "[" & oPrim.Category & "]"
            
            If InStr(1, sCategories, sCategory, vbBinaryCompare) = 0 Then
                ' Create a new "category" node if necessary
                With tvwPrimitives.Nodes.Add(, , sCategory, oPrim.Category, 1, 2)
                    '.Bold = True    ' Make the label bold
                End With
                sCategories = sCategories & " " & sCategory
            End If
            
            ' Add node for primitive
            tvwPrimitives.Nodes.Add sCategory, tvwChild, "p" & oPrim.Index, oPrim.PrimitiveName, 3, 3
        End If
    Next

End Sub

Public Sub LoadPrimitive(ByRef Source As CPrimitive)

    Dim oBlock As CBlock
    Set oBlock = New CBlock
    oBlock.Primitive = Source
    cp.Block = oBlock
    txtCode.Text = cp.Block.Text
    cp.SetFocus
    cp.FocusFirstControl

End Sub



Private Sub PositionMainControls()

    On Error Resume Next
    
    Dim iRightWidth As Integer
    
    ' Treeview
    tvwPrimitives.Height = Me.ScaleHeight - 90 '- 82
    
    ' Tab
    tabMain.Width = Me.ScaleWidth - tabMain.Left - 7
    tabMain.Height = Me.ScaleHeight - tabMain.Top - 61
    
    Dim i As Integer
    For i = picTabs.lbound To picTabs.UBound
        picTabs(i).Height = tabMain.Height - 25
        picTabs(i).Width = tabMain.Width - 6
    Next
    
    cp.BackColor = vbWindowBackground
    cp.Width = picTabs(0).Width - 6 'cp.Left
    cp.Height = picTabs(0).Height - 6 'cp.Top
    
    lblCode.Width = picTabs(1).Width
    txtCode.Width = picTabs(1).Width - 6
    txtCode.Height = picTabs(1).Height - 25
    
    
    ' Seperator lines
    Line1.X2 = Me.ScaleWidth - 5
    Line2.X2 = Line1.X2
    
    ' Buttons
    With cmdCancel
        .Top = Me.ScaleHeight - 53 '45
        .Left = Me.ScaleWidth - 97
    End With
    With cmdInsert
        .Top = cmdCancel.Top
        .Left = cmdCancel.Left - 100
    End With

End Sub



' Miscellaneous Methods
' ---------------------

Private Function GetIndexFromKey(ByVal Source As String) As Long

    Dim sVal As String
    Dim iInStr As Integer
    
    iInStr = InStr(1, Source, "a", vbBinaryCompare)
    If iInStr <> 0 Then
        sVal = Mid$(Source, iInStr + 1, Len(Source) - (iInStr))
    Else
        sVal = Mid$(Source, 2, Len(Source) - 1)
    End If
    
    If IsNumeric(sVal) Then
        GetIndexFromKey = CLng(sVal)
    Else
        GetIndexFromKey = -1
    End If

End Function

Private Function GetParamControlFromIndex(ByVal ParameterIndex As Integer) As Control

    ' Returns a reference to a control (textbox, checkbox or combo) that
    ' matches the specified parameter index.
    
    Dim cCtl As Control
    
    For Each cCtl In Me.Controls
        If TypeOf cCtl Is TextBox Or _
           TypeOf cCtl Is ComboBox Or _
           TypeOf cCtl Is CheckBox Then
        
            If cCtl.Tag = CStr(ParameterIndex) Then
                Set GetParamControlFromIndex = cCtl
            End If
        End If
    Next

End Function



' Event Handlers
' --------------

Private Sub cmdCancel_Click()
    m_bIsCancelled = True
    Me.Hide
End Sub

Private Sub cmdInsert_Click()
    m_bIsCancelled = False
    m_sCode = cp.Block.Text
    
    With FrmMain.ceCONs(FrmMain.CurrentCON)
        .InsertText m_sCode, True
        .SetFocus
    End With
    Unload Me
    'Me.Hide

End Sub


Private Sub Form_Load()
    
    SetCompatibleColours Me

    m_bIsCancelled = True
    
    Set m_oDef = New CDefinition

    Dim sFile As String

    sFile = App.Path & "\" & Settings.ReadSetting("Path_PrimitivesDef")

    If (Trim$(sFile) = "") Or (Trim$(Dir$(sFile)) = "") Then
        ' Disable everything -- no primitives!!
        tvwPrimitives.Enabled = False

        Dim iResult As Integer
        iResult = MsgBox(App.Title & " was unable to locate a Primitive Definition file. The file may have been moved or deleted or the specified path for this file may be invalid." & vbNewLine & vbNewLine & "Do you wish to open the Options window to check the setting?", vbYesNo + vbExclamation)
        If iResult = vbYes Then
            FrmOptions.Show vbModal
            If InStr(1, FrmOptions.SettingsChanged, "[Path_PrimitivesDef]") <> 0 Then
                MsgBox "The new Primitives Definition Path setting will take effect next time you open the Primitives window.", vbInformation
            End If
        End If
        'DisableControls
        Exit Sub
    End If
    
    m_oDef.LoadBinary sFile         ' Load file contents
    
    RefreshTree                     ' Refresh tree view of primitives
    
    ' Select first node
    DoEvents
    If tvwPrimitives.SelectedItem Is Nothing Then
        tvwPrimitives.Nodes(1).Selected = True
        tvwPrimitives_NodeClick tvwPrimitives.Nodes(1)
    End If

    PositionMainControls            ' Size/position controls

    cp.BackColor = vbButtonFace     ' Set background color of cp control

End Sub

Private Sub Form_Resize()

    PositionMainControls

End Sub


Private Sub tabMain_Click()

    picTabs(tabMain.SelectedItem.Index - 1).ZOrder vbBringToFront

    Select Case tabMain.SelectedItem.Index
    
        Case 1      ' Primitive
            cp.FocusFirstControl
        
        Case 2      ' Code
            txtCode.Text = cp.Block.Text
    
    End Select

End Sub

Private Sub ShowPrimitive(ByVal Index As Integer)
    Dim iWidth As Integer
    
    picTabs(0).Visible = True
    picTabs(1).Visible = True
    tabMain.Visible = True
    
    imgPrim.Visible = True
    imgCat.Visible = False
    With m_oDef.Primitives(Index)
        lblPrimitive.Caption = .PrimitiveName
        iWidth = lblNotes.Width
        lblDescription.AutoSize = True
        lblDescription.Caption = .Description
        lblDescription.ToolTipText = .Description
        lblNotes.AutoSize = True
        lblNotes.Caption = "Version " & Trim$(.DukeVersion) & " or higher"
        lblNotes.ToolTipText = lblNotes.Caption
        lblNotes.AutoSize = False
    End With
    
    LoadPrimitive m_oDef.Primitives(Index)
    m_iPrimitiveIndex = Index
End Sub

Private Sub tvwPrimitives_NodeClick(ByVal Node As ComctlLib.Node)

    Select Case Node.Image
        Case 1      ' Category
            picTabs(0).Visible = False
            picTabs(1).Visible = False
            tabMain.Visible = False
            
            imgPrim.Visible = False
            imgCat.Visible = True
            lblPrimitive.Caption = Node.Text
            lblNotes.Caption = ""
            lblNotes.ToolTipText = ""
            lblDescription.Caption = ""
            lblDescription.ToolTipText = ""
            
        Case 3      ' Primitive
            ShowPrimitive GetIndexFromKey(Node.Key)
    End Select


End Sub

Private Sub txtCode_Change()
    cp.Block.Text = txtCode.Text
    cp.Refresh
End Sub

