VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.UserControl CONTree 
   Alignable       =   -1  'True
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   ScaleHeight     =   240
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   320
   Begin ComctlLib.TreeView tvw 
      Height          =   2490
      Left            =   0
      TabIndex        =   0
      Top             =   240
      Width           =   3090
      _ExtentX        =   5450
      _ExtentY        =   4392
      _Version        =   327682
      HideSelection   =   0   'False
      Indentation     =   397
      LabelEdit       =   1
      LineStyle       =   1
      Style           =   7
      ImageList       =   "imlIcons"
      Appearance      =   1
   End
   Begin ComctlLib.ImageList imlIcons 
      Left            =   2940
      Top             =   1425
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   16711935
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   8
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "CONTree.ctx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "CONTree.ctx":0352
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "CONTree.ctx":06A4
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "CONTree.ctx":09F6
            Key             =   "Actor"
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "CONTree.ctx":0D48
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "CONTree.ctx":109A
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "CONTree.ctx":13EC
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "CONTree.ctx":173E
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Label lblTreeTitle 
      BackStyle       =   0  'Transparent
      Caption         =   "Objects"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000009&
      Height          =   255
      Left            =   0
      TabIndex        =   1
      Top             =   15
      Width           =   3000
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H80000002&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   255
      Left            =   0
      Top             =   0
      Width           =   3120
   End
   Begin VB.Menu mnuBlock 
      Caption         =   "&Block"
      Begin VB.Menu mnuBlockGoto 
         Caption         =   "&Goto"
      End
      Begin VB.Menu mnuBlockBar1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuBlockAdd 
         Caption         =   "&Add..."
      End
      Begin VB.Menu mnuBlockEdit 
         Caption         =   "&Edit..."
      End
      Begin VB.Menu mnuBlockRemove 
         Caption         =   "&Remove"
      End
   End
   Begin VB.Menu mnuFolder 
      Caption         =   "&Folder"
      Begin VB.Menu mnuFolderExpand 
         Caption         =   "&Expand"
      End
      Begin VB.Menu mnuFolderCompress 
         Caption         =   "&Compress"
      End
      Begin VB.Menu mnuFolderBar1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFolderAdd 
         Caption         =   "&Add..."
      End
   End
End
Attribute VB_Name = "CONTree"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit


' Public Events
' -------------


Public Event ObjectClick(ByRef Block As CBlock)
Public Event ObjectDblClick(ByRef Block As CBlock)
Public Event ObjectAdd(ByRef Structure As CStructure)
Public Event ObjectDelete(ByRef Block As CBlock)


' Private Variables
' -----------------

Private WithEvents m_oParser As CParser
Attribute m_oParser.VB_VarHelpID = -1

Private m_sRoot As String


' Public Properties
' -----------------

Public Property Get Parser() As CParser
    Set Parser = m_oParser
End Property

Public Property Let Parser(ByRef NewValue As CParser)
    Set m_oParser = NewValue
End Property



' Public Methods
' --------------

Public Sub SelectBlock(ByRef Block As CBlock)
    On Error GoTo ProcedureError
    
    With tvw.Nodes("b" & Block.Index)
        .Parent.Expanded = True
        .Selected = True
    End With
    Exit Sub
    
ProcedureError:
    Resume Next
End Sub


Public Sub SetRootNote(ByVal Text As String)
    m_sRoot = Text
    tvw.Nodes("Root").Text = m_sRoot
End Sub


Public Function GetRootNode() As String
    GetRootNode = m_sRoot
End Function


' Private Procedures
' ------------------


Private Sub DeleteNode(ByVal ObjectID As Long)
    tvw.Nodes.Remove "K" & ObjectID
End Sub


Public Sub InitCategories(ByVal CurrentCON As Integer)

    Dim oStruct As CStructure
    Dim sList() As String
    Dim sCategory As String
    ReDim sList(0) As String
    
    ' Get a unique list of the categories of all structures
    For Each oStruct In FrmMain.Definition.Structures 'FrmMain.ceCONs(CurrentCON).Parser.Definition.Structures
        If Not (oStruct.Primitive Is Nothing) Then
            sCategory = oStruct.Primitive.Category
            If sList(0) = "" Then
                sList(0) = sCategory
            Else
                If UBound(Filter(sList, sCategory, True, vbBinaryCompare)) = -1 Then
                    ReDim Preserve sList(LBound(sList) To UBound(sList) + 1)
                    sList(UBound(sList)) = sCategory
                End If
            End If
        End If
    Next
    
    ' Insert category nodes from the list
    Dim i As Integer
    For i = LBound(sList) To UBound(sList)
        tvw.Nodes.Add "Root", tvwChild, "c" & i, sList(i), 2, 3
    Next

End Sub

Public Sub InitRoot()
    
    ' Add the root node
    With tvw.Nodes
        .Clear
        .Add , , "Root", m_sRoot, 1, 1
        .Item("Root").Expanded = True
    End With

End Sub


Public Sub UpdateBlock(ByRef Block As CBlock)
    
    On Error GoTo ProcedureError
    
    tvw.Nodes("b" & Block.Index).Text = Block.ToString()

    Exit Sub

ProcedureError:
    If err.Number = 35601 Or err.Number = 91 Then
        ' Ignore error if the object not found
    Else
        MsgBox err.Number & vbNewLine & err.Description
    End If

End Sub


' Event Handlers
' --------------

Private Function GetCategoryNode(ByVal Category As String) As String

    Dim oSelNode As Variant

    
    If tvw.Nodes.Count > 0 Then
        Dim oNode As Variant
        Dim i As Integer
        For i = 1 To tvw.Nodes.Count
            Set oNode = tvw.Nodes.Item(i)
            If oNode.Image = 2 And oNode.Text = Category Then
                Set oSelNode = oNode
                Exit For
            End If
        Next
    End If

    If oSelNode Is Nothing Then
        Set oSelNode = tvw.Nodes.Add("Root", tvwChild, Replace$(Category, " ", ""), Category, 2, 3)
    End If
    
    GetCategoryNode = oSelNode.Key

End Function

Private Sub m_oParser_BlockAdded(Block As CBlock)
    
    Dim sKey As String
    Dim oNode As Variant
    
    If Block.Structure.Primitive Is Nothing Then
        sKey = "Root"
    Else
        sKey = GetCategoryNode(Block.Structure.Primitive.Category)
    End If
    
    Set oNode = tvw.Nodes.Add(sKey, tvwChild, "b" & Block.Index, Block.ToString(), Block.Structure.ImageID + 3, Block.Structure.ImageID + 3)
    If Not m_oParser.IsLoading Then
        oNode.Parent.Expanded = True
        oNode.Selected = True
    End If
    
End Sub

Private Sub m_oParser_BlockDeleted(ByVal BlockIndex As Integer)
    On Error GoTo ProcedureError
    
    tvw.Nodes.Remove "b" & BlockIndex
    Exit Sub
    
ProcedureError:
    If err.Number = 35601 Then
    Else
        MsgBox err.Number & vbNewLine & err.Description
    End If
End Sub

Private Sub mnuBlockAdd_Click()
    Dim oSelStruct As CStructure
    
    Select Case tvw.SelectedItem.Image
        'Case 1
        '    Set oSelStruct = m_oParser.Definition.Structures(1)
        Case 2 To 3
            Dim oStruct As CStructure
            For Each oStruct In m_oParser.Definition.Structures
                If oStruct.Primitive.Category = tvw.SelectedItem.Text Then
                    Set oSelStruct = oStruct
                    Exit For
                End If
            Next
            
            If oSelStruct Is Nothing Then
                Set oSelStruct = m_oParser.Definition.Structures(1)
            End If
        
        Case Else
            Set oSelStruct = m_oParser.Blocks(GetID(tvw.SelectedItem.Key)).Structure
            
    End Select
    
    If oSelStruct Is Nothing Then
        Set oSelStruct = m_oParser.Definition.Structures(1)
    End If
    
    RaiseEvent ObjectAdd(oSelStruct)

End Sub

Private Sub mnuBlockEdit_Click()
    If tvw.SelectedItem.Image > 3 Then
        RaiseEvent ObjectDblClick(m_oParser.Blocks(GetID(tvw.SelectedItem.Key)))
    End If
End Sub

Private Sub mnuBlockGoto_Click()
    If tvw.SelectedItem.Image > 3 Then
        FrmMain.ceCONs(FrmMain.CurrentCON).GotoBlock m_oParser.Blocks(GetID(tvw.SelectedItem.Key))
    End If
End Sub

Private Sub mnuBlockRemove_Click()
    If tvw.SelectedItem.Image > 3 Then
        RaiseEvent ObjectDelete(m_oParser.Blocks(GetID(tvw.SelectedItem.Key)))
    End If
End Sub

Private Sub mnuFolderAdd_Click()
    'RaiseEvent ObjectAdd
    If tvw.SelectedItem.Image > 1 Then
        mnuBlockAdd_Click
    End If
End Sub

Private Sub tvw_DblClick()
    If tvw.SelectedItem.Image > 3 Then
        RaiseEvent ObjectDblClick(m_oParser.Blocks(GetID(tvw.SelectedItem.Key)))
    End If
End Sub

Private Sub tvw_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)

    On Error Resume Next
    
    Select Case tvw.HitTest(x, y).Image
        Case 1 To 3     ' Root/Folder clicked
            If Button = vbRightButton Then
                PopupMenu mnuFolder, , , , mnuFolderExpand
            End If
        Case Else       ' Item clicked
            If Button = vbLeftButton Then
                'FrmMain.ceCONs(FrmMain.CurrentCON).GotoBlock m_oParser.Blocks(GetID(tvw.HitTest(x, y).Key))
                'FrmMain.ceCONs(FrmMain.CurrentCON).SetFocus
                RaiseEvent ObjectClick(m_oParser.Blocks(GetID(tvw.HitTest(x, y).Key)))
            ElseIf Button = vbRightButton Then
                PopupMenu mnuBlock, , , , mnuBlockGoto
            End If
    End Select
    
End Sub


Private Sub UserControl_Initialize()
    InitRoot
End Sub

Private Sub UserControl_Resize()

    On Error Resume Next
    
    Dim iW As Integer
    Dim iH As Integer
    
    iW = UserControl.ScaleWidth
    iH = UserControl.ScaleHeight

    tvw.Width = iW
    tvw.Height = iH - tvw.Top

    lblTreeTitle.Width = iW
    Shape1.Width = iW + 10

End Sub

