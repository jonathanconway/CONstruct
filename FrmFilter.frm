VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmFilter 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "CON Script Filter"
   ClientHeight    =   3855
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4335
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FrmFilter.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   257
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   289
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.TreeView tvwFilter 
      Height          =   2655
      Left            =   1200
      TabIndex        =   5
      Top             =   480
      Width           =   3015
      _ExtentX        =   5318
      _ExtentY        =   4683
      _Version        =   393217
      Indentation     =   529
      LineStyle       =   1
      Sorted          =   -1  'True
      Style           =   7
      Checkboxes      =   -1  'True
      SingleSel       =   -1  'True
      Appearance      =   1
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   3000
      TabIndex        =   4
      Top             =   3360
      Width           =   1215
   End
   Begin VB.CommandButton cmdFilter 
      Caption         =   "Filter"
      Default         =   -1  'True
      Height          =   375
      Left            =   1680
      TabIndex        =   3
      Top             =   3360
      Width           =   1215
   End
   Begin VB.ComboBox cboDocument 
      Height          =   315
      Left            =   1200
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   120
      Width           =   3030
   End
   Begin VB.Label lblFilterBy 
      Caption         =   "Filter By:"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   480
      Width           =   975
   End
   Begin VB.Label lblDocument 
      Caption         =   "Document:"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   975
   End
End
Attribute VB_Name = "FrmFilter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private m_bIsCancelled As Boolean
Private m_sFilterText As String


Public Property Get IsCancelled() As Boolean
    IsCancelled = m_bIsCancelled
End Property

Public Property Get FilterText() As String
    FilterText = m_sFilterText
End Property




Private Function GetCurrentCONIndex() As Integer

    ' Get CON index of source document
    If cboDocument.ListIndex <= 0 Then
        GetCurrentCONIndex = FrmMain.CurrentCON
    Else
        GetCurrentCONIndex = cboDocument.ItemData(cboDocument.ListIndex)
    End If

End Function

Private Sub RefreshTree()

    Dim iIndex As Integer

    ' CON index of source document
    iIndex = GetCurrentCONIndex()
    
    ' Add main node
    tvwFilter.Nodes.Clear
    
    ' Add all structures
    Dim oStruct As CStructure
    Dim oBlock As CBlock
    For Each oStruct In FrmMain.Definition.Structures
        tvwFilter.Nodes.Add , , "s" & oStruct.Index, oStruct.StructureName ', 2, 3
    
        ' Add all blocks for this structure
        For Each oBlock In FrmMain.ceCONs(iIndex).Parser.Blocks
            If oBlock.Structure.StructureName = oStruct.StructureName Then
                tvwFilter.Nodes.Add "s" & oStruct.Index, tvwChild, "b" & oBlock.Index, oBlock.ToString() ', 3, 3
            End If
        Next
    Next

End Sub



Private Sub EnableMainControls(ByVal Status As Boolean)

    lblDocument.Enabled = Status
    cboDocument.Enabled = Status
    lblFilterBy.Enabled = Status
    tvwFilter.Enabled = Status
    cmdFilter.Enabled = Status

End Sub



Private Sub LoadDocsCombo()

    On Error GoTo ProcedureError

    Dim bEnable As Boolean
    
    bEnable = (FrmMain.ceCONs.ubound > 0)
    
    EnableMainControls bEnable
    
    If bEnable Then
        
        With cboDocument
            .Clear
            .AddItem "[Current]"
        
            Dim i As Integer
            For i = FrmMain.ceCONs.lbound + 1 To FrmMain.ceCONs.ubound
                If Len(Trim$(FrmMain.ceCONs(i).Filename)) > 0 Then
                    .AddItem GetFilenameFromPath(FrmMain.ceCONs(i).Filename)
                    .ItemData(.ListCount - 1) = i
                End If
            Next
        
            .ListIndex = 0
        End With
        
    End If
    Exit Sub
    
ProcedureError:
    If err.Number = 340 Then
        Resume Next
    Else
        MsgBox err.Number & vbNewLine & err.Description
    End If


End Sub



Private Sub cboDocument_Click()
    Dim bEnable As Boolean
    
    bEnable = (Len(Trim$(cboDocument.Text)) <> 0)
    EnableMainControls bEnable
    If bEnable Then
        RefreshTree
    End If
    
End Sub

Private Sub cmdCancel_Click()

    m_bIsCancelled = True
    Me.Hide

End Sub


Private Sub cmdFilter_Click()
    
    Dim sText As String
    Dim sHeader As String
    Dim lCount As Long
    
    With FrmMain.ceCONs(GetCurrentCONIndex())
        
        sText = vbNewLine
        sHeader = "/* Filter Results for """ & GetFilenameFromPath(.Filename) & """" & vbNewLine
        
        Dim oStructNode As MSComctlLib.Node
        Dim oNode As MSComctlLib.Node
        
        For Each oStructNode In tvwFilter.Nodes
            If Left$(oStructNode.Key, 1) = "s" Then
                ' Structure
                lCount = 0
                For Each oNode In tvwFilter.Nodes
                    If oNode.Parent Is oStructNode Then
                        If oNode.Checked Then
                            sText = sText & .Parser.Blocks(GetID(oNode.Key)).Text & vbNewLine & vbNewLine
                            lCount = lCount + 1
                        End If
                    End If
                Next
                sHeader = sHeader & "/* " & oStructNode.Text & ": " & lCount & vbNewLine
            End If
        Next
    
        sHeader = sHeader & "*/" & vbNewLine
        sText = sHeader & sText
    
    End With
    
    m_sFilterText = sText
    
    m_bIsCancelled = False
    Me.Hide

End Sub


Private Sub Form_Load()

    m_bIsCancelled = True

    LoadDocsCombo
    'LoadStructsCombo

End Sub

Private Sub tvwFilter_NodeCheck(ByVal Node As MSComctlLib.Node)
    
    If Node.Children > 0 Then
        Dim oNode As MSComctlLib.Node
        For Each oNode In tvwFilter.Nodes
            If oNode.Parent Is Node Then
                oNode.Checked = Node.Checked
            End If
        Next
    End If

End Sub
