VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Begin VB.Form FrmSnippet 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Add Snippet"
   ClientHeight    =   3615
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4845
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FrmSnippet.frx":0000
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   241
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   323
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame fraTags 
      Caption         =   "Tags"
      Height          =   2535
      Left            =   105
      TabIndex        =   9
      Top             =   3675
      Width           =   4635
      Begin ComctlLib.ListView lvwTags 
         Height          =   1590
         Left            =   105
         TabIndex        =   18
         Top             =   840
         Width           =   2115
         _ExtentX        =   3731
         _ExtentY        =   2805
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         HideColumnHeaders=   -1  'True
         _Version        =   327682
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         Appearance      =   1
         NumItems        =   1
         BeginProperty ColumnHeader(1) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Item"
            Object.Tag             =   ""
            Text            =   "Item"
            Object.Width           =   2999
         EndProperty
      End
      Begin VB.PictureBox Picture1 
         BorderStyle     =   0  'None
         Height          =   435
         Left            =   105
         ScaleHeight     =   29
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   295
         TabIndex        =   14
         Top             =   315
         Width           =   4425
         Begin VB.CommandButton cmdTagInsert 
            Appearance      =   0  'Flat
            Caption         =   "&Insert"
            Enabled         =   0   'False
            Height          =   375
            Left            =   3015
            TabIndex        =   19
            Top             =   0
            Width           =   960
         End
         Begin VB.CommandButton cmdTagAdd 
            Appearance      =   0  'Flat
            Caption         =   "&Add"
            Enabled         =   0   'False
            Height          =   375
            Left            =   0
            TabIndex        =   17
            Top             =   0
            Width           =   960
         End
         Begin VB.CommandButton cmdTagDelete 
            Appearance      =   0  'Flat
            Caption         =   "&Delete"
            Height          =   375
            Left            =   2010
            TabIndex        =   16
            Top             =   0
            Width           =   960
         End
         Begin VB.CommandButton cmdTagSave 
            Appearance      =   0  'Flat
            Caption         =   "&Save"
            Enabled         =   0   'False
            Height          =   375
            Left            =   1005
            TabIndex        =   15
            Top             =   0
            Width           =   960
         End
      End
      Begin VB.TextBox txtTagDataType 
         Height          =   285
         Left            =   2310
         TabIndex        =   12
         Top             =   1575
         Width           =   1380
      End
      Begin VB.TextBox txtTagLabel 
         Height          =   285
         Left            =   2310
         TabIndex        =   10
         Top             =   1050
         Width           =   2220
      End
      Begin VB.Label lblTagDataType 
         Caption         =   "Data Type:"
         Height          =   225
         Left            =   2310
         TabIndex        =   13
         Top             =   1365
         Width           =   1485
      End
      Begin VB.Label lblTagLabel 
         Caption         =   "Label:"
         Height          =   225
         Left            =   2310
         TabIndex        =   11
         Top             =   840
         Width           =   1485
      End
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Save"
      Height          =   375
      Left            =   2310
      TabIndex        =   8
      Top             =   3150
      Width           =   1170
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   3570
      TabIndex        =   7
      Top             =   3150
      Width           =   1170
   End
   Begin VB.CommandButton cmdMore 
      Caption         =   "&Tags >>"
      Height          =   375
      Left            =   105
      TabIndex        =   6
      Top             =   3150
      Width           =   1170
   End
   Begin VB.TextBox txtText 
      Height          =   1905
      HideSelection   =   0   'False
      Left            =   105
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   5
      Top             =   1155
      Width           =   4635
   End
   Begin VB.TextBox txtCategory 
      Height          =   300
      Left            =   1260
      TabIndex        =   3
      Top             =   495
      Width           =   3480
   End
   Begin VB.TextBox txtLabel 
      Height          =   300
      Left            =   1260
      TabIndex        =   1
      Top             =   105
      Width           =   3480
   End
   Begin VB.Label lblText 
      Caption         =   "Text:"
      Height          =   225
      Left            =   105
      TabIndex        =   4
      Top             =   945
      Width           =   1590
   End
   Begin VB.Label lblCategory 
      Caption         =   "Category:"
      Height          =   225
      Left            =   105
      TabIndex        =   2
      Top             =   495
      Width           =   1065
   End
   Begin VB.Label lblLabel 
      Caption         =   "Label:"
      Height          =   225
      Left            =   105
      TabIndex        =   0
      Top             =   105
      Width           =   1065
   End
End
Attribute VB_Name = "FrmSnippet"
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


Private m_oSnippet As CSnippet
Private m_oOldSnippet As CSnippet

Private m_bIsCancelled As Boolean



Public Property Get IsCancelled() As Boolean
    IsCancelled = m_bIsCancelled
End Property



Public Property Get Snippet() As CSnippet
    Set Snippet = m_oSnippet
End Property

Public Property Let Snippet(ByVal NewValue As CSnippet)
    Set m_oSnippet = NewValue
    UpdateForm
    CopySnippet NewValue
End Property


Private Sub UpdateForm()

    ' Load form controls
    txtLabel.Text = m_oSnippet.Label
    txtCategory.Text = m_oSnippet.Category
    txtText.Text = m_oSnippet.Text
    
    ' Load tags
    Dim oField As CSnippetField
    For Each oField In m_oSnippet.Fields
        lvwTags.ListItems.Add , , oField.ToString()
    Next

    cmdTagInsert.Enabled = (lvwTags.ListItems.Count > 0)

End Sub


Private Sub CopySnippet(ByRef Source As CSnippet)
    ' Make a physical backup copy of the block.
    
    ' When it is necessary to search source code for the code of the
    ' original block so as to replace it, use this copy.
    
    Set m_oOldSnippet = New CSnippet
    m_oOldSnippet.Text = Source.Text
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdSave_Click()
    m_bIsCancelled = False
    
    Me.Hide
End Sub


Private Sub cmdMore_Click()

    If Me.Height > 4100 Then
        ' collapse
        Me.Height = 4100
        cmdMore.Caption = "&Tags >>"
    Else
        ' expand
        Me.Height = 6795
        cmdMore.Caption = "&Tags <<"
    End If

End Sub

Private Sub cmdTagAdd_Click()

    Dim oField As CSnippetField
    Set oField = m_oSnippet.Fields.FindItem(txtTagLabel.Text)
    If oField Is Nothing Then
    
        Set oField = New CSnippetField
        oField.FieldName = txtTagLabel.Text
        oField.DataType = FixInteger(txtTagDataType.Text)
        
        ' Add item
        m_oSnippet.Fields.Add oField            ' To collection
        
        lvwTags.ListItems.Add(, , oField.ToString()).Selected = True
    
    End If

    cmdTagInsert.Enabled = (lvwTags.ListItems.Count > 0)

End Sub

Private Sub cmdTagDelete_Click()
    If lvwTags.ListItems.Count > 0 Then
        Dim oField As CSnippetField
        Set oField = m_oSnippet.Fields.FindItem(lvwTags.SelectedItem.Text)
        If Not (oField Is Nothing) Then
            
            lvwTags.ListItems.Remove lvwTags.SelectedItem.Index
            
            m_oSnippet.Fields.Delete oField.Index
        
            If lvwTags.ListItems.Count > 0 Then
                lvwTags.ListItems(1).Selected = True
            Else
                txtTagLabel.Text = ""
                txtTagDataType.Text = "0"
            End If
        
        End If
    End If
    cmdTagInsert.Enabled = (lvwTags.ListItems.Count > 0)
End Sub

Private Sub cmdTagInsert_Click()
    If lvwTags.ListItems.Count > 0 Then
        txtText.SelText = "<" & m_oSnippet.Fields.FindItem(lvwTags.SelectedItem.Text).FieldName & ">"
        txtText.SetFocus
    End If
End Sub

Private Sub cmdTagSave_Click()

    If lvwTags.ListItems.Count > 0 Then
    
        With m_oSnippet.Fields.FindItem(lvwTags.SelectedItem.Text)
            .FieldName = txtTagLabel.Text
            .DataType = txtTagDataType.Text
            
            lvwTags.SelectedItem.Text = .FieldName
        End With
    
    End If

End Sub

Private Sub Form_Load()
    m_bIsCancelled = True
End Sub


Private Sub lvwTags_Click()
    If lvwTags.ListItems.Count > 0 Then
        Dim oField As CSnippetField
        Set oField = m_oSnippet.Fields.FindItem(lvwTags.SelectedItem.Text)
        If Not (oField Is Nothing) Then
            txtTagLabel.Text = oField.FieldName
            txtTagDataType.Text = oField.DataType
        End If
    End If
End Sub

Private Sub txtCategory_Change()
    m_oSnippet.Category = txtCategory.Text
End Sub

Private Sub txtLabel_Change()
    m_oSnippet.Label = txtLabel.Text
End Sub


Private Sub txtTagDataType_Change()
    txtTagDataType.Text = FixInteger(txtTagDataType.Text)
End Sub

Private Sub txtTagLabel_Change()
    Dim bEnable As Boolean
    bEnable = Len(Trim$(txtTagLabel.Text)) <> 0
    cmdTagAdd.Enabled = bEnable
    cmdTagSave.Enabled = bEnable
End Sub

Private Sub txtText_Change()
    m_oSnippet.Text = txtText.Text
End Sub
