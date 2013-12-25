VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Begin VB.Form FrmFindReplace 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Find and Replace"
   ClientHeight    =   2790
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5115
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FrmFindReplace.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   186
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   341
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picTabs 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1575
      Index           =   2
      Left            =   240
      ScaleHeight     =   105
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   313
      TabIndex        =   18
      Top             =   540
      Width           =   4695
      Begin VB.ComboBox cboGotoBlock 
         Height          =   315
         ItemData        =   "FrmFindReplace.frx":000C
         Left            =   630
         List            =   "FrmFindReplace.frx":000E
         Style           =   2  'Dropdown List
         TabIndex        =   22
         Top             =   420
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.TextBox txtGoto 
         Height          =   300
         Left            =   630
         TabIndex        =   21
         Top             =   420
         Width           =   1695
      End
      Begin VB.ComboBox cboGoto 
         Height          =   315
         ItemData        =   "FrmFindReplace.frx":0010
         Left            =   630
         List            =   "FrmFindReplace.frx":001D
         Style           =   2  'Dropdown List
         TabIndex        =   19
         Top             =   0
         Width           =   1695
      End
      Begin VB.Label lblGoto 
         BackStyle       =   0  'Transparent
         Caption         =   "Goto:"
         Height          =   255
         Left            =   0
         TabIndex        =   20
         Top             =   0
         Width           =   495
      End
   End
   Begin VB.PictureBox picTabs 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1575
      Index           =   1
      Left            =   240
      ScaleHeight     =   105
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   313
      TabIndex        =   11
      Top             =   540
      Width           =   4695
      Begin VB.TextBox txtReplaceReplaceWith 
         Height          =   300
         Left            =   1080
         TabIndex        =   16
         Top             =   360
         Width           =   3525
      End
      Begin VB.TextBox txtReplaceFindWhat 
         Height          =   300
         Left            =   1080
         TabIndex        =   14
         Top             =   0
         Width           =   3525
      End
      Begin VB.CheckBox chkReplaceMatchCase 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Match case"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   0
         TabIndex        =   13
         Top             =   840
         Width           =   2295
      End
      Begin VB.CheckBox chkReplaceWholeWord 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Find whole words only"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   0
         TabIndex        =   12
         Top             =   1200
         Width           =   2295
      End
      Begin VB.Label lblReplaceReplaceWith 
         BackStyle       =   0  'Transparent
         Caption         =   "Replace with:"
         Height          =   255
         Left            =   0
         TabIndex        =   17
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label lblReplaceFindWhat 
         BackStyle       =   0  'Transparent
         Caption         =   "Find what:"
         Height          =   255
         Left            =   0
         TabIndex        =   15
         Top             =   0
         Width           =   1095
      End
   End
   Begin VB.PictureBox picTabs 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1575
      Index           =   0
      Left            =   240
      ScaleHeight     =   105
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   313
      TabIndex        =   6
      Top             =   540
      Width           =   4695
      Begin VB.CheckBox chkFindWholeWord 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Find whole words only"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   0
         TabIndex        =   10
         Top             =   840
         Width           =   2295
      End
      Begin VB.CheckBox chkFindMatchCase 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Match case"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   0
         TabIndex        =   9
         Top             =   480
         Width           =   2295
      End
      Begin VB.TextBox txtFindFindWhat 
         Height          =   300
         Left            =   1080
         TabIndex        =   7
         Top             =   0
         Width           =   3525
      End
      Begin VB.Label lblFindFindWhat 
         BackStyle       =   0  'Transparent
         Caption         =   "Find what:"
         Height          =   255
         Left            =   0
         TabIndex        =   8
         Top             =   0
         Width           =   1095
      End
   End
   Begin VB.CommandButton cmdReplace 
      Caption         =   "Replace"
      Enabled         =   0   'False
      Height          =   375
      Left            =   420
      TabIndex        =   3
      Tag             =   "[ControlGroup]=Replace;"
      Top             =   2310
      Width           =   1110
   End
   Begin VB.CommandButton cmdReplaceAll 
      Caption         =   "Replace All"
      Enabled         =   0   'False
      Height          =   375
      Left            =   1575
      TabIndex        =   2
      Tag             =   "[ControlGroup]=Replace;"
      Top             =   2310
      Width           =   1110
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   3885
      TabIndex        =   0
      Tag             =   "[ControlGroup]=AlwaysVisible;"
      Top             =   2310
      Width           =   1110
   End
   Begin VB.CommandButton cmdFindNext 
      Caption         =   "Find Next"
      Default         =   -1  'True
      Height          =   375
      Left            =   2730
      TabIndex        =   1
      Tag             =   "[ControlGroup]=Find; [ControlGroup]=Replace;"
      Top             =   2310
      Width           =   1110
   End
   Begin VB.CommandButton cmdGoto 
      Caption         =   "Goto"
      Height          =   375
      Left            =   2730
      TabIndex        =   4
      Tag             =   "[ControlGroup]=Goto;"
      Top             =   2310
      Visible         =   0   'False
      Width           =   1110
   End
   Begin ComctlLib.TabStrip tabMain 
      Height          =   2085
      Left            =   120
      TabIndex        =   5
      Tag             =   "[ControlGroup]=AlwaysVisible;"
      Top             =   120
      Width           =   4890
      _ExtentX        =   8625
      _ExtentY        =   3678
      _Version        =   327682
      BeginProperty Tabs {0713E432-850A-101B-AFC0-4210102A8DA7} 
         NumTabs         =   3
         BeginProperty Tab1 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "&Find"
            Key             =   "Find"
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "&Replace"
            Key             =   "Replace"
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "&Goto"
            Key             =   "Goto"
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "FrmFindReplace"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' =============================================================================
' Module Name:      FrmFindReplace
' Module Type:      User Form
' Description:      Find/Replace/Goto tabbed dialog for Find, Replace and Goto
'                   editing functionality
' Author(s):        Jonathan A. Conway
' -----------------------------------------------------------------------------
' Log:
'
' 04 07 28 :
'   - Fixed bug of form obscuring message box when it is shown
'
' ?? ?? ?? :
' Created FrmFindReplace and coded code for displaying it from FrmMain (see
' FrmMain.ShowFindReplaceDialog())
' =============================================================================

Option Explicit


Private WithEvents m_oParser As CParser
Attribute m_oParser.VB_VarHelpID = -1


' Public Properties
' =================

Public Property Let Parser(ByRef NewValue As CParser)
    Set m_oParser = NewValue
    RefreshGotoCombo
End Property




' Public Methods
' ==============

Public Sub ChangeTab(ByVal TabIndex As Integer)
    tabMain.Tabs(TabIndex + 1).Selected = True
    tabMain_Click
End Sub


' Private Methods
' ===============

Private Sub RefreshGotoCombo()

    ' Load up "goto" combo of blocks
    cboGotoBlock.Clear
    Dim oBlock As CBlock
    For Each oBlock In m_oParser.Blocks
        cboGotoBlock.AddItem oBlock.ToString()
    Next

End Sub


' Event Handlers
' ==============

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdFindNext_Click()
    With FrmMain.ceCONs(FrmMain.CurrentCON)
        Select Case tabMain.SelectedItem.Index
            Case 1      ' Find tab
                .FindWhat = txtFindFindWhat.Text
                .FindMatchCase = (chkFindMatchCase.Value = vbChecked)
                .FindWholeWord = (chkFindWholeWord.Value = vbChecked)
            Case 2      ' Replace tab
                .FindWhat = txtReplaceFindWhat.Text
                .FindMatchCase = (chkReplaceMatchCase.Value = vbChecked)
                .FindWholeWord = (chkReplaceWholeWord.Value = vbChecked)
        End Select
        
        If .FindNext() Then
            StayOnTop Me, False
            MsgBox App.Title & " has finished searching the document.", vbInformation
            StayOnTop Me, True
        End If
        .SetFocus
    End With
End Sub

Private Sub cmdGoto_Click()
    
    Select Case cboGoto.ListIndex
    
        Case 0      ' Position
            With FrmMain.ceCONs(FrmMain.CurrentCON)
                .SelStart = FixLong(txtGoto.Text) + 1
                .SetFocus
            End With
        
        Case 1      ' Line
            With FrmMain.ceCONs(FrmMain.CurrentCON)
                .SelLine = FixLong(txtGoto.Text)
                .SetFocus
            End With
        
        Case 2      ' Block
            With FrmMain.ceCONs(FrmMain.CurrentCON)
                .GotoBlock m_oParser.Blocks.FindItem(cboGotoBlock.Text)
                .SetFocus
            End With
    
    End Select

End Sub

Private Sub cmdReplace_Click()
    With FrmMain.ceCONs(FrmMain.CurrentCON)
        .FindWhat = txtReplaceFindWhat.Text
        .FindMatchCase = (chkReplaceMatchCase.Value = vbChecked)
        .FindWholeWord = (chkReplaceWholeWord.Value = vbChecked)
        
        If .IsSearchWordSelected() Then
            .SelText = txtReplaceReplaceWith.Text
        End If
        
        
        If .FindNext() Then
            StayOnTop Me, False
            MsgBox App.Title & " has finished searching the document.", vbInformation
            StayOnTop Me, True
        End If
        .SetFocus
    End With
End Sub

Private Sub cmdReplaceAll_Click()
    Dim iCount As Integer
    With FrmMain.ceCONs(FrmMain.CurrentCON)
        .FindWhat = txtReplaceFindWhat.Text
        .FindMatchCase = (chkReplaceMatchCase.Value = vbChecked)
        .FindWholeWord = (chkReplaceWholeWord.Value = vbChecked)
        iCount = .ReplaceAll(txtReplaceReplaceWith.Text)
        StayOnTop Me, False
        MsgBox "CONstruct has completed its search of the document and has made " & iCount & " replacements."
        StayOnTop Me, True
        .SetFocus
    End With
End Sub

Private Sub Form_Load()
    StayOnTop Me, True
    Me.Show
    tabMain.Tabs(1).Selected = True
    cboGoto.ListIndex = 0
    SetCompatibleColours Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Cancel = 1
    StayOnTop Me, False
    Me.Hide
End Sub


Private Sub m_oParser_BlockAdded(Block As CBlock)
    RefreshGotoCombo
End Sub

Private Sub m_oParser_BlockDeleted(ByVal BlockIndex As Integer)
    RefreshGotoCombo
End Sub

Private Sub tabMain_Click()
    
    Static iPrevTab As Integer
    Dim bReplace As Boolean
    Dim iIndex As Integer
    
    iIndex = tabMain.SelectedItem.Index
    
    picTabs(iIndex - 1).ZOrder vbBringToFront
    
    bReplace = (iIndex = 2)
    
    cmdReplace.Enabled = bReplace
    cmdReplaceAll.Enabled = bReplace
    
    cmdGoto.Visible = (iIndex = 3)
    cmdFindNext.Visible = Not (iIndex = 3)
    
    If iPrevTab <> iIndex _
        And iIndex < 3 Then
        
        Select Case iPrevTab
            Case 1
                txtReplaceFindWhat.Text = txtFindFindWhat.Text
                chkReplaceMatchCase.Value = chkFindMatchCase.Value
                chkReplaceWholeWord.Value = chkFindWholeWord.Value
            Case 2
                txtFindFindWhat.Text = txtReplaceFindWhat.Text
                chkFindMatchCase.Value = chkReplaceMatchCase.Value
                chkFindWholeWord.Value = chkReplaceWholeWord.Value
        End Select
        
        iPrevTab = iIndex
    
    End If
    
End Sub


Private Sub cboGoto_Click()
    txtGoto.Visible = (cboGoto.ListIndex <> 2)
    cboGotoBlock.Visible = (cboGoto.ListIndex = 2)
End Sub


'Private Function ShowControls(ByVal ControlGroup As String)
'
'    Dim cctl As Control
'
'    For Each cctl In Me.Controls
'        cctl.Visible = (InStr(1, cctl.Tag, "[ControlGroup]=" & ControlGroup, vbTextCompare) <> 0 Or InStr(1, cctl.Tag, "[ControlGroup]=AlwaysVisible") <> 0)
'    Next
'
'    Select Case ControlGroup
'        Case "Find"
'            chkMatchCase.Top = 32 + 35
'            chkWholeWord.Top = 56 + 35
'            cmdFindNext.Default = True
'            txtFindWhat.SetFocus
'        Case "Replace"
'            chkMatchCase.Top = 56 + 35
'            chkWholeWord.Top = 80 + 35
'            cmdFindNext.Default = True
'            txtFindWhat.SetFocus
'        Case "Goto"
'            cmdGoto.Default = True
'            cboGoto.SetFocus
'    End Select
'
'End Function
