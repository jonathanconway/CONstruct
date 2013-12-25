VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Begin VB.Form FrmInsertSnippet 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Insert Snippet"
   ClientHeight    =   4170
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4425
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FrmInsertSnippet.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   278
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   295
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdInsert 
      Caption         =   "Insert"
      Default         =   -1  'True
      Height          =   375
      Left            =   1470
      TabIndex        =   1
      Top             =   3675
      Width           =   1380
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   2940
      TabIndex        =   0
      Top             =   3675
      Width           =   1380
   End
   Begin VB.PictureBox picTabs 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2265
      Index           =   0
      Left            =   150
      ScaleHeight     =   2265
      ScaleWidth      =   4110
      TabIndex        =   4
      Top             =   1260
      Width           =   4110
      Begin VB.ListBox lstFields 
         Height          =   1815
         Left            =   105
         TabIndex        =   6
         Top             =   315
         Width           =   1800
      End
      Begin VB.TextBox txtValue 
         Height          =   285
         Left            =   1995
         TabIndex        =   5
         Top             =   315
         Width           =   2010
      End
      Begin VB.Label lblFields 
         BackStyle       =   0  'Transparent
         Caption         =   "Fields:"
         Height          =   225
         Left            =   105
         TabIndex        =   8
         Top             =   105
         Width           =   540
      End
      Begin VB.Label lblValue 
         BackStyle       =   0  'Transparent
         Caption         =   "Value:"
         Height          =   225
         Left            =   1995
         TabIndex        =   7
         Top             =   105
         Width           =   1065
      End
   End
   Begin ComctlLib.TabStrip tabMain 
      Height          =   2745
      Left            =   105
      TabIndex        =   3
      Top             =   840
      Width           =   4215
      _ExtentX        =   7435
      _ExtentY        =   4842
      _Version        =   327682
      BeginProperty Tabs {0713E432-850A-101B-AFC0-4210102A8DA7} 
         NumTabs         =   2
         BeginProperty Tab1 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Fields"
            Key             =   ""
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Code"
            Key             =   ""
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox picTabs 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2265
      Index           =   1
      Left            =   150
      ScaleHeight     =   2265
      ScaleWidth      =   4005
      TabIndex        =   9
      Top             =   1260
      Width           =   4005
      Begin VB.TextBox txtPreview 
         Height          =   1800
         Left            =   105
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   11
         Top             =   315
         Width           =   3900
      End
      Begin VB.Label lblPreview 
         BackStyle       =   0  'Transparent
         Caption         =   "Code Preview:"
         Height          =   225
         Left            =   105
         TabIndex        =   10
         Top             =   105
         Width           =   1170
      End
   End
   Begin VB.Label Label1 
      Caption         =   $"FrmInsertSnippet.frx":058A
      Height          =   645
      Left            =   105
      TabIndex        =   2
      Top             =   105
      Width           =   4110
   End
End
Attribute VB_Name = "FrmInsertSnippet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private m_oSnippet As CSnippet
Private m_bIsCancelled As Boolean
Private m_sInsertText As String

Private m_oValues As CList



Public Property Get InsertText() As String
    InsertText = m_sInsertText
End Property


Public Property Get IsCancelled() As Boolean
    IsCancelled = m_bIsCancelled
End Property


Public Property Let Snippet(ByRef NewValue As CSnippet)
    Set m_oSnippet = NewValue
    
    Set m_oValues = New CList
    
    ' Populate listbox
    lstFields.Clear
    Dim oField As CSnippetField
    For Each oField In m_oSnippet.Fields
        lstFields.AddItem oField.FieldName
        m_oValues.AddItem oField.FieldName, ""
    Next
    
    ' Select first list item
    lstFields.ListIndex = 0
    
    ' Update insert text
    UpdateInsertText
    
End Property



Private Sub UpdateInsertText()

    m_sInsertText = m_oSnippet.Text
    InsertFieldValues
    txtPreview.Text = m_sInsertText

End Sub

Private Sub InsertFieldValues()

    Dim oField As CSnippetField
    Dim sField As String
    
    For Each oField In m_oSnippet.Fields
        sField = "<" & oField.FieldName & ">"
        Do Until InStr32(1, m_sInsertText, sField, vbTextCompare) = 0
            m_sInsertText = Replace$(m_sInsertText, sField, m_oValues(oField.FieldName).Value, , , vbTextCompare)
        Loop
    Next

End Sub

Private Sub cmdCancel_Click()
    m_bIsCancelled = True
    Me.Hide
End Sub

Private Sub cmdInsert_Click()
    m_bIsCancelled = False
    Me.Hide
End Sub

Private Sub Form_Load()
    m_bIsCancelled = True
    SetCompatibleColours Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Me.Hide
End Sub

Private Sub lstFields_Click()
    With txtValue
        .Text = m_oValues(lstFields.Text).Value
        .SelStart = 0
        .SelLength = Len(txtValue.Text)
        If Me.Visible Then
            .SetFocus
        End If
    End With
End Sub


Private Sub tabMain_Click()
    picTabs(tabMain.SelectedItem.Index - 1).ZOrder vbBringToFront
End Sub

Private Sub txtValue_Change()
     With m_oValues
        If .HasItem(lstFields.Text) Then
            m_oValues(lstFields.Text).Value = txtValue.Text
        Else
            .AddItem lstFields.Text, txtValue.Text
        End If
    End With
    UpdateInsertText
End Sub







'Private Sub RemoveTags()
'
'    Dim bIsDone As Boolean
'    Dim lBegin As Long
'    Dim lEnd As Long
'    Dim lPos As Long
'
'    lPos = 1
'    m_sInsertText = m_oSnippet.Text
'    bIsDone = False
'
'    Do Until bIsDone
'        lBegin = InStr32(lPos, m_sInsertText, "/**", vbTextCompare)
'        If lBegin > 0 Then
'            lEnd = InStr32(lBegin, m_sInsertText, "*/", vbBinaryCompare)
'            If lEnd > 0 Then
'                m_sInsertText = Mid$(m_sInsertText, 1, lBegin - 1) & Mid$(m_sInsertText, lEnd + 4, Len(m_sInsertText) - (lEnd + 1))
'            Else
'                lPos = lBegin + 3
'            End If
'        Else
'            bIsDone = True
'        End If
'    Loop
'
'End Sub
