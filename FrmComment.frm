VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form FrmComment 
   Caption         =   "Comment"
   ClientHeight    =   5175
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5640
   Icon            =   "FrmComment.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   345
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   376
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdHelp 
      Caption         =   "&Help"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4320
      TabIndex        =   6
      Top             =   4680
      Width           =   1215
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3000
      TabIndex        =   5
      Top             =   4680
      Width           =   1215
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1680
      TabIndex        =   4
      Top             =   4680
      Width           =   1215
   End
   Begin VB.PictureBox picTabs 
      BorderStyle     =   0  'None
      Height          =   3255
      Index           =   0
      Left            =   240
      ScaleHeight     =   3255
      ScaleWidth      =   5175
      TabIndex        =   7
      Top             =   1200
      Width           =   5175
      Begin VB.PictureBox picStyle 
         BorderStyle     =   0  'None
         Height          =   615
         Index           =   0
         Left            =   120
         ScaleHeight     =   41
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   153
         TabIndex        =   8
         Top             =   360
         Width           =   2295
         Begin VB.OptionButton optCStyle 
            Caption         =   "C Style Comment (/*...*/)"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   0
            TabIndex        =   10
            Top             =   45
            Width           =   2295
         End
         Begin VB.OptionButton optCPPStyle 
            Caption         =   "C++ Style Comment (//...)"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   0
            TabIndex        =   9
            Top             =   360
            Value           =   -1  'True
            Width           =   2295
         End
      End
      Begin VB.Frame fraStyle 
         Caption         =   "Comment Style"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   0
         TabIndex        =   11
         Top             =   120
         Width           =   2535
      End
   End
   Begin MSComctlLib.TabStrip tabstrip 
      Height          =   3735
      Left            =   120
      TabIndex        =   3
      Top             =   840
      Width           =   5415
      _ExtentX        =   9551
      _ExtentY        =   6588
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   2
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "&General"
            Key             =   "General"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "&Code"
            Key             =   "Code"
            ImageVarType    =   2
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
   Begin VB.PictureBox picTabs 
      BorderStyle     =   0  'None
      Height          =   3255
      Index           =   1
      Left            =   240
      ScaleHeight     =   3255
      ScaleWidth      =   5175
      TabIndex        =   12
      Top             =   1200
      Width           =   5175
      Begin RichTextLib.RichTextBox rtbCode 
         Height          =   3135
         Left            =   0
         TabIndex        =   13
         Top             =   120
         Width           =   5175
         _ExtentX        =   9128
         _ExtentY        =   5530
         _Version        =   393217
         BorderStyle     =   0
         Enabled         =   -1  'True
         TextRTF         =   $"FrmComment.frx":1442
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
   Begin VB.Label lblType 
      BackStyle       =   0  'Transparent
      Caption         =   "C++ Style Comment"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   960
      TabIndex        =   2
      Top             =   240
      Width           =   2655
   End
   Begin VB.Label lblPosition 
      BackStyle       =   0  'Transparent
      Caption         =   "Position: [N/A] to [N/A]"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   960
      TabIndex        =   1
      Top             =   480
      Width           =   2655
   End
   Begin VB.Label lblName 
      BackStyle       =   0  'Transparent
      Caption         =   "[New Comment]"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   960
      TabIndex        =   0
      Top             =   0
      Width           =   2655
   End
   Begin VB.Image Image1 
      Height          =   720
      Left            =   120
      Picture         =   "FrmComment.frx":14C2
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
Attribute VB_Name = "FrmComment"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private m_bAddNew As Boolean
Private m_oBlock As Object
Private m_bIsCancelled As Boolean



Public Property Get AddNew() As Boolean
    AddNew = m_bAddNew
End Property
Public Property Let AddNew(ByVal NewValue As Boolean)
    m_bAddNew = NewValue
End Property


Public Property Get IsCancelled() As Boolean
    IsCancelled = m_bIsCancelled
End Property


Public Property Get Block() As Object
    Set Block = m_oBlock
End Property
Public Property Let Block(ByRef NewValue As Object)
    Set m_oBlock = NewValue
End Property


Public Sub Init()
    If m_bAddNew Then
        rtbCode.Text = "// New Comment"
    Else
        Me.Caption = m_oBlock.ToString() & " - Comment"
        lblName.Caption = m_oBlock.ToString()
        If m_oBlock.TypeID = 1 Then
            lblType.Caption = "C Style Comment"
            optCStyle.Value = True
        Else
            lblType.Caption = "C++ Style Comment"
            optCPPStyle.Value = True
        End If
        lblPosition.Caption = "Position: " & m_oBlock.BeginPos & " to " & m_oBlock.EndPos
        rtbCode.Text = m_oBlock.GetCode
    End If
End Sub


Private Sub cmdCancel_Click()

    m_bIsCancelled = True
    Me.Hide

End Sub

Private Sub cmdSave_Click()

    If m_bAddNew Then
        Set m_oBlock = New CComment
    End If
    m_oBlock.SetCode rtbCode.Text

    m_bIsCancelled = False
    Me.Hide

End Sub

Private Sub ChangeToCPPStyle()

    Dim s As String
    
    s = rtbCode.Text
    
    s = Replace(s, "/*", "", , , vbTextCompare)
    s = Replace(s, "*/", "", , , vbTextCompare)
    s = "//" & s
    
    s = Replace(s, Chr(13) & Chr(10), Chr(13), , , vbTextCompare)
    s = Replace(s, Chr(13), " ", , , vbTextCompare)
    ''s = Replace(s, Chr(13), Chr(13) & Chr(10) & "//", , , vbTextCompare)

    rtbCode.Text = s

End Sub

Private Sub ChangeToCStyle()

    Dim s As String
    
    s = rtbCode.Text
    
    s = Replace(s, "//", "", , , vbTextCompare)
    s = "/*" & s & "*/"

    rtbCode.Text = s

End Sub

Private Sub Form_Load()

    tabstrip_Click

End Sub


Private Sub Form_Unload(Cancel As Integer)

    m_bIsCancelled = True
    Me.Hide
    Cancel = 1

End Sub

Private Sub optCPPStyle_Click()
    
    If Mid$(LTrim$(rtbCode.Text), 1, 2) = "/*" Then _
        ChangeToCPPStyle
    'End If

End Sub

Private Sub optCStyle_Click()

    'If m_bAddNew Then
    '    rtbCode.Text = "/* New Comment */"
    'Else
        If Mid$(LTrim$(rtbCode.Text), 1, 2) <> "/*" Then _
            ChangeToCStyle
    'End If

End Sub

Private Sub tabstrip_Click()

    picTabs(tabstrip.SelectedItem.Index - 1).ZOrder vbBringToFront
    
End Sub

