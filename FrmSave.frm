VERSION 5.00
Begin VB.Form FrmCloseSave 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "CONstruct"
   ClientHeight    =   3750
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5670
   Icon            =   "FrmSave.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   250
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   378
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      Height          =   3735
      Left            =   0
      ScaleHeight     =   3675
      ScaleWidth      =   5595
      TabIndex        =   6
      Top             =   0
      Width           =   5655
      Begin VB.Label Label2 
         Caption         =   $"FrmSave.frx":000C
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3495
         Left            =   120
         TabIndex        =   7
         Top             =   120
         Width           =   5415
      End
   End
   Begin VB.CommandButton cmdYes 
      Caption         =   "&Yes"
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
      Height          =   350
      Left            =   840
      TabIndex        =   5
      Top             =   3240
      Width           =   1125
   End
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
      Height          =   350
      Left            =   4440
      TabIndex        =   4
      Top             =   3240
      Width           =   1125
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
      Height          =   350
      Left            =   3240
      TabIndex        =   3
      Top             =   3240
      Width           =   1125
   End
   Begin VB.CommandButton cmdNo 
      Caption         =   "&No"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   350
      Left            =   2040
      TabIndex        =   2
      Top             =   3240
      Width           =   1125
   End
   Begin VB.ListBox lstDirtyFiles 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2595
      ItemData        =   "FrmSave.frx":016A
      Left            =   120
      List            =   "FrmSave.frx":016C
      MultiSelect     =   2  'Extended
      TabIndex        =   0
      Top             =   360
      Width           =   5415
   End
   Begin VB.Label Label1 
      Caption         =   "&Save changes to the following items?"
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
      Left            =   120
      TabIndex        =   1
      Top             =   90
      Width           =   3255
   End
End
Attribute VB_Name = "FrmCloseSave"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private m_bIsCancelled As Boolean
Private m_iFilesChosen() As Integer


Public Sub SelectAll()

    Dim i As Integer
    For i = 0 To lstDirtyFiles.ListCount - 1
        lstDirtyFiles.Selected(i) = True
    Next

End Sub


Public Sub AddSaveTarget(ByVal TabIndex As Integer, ByVal DisplayName As String)

    lstDirtyFiles.AddItem DisplayName
    lstDirtyFiles.ItemData(lstDirtyFiles.ListCount - 1) = TabIndex

End Sub


Public Property Get IsSaveTargetCountZero() As Boolean
    IsSaveTargetCountZero = (lstDirtyFiles.ListCount <= 0)
End Property

Public Property Get IsCancelled() As Boolean
    IsCancelled = m_bIsCancelled
End Property

Public Property Get FilesChosen() As Integer()
    FilesChosen = m_iFilesChosen
End Property

Public Property Let FilesChosen(ByRef NewValue() As Integer)
    m_iFilesChosen = NewValue
End Property


Private Sub ReturnFilesChosen()

    Dim i As Integer
    For i = 0 To lstDirtyFiles.ListCount - 1
        If lstDirtyFiles.Selected(i) = True Then
            ReDim Preserve m_iFilesChosen(UBound(m_iFilesChosen) + 1)
            m_iFilesChosen(UBound(m_iFilesChosen)) = lstDirtyFiles.ItemData(i)
        End If
    Next

End Sub


Private Sub cmdCancel_Click()
    m_bIsCancelled = True       ' Cancel button *was* pressed
    Unload Me                   ' Unload form
End Sub

Private Sub cmdNo_Click()
    lstDirtyFiles.Clear
    m_bIsCancelled = False      ' Cancel button was not pressed
    Unload Me                   ' Unload form (leaving array blank)
End Sub

Private Sub cmdYes_Click()
    ReturnFilesChosen           ' Populate array with chosen items
    m_bIsCancelled = False      ' Cancel button was not clicked
    Unload Me                   ' Unload form
End Sub

Private Sub Form_Load()
    ReDim m_iFilesChosen(0)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Cancel = 1
    Me.Hide
End Sub

