VERSION 5.00
Begin VB.Form FrmAddBookmark 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Add Bookmark"
   ClientHeight    =   1575
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4815
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FrmAddBookmark.frx":0000
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   105
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   321
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   3360
      TabIndex        =   5
      Top             =   1080
      Width           =   1335
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "Add"
      Default         =   -1  'True
      Height          =   375
      Left            =   1920
      TabIndex        =   4
      Top             =   1080
      Width           =   1335
   End
   Begin VB.TextBox txtCharPosition 
      BackColor       =   &H8000000F&
      Height          =   300
      Left            =   1440
      Locked          =   -1  'True
      TabIndex        =   3
      Top             =   480
      Width           =   1575
   End
   Begin VB.TextBox txtLabel 
      Height          =   300
      Left            =   1440
      TabIndex        =   1
      Top             =   120
      Width           =   3255
   End
   Begin VB.Label lblCharPosition 
      Caption         =   "Char Position:"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   480
      Width           =   1215
   End
   Begin VB.Label lblLabel 
      Caption         =   "Label:"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1215
   End
End
Attribute VB_Name = "FrmAddBookmark"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Option Explicit



Private m_bIsCancelled As Boolean



Public Property Get IsCancelled() As Boolean
    IsCancelled = m_bIsCancelled
End Property



Private Sub cmdAdd_Click()
    m_bIsCancelled = False
    Me.Hide
End Sub

Private Sub cmdCancel_Click()
    Me.Hide
End Sub

Private Sub Form_Load()
    m_bIsCancelled = True
End Sub
