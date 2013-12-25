VERSION 5.00
Begin VB.Form FrmPrimitiveHelpers 
   ClientHeight    =   4950
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5715
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FrmPrimitiveHelpers.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4950
   ScaleWidth      =   5715
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox pic 
      Height          =   4215
      Index           =   0
      Left            =   0
      ScaleHeight     =   4155
      ScaleWidth      =   5625
      TabIndex        =   4
      Top             =   0
      Width           =   5685
      Begin VB.Label lbl1Actors 
         Caption         =   "##############"
         Height          =   225
         Left            =   105
         TabIndex        =   5
         Top             =   105
         Width           =   2220
      End
   End
   Begin VB.TextBox txt 
      Height          =   330
      Left            =   105
      TabIndex        =   2
      Top             =   4620
      Width           =   3585
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   2730
      TabIndex        =   1
      Top             =   4410
      Width           =   1380
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   4200
      TabIndex        =   0
      Top             =   4410
      Width           =   1380
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   435
      Left            =   735
      TabIndex        =   3
      Top             =   4515
      Width           =   1905
   End
End
Attribute VB_Name = "FrmPrimitiveHelpers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private m_bIsCancelled As Boolean
Private m_iBuilder As Integer
Private m_sData As String


Public Property Get IsCancelled() As Boolean
    IsCancelled = m_bIsCancelled
End Property


Private Sub cmdCancel_Click()
    Me.Hide
End Sub

Private Sub cmdOK_Click()
    m_bIsCancelled = False
    Me.Hide
End Sub

Private Sub Form_Load()
    m_bIsCancelled = True
End Sub



Public Property Get Data() As String
    Data = m_sData
End Property

Public Property Let Data(ByVal NewValue As String)
    m_sData = NewValue
End Property



Public Property Get Builder() As Integer
    Builder = m_iBuilder
End Property

Public Property Let Builder(ByVal NewValue As Integer)
    m_iBuilder = NewValue
    pic(m_iBuilder).ZOrder vbBringToFront
End Property

