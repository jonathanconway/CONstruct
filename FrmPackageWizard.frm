VERSION 5.00
Begin VB.Form FrmPackageWizard 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Batch File Wizard"
   ClientHeight    =   5220
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5850
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FrmPackageWizard.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   348
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   390
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picPanel 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   750
      Left            =   1095
      ScaleHeight     =   50
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   316
      TabIndex        =   8
      Top             =   4560
      Width           =   4740
      Begin VB.CommandButton cmdNext 
         Caption         =   "&Next >"
         Height          =   345
         Left            =   1395
         TabIndex        =   12
         Top             =   165
         Width           =   1050
      End
      Begin VB.CommandButton cmdPrevious 
         Caption         =   "< &Previous"
         Enabled         =   0   'False
         Height          =   345
         Left            =   315
         TabIndex        =   11
         Top             =   165
         Width           =   1050
      End
      Begin VB.CommandButton cmdFinish 
         Caption         =   "&Finish"
         Enabled         =   0   'False
         Height          =   345
         Left            =   2475
         TabIndex        =   10
         Top             =   165
         Width           =   1050
      End
      Begin VB.CommandButton cmdClose 
         Cancel          =   -1  'True
         Caption         =   "Close"
         Height          =   345
         Left            =   3555
         TabIndex        =   9
         Top             =   165
         Width           =   1050
      End
   End
   Begin VB.PictureBox picPage 
      BorderStyle     =   0  'None
      Height          =   4350
      Index           =   0
      Left            =   1200
      ScaleHeight     =   4350
      ScaleWidth      =   4560
      TabIndex        =   0
      Top             =   105
      Width           =   4560
      Begin VB.Label Label3 
         Caption         =   "Click Next to continue."
         Height          =   255
         Left            =   0
         TabIndex        =   3
         Top             =   1200
         Width           =   4455
      End
      Begin VB.Label Label2 
         Caption         =   $"FrmPackageWizard.frx":038A
         Height          =   615
         Left            =   0
         TabIndex        =   2
         Top             =   480
         Width           =   4455
      End
      Begin VB.Label Label1 
         Caption         =   "Welcome to the Batch File Wizard"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   0
         TabIndex        =   1
         Top             =   0
         Width           =   4455
      End
   End
   Begin VB.PictureBox picPage 
      BorderStyle     =   0  'None
      Height          =   4365
      Index           =   1
      Left            =   1200
      ScaleHeight     =   4365
      ScaleWidth      =   4455
      TabIndex        =   4
      Top             =   105
      Width           =   4455
      Begin VB.Label Label6 
         Caption         =   "Welcome to the Batch File Wizard"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   7
         Top             =   1800
         Width           =   4455
      End
      Begin VB.Label Label5 
         Caption         =   $"FrmPackageWizard.frx":043B
         Height          =   615
         Left            =   -120
         TabIndex        =   6
         Top             =   2640
         Width           =   4455
      End
      Begin VB.Label Label4 
         Caption         =   "Click Next to continue."
         Height          =   255
         Left            =   0
         TabIndex        =   5
         Top             =   1200
         Width           =   4455
      End
   End
   Begin VB.Image Image1 
      Height          =   5295
      Left            =   0
      Picture         =   "FrmPackageWizard.frx":04EC
      Top             =   0
      Width           =   1095
   End
End
Attribute VB_Name = "FrmPackageWizard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private m_iSelTab As Integer


Private Sub NextTab()

    m_iSelTab = m_iSelTab + 1
    picPage(m_iSelTab).ZOrder vbBringToFront

    Dim bLast As Boolean
    bLast = (m_iSelTab = picPage.UBound)
    cmdPrevious.Enabled = True
    cmdNext.Enabled = Not bLast
    cmdFinish.Enabled = bLast

End Sub


Private Sub PreviousTab()

    m_iSelTab = m_iSelTab - 1
    picPage(m_iSelTab).ZOrder vbBringToFront
    
    Dim bFirst As Boolean
    bFirst = (m_iSelTab = 0)
    cmdPrevious.Enabled = Not bFirst
    cmdNext.Enabled = True
    cmdFinish.Enabled = False

End Sub


Private Sub cmdClose_Click()
    If m_iSelTab > 0 Then
        If MsgBox("Are you sure to want to quit this wizard?", vbQuestion + vbYesNo) = vbNo Then
            Exit Sub
        End If
    End If
    
    Unload Me
    Set FrmPackageWizard = Nothing
End Sub

Private Sub cmdNext_Click()
    NextTab
End Sub

Private Sub cmdPrevious_Click()
    PreviousTab
End Sub

Private Sub Form_Load()

    Dim lR As Integer, lG As Integer, lB As Integer
    ColorToRGB Point(0, 0), lR, lG, lB
    picPanel.BackColor = RGB(lR - 20, lG - 20, lB - 20)
    cmdClose.BackColor = picPanel.BackColor
    cmdFinish.BackColor = picPanel.BackColor
    cmdNext.BackColor = picPanel.BackColor
    cmdPrevious.BackColor = picPanel.BackColor

End Sub
