VERSION 5.00
Begin VB.Form FrmSplash 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   3750
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7500
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   250
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   500
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer tmrHide 
      Interval        =   3500
      Left            =   2640
      Top             =   2640
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "http://jaconline.5u.com"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF8080&
      Height          =   390
      Left            =   4410
      TabIndex        =   1
      Top             =   525
      Width           =   2805
   End
   Begin VB.Label lblVersion 
      BackStyle       =   0  'Transparent
      Caption         =   "Version 1.0"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF8080&
      Height          =   390
      Left            =   4410
      TabIndex        =   0
      Top             =   210
      Width           =   2700
   End
   Begin VB.Image Image1 
      Height          =   3750
      Left            =   0
      Picture         =   "FrmSplash.frx":0000
      Top             =   0
      Width           =   7500
   End
End
Attribute VB_Name = "FrmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' =============================================================================
' Module Name:      FrmSplash
' Module Type:      User Form
' Description:      Splash screen for CONstruct; shows logo, version, etc. and
'                   waits for a second or two, then closes.
' Author(s):        Jonathan A. Conway
' -----------------------------------------------------------------------------
' Log:
'
' ?? ?? ?? :
'   Created FrmSplash form
' =============================================================================

Option Explicit


' Event Handlers
' ==============

Private Sub Form_Click()
    Unload Me
End Sub

Private Sub Form_Load()

    lblVersion.Caption = "Version " & App.Major & "." & App.Minor & "." & App.Revision & IIf(App.Major = 0, " Beta", "")
    Me.Left = (Screen.Width / 2) - (Me.Width / 2)
    Me.Top = (Screen.Height / 2) - (Me.Height / 2)

End Sub

Private Sub Image1_Click()
    Unload Me
End Sub

Private Sub lblVersion_Click()
    Unload Me
End Sub

Private Sub tmrHide_Timer()
    Unload Me
End Sub
