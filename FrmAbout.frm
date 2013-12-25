VERSION 5.00
Begin VB.Form FrmAbout 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "About CONstruct"
   ClientHeight    =   30
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   570
   Icon            =   "FrmAbout.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2
   ScaleMode       =   0  'User
   ScaleWidth      =   38
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox picMask 
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      Height          =   6720
      Left            =   0
      ScaleHeight     =   6720
      ScaleWidth      =   8295
      TabIndex        =   13
      Top             =   0
      Width           =   8295
   End
   Begin VB.Timer tmrStart 
      Interval        =   1
      Left            =   1080
      Top             =   5355
   End
   Begin VB.Timer tmrBurning 
      Interval        =   100
      Left            =   3645
      Top             =   5580
   End
   Begin VB.CommandButton cmdSystemInfo 
      Caption         =   "System Info"
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
      Left            =   4395
      TabIndex        =   9
      Top             =   5220
      Width           =   1215
   End
   Begin VB.CommandButton cmdFunky 
      Cancel          =   -1  'True
      Caption         =   "Funky!"
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
      Left            =   3060
      TabIndex        =   8
      Top             =   5220
      Width           =   1215
   End
   Begin VB.Image imgBurning 
      Height          =   735
      Index           =   0
      Left            =   0
      Picture         =   "FrmAbout.frx":000C
      Stretch         =   -1  'True
      Top             =   0
      Width           =   735
   End
   Begin VB.Image imgBurning 
      Height          =   735
      Index           =   4
      Left            =   0
      Picture         =   "FrmAbout.frx":304E
      Stretch         =   -1  'True
      Top             =   0
      Width           =   735
   End
   Begin VB.Image imgBurning 
      Height          =   735
      Index           =   3
      Left            =   0
      Picture         =   "FrmAbout.frx":6090
      Stretch         =   -1  'True
      Top             =   0
      Width           =   735
   End
   Begin VB.Image imgBurning 
      Height          =   735
      Index           =   2
      Left            =   0
      Picture         =   "FrmAbout.frx":90D2
      Stretch         =   -1  'True
      Top             =   0
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   90
      Picture         =   "FrmAbout.frx":C114
      Top             =   120
      Width           =   480
   End
   Begin VB.Image imgBurning 
      Height          =   735
      Index           =   5
      Left            =   0
      Picture         =   "FrmAbout.frx":CF56
      Stretch         =   -1  'True
      Top             =   0
      Width           =   735
   End
   Begin VB.Image imgBurning 
      Height          =   735
      Index           =   1
      Left            =   0
      Picture         =   "FrmAbout.frx":FF98
      Stretch         =   -1  'True
      Top             =   0
      Width           =   735
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Online:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   840
      TabIndex        =   12
      Top             =   1530
      Width           =   825
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "A free game script editor for Duke Nukem 3D"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   840
      TabIndex        =   11
      Top             =   675
      Width           =   4770
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Contact Me:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   840
      TabIndex        =   10
      Top             =   3780
      Width           =   4770
   End
   Begin VB.Label lblContact 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   675
      Left            =   840
      TabIndex        =   7
      Top             =   4005
      Width           =   4770
   End
   Begin VB.Label lblThanksTo 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1365
      Left            =   840
      TabIndex        =   6
      Top             =   2415
      Width           =   4770
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Thanks to:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   840
      TabIndex        =   5
      Top             =   2205
      Width           =   4770
   End
   Begin VB.Label lblWebsite 
      BackStyle       =   0  'Transparent
      Caption         =   "http://jaconline.5u.com"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   1665
      TabIndex        =   4
      Top             =   1530
      Width           =   3885
   End
   Begin VB.Label lblTrademark 
      BackStyle       =   0  'Transparent
      Caption         =   "All Rights Reserved"
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
      Left            =   840
      TabIndex        =   3
      Top             =   1260
      Width           =   4770
   End
   Begin VB.Label lblCopyright 
      BackStyle       =   0  'Transparent
      Caption         =   "Copyright � 2004, Jonathan A. Conway"
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
      Left            =   840
      TabIndex        =   2
      Top             =   1035
      Width           =   4770
   End
   Begin VB.Label lblVersion 
      BackStyle       =   0  'Transparent
      Caption         =   "Version 0.2 Beta"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   840
      TabIndex        =   1
      Top             =   405
      Width           =   4770
   End
   Begin VB.Label lblProduct 
      BackStyle       =   0  'Transparent
      Caption         =   "CONstruct"
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
      Left            =   840
      TabIndex        =   0
      Top             =   120
      Width           =   4770
   End
   Begin VB.Shape shpPanel 
      BackColor       =   &H8000000F&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   6135
      Left            =   0
      Top             =   0
      Width           =   750
   End
   Begin VB.Image Image3 
      Height          =   120
      Left            =   720
      Picture         =   "FrmAbout.frx":12FDA
      Top             =   4905
      Width           =   7680
   End
   Begin VB.Image Image2 
      Height          =   120
      Left            =   720
      Picture         =   "FrmAbout.frx":1601C
      Top             =   1935
      Width           =   7680
   End
   Begin VB.Menu popup 
      Caption         =   "popup"
      Visible         =   0   'False
      Begin VB.Menu popupCopy 
         Caption         =   "&Copy"
      End
   End
End
Attribute VB_Name = "FrmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' =============================================================================
' Module Name:      FrmAbout
' Module Type:      User Form
' Description:      About window for displaying program name, version, author,
'                   etc.
' Author(s):        Jonathan A. Conway
' -----------------------------------------------------------------------------
' Log:
'
' ?? ?? ?? :
'   Created About dialog
' =============================================================================


Option Explicit

Private m_iTop As Integer
Private m_iHeight As Integer

Private m_iLeft As Integer
Private m_iWidth As Integer


' Event Handlers
' ==============

Private Sub cmdFunky_Click()
    Unload Me
    Set FrmAbout = Nothing
End Sub

Private Sub cmdSystemInfo_Click()
    ' Open System Information window
    ModSystemInfo.StartSysInfo
End Sub

Private Sub Form_Load()
    
    m_iTop = Me.Top + 3090
    m_iLeft = Me.Left + 2903
    
    Dim lr As Integer, lg As Integer, lb As Integer
    ColorToRGB Point(0, 0), lr, lg, lb
    shpPanel.BackColor = RGB(lr - 20, lg - 20, lb - 20)

    lblProduct.Caption = App.ProductName
    lblVersion.Caption = "Version " & App.Major & "." & App.Minor & "." & App.Revision & IIf(App.Major = 0, " Beta", "")
    lblCopyright.Caption = App.LegalCopyright
    lblTrademark.Caption = App.LegalTrademarks

    ' =========================================================================
    ' TO ANYONE WHOSE EYE CHANCES TO FALL ON THIS SECTION OF CODE:
    ' -------------------------------------------------------------------------
    '
    ' Through fate or design you find yourself in the heart of this project --
    ' at its very ego. The creators of the created inscribe their names and
    ' titles here to be honored and admired by fellow programmers.
    ' At this time, an utterly, repulsively evil thought may have chanced on
    ' your mind -- that of daring to step on holy ground and modify that which
    ' belongs to the Divine as to have your own name inscribed in the book of
    ' heaven. If such a horrendous, unforgivable urge has chanced to prey on
    ' your mind, reject it with all of your might, and prove your faith by
    ' instead inscribing your name at the footstoole of the Lord's throne.
    '
    ' ~ IN OTHER WORDS: ~
    '
    ' If you change the code and release it, we want some credit for starting
    ' the whole damn thing so keep all the names in -- or in the words of ID
    ' software: "burn in hell"!!
    '
    ' =========================================================================

    ' Credits
    lblThanksTo.Caption = _
    "� Repository Guy, for helping me market CONstruct" & vbNewLine & _
    "� Creators of Duke3D, for making CONstruct necessary in the first place!" & vbNewLine & _
    "� The Devastator, for a *lot* of inspiration and groundwork." & vbNewLine & _
    "� Ken Silverman, the creator of Build -- what more can I say?" & vbNewLine & _
    "� My family for taking care of all the ""human"" elements of my life!!"

    ' Contact details
    lblContact.Caption = _
    "Email: conwayj@eudoramail.com" & vbNewLine & _
    "3D Realms Forums: DrFunkenstien" & vbNewLine & _
    "Website: http://jaconline.5u.com"
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Unload Me
    Set FrmAbout = Nothing
End Sub

Private Sub lblWebsite_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = vbRightButton Then
        PopupMenu popup
    End If
End Sub

Private Sub popupCopy_Click()
    Clipboard.SetText lblWebsite.Caption
End Sub

Private Sub tmrBurning_Timer()

    Static i As Integer
    
    If i > 5 Then i = 0
    imgBurning(i).ZOrder vbBringToFront
    Image1.ZOrder vbBringToFront
    i = i + 1

End Sub

Private Sub tmrStart_Timer()
    '5805
    
    If m_iHeight >= 6180 Then
        Me.Height = 6180
        Me.Width = 5805
        tmrStart.Enabled = False
        picMask.Visible = False
    Else
        m_iHeight = m_iHeight + 400
        Me.Height = m_iHeight
        m_iTop = m_iTop - 80
        
        m_iWidth = 5805 / (6180 / m_iHeight)
        Me.Width = m_iWidth
        Me.Left = (Screen.Width / 2) - (m_iWidth / 2)
        
        Me.Top = m_iTop
    End If

End Sub
