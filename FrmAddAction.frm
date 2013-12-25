VERSION 5.00
Begin VB.Form FrmAddAction 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Add Action to the CON"
   ClientHeight    =   3720
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   6000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3720
   ScaleWidth      =   6000
   ShowInTaskbar   =   0   'False
   Begin VB.Frame frameaddaaction 
      Caption         =   "Action Creation"
      Height          =   3615
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   4455
      Begin VB.TextBox Anname 
         Height          =   285
         Left            =   240
         TabIndex        =   8
         Top             =   720
         Width           =   2055
      End
      Begin VB.TextBox offset 
         Height          =   285
         Left            =   240
         TabIndex        =   7
         Top             =   1200
         Width           =   615
      End
      Begin VB.TextBox nof 
         Height          =   285
         Left            =   240
         TabIndex        =   6
         Top             =   1680
         Width           =   615
      End
      Begin VB.TextBox DirectionAn 
         Height          =   285
         Left            =   240
         TabIndex        =   5
         Top             =   2640
         Width           =   615
      End
      Begin VB.TextBox skipspr 
         Height          =   285
         Left            =   240
         TabIndex        =   4
         Top             =   2160
         Width           =   615
      End
      Begin VB.TextBox delayspr 
         Height          =   285
         Left            =   240
         TabIndex        =   3
         Top             =   3120
         Width           =   615
      End
      Begin VB.Label Label26 
         Caption         =   "Off-set"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   960
         TabIndex        =   14
         Top             =   1200
         Width           =   615
      End
      Begin VB.Label Label27 
         Caption         =   "Number of frames in action"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   960
         TabIndex        =   13
         Top             =   1680
         Width           =   2895
      End
      Begin VB.Label Label28 
         Caption         =   "Direction ( left none right --  -1 0 1 )"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   960
         TabIndex        =   12
         Top             =   2640
         Width           =   3255
      End
      Begin VB.Label Label29 
         Caption         =   "Frames to skip between shown frames"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   960
         TabIndex        =   11
         Top             =   2160
         Width           =   3375
      End
      Begin VB.Label Label30 
         Caption         =   "Delay between frames"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   960
         TabIndex        =   10
         Top             =   3120
         Width           =   2055
      End
      Begin VB.Label Label31 
         Alignment       =   2  'Центровка
         Caption         =   "Action Name"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   9
         Top             =   480
         Width           =   2055
      End
   End
   Begin VB.CommandButton CancelButton 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   4680
      TabIndex        =   1
      Top             =   600
      Width           =   1215
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "OK"
      Height          =   375
      Left            =   4680
      TabIndex        =   0
      Top             =   120
      Width           =   1215
   End
End
Attribute VB_Name = "FrmAddAction"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub Craction_Click()

End Sub

Private Sub Label1_Click()

End Sub

Private Sub CancelButton_Click()
FrmAddAction.Hide
End Sub

Private Sub OKButton_Click()
If Anname = "" Then
  MsgBox ("Action must have a name!")
  GoTo 668
End If
If offset.Text = "" Then offset.Text = "0"
If skipspr.Text = "" Then skipspr.Text = "0"
If delayspr.Text = "" Then delayspr.Text = "0"
If nof.Text = "" Then nof.Text = "0"
If DirectionAn.Text = "" Then DirectionAn.Text = "0"
 '!!!!!!!!!!!!!!!! next line needs to be fixed too \/ m_iCurrentCON dosen't work, right part of equation is OK
 FrmMain.ceCONs(m_iCurrentCON).SelText = "action " + UCase$(Anname.Text) + " " + offset.Text + " " + nof.Text + " " + skipspr.Text + " " + DirectionAn.Text + " " + delayspr.Text
FrmAddAction.Hide

668
End Sub
