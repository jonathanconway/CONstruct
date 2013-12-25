VERSION 5.00
Begin VB.Form FrmAddAI 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Add AI to the CON"
   ClientHeight    =   3555
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   6120
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3555
   ScaleWidth      =   6120
   ShowInTaskbar   =   0   'False
   Begin VB.Frame AIc 
      Caption         =   "AI creation"
      Height          =   3495
      Left            =   120
      TabIndex        =   2
      Top             =   0
      Width           =   4575
      Begin VB.ComboBox AImoveCombo 
         Height          =   315
         Left            =   120
         TabIndex        =   10
         Top             =   1200
         Width           =   2295
      End
      Begin VB.ComboBox AIactionCombo 
         Height          =   315
         Left            =   120
         TabIndex        =   9
         Top             =   840
         Width           =   2295
      End
      Begin VB.TextBox ainame 
         Height          =   285
         Left            =   120
         TabIndex        =   4
         Top             =   480
         Width           =   2295
      End
      Begin VB.ListBox RoutineList 
         Height          =   1410
         ItemData        =   "FrmAddAI.frx":0000
         Left            =   120
         List            =   "FrmAddAI.frx":0028
         Style           =   1  'Checkbox
         TabIndex        =   3
         Top             =   1800
         Width           =   2295
      End
      Begin VB.Label Label1 
         BackColor       =   &H0000FFFF&
         Caption         =   $"FrmAddAI.frx":00B7
         Height          =   1935
         Left            =   2520
         TabIndex        =   11
         Top             =   1440
         Width           =   2055
      End
      Begin VB.Label Label38 
         Caption         =   "AI name"
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
         Left            =   2520
         TabIndex        =   8
         Top             =   480
         Width           =   735
      End
      Begin VB.Label Label39 
         Caption         =   "Action name"
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
         Left            =   2520
         TabIndex        =   7
         Top             =   840
         Width           =   1095
      End
      Begin VB.Label Label40 
         Alignment       =   2  'Центровка
         Caption         =   "AI routines"
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
         Left            =   120
         TabIndex        =   6
         Top             =   1560
         Width           =   2295
      End
      Begin VB.Label Label41 
         Caption         =   "Move name"
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
         Left            =   2520
         TabIndex        =   5
         Top             =   1200
         Width           =   1095
      End
   End
   Begin VB.CommandButton CancelButton 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   4800
      TabIndex        =   1
      Top             =   600
      Width           =   1215
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "OK"
      Height          =   375
      Left            =   4800
      TabIndex        =   0
      Top             =   120
      Width           =   1215
   End
End
Attribute VB_Name = "FrmAddAI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Private Sub CancelButton_Click()
FrmAddAI.Hide
End Sub

Private Sub OKButton_Click()
Dim RRnum As Integer
Dim routname(12) As String
    
    routname(1) = "faceplayer"
    routname(2) = "geth"
    routname(3) = "getv"
    routname(4) = "randomangle"
    routname(5) = "faceplayerslow"
    routname(6) = "spin"
    routname(7) = "faceplayersmart"
    routname(8) = "fleeenemy"
    routname(9) = "jumptoplayer"
    routname(10) = "seekplayer"
    routname(11) = "furthestdir"
    routname(12) = "dodgebullet"
If ainame.Text = "" Then
  MsgBox ("AI must have a name!")
  GoTo 700
End If
Dim AIstr As String

For RRnum = 0 To 11
 If RoutineList.Selected(RRnum) = True Then SString = SString + " " + routname(RRnum + 1)
 
Next
 '!!!!!!!!!!!!!!!! next line needs to be fixed \/ m_iCurrentCON dosen't work, right part of equation is OK
 FrmMain.ceCONs(m_iCurrentCON).SelText = "ai " + UCase$(ainame.Text) + " " + UCase(AIactionCombo.Text) + " " + UCase(AImoveCombo.Text) + " " + SString
 FrmAddAI.Hide

700
End Sub
