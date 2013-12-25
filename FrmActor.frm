VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmActor 
   Caption         =   "Actor"
   ClientHeight    =   5175
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5640
   Icon            =   "FrmActor.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5175
   ScaleWidth      =   5640
   Begin VB.CommandButton cmdHelp 
      Caption         =   "&Help"
      CausesValidation=   0   'False
      Height          =   375
      Left            =   4320
      TabIndex        =   3
      Top             =   4680
      Width           =   1215
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      CausesValidation=   0   'False
      Height          =   375
      Left            =   3000
      TabIndex        =   2
      Top             =   4680
      Width           =   1215
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      Default         =   -1  'True
      Height          =   375
      Left            =   1680
      TabIndex        =   1
      Top             =   4680
      Width           =   1215
   End
   Begin VB.PictureBox picTabs 
      BorderStyle     =   0  'None
      Height          =   3255
      Index           =   0
      Left            =   240
      ScaleHeight     =   217
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   345
      TabIndex        =   0
      Top             =   1200
      Width           =   5175
   End
   Begin MSComctlLib.TabStrip tabstrip 
      Height          =   3735
      Left            =   120
      TabIndex        =   4
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
      TabIndex        =   5
      Top             =   1200
      Width           =   5175
      Begin CONstruct.CONEditor ceCode 
         Height          =   3135
         Left            =   0
         TabIndex        =   6
         Top             =   120
         Width           =   5175
         _ExtentX        =   9128
         _ExtentY        =   5530
      End
   End
   Begin VB.Label lblType 
      BackStyle       =   0  'Transparent
      Caption         =   "Actor/UserActor"
      Height          =   255
      Left            =   960
      TabIndex        =   9
      Top             =   240
      Width           =   2655
   End
   Begin VB.Label lblPosition 
      BackStyle       =   0  'Transparent
      Caption         =   "Position: # to #"
      Height          =   255
      Left            =   960
      TabIndex        =   8
      Top             =   480
      Width           =   2655
   End
   Begin VB.Label lblName 
      BackStyle       =   0  'Transparent
      Caption         =   "LIZSPITID"
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
      TabIndex        =   7
      Top             =   0
      Width           =   2655
   End
   Begin VB.Image Image1 
      Height          =   720
      Left            =   120
      Picture         =   "FrmActor.frx":058A
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
Attribute VB_Name = "FrmActor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

