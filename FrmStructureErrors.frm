VERSION 5.00
Begin VB.Form FrmStructureErrors 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Errors Found"
   ClientHeight    =   3090
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4680
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FrmStructureErrors.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   206
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   312
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdDone 
      Cancel          =   -1  'True
      Caption         =   "Done"
      Default         =   -1  'True
      Height          =   375
      Left            =   3255
      TabIndex        =   2
      Top             =   2625
      Width           =   1275
   End
   Begin VB.TextBox txtErrors 
      Height          =   2115
      Left            =   105
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Top             =   420
      Width           =   4425
   End
   Begin VB.Label Label1 
      Caption         =   "The following errors were found:"
      Height          =   225
      Left            =   105
      TabIndex        =   0
      Top             =   105
      Width           =   2430
   End
End
Attribute VB_Name = "FrmStructureErrors"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdDone_Click()
    Unload Me
End Sub

