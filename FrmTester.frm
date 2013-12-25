VERSION 5.00
Begin VB.Form FrmTester 
   Caption         =   "Form1"
   ClientHeight    =   7290
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8235
   Icon            =   "FrmTester.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7290
   ScaleWidth      =   8235
   StartUpPosition =   3  'Windows Default
   Begin CONstruct.CONEditor CONEditor1 
      Height          =   3060
      Left            =   3570
      TabIndex        =   8
      Top             =   3150
      Width           =   4320
      _ExtentX        =   7620
      _ExtentY        =   5398
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Change"
      Height          =   375
      Left            =   6600
      TabIndex        =   7
      Top             =   1440
      Width           =   1215
   End
   Begin VB.TextBox txtMax 
      Height          =   375
      Left            =   5280
      TabIndex        =   6
      Text            =   "4"
      Top             =   1440
      Width           =   1215
   End
   Begin VB.CommandButton Command3 
      Caption         =   "List => String"
      Height          =   495
      Left            =   1680
      TabIndex        =   5
      Top             =   5880
      Width           =   1455
   End
   Begin VB.CommandButton Command2 
      Caption         =   "String => List"
      Height          =   495
      Left            =   120
      TabIndex        =   4
      Top             =   5880
      Width           =   1455
   End
   Begin VB.TextBox txtString 
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   5400
      Width           =   3135
   End
   Begin VB.TextBox txt 
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   360
      Width           =   2895
   End
   Begin VB.ListBox lst 
      Height          =   4155
      Left            =   120
      TabIndex        =   1
      Top             =   1080
      Width           =   4455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Add"
      Height          =   375
      Left            =   3120
      TabIndex        =   0
      Top             =   360
      Width           =   1455
   End
End
Attribute VB_Name = "FrmTester"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit


Private m_oMRUList As CMRUList


Private Sub RefreshList()

    lst.Clear
    
    Dim sItems() As String
    Dim i As Integer
    
    sItems = m_oMRUList.Items
    
    For i = LBound(sItems) To UBound(sItems)
        lst.AddItem i & " - """ & sItems(i) & """"
    Next

End Sub


Private Sub Command1_Click()

    m_oMRUList.AddMRUItem txt
    RefreshList

End Sub





Private Sub Command2_Click()

    m_oMRUList.LoadFromString txtString.Text
    RefreshList

End Sub

Private Sub Command3_Click()
    
    txtString.Text = m_oMRUList.ToString()

End Sub

Private Sub Command4_Click()
    m_oMRUList.MaxItems = FixInteger(txtMax.Text)
End Sub

Private Sub Form_Load()
    Set m_oMRUList = New CMRUList
    CONEditor1.Locked = False
    CONEditor1.Enabled = True
End Sub
