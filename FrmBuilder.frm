VERSION 5.00
Object = "{FE0065C0-1B7B-11CF-9D53-00AA003C9CB6}#1.1#0"; "COMCT232.OCX"
Begin VB.Form FrmBuilder 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Builder"
   ClientHeight    =   4320
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4950
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FrmBuilder.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   288
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   330
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox pic 
      BorderStyle     =   0  'None
      Height          =   1065
      Index           =   2
      Left            =   105
      ScaleHeight     =   71
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   316
      TabIndex        =   10
      Tag             =   "3"
      Top             =   105
      Width           =   4740
      Begin VB.Frame fra03Time 
         Caption         =   "Time"
         Height          =   960
         Left            =   0
         TabIndex        =   11
         Top             =   0
         Width           =   2640
         Begin ComCtl2.UpDown upd03Sec 
            Height          =   285
            Left            =   1380
            TabIndex        =   18
            Top             =   525
            Width           =   255
            _ExtentX        =   450
            _ExtentY        =   503
            _Version        =   327681
            BuddyControl    =   "txt03Sec"
            BuddyDispid     =   196624
            OrigLeft        =   2205
            OrigTop         =   525
            OrigRight       =   2460
            OrigBottom      =   855
            Max             =   59
            SyncBuddy       =   -1  'True
            BuddyProperty   =   0
            Enabled         =   -1  'True
         End
         Begin VB.TextBox txt03Sec 
            Height          =   285
            Left            =   945
            Locked          =   -1  'True
            TabIndex        =   15
            Text            =   "0"
            Top             =   525
            Width           =   420
         End
         Begin VB.TextBox txt03Min 
            Height          =   285
            Left            =   105
            Locked          =   -1  'True
            TabIndex        =   14
            Text            =   "0"
            Top             =   525
            Width           =   420
         End
         Begin ComCtl2.UpDown upd03Min 
            Height          =   285
            Left            =   540
            TabIndex        =   17
            Top             =   525
            Width           =   255
            _ExtentX        =   450
            _ExtentY        =   503
            _Version        =   327681
            BuddyControl    =   "txt03Min"
            BuddyDispid     =   196623
            OrigLeft        =   525
            OrigTop         =   525
            OrigRight       =   780
            OrigBottom      =   750
            Max             =   59
            SyncBuddy       =   -1  'True
            BuddyProperty   =   0
            Enabled         =   -1  'True
         End
         Begin VB.Label lbl03Time 
            Caption         =   "0:0"
            Height          =   225
            Left            =   1785
            TabIndex        =   16
            Top             =   525
            Width           =   750
         End
         Begin VB.Label lbl03Sec 
            Caption         =   "Sec"
            Height          =   225
            Left            =   945
            TabIndex        =   13
            Top             =   315
            Width           =   330
         End
         Begin VB.Label lbl03Min 
            Caption         =   "Min"
            Height          =   225
            Left            =   105
            TabIndex        =   12
            Top             =   315
            Width           =   330
         End
      End
   End
   Begin VB.PictureBox pic 
      BorderStyle     =   0  'None
      Height          =   2550
      Index           =   1
      Left            =   105
      ScaleHeight     =   170
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   316
      TabIndex        =   7
      Tag             =   "2"
      Top             =   105
      Width           =   4740
      Begin VB.ListBox lst02ActorFlags 
         Height          =   2340
         IntegralHeight  =   0   'False
         ItemData        =   "FrmBuilder.frx":058A
         Left            =   0
         List            =   "FrmBuilder.frx":05AC
         Style           =   1  'Checkbox
         TabIndex        =   9
         Top             =   210
         Width           =   2850
      End
      Begin VB.Label lbl02ActorFlags 
         Caption         =   "Actor Flags:"
         Height          =   225
         Left            =   0
         TabIndex        =   8
         Top             =   0
         Width           =   960
      End
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   3570
      TabIndex        =   4
      Top             =   3885
      Width           =   1275
   End
   Begin VB.TextBox txtValue 
      Height          =   540
      Left            =   1260
      Locked          =   -1  'True
      TabIndex        =   2
      Top             =   3150
      Width           =   3585
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   2205
      TabIndex        =   1
      Top             =   3885
      Width           =   1275
   End
   Begin VB.PictureBox pic 
      BorderStyle     =   0  'None
      Height          =   330
      Index           =   0
      Left            =   105
      ScaleHeight     =   22
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   316
      TabIndex        =   0
      Tag             =   "1"
      Top             =   105
      Visible         =   0   'False
      Width           =   4740
      Begin VB.ComboBox cbo01Actors 
         Height          =   315
         ItemData        =   "FrmBuilder.frx":0698
         Left            =   1155
         List            =   "FrmBuilder.frx":069A
         TabIndex        =   6
         Top             =   0
         Width           =   3585
      End
      Begin VB.Label lbl01Actors 
         Caption         =   "Actor:"
         Height          =   225
         Left            =   0
         TabIndex        =   5
         Top             =   0
         Width           =   1065
      End
   End
   Begin VB.Label lblValue 
      Caption         =   "Value:"
      Height          =   225
      Left            =   105
      TabIndex        =   3
      Top             =   3150
      Width           =   1080
   End
End
Attribute VB_Name = "FrmBuilder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private m_bIsCancelled As Boolean
Private m_iBuilder As Integer
Private m_bError As Boolean


Public Property Get IsCancelled() As Boolean
    IsCancelled = m_bIsCancelled
End Property

Public Property Get Error() As Boolean
    Error = m_bError
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
    m_bError = False
End Sub



Public Property Get Data() As String
    Data = txtValue.Text
End Property

Public Property Let Data(ByVal NewValue As String)
    txtValue.Text = NewValue
End Property



Public Property Get Builder() As Integer
    Builder = m_iBuilder
End Property

Public Property Let Builder(ByVal NewValue As Integer)
    
    m_iBuilder = NewValue
    
    Dim oPic As PictureBox
    Set oPic = GetPicForBuilder(m_iBuilder)
    If Not (oPic Is Nothing) Then
        ShowPic oPic
        LoadControls m_iBuilder
        SizeForm oPic
    Else
        ' TODO : Put better error message here!!
        MsgBox "Builder #" & m_iBuilder & " not found." & vbNewLine & "Please alter your definition file.", vbExclamation
        m_bError = True
    End If

End Property



Private Sub ShowPic(ByRef SelPic As PictureBox)

    Dim i As Integer
    For i = pic.LBound To pic.UBound
        pic(i).Visible = (pic(i) Is SelPic)
    Next

End Sub

Private Function GetPicForBuilder(ByVal Builder As Integer) As PictureBox

    Dim i As Integer
    For i = pic.LBound To pic.UBound
        If FixInteger(pic(i).Tag) = Builder Then
            Set GetPicForBuilder = pic(i)
            Exit Function
        End If
    Next

End Function

Private Sub SizeForm(ByRef CurrentPic As PictureBox)

    Me.Height = (CurrentPic.Top + CurrentPic.Height + 114) * Screen.TwipsPerPixelY
    
    lblValue.Top = CurrentPic.Top + CurrentPic.Height + 7
    lblValue.ZOrder vbBringToFront
    txtValue.Top = lblValue.Top
    
    cmdOK.Top = txtValue.Top + txtValue.Height + 7
    cmdCancel.Top = cmdOK.Top

End Sub

Private Sub LoadControls(ByVal Builder As Integer)

    Select Case Builder
    
        Case 1      ' Actor List
    
            Dim oStruct As CStructure
            Set oStruct = FrmMain.ceCONs(FrmMain.CurrentCON).Parser.Definition.Structures.FindItem("Actor")
            
            Dim oBlock As CBlock
            cbo01Actors.Clear
            For Each oBlock In FrmMain.ceCONs(FrmMain.CurrentCON).Parser.Blocks
                If oBlock.Structure Is oStruct Then
                    cbo01Actors.AddItem oBlock.ToString()
                End If
            Next
            
            cbo01Actors.Text = txtValue.Text
        
        Case 2      ' Actor Types (cstat, cstator)
            
            With lst02ActorFlags
                .ItemData(0) = 1
                .ItemData(1) = 2
                .ItemData(2) = 4
                .ItemData(3) = 8
                .ItemData(4) = 16
                .ItemData(5) = 32
                .ItemData(6) = 64
                .ItemData(7) = 128
                .ItemData(8) = 256
                .ItemData(9) = 32768
            End With
    
        Case 3      ' Build Time values
        
            Dim lTime As Long
            lTime = FixLong(txtValue.Text)
            lTime = lTime / 30
            
            Dim dTime As Date
            dTime = TimeSerial(0, lTime / 60, lTime Mod 60)
            lbl03Time.Caption = Minute(dTime) & ":" & Second(dTime)
            'txt03Hours.Text = Hour(dTime)
            txt03Min.Text = Minute(dTime)
            txt03Sec.Text = Second(dTime)
            
    End Select

End Sub








' 01 - Actor List
' ---------------

Private Sub cbo01Actors_Click()
    Me.Data = cbo01Actors.Text
End Sub

Private Sub lst02ActorFlags_Click()
    
    Dim lValue As Long
    Dim i As Integer
    
    For i = 0 To lst02ActorFlags.ListCount - 1
        If lst02ActorFlags.Selected(i) Then
            lValue = lValue + lst02ActorFlags.ItemData(i)
        End If
    Next
    
    Me.Data = lValue
    
End Sub

Private Sub lst02ActorFlags_ItemCheck(Item As Integer)
    lst02ActorFlags_Click
End Sub

Private Sub p03RefreshTime()
    Dim dTime As Date
    'Dim iHour As Integer,
    Dim iMin As Long, iSec As Long
    
    'iHour = FixInteger(txt03Hours.Text)
    iMin = FixInteger(txt03Min.Text)
    iSec = FixInteger(txt03Sec.Text)
    
    dTime = TimeSerial(0, _
                       iMin, _
                       iSec)
    
    lbl03Time.Caption = Minute(dTime) & ":" & Second(dTime)
    txtValue.Text = (iSec + (iMin * 60)) * 30   ' (iSec + (iMin * 60) + (iHour * 3600)) * 30
End Sub

Private Sub txt03Min_Change()
    p03RefreshTime
End Sub

Private Sub txt03Sec_Change()
    p03RefreshTime
End Sub
