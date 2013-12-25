VERSION 5.00
Begin VB.UserControl CONTimeEdit 
   ClientHeight    =   390
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1230
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ScaleHeight     =   26
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   82
   Begin VB.PictureBox pic 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   300
      Left            =   30
      ScaleHeight     =   20
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   73
      TabIndex        =   0
      Top             =   30
      Width           =   1095
      Begin VB.TextBox txtMin 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   510
         TabIndex        =   3
         Top             =   15
         Width           =   480
      End
      Begin VB.TextBox txtHour 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   15
         TabIndex        =   1
         Top             =   15
         Width           =   480
      End
      Begin VB.Label lblColon 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   ":"
         Height          =   255
         Left            =   480
         TabIndex        =   2
         Top             =   0
         Width           =   75
      End
   End
   Begin VB.TextBox Text1 
      Enabled         =   0   'False
      Height          =   375
      Left            =   0
      Locked          =   -1  'True
      TabIndex        =   4
      Top             =   0
      Width           =   1215
   End
End
Attribute VB_Name = "CONTimeEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit





Private Sub txtHour_Change()
    Static sTextCache As String
    If Trim(txtHour.Text) = "" Then txtHour.Text = "0"
    If IsNumeric(txtHour.Text) = False Then
        txtHour.Text = sTextCache
    Else
        If Len(txtHour.Text) > 2 Then
            txtHour.Text = Left$(txtHour.Text, 2)
            txtMin.Text = Mid$(txtHour.Text, 2, 2)
            txtMin.SelStart = 2
            txtMin.SetFocus
        End If
        sTextCache = txtHour.Text
    End If
End Sub

Private Sub txtHour_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyRight Then
        If txtHour.SelStart = 2 Then
            With txtMin
                .SelStart = 0
                .SelLength = 0
                .SetFocus
            End With
        End If
    End If
End Sub

Private Sub txtMin_Change()
    Static sTextCache As String
    If Trim(txtMin.Text) = "" Then txtMin.Text = "0"
    If Len(txtMin.Text) > 2 Or IsNumeric(txtMin.Text) = False _
        Or (IsNumeric(txtMin.Text) And CInt(txtMin.Text) > 59) Then
        txtMin.Text = sTextCache
    Else
        sTextCache = txtMin.Text
    End If
End Sub

Private Sub txtMin_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyLeft Then
        If txtMin.SelStart = 0 Then
            With txtHour
                .SelStart = 2
                .SelLength = 0
                .SetFocus
            End With
        End If
    End If
End Sub

Private Sub UserControl_Resize()
    pic.Width = UserControl.ScaleWidth - 4
    pic.Height = UserControl.ScaleHeight - 4
    
    txtHour.Width = ((pic.Width / 2) - (lblColon.Width / 2)) - 1
    txtHour.Height = pic.Height
    
    lblColon.Left = txtHour.Width + 1
    
    txtMin.Left = lblColon.Left + lblColon.Width
    txtMin.Width = (pic.Width / 2) - (lblColon.Width / 2)
    txtMin.Height = pic.Height
    
    
    Text1.Width = UserControl.ScaleWidth
    Text1.Height = UserControl.ScaleHeight
    Text1.ZOrder vbSendToBack
End Sub

Public Property Get Text() As String
    Text = txtHour.Text & ":" & txtMin.Text
End Property

Public Property Let Text(ByVal NewValue As String)
    Dim iInstr As Integer
    Dim sValue As String
    sValue = NewValue
    If Len(Trim(sValue)) = 0 Then sValue = "0:0"
    If IsWordTime(sValue) Then
        iInstr = InStr(1, sValue, ":")
        txtHour.Text = Left$(sValue, iInstr - 1)
        txtMin.Text = Right$(sValue, Len(sValue) - iInstr)
    Else
        txtHour.Text = "0"
        txtMin.Text = "0"
    End If
End Property
