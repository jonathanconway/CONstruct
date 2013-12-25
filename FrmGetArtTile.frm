VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form FrmGetArtTile 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Get Art Tile Number"
   ClientHeight    =   5910
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6015
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FrmGetArtTile.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   394
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   401
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtTile 
      Enabled         =   0   'False
      Height          =   300
      Left            =   3240
      TabIndex        =   14
      Top             =   990
      Width           =   615
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   4560
      TabIndex        =   13
      Top             =   600
      Width           =   1335
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   4560
      TabIndex        =   12
      Top             =   120
      Width           =   1335
   End
   Begin VB.ComboBox cboTiles 
      Enabled         =   0   'False
      Height          =   315
      Left            =   1080
      Style           =   2  'Dropdown List
      TabIndex        =   11
      Top             =   990
      Width           =   2055
   End
   Begin VB.CommandButton cmdPal 
      Caption         =   "..."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   4035
      TabIndex        =   6
      Top             =   120
      Width           =   300
   End
   Begin VB.TextBox txtPal 
      Height          =   300
      Left            =   840
      TabIndex        =   5
      Tag             =   "Path_Duke3D"
      Top             =   120
      Width           =   3135
   End
   Begin VB.CommandButton cmdPrev 
      Caption         =   "<-"
      Enabled         =   0   'False
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   960
      Width           =   375
   End
   Begin VB.CommandButton cmdNext 
      Caption         =   "->"
      Enabled         =   0   'False
      Height          =   375
      Left            =   3960
      TabIndex        =   3
      Top             =   960
      Width           =   375
   End
   Begin VB.TextBox txtFile 
      Enabled         =   0   'False
      Height          =   300
      Left            =   840
      TabIndex        =   1
      Tag             =   "Path_Duke3D"
      Top             =   480
      Width           =   3135
   End
   Begin VB.CommandButton cmdFile 
      Caption         =   "..."
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   4035
      TabIndex        =   0
      Top             =   480
      Width           =   300
   End
   Begin MSComDlg.CommonDialog cdl 
      Left            =   6600
      Top             =   6000
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame fraPreview 
      Caption         =   "Preview"
      Enabled         =   0   'False
      Height          =   4335
      Left            =   120
      TabIndex        =   8
      Top             =   1440
      Width           =   4215
      Begin VB.PictureBox Picture1 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   3975
         Left            =   120
         ScaleHeight     =   265
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   265
         TabIndex        =   9
         Top             =   240
         Width           =   3975
      End
   End
   Begin VB.Label Label1 
      Caption         =   "Tiles:"
      Enabled         =   0   'False
      Height          =   255
      Left            =   600
      TabIndex        =   10
      Top             =   1020
      Width           =   375
   End
   Begin VB.Label lblPal 
      Caption         =   "Palette:"
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   120
      Width           =   615
   End
   Begin VB.Label lblFile 
      Caption         =   "Art File:"
      Enabled         =   0   'False
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   480
      Width           =   615
   End
End
Attribute VB_Name = "FrmGetArtTile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' =============================================================================
' Module Name:      FrmGetArtTile
' Module Type:      User Form
' Description:      Main application form of CONstruct
' Author(s):        Jonathan A. Conway
' -----------------------------------------------------------------------------
' Log:
'
' 04 05 11 :
'   - Created FrmGetArtTile and made it fully functional! Yay!!
' =============================================================================


Option Explicit

Private iCurrentTile As Long
Private m_bCancelled As Boolean



Public Property Get TileChosen() As Long
    TileChosen = iCurrentTile
End Property

Public Property Get Cancelled() As Boolean
    Cancelled = m_bCancelled
End Property



Private Sub cboTiles_Change()
    Beep
    Stop
End Sub

Private Sub cboTiles_Click()
    GotoTile cboTiles.ItemData(cboTiles.ListIndex)
End Sub

Private Sub cmdCancel_Click()
    m_bCancelled = True
    Unload Me
End Sub

Private Sub cmdFile_Click()

    On Error GoTo cmdFile_Click_err
    
    With cdl
        .CancelError = True
        .InitDir = Settings.ReadSetting("Path_Art")
        .Filter = "Duke Nukem 3D Art Files (*.art)|*.art|All Files (*.*)|*.*"
        .Flags = cdlOFNFileMustExist + cdlOFNHideReadOnly
        .ShowOpen
    End With

    With txtFile
        .Text = GetFilenameFromPath(cdl.Filename)
        .SelStart = 0   ' Select all and focus
        .SelLength = Len(.Text)
        .SetFocus
    End With
    ModArt.OpenArt cdl.Filename
    
    cmdNext.Enabled = True
    cmdPrev.Enabled = True
    Label1.Enabled = True
    cboTiles.Enabled = True
    fraPreview.Enabled = True
    txtTile.Enabled = True
    GotoTile 0
    
    Exit Sub

cmdFile_Click_err:
    If err.Number <> 32755 Then
        MsgBox "ERROR: " & err.Description
    End If

End Sub

Private Sub cmdNext_Click()
    GotoTile iCurrentTile + 1
End Sub

Private Sub cmdOK_Click()
    m_bCancelled = False
    Me.Hide
End Sub

Private Sub cmdPal_Click()
    
    On Error GoTo cmdPal_Click_err
    
    With cdl
        .CancelError = True
        .Filter = "Duke Nukem 3D Palette Files (*.pal)|*.pal|All Files (*.*)|*.*"
        .Flags = cdlOFNFileMustExist + cdlOFNHideReadOnly
        .ShowOpen
    End With

    With txtPal
        .Text = GetFilenameFromPath(cdl.Filename)
        .SelStart = 0   ' Select all and focus
        .SelLength = Len(.Text)
        .SetFocus
    End With
    
    ModArt.LoadPalette cdl.Filename
    
    lblFile.Enabled = True
    txtFile.Enabled = True
    cmdFile.Enabled = True
    
    Exit Sub

cmdPal_Click_err:
    If err.Number <> 32755 Then
        MsgBox "ERROR: " & err.Description
    End If

End Sub

Private Sub cmdPrev_Click()
    GotoTile iCurrentTile - 1
End Sub


Private Sub GotoTile(ByVal TargetTile As Long)

    Dim bPixels() As Byte
    
    If TargetTile < 0 Then Exit Sub
    If TargetTile > (ModArt.ArtFile.NumTiles - 1) Then Exit Sub
    
    bPixels = ModArt.ArtFile.Tiles(TargetTile + 1).PicData.Pixels

    Dim x As Long, y As Long
    Picture1.Cls
    
    For x = LBound(bPixels, 1) To UBound(bPixels, 1)
        For y = LBound(bPixels, 2) To UBound(bPixels, 2)
            Picture1.PSet (x, y), RGB(ModArt.PalColor(bPixels(x, y)).R, ModArt.PalColor(bPixels(x, y)).G, ModArt.PalColor(bPixels(x, y)).b)
        Next
    Next
    
    iCurrentTile = TargetTile
    
    With txtTile
        .Tag = Chr(0)
        .Text = iCurrentTile
        .Tag = vbEmpty
    End With

End Sub

Private Sub Form_Load()

'    Dim oObject As Object
'    For Each oObject In FrmMain.CurrentSourceTree.Objects
'        If TypeOf oObject Is CDefine Then
'            If oObject.TypeID = [gbtDefine] Then
'                cboTiles.AddItem oObject.ToString()
'                cboTiles.ItemData(cboTiles.ListCount - 1) = oObject.GetPropertyValue(1)
'            End If
'        End If
'    Next

End Sub

Private Sub txtTile_Change()
    If Len(txtTile.Tag) = 0 Then
        If IsNumeric(txtTile.Text) Then
            txtTile.Text = Trim(txtTile.Text)
            GotoTile txtTile.Text
        Else
            txtTile.Text = 0
            GotoTile 0
        End If
    End If
End Sub
