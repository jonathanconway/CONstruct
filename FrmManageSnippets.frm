VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Begin VB.Form FrmManageSnippets 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Manage Snippets"
   ClientHeight    =   4815
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4305
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FrmManageSnippets.frx":0000
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   321
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   287
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdImport 
      Caption         =   "&Import..."
      Height          =   375
      Left            =   3045
      TabIndex        =   7
      Top             =   840
      Width           =   1170
   End
   Begin MSComDlg.CommonDialog cdl 
      Left            =   3360
      Top             =   2520
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "Delete"
      Height          =   375
      Left            =   3045
      TabIndex        =   6
      Top             =   1680
      Width           =   1170
   End
   Begin VB.CommandButton cmdEdit 
      Caption         =   "Edit"
      Height          =   375
      Left            =   3045
      TabIndex        =   5
      Top             =   1260
      Width           =   1170
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "Add"
      Height          =   375
      Left            =   3045
      TabIndex        =   4
      Top             =   420
      Width           =   1170
   End
   Begin ComctlLib.ListView lvwSnippets 
      Height          =   3690
      Left            =   105
      TabIndex        =   3
      Top             =   420
      Width           =   2850
      _ExtentX        =   5027
      _ExtentY        =   6509
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      HideColumnHeaders=   -1  'True
      _Version        =   327682
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   1
      NumItems        =   1
      BeginProperty ColumnHeader(1) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Text"
         Object.Width           =   3704
      EndProperty
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "&Help"
      CausesValidation=   0   'False
      Height          =   375
      Left            =   2940
      TabIndex        =   1
      Top             =   4305
      Width           =   1215
   End
   Begin VB.CommandButton cmdDone 
      Caption         =   "Done"
      Default         =   -1  'True
      Height          =   375
      Left            =   1575
      TabIndex        =   0
      Top             =   4305
      Width           =   1215
   End
   Begin ComctlLib.ImageList imlToolbar 
      Left            =   4095
      Top             =   3255
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   16711935
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   3
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmManageSnippets.frx":058A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmManageSnippets.frx":08DC
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmManageSnippets.frx":0C2E
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Label Label1 
      Caption         =   "Snippets:"
      Height          =   225
      Left            =   105
      TabIndex        =   2
      Top             =   210
      Width           =   750
   End
End
Attribute VB_Name = "FrmManageSnippets"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub RefreshList()
    lvwSnippets.ListItems.Clear
    
    Dim oSnip As CSnippet
    For Each oSnip In Settings.Snippets
        lvwSnippets.ListItems.Add , , oSnip.ToString()
    Next
End Sub

Private Sub cmdAdd_Click()
    On Error GoTo ProcedureError
    
    Dim oSnip As CSnippet
    Set oSnip = New CSnippet
    
    Load FrmSnippet
    With FrmSnippet
        .Snippet = oSnip
        .Show vbModal
        If Not .IsCancelled Then
            Set oSnip = .Snippet
            If Len(Trim$(oSnip.ToString)) > 0 Then
                With cdl
                    .CancelError = True
                    .DialogTitle = "Save Snippet As"
                    .Filter = "Duke Nukem 3D CON Scripts (*.con)|*.con|All Files (*.*)|*.*"
                    .FilterIndex = 0
                    .ShowSave
                    
                    oSnip.SaveFile .Filename
                    
                    Do Until IsFileExistant(.Filename)
                    Loop
                    
                    Settings.Snippets.Add oSnip
                    Settings.WriteSnippets
                    
                    RefreshList
                End With
            End If
        End If
    End With
    Unload FrmSnippet

    Exit Sub
    
ProcedureError:
    If err.Number = 32755 Then
    Else
        MsgBox err.Number & vbNewLine & err.Description
    End If
End Sub

Private Sub cmdDelete_Click()
    If Not (lvwSnippets.SelectedItem Is Nothing) Then
    
        Dim lIndex As Long
        
        With Settings.Snippets.FindItem(lvwSnippets.SelectedItem.Text)
            If MsgBox(GetFilenameFromPath(.Filename) & vbNewLine & vbNewLine & "This file will be permanently deleted." & vbNewLine & "Are you sure?", vbQuestion + vbYesNo) = vbYes Then
                Kill .Filename
                lIndex = .Index
                
                Settings.Snippets.Delete lIndex
                Settings.WriteSnippets
                RefreshList
            End If
        End With
    
    End If
End Sub

Private Sub cmdDone_Click()
    Unload Me
End Sub


Private Sub cmdEdit_Click()
    
    If lvwSnippets.ListItems.Count > 0 Then
    
        Load FrmSnippet
        With FrmSnippet
            .Snippet = Settings.Snippets.FindItem(lvwSnippets.SelectedItem.Text)
            .Show vbModal
            
            If Not .IsCancelled Then
                .Snippet.SaveFile .Snippet.Filename
                
                RefreshList
            End If
        End With
        Unload FrmSnippet
    
    End If

End Sub

Private Sub cmdImport_Click()
    
    On Error GoTo ProcedureError
    
    With cdl
        .CancelError = True
        .DialogTitle = "Import Snippets"
        .Filter = "Duke Nukem 3D CON Scripts (*.con)|*.con|All Files (*.*)|*.*"
        .FilterIndex = 0
        .Flags = cdlOFNPathMustExist
        
        .ShowOpen
        
        If IsFileExistant(.Filename) Then
            
            ' Create new snippet object
            Dim oSnippet As CSnippet
            Set oSnippet = New CSnippet
            oSnippet.LoadFile .Filename
            
            Settings.Snippets.Add oSnippet      ' Add to collection
            
            Settings.WriteSnippets              ' Save snippets list
            
            RefreshList                         ' Refresh list
        
        End If
        
    End With
    
    Exit Sub
    
ProcedureError:
    If err.Number = 32755 Then
    Else
        MsgBox err.Number & vbNewLine & err.Description
    End If
    
End Sub

Private Sub Form_Load()
    RefreshList
End Sub


Private Sub lvwSnippets_DblClick()
    If Not (lvwSnippets.SelectedItem Is Nothing) Then
        cmdEdit_Click
    End If
End Sub
