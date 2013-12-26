VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.UserControl CONBookmarks 
   ClientHeight    =   3255
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2490
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ScaleHeight     =   217
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   166
   Begin ComctlLib.ListView lvwBookmarks 
      Height          =   3015
      Left            =   0
      TabIndex        =   1
      Top             =   240
      Width           =   2490
      _ExtentX        =   4392
      _ExtentY        =   5318
      View            =   3
      Arrange         =   1
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   327682
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   1
      NumItems        =   2
      BeginProperty ColumnHeader(1) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         Key             =   "Pos"
         Object.Tag             =   ""
         Text            =   "Pos"
         Object.Width           =   397
      EndProperty
      BeginProperty ColumnHeader(2) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   1
         Key             =   "Label"
         Object.Tag             =   ""
         Text            =   "Label"
         Object.Width           =   397
      EndProperty
   End
   Begin VB.Label lblTreeTitle 
      BackStyle       =   0  'Transparent
      Caption         =   "Bookmarks"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000009&
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   15
      Width           =   2535
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H80000002&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   255
      Left            =   0
      Top             =   0
      Width           =   2520
   End
   Begin VB.Menu mnuItem 
      Caption         =   "Item"
      Visible         =   0   'False
      Begin VB.Menu mnuItemGoto 
         Caption         =   "&Goto"
      End
      Begin VB.Menu mnuItemRemove 
         Caption         =   "&Remove"
      End
      Begin VB.Menu mnuItemBar1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuItemClear 
         Caption         =   "&Clear"
      End
      Begin VB.Menu mnuItemImportExport 
         Caption         =   "&Import/Export..."
         Enabled         =   0   'False
      End
   End
End
Attribute VB_Name = "CONBookmarks"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit


Public Event BookmarkClicked(ByVal Value As Long)


Private m_sFilename As String
Private m_oList As CList              ' String-list
Private m_bClickEngaged As Boolean



Public Function DoesBookmarkExist(ByVal Value As Long) As Boolean

    DoesBookmarkExist = False
    
    Dim oListItem As CListItem
    For Each oListItem In m_oList
        If FixLong(oListItem.Label) = Value Then
            DoesBookmarkExist = True
            Exit For
        End If
    Next

End Function


Public Sub AddBookmark(ByVal BookmarkName As String, ByVal Value As Long)
    
    Dim oNewItem As Variant

    lvwBookmarks.Sorted = False     ' Turn off sorting while adding item
    
    m_oList.AddItem Value, BookmarkName
    Settings.WriteBookmarks m_sFilename, m_oList.GenerateString()
    
    Set oNewItem = lvwBookmarks.ListItems.Add(, , Value)
    oNewItem.SubItems(1) = BookmarkName

    lvwBookmarks.Sorted = True      ' Turn sorting back on

End Sub



Private Sub InitializeBookmarks()

    Dim oListItem As CListItem      ' String-list item
    Dim oNewItem As Variant         ' Item for listview control

    lvwBookmarks.ListItems.Clear
    m_oList.ProcessString m_oList.GenerateString() & ", " & Settings.ReadBookmarks(m_sFilename)
    Settings.WriteBookmarks m_sFilename, m_oList.GenerateString()
    
    For Each oListItem In m_oList
        Set oNewItem = lvwBookmarks.ListItems.Add(, , FixLong(oListItem.Label))
        oNewItem.SubItems(1) = oListItem.Value
    Next

End Sub

Private Sub ChangeFilename(ByVal NewFilename As String)

    ' Update old filename to new filename
    
    ' Example of usage:
    '   1. User starts new document entitled "Untitled 1"
    '   2. User adds several new bookmarks to "Untitled 1"
    '   3. User decides to save the file as "game.con"
    '   4. Bookmarks filename must be updated from "Untitled 1" to "game.con"
    
    ' (See code in Property Let for Filename property)
    
    Settings.ChangeBookmarks m_sFilename, NewFilename
    m_sFilename = NewFilename

End Sub


Private Sub lvwBookmarks_Click()
    On Error GoTo ProcedureError
    
    If m_bClickEngaged Then
        m_bClickEngaged = False
    Else
        RaiseEvent BookmarkClicked(FixLong(lvwBookmarks.SelectedItem.Text))
    End If
    Exit Sub
    
ProcedureError:
    If err.Number = 91 Then
    Else
        MsgBox err.Number & vbNewLine & err.Description
    End If
End Sub


Public Property Get Filename() As String
    Filename = m_sFilename
End Property

Public Property Let Filename(ByVal Value As String)
    If Value <> m_sFilename Then
        If Len(Trim$(m_sFilename)) = 0 Then
            m_sFilename = Value
            'MsgBox "calling InitializeBookmarks(), filename value has already been changed"
            InitializeBookmarks
        Else
            'MsgBox "calling ChangeFilename()"
            ChangeFilename Value
        End If
    End If
End Property

Private Sub lvwBookmarks_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    On Error GoTo ProcedureError
    If Button = vbRightButton Then
        m_bClickEngaged = True
        
        If Not (lvwBookmarks.HitTest(x, y) Is Nothing) Then
            PopupMenu mnuItem, , , , mnuItemGoto
        End If
    
    End If
    Exit Sub
    
ProcedureError:
    If err.Number = 91 Then
    Else
        MsgBox err.Number & vbNewLine & err.Description
    End If
End Sub



Private Sub mnuItemClear_Click()
    ' Clear all bookmarks
    m_oList.Clear
    Settings.WriteBookmarks m_sFilename, m_oList.GenerateString()
    InitializeBookmarks
End Sub

Private Sub mnuItemGoto_Click()
    RaiseEvent BookmarkClicked(FixLong(lvwBookmarks.SelectedItem.Text))
End Sub

Private Sub mnuItemRemove_Click()
    Dim sLabel As String
    sLabel = lvwBookmarks.SelectedItem.Text
    m_oList.Delete sLabel
    lvwBookmarks.ListItems.Remove lvwBookmarks.SelectedItem.Index
    Settings.WriteBookmarks m_sFilename, m_oList.GenerateString()
End Sub

Private Sub UserControl_Initialize()
    m_bClickEngaged = False
    Set m_oList = New CList
End Sub

Private Sub UserControl_Resize()

    On Error GoTo ProcedureError

    With lvwBookmarks
        .Width = UserControl.ScaleWidth
        .Height = UserControl.ScaleHeight - .Top
        .ColumnHeaders("Label").Width = .Width - 60
    End With

    lblTreeTitle.Width = UserControl.ScaleWidth
    Shape1.Width = UserControl.ScaleWidth + 10
    Exit Sub
    
ProcedureError:
    If err.Number = 380 Then
    Else
        MsgBox err.Number & vbNewLine & err.Description
    End If

End Sub



