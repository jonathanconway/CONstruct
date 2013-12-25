VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MsComCtl.ocx"
Begin VB.Form FrmMain 
   ClientHeight    =   6270
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   8640
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FrmMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   418
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   576
   StartUpPosition =   3  'Windows Default
   Tag             =   "[CmdGroup]=OpenFileRequired"
   Begin ComctlLib.Toolbar tlbCode 
      Align           =   1  'Align Top
      Height          =   390
      Left            =   0
      TabIndex        =   11
      Top             =   390
      Width           =   8640
      _ExtentX        =   15240
      _ExtentY        =   688
      ButtonWidth     =   609
      ButtonHeight    =   582
      Appearance      =   1
      ImageList       =   "imlCodeToolbar"
      _Version        =   327682
      BeginProperty Buttons {0713E452-850A-101B-AFC0-4210102A8DA7} 
         NumButtons      =   10
         BeginProperty Button1 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "InsertStructure"
            Object.ToolTipText     =   "Insert Structure"
            Object.Tag             =   "[CmdGroup]=OpenFileRequired; "
            ImageKey        =   "Structure"
         EndProperty
         BeginProperty Button2 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "InsertPrimitive"
            Object.ToolTipText     =   "Insert Primitive"
            Object.Tag             =   "[CmdGroup]=OpenFileRequired; "
            ImageKey        =   "Primitive"
         EndProperty
         BeginProperty Button3 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button4 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Bookmark"
            Object.ToolTipText     =   "Insert Bookmark"
            Object.Tag             =   "[CmdGroup]=OpenFileRequired; "
            ImageKey        =   "Bookmark"
         EndProperty
         BeginProperty Button5 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Snippet"
            Object.ToolTipText     =   "Snippet"
            Object.Tag             =   ""
            ImageKey        =   "Snippet"
         EndProperty
         BeginProperty Button6 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button7 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Outdent"
            Object.ToolTipText     =   "Outdent"
            Object.Tag             =   "[CmdGroup]=OpenFileRequired; "
            ImageKey        =   "Outdent"
         EndProperty
         BeginProperty Button8 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Indent"
            Object.ToolTipText     =   "Indent"
            Object.Tag             =   "[CmdGroup]=OpenFileRequired; "
            ImageKey        =   "Indent"
         EndProperty
         BeginProperty Button9 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Comment"
            Object.ToolTipText     =   "Comment"
            Object.Tag             =   "[CmdGroup]=OpenFileRequired; "
            ImageKey        =   "Comment"
         EndProperty
         BeginProperty Button10 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Uncomment"
            Object.ToolTipText     =   "Uncomment"
            Object.Tag             =   "[CmdGroup]=OpenFileRequired; "
            ImageKey        =   "Uncomment"
         EndProperty
      EndProperty
      Begin VB.CommandButton cmdClose 
         Caption         =   "X"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   8175
         TabIndex        =   12
         Tag             =   "[CmdGroup]=OpenFileRequired; "
         ToolTipText     =   "Close"
         Top             =   75
         Width           =   315
      End
   End
   Begin ComctlLib.Toolbar tlbMain 
      Align           =   1  'Align Top
      Height          =   390
      Left            =   0
      TabIndex        =   10
      Top             =   0
      Width           =   8640
      _ExtentX        =   15240
      _ExtentY        =   688
      ButtonWidth     =   609
      ButtonHeight    =   582
      ImageList       =   "imlMainToolbar"
      _Version        =   327682
      BeginProperty Buttons {0713E452-850A-101B-AFC0-4210102A8DA7} 
         NumButtons      =   15
         BeginProperty Button1 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "New"
            Object.ToolTipText     =   "New"
            Object.Tag             =   ""
            ImageKey        =   "New"
         EndProperty
         BeginProperty Button2 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Open"
            Object.ToolTipText     =   "Open"
            Object.Tag             =   ""
            ImageKey        =   "Open"
         EndProperty
         BeginProperty Button3 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Save"
            Object.ToolTipText     =   "Save"
            Object.Tag             =   "[CmdGroup]=OpenFileRequired; "
            ImageKey        =   "Save"
         EndProperty
         BeginProperty Button4 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "SaveAll"
            Object.ToolTipText     =   "Save All"
            Object.Tag             =   "[CmdGroup]=OpenFileRequired; "
            ImageKey        =   "SaveAll"
         EndProperty
         BeginProperty Button5 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button6 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Cut"
            Object.ToolTipText     =   "Cut"
            Object.Tag             =   "[CmdGroup]=OpenFileRequired; "
            ImageKey        =   "Cut"
         EndProperty
         BeginProperty Button7 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Copy"
            Object.ToolTipText     =   "Copy"
            Object.Tag             =   "[CmdGroup]=OpenFileRequired; "
            ImageKey        =   "Copy"
         EndProperty
         BeginProperty Button8 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Paste"
            Object.ToolTipText     =   "Paste"
            Object.Tag             =   "[CmdGroup]=OpenFileRequired; "
            ImageKey        =   "Paste"
         EndProperty
         BeginProperty Button9 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Find"
            Object.ToolTipText     =   "Find"
            Object.Tag             =   "[CmdGroup]=OpenFileRequired; "
            ImageKey        =   "Find"
         EndProperty
         BeginProperty Button10 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button11 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Undo"
            Object.ToolTipText     =   "Undo"
            Object.Tag             =   "[CmdGroup]=OpenFileRequired; "
            ImageKey        =   "Undo"
         EndProperty
         BeginProperty Button12 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Redo"
            Object.ToolTipText     =   "Redo"
            Object.Tag             =   "[CmdGroup]=OpenFileRequired; "
            ImageKey        =   "Redo"
         EndProperty
         BeginProperty Button13 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button14 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "RunDuke"
            Object.ToolTipText     =   "Run Duke Nukem 3D"
            Object.Tag             =   ""
            ImageKey        =   "RunDuke"
         EndProperty
         BeginProperty Button15 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "CONFilter"
            Object.ToolTipText     =   "CON Filter"
            Object.Tag             =   "[CmdGroup]=OpenFileRequired; "
            ImageKey        =   "Filter"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imlStructIcons 
      Left            =   8295
      Top             =   1365
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   16711935
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMain.frx":058A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMain.frx":08DC
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMain.frx":0C2E
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMain.frx":0F80
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Timer tmrFlashStatus 
      Enabled         =   0   'False
      Interval        =   3000
      Left            =   8295
      Top             =   4620
   End
   Begin CONstruct.LongTimer tmrAutoSave 
      Left            =   8295
      Top             =   2100
      _ExtentX        =   741
      _ExtentY        =   741
      Enabled         =   0   'False
   End
   Begin VB.Timer tmrDoUpdate 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   8190
      Top             =   4830
   End
   Begin VB.Timer tmrCodeUpdate 
      Interval        =   2000
      Left            =   8160
      Top             =   2520
   End
   Begin CONstruct.CONEditor ceCONs 
      Height          =   4470
      Index           =   0
      Left            =   2745
      TabIndex        =   9
      Top             =   1275
      Visible         =   0   'False
      Width           =   5700
      _ExtentX        =   10054
      _ExtentY        =   7885
   End
   Begin ComctlLib.ProgressBar prg 
      Height          =   240
      Left            =   0
      TabIndex        =   8
      Top             =   6000
      Visible         =   0   'False
      Width           =   4290
      _ExtentX        =   7567
      _ExtentY        =   423
      _Version        =   327682
      Appearance      =   0
   End
   Begin ComctlLib.StatusBar barStatus 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   6
      Top             =   5910
      Width           =   8640
      _ExtentX        =   15240
      _ExtentY        =   635
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   5
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            AutoSize        =   1
            Object.Width           =   9499
            MinWidth        =   5292
            Text            =   "Ready"
            TextSave        =   "Ready"
            Key             =   "Status"
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel2 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Object.Width           =   2646
            MinWidth        =   2646
            Key             =   "Cursor"
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel3 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Style           =   1
            Enabled         =   0   'False
            Object.Width           =   926
            MinWidth        =   926
            TextSave        =   "CAPS"
            Key             =   "Caps"
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel4 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Style           =   2
            Object.Width           =   847
            MinWidth        =   847
            TextSave        =   "NUM"
            Key             =   "Num"
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel5 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Style           =   3
            Enabled         =   0   'False
            Object.Width           =   688
            MinWidth        =   688
            TextSave        =   "INS"
            Key             =   "Ins"
            Object.Tag             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComDlg.CommonDialog cdl 
      Left            =   4200
      Top             =   360
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.PictureBox picProp 
      BorderStyle     =   0  'None
      Height          =   2415
      Left            =   0
      ScaleHeight     =   161
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   169
      TabIndex        =   1
      Top             =   3450
      Width           =   2535
      Begin CONstruct.CONBookmarks cbBookmarks 
         Height          =   2340
         Index           =   0
         Left            =   0
         TabIndex        =   5
         Top             =   0
         Visible         =   0   'False
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   4128
      End
   End
   Begin VB.PictureBox picTree 
      BorderStyle     =   0  'None
      Height          =   2655
      Left            =   0
      ScaleHeight     =   177
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   169
      TabIndex        =   0
      Top             =   720
      Width           =   2535
      Begin CONstruct.CONTree ctTrees 
         Height          =   2505
         Index           =   0
         Left            =   0
         TabIndex        =   4
         Top             =   75
         Visible         =   0   'False
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   4419
      End
   End
   Begin ComctlLib.TabStrip tabFiles 
      Height          =   5040
      Left            =   2625
      TabIndex        =   7
      Top             =   825
      Visible         =   0   'False
      Width           =   5940
      _ExtentX        =   10478
      _ExtentY        =   8890
      _Version        =   327682
      BeginProperty Tabs {0713E432-850A-101B-AFC0-4210102A8DA7} 
         NumTabs         =   1
         BeginProperty Tab1 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
   Begin ComctlLib.ImageList imlMainToolbar 
      Left            =   8160
      Top             =   3840
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   16711935
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   12
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmMain.frx":12D2
            Key             =   "New"
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmMain.frx":1624
            Key             =   "Open"
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmMain.frx":1976
            Key             =   "Save"
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmMain.frx":1CC8
            Key             =   "SaveAll"
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmMain.frx":201A
            Key             =   "Cut"
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmMain.frx":236C
            Key             =   "Copy"
         EndProperty
         BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmMain.frx":26BE
            Key             =   "Paste"
         EndProperty
         BeginProperty ListImage8 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmMain.frx":2A10
            Key             =   "Find"
         EndProperty
         BeginProperty ListImage9 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmMain.frx":2D62
            Key             =   "Undo"
         EndProperty
         BeginProperty ListImage10 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmMain.frx":30B4
            Key             =   "Redo"
         EndProperty
         BeginProperty ListImage11 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmMain.frx":3406
            Key             =   "RunDuke"
         EndProperty
         BeginProperty ListImage12 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmMain.frx":3758
            Key             =   "Filter"
         EndProperty
      EndProperty
   End
   Begin ComctlLib.ImageList imlCodeToolbar 
      Left            =   8160
      Top             =   3000
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   16711935
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   8
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmMain.frx":3AAA
            Key             =   "Structure"
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmMain.frx":3DFC
            Key             =   "Primitive"
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmMain.frx":414E
            Key             =   "Bookmark"
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmMain.frx":44A0
            Key             =   "Snippet"
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmMain.frx":47F2
            Key             =   "Outdent"
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmMain.frx":4B44
            Key             =   "Indent"
         EndProperty
         BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmMain.frx":4E96
            Key             =   "Comment"
         EndProperty
         BeginProperty ListImage8 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmMain.frx":51E8
            Key             =   "Uncomment"
         EndProperty
      EndProperty
   End
   Begin VB.Label lblVSplitter 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5175
      Left            =   2490
      MousePointer    =   9  'Size W E
      TabIndex        =   3
      Top             =   720
      Width           =   105
   End
   Begin VB.Label lblHSplitter 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Left            =   0
      MousePointer    =   7  'Size N S
      TabIndex        =   2
      Top             =   3315
      Width           =   2535
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileNew 
         Caption         =   "&New"
         Shortcut        =   ^N
      End
      Begin VB.Menu mnuFileOpen 
         Caption         =   "&Open..."
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuFileClose 
         Caption         =   "&Close"
         Shortcut        =   ^{F4}
      End
      Begin VB.Menu mnuFileCloseAll 
         Caption         =   "Clos&e All"
      End
      Begin VB.Menu mnuFileBar1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileSave 
         Caption         =   "&Save"
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuFileSaveAs 
         Caption         =   "Save &As..."
      End
      Begin VB.Menu mnuFileSaveCopyAs 
         Caption         =   "Save &Copy As..."
      End
      Begin VB.Menu mnuFileSaveAll 
         Caption         =   "Save A&ll"
      End
      Begin VB.Menu mnuFileMRU 
         Caption         =   "-"
         Index           =   0
         Visible         =   0   'False
      End
      Begin VB.Menu mnuFileBar2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "&Edit"
      Begin VB.Menu mnuEditUndo 
         Caption         =   "&Undo"
         Shortcut        =   ^Z
      End
      Begin VB.Menu mnuEditRedo 
         Caption         =   "&Redo"
         Shortcut        =   ^Y
      End
      Begin VB.Menu mnuEditBar1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditCut 
         Caption         =   "Cu&t"
      End
      Begin VB.Menu mnuEditCopy 
         Caption         =   "&Copy"
      End
      Begin VB.Menu mnuEditPaste 
         Caption         =   "&Paste"
      End
      Begin VB.Menu mnuEditClear 
         Caption         =   "Cle&ar"
         Shortcut        =   {DEL}
      End
      Begin VB.Menu mnuEditBar5 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditSelectAll 
         Caption         =   "&Select All"
         Shortcut        =   ^A
      End
      Begin VB.Menu mnuEditSelectLine 
         Caption         =   "Select &Line"
      End
      Begin VB.Menu mnuEditBar2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditAddBookmark 
         Caption         =   "&Add Bookmark..."
         Shortcut        =   ^B
      End
      Begin VB.Menu mnuEditBar3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditAdvanced 
         Caption         =   "&Advanced"
         Begin VB.Menu mnuEditAdvancedToUpper 
            Caption         =   "To &Uppercase"
            Shortcut        =   ^U
         End
         Begin VB.Menu mnuEditAdvancedToLower 
            Caption         =   "To &Lowercase"
            Shortcut        =   ^L
         End
         Begin VB.Menu mnuEditAdvancedComment 
            Caption         =   "&Comment Selection"
         End
         Begin VB.Menu mnuEditAdvancedUncomment 
            Caption         =   "U&ncomment Selection"
         End
         Begin VB.Menu mnuEditAdvancedIndent 
            Caption         =   "&Increase Line Indent"
         End
         Begin VB.Menu mnuEditAdvancedOutdent 
            Caption         =   "&Decrease Line Indent"
         End
      End
      Begin VB.Menu mnuEditBar4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditFind 
         Caption         =   "&Find..."
         Shortcut        =   ^F
      End
      Begin VB.Menu mnuEditFindNext 
         Caption         =   "Find &Next"
         Shortcut        =   {F3}
      End
      Begin VB.Menu mnuEditReplace 
         Caption         =   "R&eplace..."
         Shortcut        =   ^H
      End
      Begin VB.Menu mnuEditGoto 
         Caption         =   "&Goto..."
      End
   End
   Begin VB.Menu mnuInsert 
      Caption         =   "&Insert"
      Begin VB.Menu mnuInsertStructure 
         Caption         =   "&Structure..."
      End
      Begin VB.Menu mnuInsertPrimitive 
         Caption         =   "&Primitive..."
      End
      Begin VB.Menu mnuInsertBar2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuInsertDateTime 
         Caption         =   "&Date/Time"
         Shortcut        =   {F5}
      End
      Begin VB.Menu mnuInsertCommentHeader 
         Caption         =   "&Comment Header"
      End
      Begin VB.Menu mnuInsertBar1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuInsertFromFile 
         Caption         =   "&From File..."
      End
   End
   Begin VB.Menu mnuSnippets 
      Caption         =   "&Snippets"
      Begin VB.Menu mnuSnippet 
         Caption         =   "[None]"
         Enabled         =   0   'False
         Index           =   0
      End
      Begin VB.Menu mnuSnippetsBar1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSnippetsManage 
         Caption         =   "&Manage Snippets..."
      End
   End
   Begin VB.Menu mnuView 
      Caption         =   "&View"
      Begin VB.Menu mnuViewMainToolbar 
         Caption         =   "&Main Toolbar"
      End
      Begin VB.Menu mnuViewCodeToolbar 
         Caption         =   "&Code Toolbar"
      End
      Begin VB.Menu mnuViewStatusBar 
         Caption         =   "&Statusbar"
      End
      Begin VB.Menu mnuViewBar1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewAutoFormatting 
         Caption         =   "&AutoFormatting"
         Begin VB.Menu mnuViewAutoFormattingEnable 
            Caption         =   "&Enable AutoFormat"
         End
         Begin VB.Menu mnuViewAutoFormattingReFormatCode 
            Caption         =   "&Re-format Code"
         End
         Begin VB.Menu mnuViewAutoFormattingClear 
            Caption         =   "&Clear Formatting"
         End
      End
   End
   Begin VB.Menu mnuTools 
      Caption         =   "&Tools"
      Begin VB.Menu mnuToolsBulkIndenter 
         Caption         =   "&Bulk Indenter..."
      End
      Begin VB.Menu mnuToolsFilter 
         Caption         =   "&CON Filter..."
      End
      Begin VB.Menu mnuToolsBar1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuToolsRun 
         Caption         =   "&Run Duke Nukem 3D"
         Shortcut        =   {F8}
      End
      Begin VB.Menu mnuToolsDefinitionEditor 
         Caption         =   "&Definition Editor..."
      End
      Begin VB.Menu mnuToolsBar2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuToolsOptions 
         Caption         =   "&Options..."
      End
   End
   Begin VB.Menu mnuBlock 
      Caption         =   "&Block"
      Visible         =   0   'False
      Begin VB.Menu mnuBlockGoto 
         Caption         =   "&Goto"
      End
      Begin VB.Menu mnuBlockBar1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuBlockAdd 
         Caption         =   "&Add..."
      End
      Begin VB.Menu mnuBlockEdit 
         Caption         =   "&Edit..."
      End
      Begin VB.Menu mnuBlockRemove 
         Caption         =   "&Remove"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuHelpAbout 
         Caption         =   "&About..."
      End
   End
End
Attribute VB_Name = "FrmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' =============================================================================
' Module Name:      FrmMain
' Module Type:      User Form
' Description:      Main application form of CONstruct
' Author(s):        Jonathan A. Conway
' -----------------------------------------------------------------------------
' Log:
'
' 04 08 11 :
'   - Implemented MRU (Most-Recently-Used) list functionality, which uses the
'     class CMRUList.
'
' 04 07 27 :
'   - Made the form totally XP look-n-feel. Major modifications to controls
'     and some less major modifications to the form code.
'   - Cleaned up some of the resizing/positioning code, fixed left-panel
'     skewing on form maximize bug.
'
' 04 07 24 :
'   - Implemented bookmarks functionality and connected CONBookmarks control
'       - Added cbBookmarks control instance, index=0
'       - Added "BookmarkClicked" event handler for the control
'       - Added all necessary supportive code for cbBookmarks (i.e. create,
'         destroy, resize, etc.)
'
' 04 07 18 :
'   - Fixed a problem in SaveCON() that made the Save dialog appear when
'     saving a pre-existing file. The mechanism isn't perfect yet -- see
'     in-code comments for details...
'
' 04 07 16 :
'   - Removed unloading code for FrmPrimitive as to preserve previous values
'
' 04 07 14 :
'   - Did a "face-lift" of toolbar button click code
'   - Added a new "code" toolbar with a button...
'   - ...and some code to launch FrmPrimitive when button is clicked
'
' 04 05 11 :
'   - Added CurrentSourceTree property (readonly)
'   - Open/Save dialogs now initialize in the folder specified for CONs.
'
' 04 05 05 :
'   - Fixed bugs in Find/Replace/Goto
'   - Added method for launching Duke Nukem 3D
'   - Added method for showing the Options form
'   - Misc. mods and fixes.
' =============================================================================

Option Explicit


' Private Enumerations
' ====================

Private Enum eFindReplaceMode
    [frFind] = 0
    [frReplace] = 1
    [frGoto] = 2
End Enum


' Private Variables
' =================

Private m_bTabStop() As Boolean         ' Tab-stop setting of each control
Private m_iCurrentCON As Integer        ' Index of current CON

Private m_oDefinition As CDefinition    ' Parser definition object

Private m_oMRUList As CMRUList

Private m_bBackgroundParsing As Boolean
Private m_iAutoSaveInterval As Integer

Private m_sPreviousStatusText As String



' Public Properties
' =================

Public Property Get CurrentCON() As Integer
    CurrentCON = m_iCurrentCON
End Property

Public Property Get Definition() As CDefinition
    If m_oDefinition Is Nothing Then InitDefinition
    Set Definition = m_oDefinition
End Property



' Private Procedures
' ==================

' Form/Controls
' -------------

Private Sub FlashStatus(ByVal Message As String)

    m_sPreviousStatusText = barStatus.Panels(1).Text
    barStatus.Panels(1).Text = Message
    tmrFlashStatus.Enabled = True

End Sub

Private Sub UpdateSelStatus(ByVal SelLine As Long, ByVal SelColumn As Integer)

    ' TODO : Speed this up a bit somehow to prevent visible re-drawing
    '        each time the user changes the cursor position
    Dim sText As String
    sText = "Row " & SelLine & "; Col " & SelColumn
    With barStatus.Panels("Cursor")
        If .Text <> sText Then .Text = sText
        If Not .Enabled Then .Enabled = True
    End With
    
    'UpdateCurrentObject
    'tmrDoUpdate.Enabled = True
    
End Sub

Private Sub UpdateCurrentObject(ByRef Block As CBlock)

    'With ceCONs(m_iCurrentCON)
    '    Set oBlock = .Parser.GetObjectAtPos(.Text, .SelStart + 1)
    'End With

    If Not (Block Is Nothing) Then
        With barStatus.Panels("Status")
            .Text = Block.Structure.StructureName & ": " & Block.ToString()
            .Picture = imlStructIcons.ListImages(Block.Structure.ImageID).ExtractIcon()
        End With
        'barStatus.Panels
        ctTrees(m_iCurrentCON).SelectBlock Block
    Else
        With barStatus.Panels("Status")
            .Text = "Ready"
            .Picture = Nothing
        End With
    End If

End Sub

Private Sub UpdateEditMenu(ByVal CONIndex As Integer)
    
    Dim bCanUndo As Boolean
    Dim sCaption As String
    
    With ceCONs(CONIndex)
    
        ' Undo options
        bCanUndo = .CanUndo
        If (bCanUndo) Then
            sCaption = "&Undo " & TranslateUndoType(.UndoType)
        Else
            sCaption = "&Undo"
        End If
        If mnuEditUndo.Caption <> sCaption Then mnuEditUndo.Caption = sCaption
        If mnuEditUndo.Enabled <> bCanUndo Then mnuEditUndo.Enabled = bCanUndo
        If tlbMain.Buttons("Undo").Enabled <> bCanUndo Then _
            tlbMain.Buttons("Undo").Enabled = bCanUndo
        
        ' Redo options
        bCanUndo = .CanRedo
        If (bCanUndo) Then
            sCaption = "&Redo " & TranslateUndoType(.RedoType)
        Else
            sCaption = "&Redo"
        End If
        If mnuEditRedo.Caption <> sCaption Then mnuEditRedo.Caption = sCaption
        If mnuEditRedo.Enabled <> bCanUndo Then mnuEditRedo.Enabled = bCanUndo
        If tlbMain.Buttons("Redo").Enabled <> bCanUndo Then _
            tlbMain.Buttons("Redo").Enabled = bCanUndo

    End With

End Sub

Private Sub RefreshTabStopArray()
    ' Refreshed the module-level array bTabStop so that it represents
    ' the TabStop value of each control on the form. This is part of
    ' the hack for getting TAB key to work with RichTextBox.
    On Error Resume Next
    Dim i As Integer
    ReDim bTabStop(0 To Controls.Count - 1) As Boolean
    For i = 0 To Controls.Count - 1
       bTabStop(i) = Controls(i).TabStop
       Controls(i).TabStop = False
    Next
End Sub

Private Sub EnableDisableCommands(ByVal CommandGroup As String, ByVal Value As Boolean, ByVal ChangeEnabledProperty As Boolean)

    Dim ctl As Control
    Dim btn As ComctlLib.Button
    For Each ctl In Controls
        If TypeOf ctl Is Menu Or TypeOf ctl Is CommandButton Then
            If InStr(1, ctl.Tag, "[CmdGroup]=" & CommandGroup, vbTextCompare) <> 0 Then
                If ChangeEnabledProperty Then
                    ctl.Enabled = Value
                Else
                    ctl.Visible = Value
                End If
            End If
        ElseIf TypeOf ctl Is ComctlLib.Toolbar Then
            For Each btn In ctl.Buttons
                If InStr(1, btn.Tag, "[CmdGroup]=" & CommandGroup, vbTextCompare) <> 0 Then
                    If ChangeEnabledProperty Then
                        btn.Enabled = Value
                    Else
                        btn.Visible = Value
                    End If
                End If
            Next
        End If
    Next

End Sub

Private Sub UpdateTabCaption()

    On Error GoTo ProcedureError

    Dim sCaption As String
    
    sCaption = ceCONs(m_iCurrentCON).Caption
    If ceCONs(m_iCurrentCON).IsDirty Then sCaption = sCaption & "*"
    tabFiles.Tabs("K" & m_iCurrentCON).Caption = sCaption

    ctTrees(m_iCurrentCON).SetRootNote sCaption
    Me.Caption = sCaption & " - " & App.Title & " " & App.Major & "." & App.Minor & "." & App.Revision
    ' ****
    cbBookmarks(m_iCurrentCON).Filename = ceCONs(m_iCurrentCON).Filename
    Exit Sub
    
ProcedureError:
    If err.Number = 35601 Or err.Number = 340 Then
    Else
        MsgBox err.Number & vbNewLine & err.Description
    End If

End Sub


' Initialization Routines
' -----------------------


Private Sub InitAutoSave()

    Dim bState As Boolean
    
    bState = (Settings.ReadSetting("Code_AutoSave") = "yes")
    tmrAutoSave.Enabled = bState
    
    If bState Then
    
        m_iAutoSaveInterval = FixInteger(Settings.ReadSetting("Code_AutoSaveInterval"))
        If m_iAutoSaveInterval <= 0 Then
            m_iAutoSaveInterval = 1
            Settings.WriteSetting "Code_AutoSaveInterval", "1"
        End If
    
    End If
    
    ' Enable AutoSave timer?
    If Settings.ReadSetting("Code_EnableBackgroundParsing") = "yes" Then
        m_bBackgroundParsing = True
        'tmrCodeUpdate.Enabled = True
        
        ' Set timer interval
        Dim iInterval As Integer
        iInterval = FixInteger(Settings.ReadSetting("Code_BackgroundParsingInterval"))
        If iInterval = 0 Then iInterval = 2
        iInterval = iInterval * 1000
        tmrCodeUpdate.Interval = iInterval
    Else
        m_bBackgroundParsing = False
        'tmrCodeUpdate.Enabled = False
    End If

End Sub

Private Sub UpdateBackgroundParsing()

    ' Enable background parsing timer?
    If Settings.ReadSetting("Code_EnableBackgroundParsing") = "yes" Then
        m_bBackgroundParsing = True
        'tmrCodeUpdate.Enabled = True
        
        ' Set timer interval
        Dim iInterval As Integer
        iInterval = FixInteger(Settings.ReadSetting("Code_BackgroundParsingInterval"))
        If iInterval = 0 Then iInterval = 2
        iInterval = iInterval * 1000
        tmrCodeUpdate.Interval = iInterval
    Else
        m_bBackgroundParsing = False
        'tmrCodeUpdate.Enabled = False
    End If

End Sub

Private Sub InitSnippets()

    Dim oSnip As CSnippet
    
    Dim i As Integer
    For i = mnuSnippet.lbound + 1 To mnuSnippet.UBound
        Unload mnuSnippet(i)
    Next
    
    If Settings.Snippets.Count > 0 Then
        mnuSnippet(0).Visible = False
    
        For Each oSnip In Settings.Snippets
            Load mnuSnippet(mnuSnippet.UBound + 1)
            With mnuSnippet(mnuSnippet.UBound)
                .Caption = oSnip.ToString()
                .Enabled = True
                .Visible = True
                .Tag = "[CmdGroup]=OpenFileRequired; "
            End With
        Next
    Else
        mnuSnippet(mnuSnippet.lbound).Visible = True
    End If

    EnableDisableCommands "OpenFileRequired", (ceCONs.UBound > 0), True

End Sub

Private Sub InitDefinition()

    Set m_oDefinition = New CDefinition
    
    Dim sFile As String
    
    sFile = App.Path & "\" & Settings.ReadSetting("Path_PrimitivesDef")
    
    If IsFileExistant(sFile) Then
        m_oDefinition.LoadBinary sFile
    End If

End Sub


Private Sub InitMenuFlags()

    mnuViewAutoFormattingEnable.Checked = (Settings.ReadSetting("Code_AutoFormat") = "yes")

End Sub

Private Sub InitMenuTags()

    Dim sTag As String
    sTag = "[CmdGroup]=OpenFileRequired; "
    
    mnuFileClose.Tag = sTag                     ' File menu
    mnuFileCloseAll.Tag = sTag
    mnuFileSave.Tag = sTag
    mnuFileSaveAs.Tag = sTag
    mnuFileSaveCopyAs.Tag = sTag
    mnuFileSaveAll.Tag = sTag
    
    mnuEditClear.Tag = sTag                     ' Edit menu
    mnuEditCopy.Tag = sTag
    mnuEditCut.Tag = sTag
    mnuEditAddBookmark.Tag = sTag
    mnuEditFind.Tag = sTag
    mnuEditFindNext.Tag = sTag
    mnuEditGoto.Tag = sTag
    mnuEditPaste.Tag = sTag
    mnuEditRedo.Tag = sTag
    mnuEditReplace.Tag = sTag
    mnuEditSelectAll.Tag = sTag
    mnuEditSelectLine.Tag = sTag
    mnuEditUndo.Tag = sTag
    
    mnuEditAdvancedComment.Tag = sTag           ' Edit -> Advanced menu
    mnuEditAdvancedIndent.Tag = sTag
    mnuEditAdvancedOutdent.Tag = sTag
    mnuEditAdvancedUncomment.Tag = sTag
    mnuEditAdvancedToLower.Tag = sTag
    mnuEditAdvancedToUpper.Tag = sTag
    
    mnuViewAutoFormattingReFormatCode.Tag = sTag ' View -> AutoFormatting
    mnuViewAutoFormattingClear.Tag = sTag
    
    mnuInsertDateTime.Tag = sTag                ' Insert menu
    mnuInsertCommentHeader.Tag = sTag
    mnuInsertFromFile.Tag = sTag
    
    'mnuSnippets.Tag = sTag                      ' Snippets menu
    'mnuSnippet(mnuSnippet.LBound).Tag = sTag
    
    mnuToolsBulkIndenter.Tag = sTag             ' Tools menu
    mnuToolsFilter.Tag = sTag

End Sub

Private Sub UpdateEditorFonts()

    On Error GoTo ProcedureError

    Dim i As Integer
    For i = ceCONs.lbound + 1 To ceCONs.UBound
        With ceCONs(i)
            .SetFont Settings.ReadSetting("EditorFontName"), Settings.ReadSetting("EditorFontSize")
        End With
    Next

    Exit Sub
    
ProcedureError:
    If err.Number = 340 Then
        Resume Next
    Else
        MsgBox err.Number & vbNewLine & err.Description
    End If

End Sub

Private Sub UpdateEditorColors()
    
    On Error GoTo ProcedureError
    
    Dim bAutoFormat As Boolean
    bAutoFormat = (Settings.ReadSetting("Code_AutoFormat") = "yes")
    mnuViewAutoFormattingEnable.Checked = bAutoFormat
    
    Dim i As Integer
    For i = ceCONs.lbound + 1 To ceCONs.UBound
        With ceCONs(i)
            If bAutoFormat Then
                .FormatCode
            Else
                .ClearColoring
            End If
        End With
    Next
    Exit Sub
    
ProcedureError:
    If err.Number = 340 Then
        Resume Next
    Else
        MsgBox err.Number & vbNewLine & err.Description
    End If
    
End Sub


' File Related
' ------------

Public Function GetCONIndex(ByVal Filename As String) As Integer

    ' Returns the index of the CON matching the specified filename.
    ' Returns '-1' if not found

    On Error GoTo ProcedureError

    GetCONIndex = -1

    If ceCONs.Count > 0 Then
        Dim i As Integer
        For i = ceCONs.lbound + 1 To ceCONs.UBound
            If UCase$(Trim$(ceCONs(i).Filename)) = UCase$(Trim$(Filename)) Then
                GetCONIndex = i
                Exit Function
            End If
        Next
    End If

    Exit Function

ProcedureError:
    If err.Number = 340 Then
        Resume Next
    Else
        MsgBox err.Number & vbNewLine & err.Description
    End If

End Function


Private Function NewCON() As Integer

    ' Allocates a new CON (SourceTree, CONEdit, CONTree)
    ' [edit]: Allocated a new CON -> CONEdit, CONTree, CONBookmarks
    
    Static iIndex As Integer
    
    ' Increment index for a new CON
    iIndex = iIndex + 1
    
    ' Create new tab and select it
    With tabFiles
        .Visible = True
        .Tabs.Add , "K" & iIndex
    End With
    
    ' Dimension new CONEditor
    Load ceCONs(iIndex)
    With ceCONs(iIndex)
        .Visible = True
        .Locked = False
        .ZOrder vbBringToFront
        .SetFont Settings.ReadSetting("EditorFontName"), Settings.ReadSetting("EditorFontSize")
        .SetFocus
        
        ' Set parser's definition
        .Parser.Definition = m_oDefinition
    End With
    
    DoEvents
    
    ' Dimension new CONTree
    Load ctTrees(iIndex)
    With ctTrees(iIndex)
        ' Set parser object
        .Parser = ceCONs(iIndex).Parser
        .InitCategories iIndex
        .Visible = True
        .ZOrder vbBringToFront
    End With
    
    ' ****
    ' Dimension new CONBookmark
    Load cbBookmarks(iIndex)
    With cbBookmarks(iIndex)
        .Visible = True
        .ZOrder vbBringToFront
    End With
    
    UpdateTabCaption
    
    tabFiles.Tabs("K" & iIndex).Selected = True
    
    EnableDisableCommands "OpenFileRequired", True, True
    
    RefreshTabStopArray     ' Reload array of tabstop values
    
    AlignControls           ' Position/size new CONedit control
    
    UpdateEditMenu iIndex   ' Update Undo/Redo menus/toolbars
    
    UpdateSelStatus 1, 1
    
    NewCON = iIndex         ' Return new index

End Function

Private Sub RemoveCON(ByVal CONIndex As Integer)

    ' Deallocates a CON (removes appropriate SourceTree, CONEdit and CONTree)
    
    If CONIndex = 0 Then Exit Sub
    
    Unload ceCONs(CONIndex)                 ' Remove CONEditor
    Unload ctTrees(CONIndex)                ' Remove CONTree
    ' ****
    Unload cbBookmarks(CONIndex)            ' Remove CONBookmarks
    
    tabFiles.Tabs.Remove "K" & CONIndex     ' Remove tab
    
    RefreshTabStopArray
    
    If tabFiles.Tabs.Count > 0 Then
        tabFiles_Click
    Else
        ' If there are now zero open CONs...
        tabFiles.Visible = False
        EnableDisableCommands "OpenFileRequired", False, True
        Me.Caption = App.Title & " " & App.Major & "." & App.Minor & "." & App.Revision
        With barStatus.Panels("Status")
            .Text = "Ready"
            Set .Picture = Nothing
        End With
        With barStatus.Panels("Cursor")
            .Text = "Row 0; Col 0"
            .Enabled = False
        End With
    End If

End Sub

Private Function NewBlankCON() As Integer

    ' Creates a new blank CON with an 'untitled' caption
    
    Static lNewCount As Long        ' Untitled count (one-based)
    Dim iIndex As Integer           ' Index of new CON
    
    lNewCount = lNewCount + 1       ' Increment untitled count
    
    iIndex = NewCON()               ' Create new con
    
    ' Set tab caption
    With ceCONs(iIndex)
        .UntitledID = lNewCount
        .SetFocus
    End With
    
    m_iCurrentCON = iIndex
    NewBlankCON = iIndex
    
    UpdateTabCaption
    
End Function

Public Sub LoadCON(ByVal Filename As String)

    Dim iIndex As Integer
    
    iIndex = NewCON()
    With ceCONs(iIndex)
        .LoadFile Filename
        If Len(.Text) > 32767 Then
            If Settings.ReadSetting("Code_EnableBackgroundParsing") = "yes" Then
                If MsgBox("Background Parsing feature is currently enabled. This setting is not recommended when editing large files." & vbNewLine & vbNewLine & "Do you wish to turn off Background Parsing for now?", vbQuestion + vbYesNo) = vbYes Then
                    Settings.WriteSetting "Code_EnableBackgroundParsing", "no"
                    Settings.WriteSetting "Code_AutoFormat", "no"
                    Settings.WriteSetting "Code_DynamicParsing", "no"
                    UpdateBackgroundParsing
                End If
            End If
        End If
        .SetFocus
    End With

    AddRecentItem Filename

End Sub

Private Sub OpenCON()

    ' Opens a new CON from a filename

    On Error GoTo OpenCON_err
    
    With cdl
        .CancelError = True
        .InitDir = Settings.ReadSetting("Path_Cons")
        .Filter = "Duke Nukem 3D CON Scripts (*.con)|*.con|All Files (*.*)|*.*"
        .Flags = cdlOFNFileMustExist + cdlOFNHideReadOnly
        
        .ShowOpen
    
        LoadCON .Filename
        
    End With

    Exit Sub

OpenCON_err:
    If err.Number = 32755 Then
    Else
        MsgBox err.Number & vbNewLine & err.Description
    End If

End Sub

Private Function SaveCON(ByVal CONIndex As Integer) As Boolean

    ' Saves the specified CON
    If IsFileExistant(ceCONs(CONIndex).Filename) Then
        ceCONs(CONIndex).Save
        FlashStatus """" & GetFilenameFromPath(ceCONs(CONIndex).Filename) & """ saved"
        SaveCON = True
    Else
        SaveCON = SaveAsCON(CONIndex)
    End If

End Function

Private Function SaveAsCON(ByVal CONIndex As Integer) As Boolean

    ' Performs a 'Save As' on the specified CON

    On Error GoTo SaveAsCON_err
    
    With cdl
        .CancelError = True
        .InitDir = Settings.ReadSetting("Path_Cons")
        .Filter = "Duke Nukem 3D CON Scripts (*.con)|*.con|All Files (*.*)|*.*"
        .Flags = cdlOFNOverwritePrompt
        .ShowSave
    End With

    With ceCONs(CONIndex)
        .SaveAs cdl.Filename
        .SetFocus
    End With

    SaveAsCON = True

    Exit Function

SaveAsCON_err:
    If err.Number = 32755 Then
        SaveAsCON = False
    Else
        MsgBox "ERROR!!!" & err.Description
    End If

End Function

Private Sub SaveCopyAsCON(CONIndex)

    ' Performs a 'Save Copy As' on the specified CON

    On Error GoTo SaveCopyAsCON_err
    
    With cdl
        .InitDir = Settings.ReadSetting("Path_Cons")
        .Filter = "Duke Nukem 3D CON Scripts (*.con)|*.con|All Files (*.*)|*.*"
        .Flags = cdlOFNOverwritePrompt
        .DialogTitle = "Save Copy As"
        .ShowSave
        ceCONs(CONIndex).SaveCopyAs .Filename
    End With

    Exit Sub

SaveCopyAsCON_err:
    If err.Number <> 32755 Then
        MsgBox "ERROR!!!" & err.Description
    End If

End Sub

Private Sub SaveAllCON()

    ' Attempts to save every open CON

    On Error Resume Next
    
    Dim i As Integer
    For i = 1 To ceCONs.UBound
        If ceCONs(i).IsDirty Then
            SaveCON i
        End If
    Next

End Sub


Private Function CloseCON(ByVal CONIndex As Integer) As Boolean

    ' Closes a CON "nicely", asking the user to save if necessary

    On Error GoTo ProcedureError

    Dim iMsg As Integer
    Dim bResult As Boolean

    With ceCONs(CONIndex)
        If .IsDirty Then
            iMsg = MsgBox(.Caption & vbNewLine & _
                        vbNewLine & "Do you want to save the changes to this document?" _
                        , vbQuestion + vbYesNoCancel)
            
            If iMsg = vbCancel Then
                CloseCON = False
                Exit Function
            End If
                
            If iMsg = vbYes Then
                bResult = SaveCON(CONIndex)
                CloseCON = bResult
                If Not bResult Then Exit Function
            End If
        End If
        
        RemoveCON CONIndex
    End With

    CloseCON = True
    Exit Function
    
ProcedureError:
    If err.Number = 340 Then
        CloseCON = True
        Exit Function
    Else
        MsgBox err.Number & vbNewLine & err.Description
    End If

End Function

Private Function CloseAllCON() As Boolean

    ' Closes every CON "nicely", attempting to save each
    ' Returns Boolean; True=successful, False=Cancelled by user

    Dim i As Integer
    
    For i = ceCONs.lbound To ceCONs.UBound
        If CloseCON(i) = False Then
            CloseAllCON = False
            Exit Function
        End If
    Next
    
    CloseAllCON = True

End Function


Private Sub RefreshRecentList()

    If m_oMRUList Is Nothing Then LoadRecentList

    Dim i As Integer
    Dim sItems() As String
    
    ' Clear all MRU menu items
    For i = mnuFileMRU.lbound + 1 To mnuFileMRU.UBound
        Unload mnuFileMRU(i)
    Next
    
    sItems = m_oMRUList.Items       ' Get array of MRU items
    
    ' Load MRU menu items per array items
    For i = 0 To UBound(sItems)     ' Assumes sItems is zero-based
        If mnuFileMRU(0).Visible = False Then mnuFileMRU(0).Visible = True
        Load mnuFileMRU(i + 1)
        With mnuFileMRU(i + 1)
            .Caption = "&" & i + 1 & " " & sItems(i)
            .Tag = sItems(i)
            .Visible = Not (Len(sItems(i)) = 0)
        End With
    Next

End Sub

Private Sub AddRecentItem(ByVal Filename As String)

    If m_oMRUList Is Nothing Then LoadRecentList
    
    If UBound(Filter(m_oMRUList.Items, Filename, True)) <> -1 Then Exit Sub
    
    m_oMRUList.AddMRUItem Filename
    
    ' Save contents to settings file
    Settings.WriteSetting "Recent_List", m_oMRUList.ToString()
    
    RefreshRecentList

End Sub

Private Sub LoadRecentList()

    Set m_oMRUList = New CMRUList
    m_oMRUList.MaxItems = FixInteger(Settings.ReadSetting("Recent_Count"))
    m_oMRUList.LoadFromString Settings.ReadSetting("Recent_List")
    
    RefreshRecentList

End Sub


' Code
' ----

Private Sub InsertBookmark()

    Dim lValue As Long
    
    lValue = ceCONs(m_iCurrentCON).SelStart
    
    With cbBookmarks(m_iCurrentCON)
        If Not .DoesBookmarkExist(lValue) Then
            
            Load FrmAddBookmark
            With FrmAddBookmark
                .txtCharPosition.Text = lValue
                .txtLabel.Text = ceCONs(m_iCurrentCON).SelLineText
                
                .Show vbModal
                
                If Not .IsCancelled Then
                     cbBookmarks(m_iCurrentCON).AddBookmark .txtLabel.Text, lValue
                End If
            End With
            Unload FrmAddBookmark
            Set FrmAddBookmark = Nothing
        
        End If
    End With

End Sub

Private Sub InsertStructure(ByRef Structure As CStructure)

    tmrCodeUpdate.Enabled = False

    Load FrmStructure
    With FrmStructure
        .IsAdding = True
        
        Dim oBlock As CBlock
        Set oBlock = New CBlock
        oBlock.Structure = Structure
        oBlock.Parent = ceCONs(m_iCurrentCON).Parser
        .Block = oBlock

        .Show vbModal

        If Not .IsCancelled Then
            With ceCONs(m_iCurrentCON)
                .Parser.Blocks.Add FrmStructure.Block
            End With
        End If
    End With
    Unload FrmStructure
    Set FrmStructure = Nothing

    tmrCodeUpdate.Enabled = True

End Sub

Private Sub InsertPrimitive()
    Load FrmPrimitive
    With FrmPrimitive
        .Show 'vbModal
        
        'If Not .IsCancelled Then
        '    ceCONs(m_iCurrentCON).InsertText .Code, True 'SelText = .Code
        'End If
        'ceCONs(m_iCurrentCON).SetFocus
    End With
    'Unload FrmPrimitive
End Sub


Private Sub RunCleanUp()

    On Error GoTo ProcedureError

    If m_iCurrentCON <> 0 Then
        ceCONs(m_iCurrentCON).Parser.CleanUpBlocks ceCONs(m_iCurrentCON).Text
    End If

    Exit Sub
    
ProcedureError:
    If err.Number = 340 Then
        ' Ignore "element not found" error; this sometimes happens just
        ' between the ceCONs element being unloaded and the m_iCurrentCON
        ' variable being updated to reflect that.
    Else
        MsgBox err.Number & vbNewLine & err.Description
    End If

End Sub



' Control Positioning/Sizing
' --------------------------

Private Function GetClientTop() As Integer

    ' Get position of client's top (form.top + toolbar + etc...)

    Dim cCtl As Control
    Dim iTop As Integer
    
    For Each cCtl In Me.Controls
        If cCtl.Visible = True _
            And InStr(1, cCtl.Tag, "[IsHeader]", vbTextCompare) <> 0 Then
            
            iTop = iTop + cCtl.Height
        End If
    Next

    ' Add padding
    iTop = iTop + 8

End Function

Private Function GetPanelsHeight(ByVal IsHeader As Boolean) As Integer

    Dim ctl As Control
    Dim sSearch As String
    Dim i As Integer
    
    If IsHeader Then
        sSearch = "[IsHeader]"
    Else
        sSearch = "[IsFooter]"
    End If
    
    For Each ctl In Me.Controls
        If InStr(1, ctl.Tag, sSearch) <> 0 Then
            If ctl.Visible = True Then
                i = i + ctl.Height
            End If
        End If
    Next

    GetPanelsHeight = i

End Function

Private Sub AlignControls()

    On Error Resume Next

    Dim iW As Integer, iH As Integer
    Static iPrevH As Integer
    
    ' Cache form size for fast access
    iW = Me.ScaleWidth
    iH = Me.ScaleHeight - (GetPanelsHeight(True) + GetPanelsHeight(False)) - 70
    
    ' Size/position tab control
    tabFiles.Top = 1
    If tlbMain.Visible = True Then tabFiles.Top = tabFiles.Top + tlbMain.Height
    If tlbCode.Visible = True Then tabFiles.Top = tabFiles.Top + tlbCode.Height
    
    tabFiles.Left = lblVSplitter.Left + 8
    tabFiles.Width = iW - (picTree.Width + 16)
    tabFiles.Height = iH - 13
    
    cmdClose.Left = Me.Width - (40 * Screen.TwipsPerPixelX)
    
    Dim i As Integer
    For i = ceCONs.lbound + 1 To ceCONs.UBound
        With ceCONs(i)      ' Size code window
            .Left = tabFiles.Left + 8
            .Top = tabFiles.Top + 32
            .Width = tabFiles.Width - 16
            .Height = tabFiles.Height - 40
        End With
        
        With ctTrees(i)     ' Size code tree
        'With pic
            .Width = picTree.Width - 4
            .Height = (picTree.Height - .Top) - 4
        End With
    
        ' ****
        With cbBookmarks(i) ' Size bookmarks
            .Width = picProp.Width - 4
            .Height = (picProp.Height - .Top) - 4
        End With
    Next

    ' Size splitters
    lblHSplitter.Width = picTree.Width
    lblVSplitter.Height = tabFiles.Height

    ' Size controls according to splitter positions
    If iPrevH <> 0 Then
        picTree.Height = iH / (iPrevH / picTree.Height)
        picTree.Top = tabFiles.Top - 3
        picProp.Top = picTree.Top + picTree.Height + 8
        picProp.Height = iH - (picTree.Height + 16)
        lblHSplitter.Top = picProp.Top - 16
        iPrevH = iH
    Else
        iPrevH = iH
        AlignControls       ' * Recursion
    End If
    
    ' Align progress bar
    With prg
        If .Visible Then
            .Left = 2
            .Width = barStatus.Panels(1).Width - 4
            .Top = barStatus.Top + 4
            .Height = barStatus.Height - 6
        End If
    End With

End Sub


' Editing Functions
' -----------------

Private Sub Undo()
    ceCONs(m_iCurrentCON).Undo
    UpdateEditMenu m_iCurrentCON
End Sub

Private Sub Redo()
    ceCONs(m_iCurrentCON).Redo
    UpdateEditMenu m_iCurrentCON
End Sub

Private Sub RunDuke()

    Dim sPath As String
    Dim iResult As Integer
    sPath = Settings.ReadSetting("Path_Duke3D") & "\duke3d.exe"
    If Dir(sPath) <> "" Then
        
        If Settings.ReadSetting("General_RunDukeWithCON") = "yes" Then
            If ceCONs.UBound > 0 Then
                sPath = sPath & " /X" & GetFilenameFromPath(ceCONs(m_iCurrentCON).Filename)
            End If
        End If
        
        ChDir Settings.ReadSetting("Path_Duke3D")
        Shell sPath, vbNormalFocus
    Else
        iResult = MsgBox("The Duke Nukem 3D program file, ""duke3d.exe"" could not be located in the specified path." & vbNewLine & vbNewLine & "Would you like to open the Options window to change this setting?", vbQuestion + vbYesNo)
        If iResult = vbYes Then
            Load FrmOptions
            FrmOptions.tabProps.Tabs(2).Selected = True
            FrmOptions.Show vbModal
        End If
    End If

End Sub


' Dialogs
' -------

Private Sub ShowFindReplaceDialog(ByVal Mode As eFindReplaceMode)
    Load FrmFindReplace
    With FrmFindReplace
        Select Case Mode
            Case frFind To frReplace
                If ceCONs(m_iCurrentCON).SelLength > 0 Then _
                    .txtFindFindWhat.Text = ceCONs(m_iCurrentCON).SelText
            Case frGoto
                ' [...] TODO
        End Select
        .Parser = ceCONs(m_iCurrentCON).Parser
        .Show
        .ChangeTab Mode
    End With
End Sub



' Event Handlers
' ==============

' Form
' ----

Private Sub Form_Load()

    InitDefinition          ' Initialize parser definition object
    
    ' Set form caption: [Title] [Version] [Beta?]
    Me.Caption = App.Title & " " & "v" & App.Major & "." & App.Minor & "." & App.Revision & IIf(App.Major = 0, " Beta", "")
    
    InitMenuTags            ' Initialize menu tags
    InitMenuFlags
    
    InitSnippets            ' Initialize snippets menu
    
    UpdateBackgroundParsing ' Initialize background parsing timer
    
    InitAutoSave            ' Initialize AutoSave functionality
    
    LoadRecentList          ' Initialize Recent List (in the File menu)
    
    tabFiles.Tabs.Clear     ' Clear all tabs
    
    EnableDisableCommands "OpenFileRequired", False, True

    Me.Show                 ' Make the form visible

    If Settings.ReadSetting("DefaultDocumentOnStartup") = "yes" Then
        NewBlankCON         ' Create new blank file if specified
    End If

    UpdateEditorFonts       ' Set fonts per setting

End Sub

Private Sub Form_Resize()

    Static iOldWindowState As Integer       ' Last state of form
    
    ' Align & size form controls
    AlignControls
    
    ' Prevent left side-panels from skewing if the form is
    ' maximized or restored after being maximized
    If iOldWindowState <> Me.WindowState Then
        lblHSplitter_MouseMove vbLeftButton, 0, lblHSplitter.Left, lblHSplitter.Top
        AlignControls
        iOldWindowState = Me.WindowState
    End If

End Sub


Private Sub Form_Unload(Cancel As Integer)
    
    Dim bResult As Boolean
    
    bResult = CloseAllCON()
    If bResult = True Then
        Me.Hide                 ' If successful, hide the form
    End If
    
    Set m_oDefinition = Nothing ' Garbage collection (sort of...)
    
    Cancel = 1                  ' Either way, don't unload the form

End Sub


' Splitters
' ---------

Private Sub lblHSplitter_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    lblHSplitter.BackColor = vb3DShadow
End Sub

Private Sub lblHSplitter_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = vbLeftButton Then
        
        'On Error Resume Next
        
        Dim iY As Integer, iHalf As Integer, i As Integer
        
        iY = y / Screen.TwipsPerPixelY
        iHalf = lblHSplitter.Height / 2
        
        If iY < iHalf Then
            i = lblHSplitter.Top - (iHalf - iY)
        Else
            i = lblHSplitter.Top + (iY - iHalf)
        End If
    
        lblHSplitter.Top = i
        picTree.Height = i - ((GetPanelsHeight(True) + GetPanelsHeight(False)) - 16) - 65
        
        AlignControls
    
    End If
End Sub

Private Sub lblHSplitter_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    lblHSplitter.BackColor = vb3DFace
End Sub

Private Sub lblVSplitter_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    lblVSplitter.BackColor = vb3DShadow
End Sub

Private Sub lblVSplitter_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    
    If Button = vbLeftButton Then
        
        On Error Resume Next
        
        Dim iX As Integer, iHalf As Integer, i As Integer
        
        iX = x / Screen.TwipsPerPixelX
        iHalf = lblVSplitter.Width / 2
        
        If iX < iHalf Then
            i = lblVSplitter.Left - (iHalf - iX)
        Else
            i = lblVSplitter.Left + (iX - iHalf)
        End If
    
        lblVSplitter.Left = i
        picTree.Width = i
        picProp.Width = i
    
        AlignControls
    
    End If
    
End Sub

Private Sub lblVSplitter_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    lblVSplitter.BackColor = vb3DFace
End Sub


Private Sub LongTimer1_Tick()
    MsgBox "hey"
End Sub

' Menus
' -----

Private Sub mnuEdit_Click()
    If tabFiles.Tabs.Count <> 0 Then _
        UpdateEditMenu m_iCurrentCON
End Sub

Private Sub mnuEditAddBookmark_Click()
    InsertBookmark
End Sub

Private Sub mnuEditAdvancedComment_Click()
    ceCONs(m_iCurrentCON).SelComment
End Sub

Private Sub mnuEditAdvancedIndent_Click()
    ceCONs(m_iCurrentCON).SelIndent
End Sub

Private Sub mnuEditAdvancedOutdent_Click()
    ceCONs(m_iCurrentCON).SelOutdent
End Sub

Private Sub mnuEditAdvancedToLower_Click()
    ceCONs(m_iCurrentCON).SelCase = TC_LOWERCASE
End Sub

Private Sub mnuEditAdvancedToUpper_Click()
    ceCONs(m_iCurrentCON).SelCase = TC_UPPERCASE
End Sub

Private Sub mnuEditAdvancedUncomment_Click()
    ceCONs(m_iCurrentCON).SelUncomment
End Sub

Private Sub mnuEditClear_Click()
    ceCONs(m_iCurrentCON).SelClear
End Sub

Private Sub mnuEditCopy_Click()
    ceCONs(m_iCurrentCON).SelCopy
End Sub

Private Sub mnuEditCut_Click()
    ceCONs(m_iCurrentCON).SelCut
End Sub

Private Sub mnuEditFind_Click()
    ShowFindReplaceDialog frFind
End Sub

Private Sub mnuEditFindNext_Click()
    If ceCONs(m_iCurrentCON).FindWhat = "" Then
        ShowFindReplaceDialog frFind
    Else
        If ceCONs(m_iCurrentCON).FindNext() Then
            MsgBox App.Title & " has finished searching the document.", vbInformation
        End If
    End If
End Sub

Private Sub mnuEditGoto_Click()
    ShowFindReplaceDialog frGoto
End Sub

Private Sub mnuEditPaste_Click()
    ceCONs(m_iCurrentCON).SelPaste
End Sub

Private Sub mnuEditRedo_Click()
    Redo
End Sub

Private Sub mnuEditReplace_Click()
    ShowFindReplaceDialog frReplace
End Sub

Private Sub mnuEditSelectAll_Click()
    ceCONs(m_iCurrentCON).SelectAll
End Sub

Private Sub mnuEditSelectLine_Click()
    ceCONs(m_iCurrentCON).SelectLine ceCONs(m_iCurrentCON).SelLine
End Sub

Private Sub mnuEditUndo_Click()
    Undo
End Sub

Private Sub mnuFileClose_Click()
    CloseCON m_iCurrentCON
End Sub

Private Sub mnuFileCloseAll_Click()
    CloseAllCON
End Sub

Private Sub mnuFileExit_Click()
    Unload Me
End Sub

Private Sub mnuFileMRU_Click(Index As Integer)

    Dim iIndex As Integer
    Dim sFilename As String
    
    sFilename = CStr(mnuFileMRU(Index).Tag)

    If IsFileExistant(sFilename) Then
        iIndex = NewCON()
        With ceCONs(iIndex)
            .LoadFile sFilename
            If Len(.Text) > 32767 Then
                If Settings.ReadSetting("Code_EnableBackgroundParsing") = "yes" Then
                    If MsgBox("Background Parsing feature is currently enabled. This setting is not recommended when editing large files." & vbNewLine & vbNewLine & "Do you wish to turn off Background Parsing for now?", vbQuestion + vbYesNo) = vbYes Then
                        Settings.WriteSetting "Code_EnableBackgroundParsing", "no"
                        UpdateBackgroundParsing
                    End If
                End If
            Else
                If Settings.ReadSetting("Code_AutoFormat") = "yes" Then
                    .FormatCode
                End If
            End If
            .SetFocus
        End With
    End If

End Sub

Private Sub mnuFileNew_Click()
    NewBlankCON
End Sub

Private Sub mnuFileOpen_Click()
    OpenCON
End Sub

Private Sub mnuFileSave_Click()
    SaveCON m_iCurrentCON
End Sub

Private Sub mnuFileSaveAll_Click()
    SaveAllCON
End Sub

Private Sub mnuFileSaveAs_Click()
    SaveAsCON m_iCurrentCON
End Sub

Private Sub mnuFileSaveCopyAs_Click()
    SaveCopyAsCON m_iCurrentCON
End Sub

Private Sub mnuHelpAbout_Click()
    FrmAbout.Show vbModal
End Sub

Private Sub mnuInsertCommentHeader_Click()
    Dim sText As String
    sText = "/*" & vbNewLine & _
            vbTab & "Name: " & vbNewLine & _
            vbTab & "Author: " & vbNewLine & _
            vbTab & "Description: " & vbNewLine & _
            vbTab & "Date: " & Date & " " & Time & vbNewLine & _
            vbTab & "Copyright: " & vbNewLine & _
            "*/"

    ceCONs(m_iCurrentCON).InsertText sText, False
End Sub

Private Sub mnuInsertDateTime_Click()
    ceCONs(m_iCurrentCON).InsertText Date & " " & Time, False
End Sub

Private Sub mnuInsertFromFile_Click()

    On Error GoTo ProcedureError

    With cdl
    
        .CancelError = True
        .DialogTitle = "Insert From"
        .Filter = "Duke Nukem 3D CON Scripts (*.con)|*.con|Text Files (*.txt)|*.txt|All Files (*.*)|*.*"
        .FilterIndex = 0
        .Flags = cdlOFNFileMustExist
        
        .ShowOpen
        
        Dim lHnd As Long
        Dim sText As String
        
        lHnd = FreeFile
        
        Open .Filename For Input As #lHnd
        sText = Input(LOF(lHnd), #lHnd)     ' Read file contents
        Close #lHnd
    
        With ceCONs(m_iCurrentCON)
            .InsertText sText, True
            .SetFocus
        End With
    
    End With

    Exit Sub

ProcedureError:
    If err.Number = 32755 Then
    Else
        MsgBox err.Number & vbNewLine & err.Description
    End If

End Sub

Private Sub mnuInsertPrimitive_Click()
    InsertPrimitive
End Sub

Private Sub mnuInsertStructure_Click()
    InsertStructure Me.Definition.Structures(1)
End Sub

Private Sub mnuSnippet_Click(Index As Integer)

    ' Insert clicked snippet into the current document

    Dim oSnip As CSnippet
    Set oSnip = Settings.Snippets.FindItem(mnuSnippet(Index).Caption)
    
    If Not (oSnip Is Nothing) Then
    
        If oSnip.Fields.Count > 0 Then
        
            ' Prompt user for field values
            Load FrmInsertSnippet
            With FrmInsertSnippet
                .Snippet = oSnip
                .Show vbModal
                If Not .IsCancelled Then
                    ceCONs(m_iCurrentCON).InsertText .InsertText, True
                End If
            End With
        
        Else
        
            ' No fields, just insert
            ceCONs(m_iCurrentCON).InsertText oSnip.Text, True
        End If
    
    End If

End Sub

Private Sub mnuSnippetsManage_Click()
    FrmManageSnippets.Show vbModal
    InitSnippets
End Sub

Private Sub mnuToolsBulkIndenter_Click()
    If MsgBox("This action will alter all indentation in the current document and is irreversible." & vbNewLine & vbNewLine & "Do you wish to proceed?", vbQuestion + vbYesNo) = vbYes Then
        ceCONs(m_iCurrentCON).BulkIndent
    End If
End Sub

Private Sub mnuToolsDefinitionEditor_Click()

    Dim sPath As String
    
    sPath = App.Path & "\" & DEFINITION_EDITOR_FILENAME

    If IsFileExistant(sPath) Then
        Shell """" & sPath & """", vbNormalFocus
    Else
        MsgBox sPath & vbNewLine & vbNewLine & "Unable to run Definition Editor. This file could not be found. Please try un-installing then re-installing CONstruct.", vbExclamation
    End If

End Sub

Private Sub mnuToolsFilter_Click()
    
    On Error GoTo ProcedureError
    
    Load FrmFilter
    With FrmFilter
    
        .Show vbModal
        
        If Not (.IsCancelled) Then
            ceCONs(NewBlankCON()).Text = .FilterText
        End If
    
    End With
    Unload FrmFilter
    
    Exit Sub
    
ProcedureError:
    If err.Number = 340 Then
        Resume Next
    Else
        MsgBox err.Number & vbNewLine & err.Description
    End If

End Sub

Private Sub mnuToolsOptions_Click()
    Load FrmOptions
    With FrmOptions
        .Show vbModal
        
        UpdateEditorFonts
        UpdateEditorColors
        InitSnippets
        
        If InStr(1, .SettingsChanged, "[Code_EnableBackgroundParsing]", vbBinaryCompare) <> 0 Then
            UpdateBackgroundParsing
        End If
        If InStr(1, .SettingsChanged, "[Code_BackgroundParsingInterval]", vbBinaryCompare) <> 0 Then
            UpdateBackgroundParsing
        End If
    End With
End Sub

Private Sub mnuToolsRun_Click()
    RunDuke
End Sub



' Other Events
' ------------

Private Sub tabFiles_Click()
    
    On Error GoTo tabFiles_Click_err
    
    Dim iCONIndex As Integer
    
    iCONIndex = CInt(Right$(tabFiles.SelectedItem.Key, Len(tabFiles.SelectedItem.Key) - 1))
    
    UpdateEditMenu iCONIndex
    
    With ceCONs(iCONIndex)
        .ZOrder vbBringToFront
        .SetFocus
    End With
    
    ctTrees(iCONIndex).ZOrder vbBringToFront
    cbBookmarks(iCONIndex).ZOrder vbBringToFront
    
    Dim frm As Form
    For Each frm In Forms
        If frm.Name = "FrmFindReplace" Then
            'If FrmFindReplace.Visible = True Then
            FrmFindReplace.Parser = ceCONs(iCONIndex).Parser
            'End If
            Exit For
        End If
    Next
    
    m_iCurrentCON = iCONIndex
    Me.Caption = tabFiles.SelectedItem.Caption & " - " & App.Title & " " & App.Major & "." & App.Minor & "." & App.Revision
    With barStatus.Panels("Status")
        .Text = "Ready"
        Set .Picture = Nothing
    End With
    
    Exit Sub

tabFiles_Click_err:

    If err.Number = 5 Then
        ' No Error
    Else
        MsgBox "ERROR!!!" & err.Description
        Stop
    End If
End Sub

Private Sub cmdRefresh_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    picTree.SetFocus
End Sub

Private Sub ceCONs_GotFocus(Index As Integer)
    ' Store TabStop property setting of each control on the form
    ' then reset them all to False to that TAB has the right effect
    ' on the RichTextBox control
    On Error GoTo ProcedureError

    Dim i As Integer
    ReDim m_bTabStop(0 To Controls.Count - 1) As Boolean
    For i = 0 To Controls.Count - 1
       m_bTabStop(i) = Controls(i).TabStop
       Controls(i).TabStop = False
    Next
    Exit Sub

ProcedureError:
    If err.Number = 438 Or err.Number = 9 Then
        Resume Next
    Else
        MsgBox err.Number & vbNewLine & err.Description, , "ceCONs_GotFocus()"
    End If
End Sub

Private Sub ceCONs_LostFocus(Index As Integer)
    'Restore the Tabstop property for each control on the form
    On Error GoTo ProcedureError
    
    Dim i As Integer
    For i = 0 To Controls.Count - 1
       Controls(i).TabStop = m_bTabStop(i)
    Next
    Exit Sub

ProcedureError:
    If err.Number = 438 Or err.Number = 9 Then
        Resume Next
    Else
        MsgBox err.Number & vbNewLine & err.Description, , "ceCONs_LostFocus()"
    End If
End Sub

Private Sub ceCONs_DirtyStateChanged(Index As Integer)
    UpdateTabCaption
End Sub

Private Sub ceCONs_FileLoadBegin(Index As Integer)
    prg.Visible = True
    AlignControls
    EnableDisableCommands "OpenFileRequired", False, True
End Sub

Private Sub ceCONs_FileLoadEnd(Index As Integer)
    prg.Visible = False
    EnableDisableCommands "OpenFileRequired", True, True
    UpdateTabCaption
End Sub

Private Sub m_oParser_ParseProgress(ByVal CurrentPosition As Long, ByVal TotalLength As Long)
    Static i As Integer
    i = 100 / (TotalLength / CurrentPosition)
    If i > 100 Then i = 100
    prg.Value = i
End Sub

Private Sub ceCONs_FilenameChanged(Index As Integer)
    UpdateTabCaption
End Sub


Private Sub mnuViewMainToolbar_Click()
    'FeatureNotImplemented
    tlbMain.Visible = Not tlbMain.Visible
    mnuViewMainToolbar.Checked = tlbMain.Visible
    AlignControls
End Sub

Private Sub ceCONs_Change(Index As Integer)
    UpdateEditMenu Index
End Sub

Private Sub ceCONs_SelChange(Index As Integer, ByVal SelLine As Long, ByVal SelColumn As Integer)
    UpdateSelStatus SelLine, SelColumn
End Sub

Private Sub cbBookmarks_BookmarkClicked(Index As Integer, ByVal Value As Long)
    ' Select the appropriate line number when a bookmark is clicked
    With ceCONs(m_iCurrentCON)
        .SelStart = Value
        .SetFocus
    End With
End Sub

Private Sub tlbCode_ButtonClick(ByVal Button As ComctlLib.Button)
    
    Select Case Button.Key
        Case "InsertPrimitive": InsertPrimitive
        Case "InsertStructure": InsertStructure Me.Definition.Structures(1)
        Case "Bookmark":        InsertBookmark
        Case "Snippet"
            PopupMenu mnuSnippets, , tlbCode.Left + Button.Left, tlbCode.Top + Button.Top + Button.Height
        Case "Indent":          ceCONs(m_iCurrentCON).SelIndent
        Case "Outdent":         ceCONs(m_iCurrentCON).SelOutdent
        Case "Comment":         ceCONs(m_iCurrentCON).SelComment
        Case "Uncomment":       ceCONs(m_iCurrentCON).SelUncomment
    
    End Select

End Sub



Private Sub tlbMain_ButtonClick(ByVal Button As ComctlLib.Button)

    Select Case Button.Key
        Case "New":             NewBlankCON
        Case "Open":            OpenCON
        Case "Save":            SaveCON m_iCurrentCON
        Case "SaveAll":         SaveAllCON
        Case "Cut":             ceCONs(m_iCurrentCON).SelCut
        Case "Copy":            ceCONs(m_iCurrentCON).SelCopy
        Case "Paste":           ceCONs(m_iCurrentCON).SelPaste
        Case "Find":            ShowFindReplaceDialog frFind
        Case "Undo":            Undo
        Case "Redo":            Redo
        Case "RunDuke":         RunDuke
        Case "CONFilter":       mnuToolsFilter_Click
    End Select

End Sub


Private Sub cmdClose_Click()
    CloseCON m_iCurrentCON
End Sub

Private Sub ctTrees_ObjectClick(Index As Integer, Block As CBlock)
    On Error GoTo ProcedureError
    
    ceCONs(m_iCurrentCON).GotoBlock Block
    ceCONs(m_iCurrentCON).SetFocus
    
    Exit Sub
    
ProcedureError:
    If err.Number = 5 Then
        ' This sometimes happens when SetFocus() is called, dunno why
        ' TODO : Stop this error from happening if possible
        Resume Next
    Else
        MsgBox err.Number & vbNewLine & err.Description
    End If
End Sub

Private Sub ctTrees_ObjectDblClick(Index As Integer, Block As CBlock)

    ' Run code cleanup to ensure that block object is up-to-date
    RunCleanUp

    tmrCodeUpdate.Enabled = False

    Load FrmStructure
    With FrmStructure
        .IsAdding = False
        .Block = Block
        
        .Show vbModal
        
        If Not .IsCancelled Then
            ctTrees(Index).UpdateBlock Block
        End If
    End With
    Unload FrmStructure
    Set FrmStructure = Nothing

    tmrCodeUpdate.Enabled = True

End Sub


Private Sub ceCONs_ObjectEdit(Index As Integer, Block As CBlock)
    ctTrees_ObjectDblClick Index, Block
End Sub


Private Sub ceCONs_KeyWordEdit(Index As Integer, Primitive As CPrimitive, ByVal Text As String)
    Load FrmPrimitive
    With FrmPrimitive
        .SelectedPrimitive = Primitive
        .txtCode.Text = Text
        '.LoadPrimitive Primitive
        '.txtCode.Text = Text
        '.tabMain.Tabs(1).Selected = True
        '.Initialize
        '.Show vbModal
    End With
    'Unload FrmPrimitive
End Sub


Private Sub ctTrees_ObjectAdd(Index As Integer, ByRef Structure As CStructure)
        
    InsertStructure Structure
    
End Sub


Private Sub ctTrees_ObjectDelete(Index As Integer, Block As CBlock)

    ceCONs(m_iCurrentCON).Parser.Blocks.Delete Block.Index

End Sub

Private Sub tmrAutoSave_Tick()
    If tmrAutoSave.CurrentMinute >= m_iAutoSaveInterval Then
        tmrAutoSave.Reset
        
        ' Perform AutoSave operation
        If ceCONs.UBound > 0 Then
            If ceCONs(m_iCurrentCON).IsDirty Then
                If Settings.ReadSetting("Code_AutoSaveSaveAs") = "yes" Then
                    ' Always save, whatever the case
                    SaveCON m_iCurrentCON
                Else
                    ' Only save if a valid filename is registered
                    If IsFileExistant(ceCONs(m_iCurrentCON).Filename) Then
                        ceCONs(m_iCurrentCON).Save
                    End If
                End If
            End If
        End If
        
    End If
End Sub

Private Sub tmrCodeUpdate_Timer()
        
    If m_bBackgroundParsing Then
        If ceCONs.UBound > 0 Then
            If Not ceCONs(m_iCurrentCON).Parser.IsLoading Then
                RunCleanUp
            End If
        End If
    End If

End Sub


Private Sub ceCONs_ParseProgress(Index As Integer, ByVal CurrentPosition As Long, ByVal TotalLength As Long)

    If CurrentPosition < TotalLength Then
        prg.Max = TotalLength
        prg.Value = CurrentPosition
    End If

End Sub


Private Sub ceCONs_SelBlockChange(Index As Integer, Block As CBlock)
    UpdateCurrentObject Block
End Sub



Private Sub mnuViewAutoFormattingReFormatCode_Click()
    If ceCONs.UBound > 0 Then
        ceCONs(m_iCurrentCON).FormatCode
    End If
End Sub

Private Sub mnuViewAutoFormattingClear_Click()
    If ceCONs.UBound > 0 Then
        ceCONs(m_iCurrentCON).ClearColoring
    End If
End Sub


Private Sub mnuViewAutoFormattingEnable_Click()

    ' Toggle AutoFormat settings
    Dim bAutoFormat As Boolean
    bAutoFormat = (Settings.ReadSetting("Code_AutoFormat") = "yes")
    bAutoFormat = Not bAutoFormat
    Settings.WriteSetting "Code_AutoFormat", IIf(bAutoFormat, "yes", "no")
    
    ' Refresh everything
    UpdateEditorColors

End Sub



Private Sub tmrFlashStatus_Timer()
    barStatus.Panels(1).Text = m_sPreviousStatusText
    tmrFlashStatus.Enabled = False
End Sub

