VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{FE0065C0-1B7B-11CF-9D53-00AA003C9CB6}#1.1#0"; "COMCT232.OCX"
Begin VB.Form FrmOptions 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Options"
   ClientHeight    =   4650
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5535
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FrmOptions.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   310
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   369
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picTabs 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   3495
      Index           =   1
      Left            =   165
      ScaleHeight     =   233
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   346
      TabIndex        =   32
      Top             =   480
      Width           =   5190
      Begin VB.Frame fraAutoSave 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "AutoSave"
         ForeColor       =   &H80000008&
         Height          =   1380
         Left            =   105
         TabIndex        =   56
         Top             =   1995
         Width           =   2955
         Begin VB.CheckBox chkAutoSaveAs 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Open ""Save As"" for new files"
            ForeColor       =   &H80000008&
            Height          =   225
            Left            =   315
            TabIndex        =   61
            Tag             =   "Code_AutoSaveSaveAs"
            Top             =   945
            Width           =   2430
         End
         Begin VB.TextBox txtAutoSaveInterval 
            Height          =   285
            Left            =   1785
            TabIndex        =   58
            Tag             =   "Code_AutoSaveInterval"
            Top             =   615
            Width           =   750
         End
         Begin VB.CheckBox chkAutoSave 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Enable AutoSave"
            ForeColor       =   &H80000008&
            Height          =   225
            Left            =   105
            TabIndex        =   57
            Tag             =   "Code_AutoSave"
            Top             =   315
            Width           =   2325
         End
         Begin ComCtl2.UpDown updAutoSaveInterval 
            Height          =   285
            Left            =   2550
            TabIndex        =   59
            Top             =   615
            Width           =   255
            _ExtentX        =   450
            _ExtentY        =   503
            _Version        =   327681
            BuddyControl    =   "txtAutoSaveInterval"
            BuddyDispid     =   196637
            OrigLeft        =   175
            OrigTop         =   63
            OrigRight       =   192
            OrigBottom      =   85
            Max             =   60
            SyncBuddy       =   -1  'True
            BuddyProperty   =   0
            Enabled         =   -1  'True
         End
         Begin VB.Label lblAutoSaveInterval 
            BackStyle       =   0  'Transparent
            Caption         =   "Interval (minutes):"
            Height          =   225
            Left            =   315
            TabIndex        =   60
            Top             =   630
            Width           =   1380
         End
      End
      Begin VB.Frame fraParsing 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Parsing"
         ForeColor       =   &H80000008&
         Height          =   1275
         Left            =   105
         TabIndex        =   50
         Top             =   525
         Width           =   2955
         Begin VB.CheckBox chkBackgroundParsing 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Enable background parsing"
            ForeColor       =   &H80000008&
            Height          =   225
            Left            =   105
            TabIndex        =   54
            Tag             =   "Code_EnableBackgroundParsing"
            Top             =   315
            Width           =   2325
         End
         Begin VB.TextBox txtBackgroundParsingInterval 
            Height          =   285
            Left            =   1785
            TabIndex        =   53
            Tag             =   "Code_BackgroundParsingInterval"
            Top             =   615
            Width           =   750
         End
         Begin VB.CheckBox chkDynamicParsing 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Enable dynamic parsing"
            ForeColor       =   &H80000008&
            Height          =   225
            Left            =   105
            TabIndex        =   51
            Tag             =   "Code_DynamicParsing"
            Top             =   945
            Width           =   2010
         End
         Begin ComCtl2.UpDown updBackgroundParsingInterval 
            Height          =   285
            Left            =   2550
            TabIndex        =   52
            Top             =   615
            Width           =   255
            _ExtentX        =   450
            _ExtentY        =   503
            _Version        =   327681
            BuddyControl    =   "txtBackgroundParsingInterval"
            BuddyDispid     =   196642
            OrigLeft        =   175
            OrigTop         =   63
            OrigRight       =   192
            OrigBottom      =   85
            Max             =   60
            SyncBuddy       =   -1  'True
            BuddyProperty   =   0
            Enabled         =   -1  'True
         End
         Begin VB.Label lblBackgroundParsingInterval 
            BackStyle       =   0  'Transparent
            Caption         =   "Interval (seconds):"
            Height          =   225
            Left            =   315
            TabIndex        =   55
            Top             =   630
            Width           =   1380
         End
      End
      Begin VB.CheckBox chkAutoIndent 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "AutoIndent code"
         ForeColor       =   &H80000008&
         Height          =   225
         Left            =   105
         TabIndex        =   33
         Tag             =   "Code_AutoIndent"
         Top             =   105
         Width           =   2325
      End
   End
   Begin VB.PictureBox picTabs 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   3495
      Index           =   0
      Left            =   165
      ScaleHeight     =   233
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   346
      TabIndex        =   3
      Top             =   525
      Width           =   5190
      Begin VB.CheckBox chkCompatibleLook 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Compatible look && feel"
         ForeColor       =   &H80000008&
         Height          =   225
         Left            =   105
         TabIndex        =   63
         Tag             =   "General_CompatibleLook"
         Top             =   1785
         Width           =   1905
      End
      Begin VB.CheckBox chkRunDukeWithCON 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Run Duke3D with current CON file"
         ForeColor       =   &H80000008&
         Height          =   225
         Left            =   105
         TabIndex        =   62
         Tag             =   "General_RunDukeWithCON"
         Top             =   1365
         Width           =   2745
      End
      Begin VB.TextBox txtRecent 
         Height          =   300
         Left            =   1605
         Locked          =   -1  'True
         TabIndex        =   30
         Tag             =   "Recent_Count"
         Top             =   840
         Width           =   600
      End
      Begin VB.CheckBox chkDefaultDoc 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Create default document on startup"
         ForeColor       =   &H80000008&
         Height          =   225
         Left            =   105
         TabIndex        =   21
         Tag             =   "DefaultDocumentOnStartup"
         Top             =   420
         Width           =   2955
      End
      Begin VB.CheckBox chkSplash 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Display splash screen on startup"
         ForeColor       =   &H80000008&
         Height          =   225
         Left            =   105
         TabIndex        =   20
         Tag             =   "SplashScreenOnStartup"
         Top             =   105
         Width           =   2640
      End
      Begin ComCtl2.UpDown updRecent 
         Height          =   300
         Left            =   2220
         TabIndex        =   29
         Top             =   840
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   529
         _Version        =   327681
         BuddyControl    =   "txtRecent"
         BuddyDispid     =   196648
         OrigLeft        =   150
         OrigTop         =   85
         OrigRight       =   167
         OrigBottom      =   106
         SyncBuddy       =   -1  'True
         BuddyProperty   =   65547
         Enabled         =   -1  'True
      End
      Begin VB.Label lblRecent 
         BackStyle       =   0  'Transparent
         Caption         =   "Max. Recent Items:"
         Height          =   255
         Left            =   105
         TabIndex        =   31
         Top             =   840
         Width           =   1485
      End
   End
   Begin MSComDlg.CommonDialog cdl 
      Left            =   600
      Top             =   3600
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "Help"
      Height          =   345
      Left            =   4305
      TabIndex        =   2
      Top             =   4140
      Width           =   1125
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   345
      Left            =   3150
      TabIndex        =   1
      Top             =   4140
      Width           =   1125
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   345
      Left            =   1995
      TabIndex        =   0
      Top             =   4140
      Width           =   1125
   End
   Begin ComctlLib.TabStrip tabProps 
      Height          =   3960
      Left            =   120
      TabIndex        =   28
      Top             =   120
      Width           =   5310
      _ExtentX        =   9366
      _ExtentY        =   6985
      _Version        =   327682
      BeginProperty Tabs {0713E432-850A-101B-AFC0-4210102A8DA7} 
         NumTabs         =   5
         BeginProperty Tab1 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "&General"
            Key             =   "General"
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "&Code"
            Key             =   "Code"
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "&AutoFormat"
            Key             =   "AutoFormat"
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab4 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "&Editor"
            Key             =   "Editor"
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab5 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "&File Locations"
            Key             =   "File Locations"
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox picTabs 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   3495
      Index           =   4
      Left            =   165
      ScaleHeight     =   233
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   346
      TabIndex        =   4
      Top             =   480
      Width           =   5190
      Begin VB.CommandButton cmdPrimitives 
         BackColor       =   &H80000005&
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
         Left            =   4785
         TabIndex        =   19
         Top             =   1830
         Width           =   300
      End
      Begin VB.TextBox txtPrimitives 
         Height          =   300
         Left            =   1665
         TabIndex        =   18
         Tag             =   "Path_PrimitivesDef"
         Top             =   1830
         Width           =   3015
      End
      Begin VB.TextBox txtFile 
         Height          =   300
         Index           =   4
         Left            =   1665
         TabIndex        =   15
         Tag             =   "Path_Art"
         Top             =   1350
         Width           =   3015
      End
      Begin VB.CommandButton cmdFile 
         BackColor       =   &H80000005&
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
         Index           =   4
         Left            =   4785
         TabIndex        =   14
         Top             =   1350
         Width           =   300
      End
      Begin VB.CommandButton cmdFile 
         BackColor       =   &H80000005&
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
         Index           =   2
         Left            =   4785
         TabIndex        =   13
         Top             =   630
         Width           =   300
      End
      Begin VB.TextBox txtFile 
         Height          =   300
         Index           =   2
         Left            =   1665
         TabIndex        =   12
         Tag             =   "Path_Maps"
         Top             =   630
         Width           =   3015
      End
      Begin VB.CommandButton cmdFile 
         BackColor       =   &H80000005&
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
         Index           =   1
         Left            =   4785
         TabIndex        =   10
         Top             =   120
         Width           =   300
      End
      Begin VB.TextBox txtFile 
         Height          =   300
         Index           =   1
         Left            =   1665
         TabIndex        =   9
         Tag             =   "Path_Duke3D"
         Top             =   120
         Width           =   3015
      End
      Begin VB.CommandButton cmdFile 
         BackColor       =   &H80000005&
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
         Index           =   0
         Left            =   4785
         TabIndex        =   7
         Top             =   990
         Width           =   300
      End
      Begin VB.TextBox txtFile 
         Height          =   300
         Index           =   0
         Left            =   1665
         TabIndex        =   6
         Tag             =   "Path_Cons"
         Top             =   990
         Width           =   3015
      End
      Begin VB.Label lblPrimitives 
         BackStyle       =   0  'Transparent
         Caption         =   "Parser Definition:"
         Height          =   225
         Left            =   105
         TabIndex        =   17
         Top             =   1830
         Width           =   1485
      End
      Begin VB.Label lblFile 
         BackStyle       =   0  'Transparent
         Caption         =   "Art Files (*.ART):"
         Height          =   255
         Index           =   4
         Left            =   105
         TabIndex        =   16
         Top             =   1350
         Width           =   1500
      End
      Begin VB.Label lblFile 
         BackStyle       =   0  'Transparent
         Caption         =   "Maps (*.MAP):"
         Height          =   255
         Index           =   2
         Left            =   105
         TabIndex        =   11
         Top             =   630
         Width           =   1500
      End
      Begin VB.Label lblFile 
         BackStyle       =   0  'Transparent
         Caption         =   "Duke Nukem 3D:"
         Height          =   255
         Index           =   1
         Left            =   105
         TabIndex        =   8
         Top             =   120
         Width           =   1500
      End
      Begin VB.Label lblFile 
         BackStyle       =   0  'Transparent
         Caption         =   "CONs (*.CON):"
         Height          =   255
         Index           =   0
         Left            =   105
         TabIndex        =   5
         Top             =   990
         Width           =   1500
      End
   End
   Begin VB.PictureBox picTabs 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   3495
      Index           =   3
      Left            =   165
      ScaleHeight     =   233
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   346
      TabIndex        =   22
      Top             =   525
      Width           =   5190
      Begin VB.Frame fraEditorFonts 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Editor Fonts"
         ForeColor       =   &H80000008&
         Height          =   1065
         Left            =   105
         TabIndex        =   23
         Top             =   75
         Width           =   4890
         Begin VB.ComboBox cboEditorFontSize 
            Appearance      =   0  'Flat
            Height          =   315
            ItemData        =   "FrmOptions.frx":0902
            Left            =   3525
            List            =   "FrmOptions.frx":0936
            TabIndex        =   27
            Tag             =   "EditorFontSize"
            Top             =   525
            Width           =   1215
         End
         Begin VB.ComboBox cboFonts 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   150
            Sorted          =   -1  'True
            TabIndex        =   25
            Tag             =   "EditorFontName"
            Top             =   525
            Width           =   3240
         End
         Begin VB.Label lblSize 
            BackStyle       =   0  'Transparent
            Caption         =   "Size:"
            Height          =   240
            Left            =   3525
            TabIndex        =   26
            Top             =   300
            Width           =   465
         End
         Begin VB.Label lblFont 
            BackStyle       =   0  'Transparent
            Caption         =   "Font:"
            Height          =   240
            Left            =   150
            TabIndex        =   24
            Top             =   300
            Width           =   465
         End
      End
   End
   Begin VB.PictureBox picTabs 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   3495
      Index           =   2
      Left            =   165
      ScaleHeight     =   233
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   346
      TabIndex        =   34
      Top             =   525
      Width           =   5190
      Begin ComctlLib.Toolbar tlbElementFormat 
         Height          =   390
         Left            =   2415
         TabIndex        =   46
         Top             =   1725
         Width           =   2640
         _ExtentX        =   4657
         _ExtentY        =   688
         ButtonWidth     =   609
         ButtonHeight    =   582
         ImageList       =   "imlElementFormat"
         _Version        =   327682
         BeginProperty Buttons {0713E452-850A-101B-AFC0-4210102A8DA7} 
            NumButtons      =   4
            BeginProperty Button1 {0713F354-850A-101B-AFC0-4210102A8DA7} 
               Key             =   "Bold"
               Object.ToolTipText     =   "Bold"
               Object.Tag             =   ""
               ImageIndex      =   1
               Style           =   1
            EndProperty
            BeginProperty Button2 {0713F354-850A-101B-AFC0-4210102A8DA7} 
               Key             =   "Italic"
               Object.ToolTipText     =   "Italic"
               Object.Tag             =   ""
               ImageIndex      =   2
               Style           =   1
            EndProperty
            BeginProperty Button3 {0713F354-850A-101B-AFC0-4210102A8DA7} 
               Key             =   "Underline"
               Object.ToolTipText     =   "Underline"
               Object.Tag             =   ""
               ImageIndex      =   3
               Style           =   1
            EndProperty
            BeginProperty Button4 {0713F354-850A-101B-AFC0-4210102A8DA7} 
               Key             =   "Strikethru"
               Object.ToolTipText     =   "Strikethru"
               Object.Tag             =   ""
               ImageIndex      =   4
               Style           =   1
            EndProperty
         EndProperty
      End
      Begin VB.CheckBox chkColorKeyWords 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Color in Keywords"
         ForeColor       =   &H80000008&
         Height          =   225
         Left            =   2400
         TabIndex        =   64
         Tag             =   "AutoFormat_EnableKeywords"
         Top             =   120
         Width           =   2325
      End
      Begin VB.CommandButton cmdElementReset 
         BackColor       =   &H80000005&
         Caption         =   "&Reset"
         Height          =   330
         Left            =   3990
         TabIndex        =   49
         Top             =   3045
         Width           =   1065
      End
      Begin VB.CommandButton cmdElementSave 
         BackColor       =   &H80000005&
         Caption         =   "&Save"
         Height          =   330
         Left            =   2835
         TabIndex        =   48
         Top             =   3045
         Width           =   1065
      End
      Begin ComctlLib.ListView lvwElements 
         Height          =   2745
         Left            =   105
         TabIndex        =   47
         Top             =   630
         Width           =   2220
         _ExtentX        =   3916
         _ExtentY        =   4842
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         HideColumnHeaders=   -1  'True
         _Version        =   327682
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         Appearance      =   1
         NumItems        =   1
         BeginProperty ColumnHeader(1) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Text"
            Object.Tag             =   ""
            Text            =   "Text"
            Object.Width           =   3307
         EndProperty
      End
      Begin VB.TextBox txtElementSample 
         Height          =   435
         Left            =   2415
         Locked          =   -1  'True
         TabIndex        =   45
         Text            =   "AaBbYyZz"
         Top             =   2415
         Width           =   2640
      End
      Begin VB.CommandButton cmdElementColor 
         BackColor       =   &H80000005&
         Caption         =   "..."
         Height          =   300
         Left            =   4725
         TabIndex        =   43
         Top             =   1290
         Width           =   300
      End
      Begin VB.ComboBox cboElementSize 
         Appearance      =   0  'Flat
         Height          =   315
         ItemData        =   "FrmOptions.frx":0978
         Left            =   2415
         List            =   "FrmOptions.frx":09AC
         TabIndex        =   39
         Top             =   1275
         Width           =   1215
      End
      Begin VB.ComboBox cboElementFont 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   2415
         Sorted          =   -1  'True
         TabIndex        =   37
         Top             =   630
         Width           =   2610
      End
      Begin VB.CheckBox chkEnableAutoFormat 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Enable AutoFormat"
         ForeColor       =   &H80000008&
         Height          =   225
         Left            =   105
         TabIndex        =   35
         Tag             =   "Code_AutoFormat"
         Top             =   105
         Width           =   2325
      End
      Begin ComctlLib.ImageList imlElementFormat 
         Left            =   2730
         Top             =   315
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   16
         ImageHeight     =   16
         MaskColor       =   16711935
         _Version        =   327682
         BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
            NumListImages   =   4
            BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "FrmOptions.frx":09F0
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "FrmOptions.frx":0B02
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "FrmOptions.frx":0C14
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "FrmOptions.frx":0D26
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin VB.Label lblSample 
         BackStyle       =   0  'Transparent
         Caption         =   "Sample:"
         Height          =   225
         Left            =   2415
         TabIndex        =   44
         Top             =   2205
         Width           =   645
      End
      Begin VB.Label lblElementColor 
         BackColor       =   &H00404040&
         Height          =   300
         Left            =   3780
         TabIndex        =   42
         Top             =   1290
         Width           =   855
      End
      Begin VB.Label lblElementColor_Label 
         BackStyle       =   0  'Transparent
         Caption         =   "Color:"
         Height          =   240
         Left            =   3780
         TabIndex        =   41
         Top             =   1050
         Width           =   465
      End
      Begin VB.Label lblElementSize 
         BackStyle       =   0  'Transparent
         Caption         =   "Size:"
         Height          =   240
         Left            =   2415
         TabIndex        =   40
         Top             =   1050
         Width           =   465
      End
      Begin VB.Label lblElementFont 
         BackStyle       =   0  'Transparent
         Caption         =   "Font:"
         Height          =   240
         Left            =   2415
         TabIndex        =   38
         Top             =   420
         Width           =   465
      End
      Begin VB.Label lblElements 
         BackStyle       =   0  'Transparent
         Caption         =   "Elements:"
         Height          =   225
         Left            =   105
         TabIndex        =   36
         Top             =   420
         Width           =   2010
      End
   End
End
Attribute VB_Name = "FrmOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' =============================================================================
' Module Name:      FrmOptions
' Module Type:      User Form
' Description:      Tabbed dialog of options for CONstruct
' Author(s):        Jonathan A. Conway
' -----------------------------------------------------------------------------
' Log:
'
' 04 07 16 :
'   - Added "Default Document on Startup" setting
'
' 04 07 14 :
'   - Added "Primitives Definition Path" setting to facilitate primitives
'     generator
'   - Added "Display spash screen" setting in case someone wants to remove
'     that annoying pop-up screen on startup!!
'
' 04 05 05 :
'   Created FrmOptions
' =============================================================================


' =============================================================================
' INFO ABOUT ADDING/REMOVING OPTIONS:
' -----------------------------------
'
' To create a new option, follow these (relatively) simple steps:
'
' 1. Choose a tab to put the option on (make one if necessary) and bring its
'    related picturebox to the front so as to add controls to it.
'
' 2. Create the PRIMARY control of the option. The control can be a:
'    a. Textbox
'    b. Checkbox
'    c. Optionbutton *
'    d. Combobox *
'    e. Listbox *
'    f. Scrollbar *
'    g. Slider *
'
'    (* = yet to be coded)
'
' 3. Set the TAG property of the control to the name of the setting you want
'    to connect it to.
'
' 4. Create any helper controls needed (e.g. builder buttons, up-downs, etc.)
' =============================================================================


Option Explicit


' Private Variables
' =================

Private WithEvents c As cBrowseForFolder
Attribute c.VB_VarHelpID = -1

Private m_sSettingsChanged As String

Private m_bIsLoading As Boolean


' Private Methods
' ===============

Private Sub LoadSettings()
    m_bIsLoading = True
    Dim cCtl As Control
    For Each cCtl In Controls
        If Trim(cCtl.Tag) <> "" Then
            If TypeOf cCtl Is TextBox _
               Or TypeOf cCtl Is ComboBox Then
                cCtl.Text = Settings.ReadSetting(cCtl.Tag)
            ElseIf TypeOf cCtl Is CheckBox Then
                cCtl.Value = IIf(Settings.ReadSetting(cCtl.Tag) = "yes", vbChecked, vbUnchecked)
            End If
        End If
    Next

    chkBackgroundParsing_Click
    chkAutoSave_Click
    chkEnableAutoFormat_Click
    
    m_bIsLoading = False
        
End Sub

Private Sub SaveSettings()

    Dim cCtl As Control
    For Each cCtl In Controls
        If Trim(cCtl.Tag) <> "" Then
            If TypeOf cCtl Is TextBox _
               Or TypeOf cCtl Is ComboBox Then
                Settings.WriteSetting cCtl.Tag, cCtl.Text
            ElseIf TypeOf cCtl Is CheckBox Then
                Settings.WriteSetting cCtl.Tag, IIf(cCtl.Value = vbChecked, "yes", "no")
            End If
        End If
    Next
    
End Sub

Private Sub UpdateElementSample()

    With txtElementSample
        .FontName = IIf(Len(Trim$(cboElementFont.Text)) > 0, cboElementFont.Text, Me.FontName)
        
        Dim lSize As Long
        lSize = FixInteger(cboElementSize.Text)
        .FontSize = IIf(lSize = 0, Me.FontSize, lSize)
        
        .ForeColor = lblElementColor.BackColor
        .FontBold = (tlbElementFormat.Buttons(1).Value = tbrPressed)
        .FontItalic = (tlbElementFormat.Buttons(2).Value = tbrPressed)
        .FontUnderline = (tlbElementFormat.Buttons(3).Value = tbrPressed)
        .FontStrikethru = (tlbElementFormat.Buttons(4).Value = tbrPressed)
    End With

End Sub

Private Sub LoadElement(ByVal ElementName As String)
    
    Dim sValue As String
    sValue = Settings.ReadSetting("AutoFormat_" & ElementName)

    If Len(Trim$(sValue)) > 0 Then
    
        cboElementFont.Text = GetTagAttribute(sValue, "fontname")
        cboElementSize.Text = GetTagAttribute(sValue, "fontsize")
        lblElementColor.BackColor = FixLong(GetTagAttribute(sValue, "fontcolor"))
        With tlbElementFormat
            .Buttons(1).Value = IIf(GetTagAttribute(sValue, "fontbold") = "yes", tbrPressed, tbrUnpressed)
            .Buttons(2).Value = IIf(GetTagAttribute(sValue, "fontitalic") = "yes", tbrPressed, tbrUnpressed)
            .Buttons(3).Value = IIf(GetTagAttribute(sValue, "fontunderline") = "yes", tbrPressed, tbrUnpressed)
            .Buttons(4).Value = IIf(GetTagAttribute(sValue, "fontstrikethru") = "yes", tbrPressed, tbrUnpressed)
        End With
    
    End If

    UpdateElementSample

End Sub



' Event Handlers
' ==============

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdElementColor_Click()
    On Error GoTo ProcedureError
    
    With cdl
        .CancelError = True
        .Color = lblElementColor.BackColor
        
        .ShowColor
        
        lblElementColor.BackColor = .Color
        UpdateElementSample
    End With
    Exit Sub
    
ProcedureError:
    If err.Number = 32755 Then
    Else
        MsgBox err.Number & vbNewLine & err.Description
    End If
End Sub

Private Sub cmdElementReset_Click()
    
    cboElementFont.Text = cboFonts.Text
    cboElementSize.Text = cboEditorFontSize.Text
    lblElementColor.BackColor = vbBlack
    With tlbElementFormat
        .Buttons(1).Value = tbrUnpressed
        .Buttons(2).Value = tbrUnpressed
        .Buttons(3).Value = tbrUnpressed
        .Buttons(4).Value = tbrUnpressed
    End With
    
    UpdateElementSample
    
End Sub

Private Sub cmdElementSave_Click()
    
    If Not (lvwElements.SelectedItem Is Nothing) Then
        Dim sValue As String
            
        sValue = SetTagAttribute(sValue, "fontname", cboElementFont.Text)
        sValue = SetTagAttribute(sValue, "fontsize", cboElementSize.Text)
        sValue = SetTagAttribute(sValue, "fontcolor", lblElementColor.BackColor)
        With tlbElementFormat
            sValue = SetTagAttribute(sValue, "fontbold", IIf(.Buttons(1).Value = tbrPressed, "yes", "no"))
            sValue = SetTagAttribute(sValue, "fontitalic", IIf(.Buttons(2).Value = tbrPressed, "yes", "no"))
            sValue = SetTagAttribute(sValue, "fontunderline", IIf(.Buttons(3).Value = tbrPressed, "yes", "no"))
            sValue = SetTagAttribute(sValue, "fontstrikethru", IIf(.Buttons(4).Value = tbrPressed, "yes", "no"))
        End With
    
        Settings.WriteSetting "AutoFormat_" & lvwElements.SelectedItem.Text, sValue
    End If

End Sub

Private Sub cmdFile_Click(Index As Integer)
    Dim s As String
    
    c.hwndOwner = Me.hWnd
    c.InitialDir = App.Path
    c.FileSystemOnly = True
    c.StatusText = True
    c.EditBox = True
    c.UseNewUI = True
    s = c.BrowseForFolder
    If Len(s) > 0 Then
        With txtFile(Index)
            .Text = s
            .SelStart = 0
            .SelLength = Len(s)
            .SetFocus
        End With
    End If
End Sub

Private Sub cmdOK_Click()
    SaveSettings
    Unload Me
End Sub

Private Sub cmdPrimitives_Click()
    On Error GoTo ProcedureError
    
    With cdl
        .CancelError = True
        .DialogTitle = "Browse for Primitives Definition"
        .Filter = "CONstruct Primitive Defenition (*.csp)|*.csp|All Files (*.*)|*.*"
        .FilterIndex = 0
        .Flags = cdlOFNFileMustExist
    
        If Len(txtPrimitives.Text) > 0 Then .Filename = txtPrimitives.Text
        
        .ShowOpen
        
        With txtPrimitives
            .Text = GetFilenameFromPath(cdl.Filename)
            .SelStart = 0
            .SelLength = Len(.Text)
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

Private Sub lvwElements_ItemClick(ByVal Item As ComctlLib.ListItem)

    'If Item.Text = "[Keywords]" Then
    '    LoadKeywordsElement
    'Else
    
    LoadElement Item.Text
    
    'End If

End Sub

Private Sub tabProps_Click()
    picTabs(tabProps.SelectedItem.Index - 1).ZOrder vbBringToFront
End Sub

Private Sub Form_Load()
    
    Set c = New cBrowseForFolder
    
    LoadFontCombos
    LoadElements
    
    LoadSettings
    tabProps_Click
    
    SetCompatibleColours Me
    
End Sub

Private Sub LoadElements()
    
    lvwElements.ListItems.Clear
    
    Dim oStruct As CStructure
    For Each oStruct In FrmMain.Definition.Structures
        lvwElements.ListItems.Add , , oStruct.StructureName
    Next

    lvwElements.ListItems.Add , , "[Keywords]"

    If lvwElements.ListItems.Count > 0 Then
        lvwElements.ListItems(1).Selected = True
        LoadElement lvwElements.ListItems(1).Text
        EnableDisableControls True
    Else
        EnableDisableControls False
    End If

End Sub

Private Sub EnableDisableControls(ByVal State As Boolean)

    lblElementFont.Enabled = State
    cboElementFont.Enabled = State
    lblElementSize.Enabled = State
    cboElementSize.Enabled = State
    lblElementColor.Enabled = State
    lblElementColor_Label.Enabled = State
    cmdElementColor.Enabled = State
    tlbElementFormat.Enabled = State
    lblSample.Enabled = State
    txtElementSample.Enabled = State
    cmdElementReset.Enabled = State
    cmdElementSave.Enabled = State

End Sub

Private Sub LoadFontCombos()

    m_bIsLoading = True
    Dim i As Integer
    Dim sFont As String
    For i = 0 To Screen.FontCount
        sFont = Screen.Fonts(i)
        If Len(Trim$(sFont)) > 0 Then
            cboFonts.AddItem sFont
            cboElementFont.AddItem sFont
        End If
    Next
    m_bIsLoading = False
    
End Sub

Private Sub cboEditorFontSize_Change()
    ValidateNumericTextBox cboEditorFontSize
    If InStr(1, m_sSettingsChanged, "[EditorFont]", vbBinaryCompare) = 0 Then
        m_sSettingsChanged = m_sSettingsChanged & " [EditorFont]"
    End If
End Sub

Private Sub cboFonts_Change()
    If Not m_bIsLoading Then
        Dim i As Integer
        For i = 0 To cboFonts.ListCount - 1
            If UCase(cboFonts.List(i)) = UCase(cboFonts.Text) Then
                cboFonts.Text = cboFonts.List(i)
                cboFonts.SelStart = Len(cboFonts.Text)
            End If
        Next
        
        If InStr(1, m_sSettingsChanged, "[EditorFont]", vbBinaryCompare) = 0 Then
            m_sSettingsChanged = m_sSettingsChanged & " [EditorFont]"
        End If
    End If
End Sub

Public Property Get SettingsChanged() As String
    SettingsChanged = m_sSettingsChanged
End Property

Private Sub tlbElementFormat_ButtonClick(ByVal Button As ComctlLib.Button)
    UpdateElementSample
End Sub

Private Sub txtBackgroundParsingInterval_Change()
    If InStr(1, m_sSettingsChanged, "[Code_BackgroundParsingInterval]", vbBinaryCompare) = 0 Then
        m_sSettingsChanged = m_sSettingsChanged & " [Code_BackgroundParsingInterval]"
    End If
End Sub

Private Sub txtPrimitives_Change()
    If InStr(1, m_sSettingsChanged, "[Path_PrimitivesDef]", vbBinaryCompare) = 0 Then
        m_sSettingsChanged = m_sSettingsChanged & " [Path_PrimitivesDef]"
    End If
End Sub

Private Sub chkBackgroundParsing_Click()
    If InStr(1, m_sSettingsChanged, "[Code_EnableBackgroundParsing]", vbBinaryCompare) = 0 Then
        m_sSettingsChanged = m_sSettingsChanged & " [Code_EnableBackgroundParsing]"
    End If
    
    Dim bState As Boolean
    bState = (chkBackgroundParsing.Value = vbChecked)
    lblBackgroundParsingInterval.Enabled = bState
    txtBackgroundParsingInterval.Enabled = bState
    updBackgroundParsingInterval.Enabled = bState
End Sub


Private Sub cboElementFont_Click()
    UpdateElementSample
End Sub

Private Sub cboElementSize_Click()
    UpdateElementSample
End Sub

Private Sub cboElementFont_Change()
    UpdateElementSample
End Sub

Private Sub cboElementSize_Change()
    UpdateElementSample
End Sub

Private Sub chkAutoSave_Click()
    Dim bState As Boolean
    bState = (chkAutoSave.Value = vbChecked)
    lblAutoSaveInterval.Enabled = bState
    txtAutoSaveInterval.Enabled = bState
    updAutoSaveInterval.Enabled = bState
End Sub

Private Sub chkEnableAutoFormat_Click()
    Dim bState As Boolean
    bState = (chkEnableAutoFormat.Value = vbChecked)
    lblElements.Enabled = bState
    lvwElements.Enabled = bState
    chkColorKeyWords.Enabled = bState
    lblElementFont.Enabled = bState
    cboElementFont.Enabled = bState
    lblElementSize.Enabled = bState
    cboElementSize.Enabled = bState
    lblElementColor.Enabled = bState
    lblElementColor_Label.Enabled = bState
    cmdElementColor.Enabled = bState
    tlbElementFormat.Enabled = bState
    lblSample.Enabled = bState
    txtElementSample.Enabled = bState
    cmdElementSave.Enabled = bState
    cmdElementReset.Enabled = bState
End Sub

