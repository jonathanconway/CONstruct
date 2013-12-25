VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Begin VB.Form FrmDefine 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Define"
   ClientHeight    =   5175
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5640
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FrmDefine.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   345
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   376
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin ComctlLib.TabStrip tabstrip 
      Height          =   3735
      Left            =   150
      TabIndex        =   62
      Top             =   840
      Width           =   5415
      _ExtentX        =   9551
      _ExtentY        =   6588
      _Version        =   327682
      BeginProperty Tabs {0713E432-850A-101B-AFC0-4210102A8DA7} 
         NumTabs         =   2
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
      EndProperty
   End
   Begin MSComDlg.CommonDialog cdl 
      Left            =   360
      Top             =   4560
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.PictureBox picTabs 
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3255
      Index           =   0
      Left            =   240
      ScaleHeight     =   217
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   345
      TabIndex        =   3
      Top             =   1200
      Width           =   5175
      Begin VB.PictureBox picDefine 
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         Height          =   2655
         Index           =   0
         Left            =   0
         ScaleHeight     =   2655
         ScaleWidth      =   5175
         TabIndex        =   11
         Top             =   600
         Width           =   5175
         Begin VB.CommandButton cmdValue 
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
            Left            =   1800
            TabIndex        =   59
            Top             =   360
            Width           =   300
         End
         Begin VB.TextBox txtValue 
            Height          =   300
            Index           =   1
            Left            =   600
            TabIndex        =   12
            Tag             =   "[typeid]=1, [numeric]=yes, [property]=1"
            Text            =   "0"
            Top             =   360
            Width           =   1095
         End
         Begin VB.TextBox txtValue 
            Height          =   300
            Index           =   0
            Left            =   600
            TabIndex        =   13
            Tag             =   "[typeid]=1, [numeric]=no, [property]=0"
            Text            =   "[empty]"
            Top             =   0
            Width           =   2655
         End
         Begin VB.Label lblString 
            BackStyle       =   0  'Transparent
            Caption         =   "String:"
            Height          =   255
            Left            =   0
            TabIndex        =   15
            Top             =   0
            Width           =   495
         End
         Begin VB.Label lblValue 
            BackStyle       =   0  'Transparent
            Caption         =   "Value:"
            Height          =   255
            Left            =   0
            TabIndex        =   14
            Top             =   360
            Width           =   495
         End
      End
      Begin VB.PictureBox picDefine 
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         Height          =   2655
         Index           =   1
         Left            =   0
         ScaleHeight     =   2655
         ScaleWidth      =   5175
         TabIndex        =   16
         Top             =   600
         Width           =   5175
         Begin VB.TextBox txtValue 
            Height          =   300
            Index           =   3
            Left            =   600
            TabIndex        =   18
            Tag             =   "[typeid]=2, [numeric]=yes, [property]=0"
            Text            =   "0"
            Top             =   360
            Width           =   1095
         End
         Begin VB.TextBox txtValue 
            Height          =   300
            Index           =   2
            Left            =   600
            TabIndex        =   17
            Tag             =   "[typeid]=2, [numeric]=no, [property]=1"
            Text            =   "[empty]"
            Top             =   0
            Width           =   2655
         End
         Begin VB.Label lblQuoValue 
            BackStyle       =   0  'Transparent
            Caption         =   "Value:"
            Height          =   255
            Left            =   0
            TabIndex        =   20
            Top             =   360
            Width           =   495
         End
         Begin VB.Label lblQuoQuote 
            BackStyle       =   0  'Transparent
            Caption         =   "Quote:"
            Height          =   255
            Left            =   0
            TabIndex        =   19
            Top             =   0
            Width           =   495
         End
      End
      Begin VB.PictureBox picDefine 
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         Height          =   2655
         Index           =   2
         Left            =   0
         ScaleHeight     =   2655
         ScaleWidth      =   5175
         TabIndex        =   21
         Top             =   600
         Width           =   5175
         Begin VB.TextBox txtValue 
            Height          =   300
            Index           =   5
            Left            =   720
            TabIndex        =   23
            Tag             =   "[typeid]=3, [numeric]=no, [property]=1"
            Text            =   "[empty]"
            Top             =   360
            Width           =   2655
         End
         Begin VB.TextBox txtValue 
            Height          =   300
            Index           =   4
            Left            =   720
            TabIndex        =   22
            Tag             =   "[typeid]=3, [numeric]=yes, [property]=0"
            Text            =   "0"
            Top             =   0
            Width           =   495
         End
         Begin VB.Label lblVolEpisode 
            BackStyle       =   0  'Transparent
            Caption         =   "Episode:"
            Height          =   255
            Left            =   0
            TabIndex        =   25
            Top             =   0
            Width           =   615
         End
         Begin VB.Label lblVolName 
            BackStyle       =   0  'Transparent
            Caption         =   "Name:"
            Height          =   255
            Left            =   0
            TabIndex        =   24
            Top             =   360
            Width           =   615
         End
      End
      Begin VB.PictureBox picDefine 
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         Height          =   2655
         Index           =   3
         Left            =   0
         ScaleHeight     =   2655
         ScaleWidth      =   5175
         TabIndex        =   26
         Top             =   600
         Width           =   5175
         Begin VB.ComboBox cboSkillSkillNo 
            Height          =   315
            ItemData        =   "FrmDefine.frx":058A
            Left            =   720
            List            =   "FrmDefine.frx":059A
            Style           =   2  'Dropdown List
            TabIndex        =   30
            Top             =   0
            Width           =   735
         End
         Begin VB.TextBox txtValue 
            Height          =   300
            Index           =   17
            Left            =   720
            TabIndex        =   27
            Tag             =   "[typeid]=4, [numeric]=no, [property]=1"
            Top             =   360
            Width           =   2655
         End
         Begin VB.Label lblSkillString 
            BackStyle       =   0  'Transparent
            Caption         =   "String:"
            Height          =   255
            Left            =   0
            TabIndex        =   29
            Top             =   360
            Width           =   615
         End
         Begin VB.Label lblSkillSkillNo 
            BackStyle       =   0  'Transparent
            Caption         =   "Skill:"
            Height          =   255
            Left            =   0
            TabIndex        =   28
            Top             =   0
            Width           =   615
         End
      End
      Begin VB.PictureBox picDefine 
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         Height          =   2655
         Index           =   4
         Left            =   0
         ScaleHeight     =   177
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   345
         TabIndex        =   31
         Top             =   600
         Width           =   5175
         Begin VB.TextBox txtValue 
            Height          =   300
            Index           =   9
            Left            =   1320
            TabIndex        =   60
            Tag             =   "[typeid]=5, [numeric]=no, [property]=5"
            Text            =   "[empty]"
            Top             =   1800
            Width           =   2535
         End
         Begin CONstruct.CONTimeEdit txtLvlParTime 
            Height          =   300
            Left            =   1320
            TabIndex        =   41
            Top             =   1080
            Width           =   735
            _ExtentX        =   1296
            _ExtentY        =   529
         End
         Begin VB.CommandButton cmdLvlMapName 
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
            Left            =   3960
            TabIndex        =   38
            Top             =   720
            Width           =   300
         End
         Begin VB.TextBox txtValue 
            Height          =   300
            Index           =   7
            Left            =   1320
            TabIndex        =   36
            Tag             =   "[typeid]=5, [numeric]=yes, [property]=1"
            Text            =   "0"
            Top             =   360
            Width           =   495
         End
         Begin VB.TextBox txtValue 
            Height          =   300
            Index           =   6
            Left            =   1320
            TabIndex        =   33
            Tag             =   "[typeid]=5, [numeric]=yes, [property]=0"
            Text            =   "0"
            Top             =   0
            Width           =   495
         End
         Begin VB.TextBox txtValue 
            Height          =   300
            Index           =   8
            Left            =   1320
            TabIndex        =   32
            Tag             =   "[typeid]=5, [numeric]=no, [property]=2"
            Text            =   "[empty]"
            Top             =   720
            Width           =   2535
         End
         Begin CONstruct.CONTimeEdit txtLvl3DRealmsTime 
            Height          =   300
            Left            =   1320
            TabIndex        =   42
            Top             =   1440
            Width           =   735
            _ExtentX        =   1296
            _ExtentY        =   529
         End
         Begin VB.Label lblLvlLevelName 
            BackStyle       =   0  'Transparent
            Caption         =   "Level Name:"
            Height          =   255
            Left            =   0
            TabIndex        =   61
            Top             =   1800
            Width           =   1215
         End
         Begin VB.Label lblLvl3DRealmsTime 
            BackStyle       =   0  'Transparent
            Caption         =   "3D Realms Time:"
            Height          =   255
            Left            =   0
            TabIndex        =   40
            Top             =   1440
            Width           =   1215
         End
         Begin VB.Label lblLvlParTime 
            BackStyle       =   0  'Transparent
            Caption         =   "Par Time:"
            Height          =   255
            Left            =   0
            TabIndex        =   39
            Top             =   1080
            Width           =   1215
         End
         Begin VB.Label lblLvlLevel 
            BackStyle       =   0  'Transparent
            Caption         =   "Level:"
            Height          =   255
            Left            =   0
            TabIndex        =   37
            Top             =   360
            Width           =   1215
         End
         Begin VB.Label lblLvlMapName 
            BackStyle       =   0  'Transparent
            Caption         =   "Map Name:"
            Height          =   255
            Left            =   0
            TabIndex        =   35
            Top             =   720
            Width           =   1215
         End
         Begin VB.Label lblLvlEpisode 
            BackStyle       =   0  'Transparent
            Caption         =   "Episode:"
            Height          =   255
            Left            =   0
            TabIndex        =   34
            Top             =   0
            Width           =   1215
         End
      End
      Begin VB.PictureBox picDefine 
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         Height          =   2655
         Index           =   5
         Left            =   0
         ScaleHeight     =   177
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   345
         TabIndex        =   43
         Top             =   600
         Width           =   5175
         Begin VB.TextBox txtValue 
            Height          =   300
            Index           =   16
            Left            =   960
            TabIndex        =   57
            Tag             =   "[typeid]=6, [numeric]=yes, [property]=6"
            Text            =   "0"
            Top             =   1560
            Width           =   495
         End
         Begin VB.TextBox txtValue 
            Height          =   300
            Index           =   14
            Left            =   960
            TabIndex        =   54
            Tag             =   "[typeid]=6, [numeric]=yes, [property]=4"
            Text            =   "0"
            Top             =   840
            Width           =   495
         End
         Begin VB.TextBox txtValue 
            Height          =   300
            Index           =   15
            Left            =   960
            TabIndex        =   53
            Tag             =   "[typeid]=6, [numeric]=yes, [property]=5"
            Text            =   "0"
            Top             =   1200
            Width           =   495
         End
         Begin VB.TextBox txtValue 
            Height          =   300
            Index           =   10
            Left            =   960
            TabIndex        =   51
            Tag             =   "[typeid]=6, [numeric]=no, [property]=0"
            Text            =   "[empty]"
            Top             =   0
            Width           =   2535
         End
         Begin VB.TextBox txtValue 
            Height          =   300
            Index           =   11
            Left            =   960
            TabIndex        =   47
            Tag             =   "[typeid]=6, [numeric]=no, [property]=1"
            Text            =   "[empty]"
            Top             =   360
            Width           =   2535
         End
         Begin VB.TextBox txtValue 
            Height          =   300
            Index           =   12
            Left            =   2760
            TabIndex        =   46
            Tag             =   "[typeid]=6, [numeric]=yes, [property]=2"
            Text            =   "0"
            Top             =   840
            Width           =   495
         End
         Begin VB.TextBox txtValue 
            Height          =   300
            Index           =   13
            Left            =   2760
            TabIndex        =   45
            Tag             =   "[typeid]=6, [numeric]=yes, [property]=3"
            Text            =   "0"
            Top             =   1200
            Width           =   495
         End
         Begin VB.CommandButton cmdSndFile 
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
            Left            =   3600
            TabIndex        =   44
            Top             =   360
            Width           =   300
         End
         Begin VB.Line Line2 
            BorderColor     =   &H80000014&
            Visible         =   0   'False
            X1              =   -8
            X2              =   336
            Y1              =   21
            Y2              =   21
         End
         Begin VB.Line Line1 
            BorderColor     =   &H80000010&
            Visible         =   0   'False
            X1              =   -8
            X2              =   336
            Y1              =   20
            Y2              =   20
         End
         Begin VB.Label lblSndVolume 
            BackStyle       =   0  'Transparent
            Caption         =   "Volume:"
            Height          =   255
            Left            =   0
            TabIndex        =   58
            Top             =   1560
            Width           =   975
         End
         Begin VB.Label lblSndPriority 
            BackStyle       =   0  'Transparent
            Caption         =   "Priority:"
            Height          =   255
            Left            =   0
            TabIndex        =   56
            Top             =   840
            Width           =   975
         End
         Begin VB.Label lblSndType 
            BackStyle       =   0  'Transparent
            Caption         =   "Type:"
            Height          =   255
            Left            =   0
            TabIndex        =   55
            Top             =   1200
            Width           =   975
         End
         Begin VB.Label lblSndString 
            BackStyle       =   0  'Transparent
            Caption         =   "String:"
            Height          =   255
            Left            =   0
            TabIndex        =   52
            Top             =   0
            Width           =   975
         End
         Begin VB.Label lblSndPitch1 
            BackStyle       =   0  'Transparent
            Caption         =   "Pitch #1:"
            Height          =   255
            Left            =   1800
            TabIndex        =   50
            Top             =   840
            Width           =   855
         End
         Begin VB.Label lblSndFile 
            BackStyle       =   0  'Transparent
            Caption         =   "Sound File:"
            Height          =   255
            Left            =   0
            TabIndex        =   49
            Top             =   360
            Width           =   975
         End
         Begin VB.Label lblSndPitch2 
            BackStyle       =   0  'Transparent
            Caption         =   "Pitch #2:"
            Height          =   255
            Left            =   1800
            TabIndex        =   48
            Top             =   1200
            Width           =   855
         End
      End
      Begin VB.ComboBox cboDefineType 
         Height          =   315
         ItemData        =   "FrmDefine.frx":05AA
         Left            =   1080
         List            =   "FrmDefine.frx":05C0
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Top             =   120
         Width           =   1935
      End
      Begin VB.Label lblDefineType 
         BackStyle       =   0  'Transparent
         Caption         =   "Define Type:"
         Height          =   255
         Left            =   0
         TabIndex        =   9
         Top             =   120
         Width           =   975
      End
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      Default         =   -1  'True
      Height          =   375
      Left            =   1680
      TabIndex        =   2
      Top             =   4680
      Width           =   1215
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      CausesValidation=   0   'False
      Height          =   375
      Left            =   3000
      TabIndex        =   1
      Top             =   4680
      Width           =   1215
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "&Help"
      CausesValidation=   0   'False
      Height          =   375
      Left            =   4320
      TabIndex        =   0
      Top             =   4680
      Width           =   1215
   End
   Begin VB.PictureBox picTabs 
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3255
      Index           =   1
      Left            =   240
      ScaleHeight     =   3255
      ScaleWidth      =   5175
      TabIndex        =   4
      Top             =   1200
      Width           =   5175
      Begin CONstruct.CONEditor ceCode 
         Height          =   3135
         Left            =   0
         TabIndex        =   5
         Top             =   120
         Width           =   5175
         _ExtentX        =   9128
         _ExtentY        =   5530
      End
   End
   Begin VB.Image Image1 
      Height          =   720
      Left            =   120
      Picture         =   "FrmDefine.frx":061A
      Top             =   0
      Width           =   720
   End
   Begin VB.Label lblName 
      BackStyle       =   0  'Transparent
      Caption         =   "[Empty]"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   960
      TabIndex        =   8
      Top             =   0
      Width           =   4575
   End
   Begin VB.Label lblInfo 
      BackStyle       =   0  'Transparent
      Caption         =   "Length: 321; Lines: 3"
      Height          =   255
      Left            =   960
      TabIndex        =   7
      Top             =   480
      Width           =   4575
   End
   Begin VB.Label lblType 
      BackStyle       =   0  'Transparent
      Caption         =   "Define"
      Height          =   255
      Left            =   960
      TabIndex        =   6
      Top             =   240
      Width           =   4575
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
Attribute VB_Name = "FrmDefine"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'' =============================================================================
'' Module Name:      FrmDefine
'' Module Type:      User Form
'' Tags:             [CodeHelper]
'' Description:      Helper dialog for creating all kinds of Define structures
'' Author(s):        Jonathan A. Conway
'' -----------------------------------------------------------------------------
'' Log:
''
'' 04 07 16 :
''   - Fixed a small bug with the builder button (cmdValue) on "define".
''   - Thinking about turning the whole "definexxx" thing into a standard
''     primitive structure and including it in the primitives box.
'' 04 07 05 :
''   - Replaced zillions of seperate text-box controls for the various
''     define properties with a control array of controls. Each control
''     now has a tag that specifies its "TypeID", its "Property" number
''     and whether or not it takes numeric values only.
''   - Fixed heaps of bugs
''
'' 04 05 05 :
''   - Added functionality for DefineSound and fixed some annoying bugs to do
''     with the Code tab.
''
'' 04 05 02 :
''   - Made major alterations to the form to make it work with the new CDefine
''     class.
''
'' ?? ?? ?? :
''   - Created FrmDefine
'' =============================================================================
'
'Option Explicit
'
'
'' Private Variables
'' =================
'
'Private m_bAddNew As Boolean
'Private m_oObject As Object
'Private m_bIsCancelled As Boolean
'Private m_bInitState As Boolean
'
'
'
'' Public Properties
'' =================
'
'Public Property Get AddNew() As Boolean
'    AddNew = m_bAddNew
'End Property
'Public Property Let AddNew(ByVal NewValue As Boolean)
'    m_bAddNew = NewValue
'End Property
'
'Public Property Get IsCancelled() As Boolean
'    IsCancelled = m_bIsCancelled
'End Property
'
'Public Property Get Object() As Object
'    Set Object = m_oObject
'End Property
'Public Property Let Object(ByRef NewValue As Object)
'    Set m_oObject = NewValue
'End Property
'
'
'' Public Methods
'' ==============
'
'Public Sub Init()
'    m_bInitState = True
'    If m_bAddNew Then
'        cboDefineType.ListIndex = 0
'        InitValues
'    Else
'        InitValues
'    End If
'    m_bInitState = False
'End Sub
'
'
'' Private Methods
'' ===============
'
'Private Sub InitValues()
'
'    InitCaptions                ' Initialize captions
'
'    ' Initialize define type combo
'    cboDefineType.ListIndex = m_oObject.TypeID - 1
'
'    InitTaggedControls          ' Populate tagged controls
'
'    InitSpecificControls        ' Hard-coded intializations
'
'End Sub
'
'Private Sub InitCaptions()
'
'    ' Set captions, titles, etc.
'    Dim sCode As String
'    sCode = m_oObject.GetCode()
'    lblInfo.Caption = "Length: " & Len(sCode) & ", Lines: " & GetLineCount(sCode) & ", " & IIf(m_oObject.HasError, "Error(s) found", "No errors found")
'    Me.Caption = """" & m_oObject.ToString() & """ (Define)"
'    lblName.Caption = """" & m_oObject.ToString() & """"
'    lblType.Caption = m_oObject.TypeString
'
'End Sub
'
'
'Private Sub InitTaggedControls()
'
'    ' Set field values for all tagged controls
'    ' TODO : Add support for other controls!!
'
'    Dim cCtl As Control
'    Dim sTag As String
'    Dim sProp As String
'
'    For Each cCtl In Me.Controls
'        ' Is the control a textbox
'        If TypeOf cCtl Is TextBox Then
'            sTag = cCtl.Tag
'            sProp = Trim$(GetTagAttribute(sTag, "typeid"))
'            If IsNumeric(sProp) Then
'                ' Is the textbox tagged with an appropriate type ID?
'                If CInt(sProp) = m_oObject.TypeID Then
'                    ' Populate the text-box with the right value
'                    cCtl.Text = m_oObject.GetPropertyValue(CInt(Trim$(GetTagAttribute(sTag, "property"))))
'                End If
'            End If
'        End If
'    Next
'
'End Sub
'
'Private Sub InitSpecificControls()
'
'    ' Set field values for specific fields per define type
'
'    Select Case m_oObject.TypeID
'        Case [gbtDefineSkillName]
'            cboSkillSkillNo.ListIndex = CInt(m_oObject.GetPropertyValue(0))
'        Case [gbtDefineLevelName]
'            txtLvlParTime.Text = m_oObject.GetPropertyValue(3)
'            txtLvl3DRealmsTime.Text = m_oObject.GetPropertyValue(4)
'    End Select
'
'End Sub
'
'
'' Event Handlers
'' ==============
'
'Private Sub cboDefineType_Click()
'    picDefine(cboDefineType.ListIndex).ZOrder vbBringToFront
'
'    If Not m_bInitState Then
'        m_oObject.TypeID = cboDefineType.ListIndex + 1
'    End If
'End Sub
'
'Private Sub cmdCancel_Click()
'    m_bIsCancelled = True
'    Me.Hide
'End Sub
'
'Private Sub cmdLvlMapName_Click()
'
'    On Error GoTo cmdLvlMapName_Click_err
'
'    With cdl
'        .CancelError = True
'        .Filter = "Duke Nukem 3D Maps (*.map)|*.map|All Files (*.*)|*.*"
'        .Flags = cdlOFNFileMustExist + cdlOFNHideReadOnly
'        .ShowOpen
'    End With
'
'    With txtValue(8)
'        .Text = GetFilenameFromPath(cdl.Filename)
'        .SelStart = 0   ' Select all and focus
'        .SelLength = Len(.Text)
'        .SetFocus
'    End With
'
'    Exit Sub
'
'cmdLvlMapName_Click_err:
'    If err.Number <> 32755 Then
'        MsgBox "ERROR: " & err.Description
'    End If
'
'End Sub
'
'
'Private Sub cmdSave_Click()
'    m_bIsCancelled = False
'    Me.Hide
'End Sub
'
'Private Sub cmdSndFile_Click()
'
'    On Error GoTo cmdSndFile_Click_err
'
'    With cdl
'        .CancelError = True
'        .Filter = "Sound Files (*.voc)|*.voc|All Files (*.*)|*.*"
'        .Flags = cdlOFNFileMustExist + cdlOFNHideReadOnly
'        .ShowOpen
'    End With
'
'    With txtValue(11)
'        .Text = GetFilenameFromPath(cdl.Filename)
'        .SelStart = 0   ' Select all and focus
'        .SelLength = Len(.Text)
'        .SetFocus
'    End With
'
'    Exit Sub
'
'cmdSndFile_Click_err:
'    If err.Number <> 32755 Then
'        MsgBox "ERROR: " & err.Description
'    End If
'
'End Sub
'
'Private Sub cmdValue_Click()
'    Load FrmGetArtTile
'    FrmGetArtTile.Show vbModal
'    If Not FrmGetArtTile.Cancelled Then
'        With txtValue(1)
'            .Text = FrmGetArtTile.TileChosen
'            .SelStart = 0
'            .SelLength = Len(.Text)
'            .SetFocus
'        End With
'    End If
'End Sub
'
'Private Sub Form_Load()
'    Dim lr As Integer, lg As Integer, lb As Integer
'    ColorToRGB Point(0, 0), lr, lg, lb
'    Shape1.BackColor = RGB(lr - 20, lg - 20, lb - 20)
'
'    Set m_oObject = New CDefine
'    tabstrip_Click
'End Sub
'
'Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
'    If UnloadMode = vbFormCode Then Exit Sub
'    m_bIsCancelled = True
'    Me.Hide
'    Cancel = 1
'End Sub
'
'Private Sub tabstrip_Click()
'
'    picTabs(tabstrip.SelectedItem.Index - 1).ZOrder vbBringToFront
'
'    If m_bInitState = False And Me.Visible = True Then
'        Select Case tabstrip.SelectedItem.Index
'            Case 1
'                m_bInitState = True
'                m_oObject.SetCode ceCode.Text
'                InitValues
'                m_bInitState = False
'            Case 2
'                ceCode.Text = m_oObject.GetCode()
'        End Select
'    End If
'
'End Sub
'
'Private Sub txtLvl3DRealmsTime_Validate(Cancel As Boolean)
'    m_oObject.SetPropertyValue 4, txtLvl3DRealmsTime.Text
'End Sub
'
'Private Sub txtLvlParTime_Validate(Cancel As Boolean)
'    m_oObject.SetPropertyValue 3, txtLvlParTime.Text
'End Sub
'
'Private Sub txtValue_Validate(Index As Integer, Cancel As Boolean)
'
'    Dim sTag As String
'    Dim sValue As String
'
'    sTag = txtValue(Index).Tag
'    If GetTagAttribute(sTag, "numeric") = "yes" Then
'        If Not IsNumeric(txtValue(Index).Text) Then txtValue(Index).Text = "0"
'    End If
'
'    sValue = txtValue(Index).Text
'    m_oObject.SetPropertyValue CInt(GetTagAttribute(sTag, "property")), sValue
'
'End Sub
'
'Private Sub cboSkillSkillNo_Validate(Cancel As Boolean)
'    Dim iValue As Integer
'    iValue = cboSkillSkillNo.ItemData(cboSkillSkillNo.ListIndex)
'    If Not IsNumeric(iValue) Then iValue = 0
'    m_oObject.SetPropertyValue 0, iValue
'End Sub
'
