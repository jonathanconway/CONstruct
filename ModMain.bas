Attribute VB_Name = "ModMain"
' =============================================================================
' Module Name:      ModMain
' Module Type:      Code Module
' Description:      Main code module for CONstruct;
'                   ** Contains Sub Main() **
' Author(s):        Jonathan A. Conway
' -----------------------------------------------------------------------------
' Log:
'
' 04 07 05 :
'   - Created eCodeStructures enumeration; this will eventually replace the
'     now inefficient eObjectTypes enumeration.
'
' ?? ?? ?? :
'   - Created ModMain
' =============================================================================

Option Explicit



' Constants
' =========

#Const APPLICATION_BETA_VERSION = 0

Public Const DEFINITION_EDITOR_FILENAME = "PrimDef\PrimDef.exe"


' Enumerations
' ============


' Commands to nodes in a CONTree control
Public Enum eNodeCommands
    [ncGoto] = 0
    [ncAdd] = 1
    [ncEdit] = 2
    [ncRemove] = 3
End Enum



Public Settings As CApplicationSettings



' Public Methods
' ==============


' Sub Main
' --------

Public Sub Main()
    
    Dim sSettingsFile As String
    
    sSettingsFile = App.Path & "\" & "CONstruct.cfg"
    
    InitCommonControlsVB        ' Windows XP Themes support
    
    Set Settings = New CApplicationSettings
    
    If Trim$(Dir$(sSettingsFile)) = "" Then
        #If APPLICATION_BETA_VERSION Then
            BetaMessage             ' Show beta message if app is in beta
        #End If
        
        Settings.InitializeSettings sSettingsFile   ' Generate new settings file
        Settings.SaveFile sSettingsFile
    End If
    
    ' Load settings file
    Settings.LoadFile sSettingsFile
    
    ' Initialize list of snippets
    Settings.ReadSnippets
    
    Load FrmMain                                ' Load main form into memory
    SetIcon FrmMain.hWnd, "CONSTRUCT", True     ' Load application icon
        
    If Settings.ReadSetting("SplashScreenOnStartup") = "yes" Then _
        FrmSplash.Show
    
    
    ' Handle command-line parameters
    HandleParameters
    
    ' Loop until main form is closed
    Do While FrmMain.Visible = True
       DoEvents
    Loop

    ' Persist configuration to storage
    Settings.SaveFile sSettingsFile
    
    End                         ' End-point of program

End Sub



' File Management
' ---------------


Public Function GetFileList(ByVal Path As String) As String()

    ' Returns a string array representing a listing of all filenames
    ' found in the specified directory path. (Ignores folder names.)

    If IsFileExistant(, Path) Then
    
        Shell "CMD /cDIR """ & Path & """ /a:-d /b > """ & Path & "\LIST.TXT""", vbHide
        DoEvents
        Dim sText As String
        
        Do Until IsFileExistant(Path & "\LIST.TXT")
            'DoEvents
        Loop
        
        Dim lHnd As Long
        lHnd = FreeFile
        Open Path & "\LIST.TXT" For Input As #lHnd
        sText = Input(LOF(lHnd), #lHnd)     ' Read file contents
        Close #lHnd
        
        Kill Path & "\LIST.TXT"
        
        GetFileList = Split(sText, vbNewLine)
    
    End If
    
End Function



' Miscellaneous
' -------------

Public Sub SetCompatibleColours(ByRef Source As Form)

    If Settings.ReadSetting("General_CompatibleLook") = "yes" Then

        ' Set background colours of controls on the source form so that
        ' they conform to the current Windows visual style. For example, if
        ' in "Windows Classic Style", tab controls have a grey background,
        ' but in "Windows XP Style" they have a white background.
    
        Dim cCtl As Control
        For Each cCtl In Source.Controls
            If TypeOf cCtl Is PictureBox Then
                
                cCtl.BackColor = vb3DFace
            
            ElseIf TypeOf cCtl Is CheckBox _
                Or TypeOf cCtl Is Frame Then
                
                cCtl.Appearance = 1
                cCtl.BackColor = vb3DFace
            
            End If
        Next

    End If
    
End Sub

Public Function GetID(ByVal Key As String) As Long
    GetID = CLng(Right$(Key, Len(Key) - 1))
End Function


Public Function BooleanToChecked(ByVal Value As Boolean) As CheckBoxConstants

    Select Case Value
        Case True
            BooleanToChecked = vbChecked
        Case False
            BooleanToChecked = vbUnchecked
        Case Else
            BooleanToChecked = vbGrayed
    End Select

End Function

Public Sub ValidateNumericTextBox(ByRef Source As Control)
    If IsNumeric(Source.Text) Then
        Source.Text = CLng(Source.Text)
    Else
        Source.Text = "0"
    End If
End Sub

Public Function FixLong(ByVal Source As String) As Long
    If IsNumeric(Source) Then
        FixLong = CLng(Source)
    Else
        FixLong = 0
    End If
End Function

Public Function FixInteger(ByVal Source As String) As Integer
    If IsNumeric(Source) Then
        FixInteger = CInt(Source)
    Else
        FixInteger = 0
    End If
End Function

Public Sub ColorToRGB(ByVal Color As Long, ByRef Red As Integer, ByRef Green As Integer, ByRef Blue As Integer)
    ' Convert a standard color (long) to r-g-b components
    
    Dim c As Long
    
    c = Color
    Red = c Mod &H100
    c = c \ &H100
    Green = c Mod &H100
    c = c \ &H100
    Blue = c Mod &H100
End Sub

Public Function FeatureNotImplemented()

    ' Display msgbox saying that this feature hasn't been implemented

    MsgBox "This feature hasn't been implemented in this release." & vbNewLine & vbNewLine & _
            "Please visit http://jaconline.5u.com regularly to check for updates to the program." & vbNewLine & _
            "(Click File -> About... to see contact details)", vbExclamation

End Function

Private Sub BetaMessage()

    ' Display a msgbox saying this is just a beta of the program
    
    MsgBox "Welcome to " & App.ProductName & " v" & App.Major & "." & App.Minor & App.Revision & IIf(App.Major = 0, " Beta", "") & vbNewLine & vbNewLine & _
           "As this is a pre-release beta version of CONstruct, please expect many anomolies, many disabled/missing features and *many* bugs!! A complete version 1 is currently in the making. In the meantime, please enjoy trying out CONstruct's (limited) features and email me any ideas you have for additional features." & vbNewLine & vbNewLine & _
           "See the About box (File --> About...) for contact details." & vbNewLine & vbNewLine & _
           "Happy Duke-ing!", vbInformation

End Sub

Public Function GetFilenameFromPath(ByVal Path As String) As String
    ' Returns a filename from a pathname
    GetFilenameFromPath = Right$(Path, Len(Path) - InStrRev(Replace(Path, "/", "\"), "\"))
End Function


' Array Operations
' ----------------



Public Function IsArrayEmpty(ByRef Source()) As Boolean
    If IsArray(Source) Then
        IsArrayEmpty = pIsArrayEmpty(Source)
    Else
        IsArrayEmpty = True
    End If
End Function

Private Function pIsArrayEmpty(ByRef Source) As Boolean
    ' Test if an array is empty by concatenating it
    pIsArrayEmpty = (Len(Join$(Source, ",")) = 0)
End Function


Public Function IsInArray(ByRef Source, ByVal Value) As Long

    ' Searches an array for the specified item
    
    ' - If item is found, return array index number of the item
    ' - If item is not found, return value of '-1'
    ' - Doesn't work with multi-dimensional arrays

    IsInArray = -1
    Dim i As Long
    For i = LBound(Source) To UBound(Source)
        If Source(i) = Value Then
            IsInArray = i
            Exit Function
        End If
    Next

End Function

Public Sub AddArrayItem(ByRef Source(), ByVal Value As Variant)

    If IsArrayEmpty(Source) Then
        ReDim Source(0)
    Else
        ReDim Preserve Source(LBound(Source) To UBound(Source) + 1)
    End If

    Source(UBound(Source)) = Value

End Sub


' Private Methods
' ===============


Private Sub HandleParameters()

    ' Load the filename specified on the command-line
    ' Only works with a single filename
    ' TODO : Make it work with multiple filenames & long filenames

    Dim sCmdLine As String
    
    sCmdLine = Command$()
    
    If Len(Trim$(sCmdLine)) > 0 Then
        sCmdLine = ExtractQuotes(sCmdLine)
        
        If IsFileExistant(sCmdLine) Then
            FrmMain.LoadCON sCmdLine
        End If
    End If

End Sub


Private Function ExtractQuotes(ByVal Source As String) As String

    Dim sReturn As String
    
    sReturn = Source
    
    If Left$(sReturn, 1) = Chr(34) Then
        If Right$(sReturn, 1) = Chr(34) Then
            sReturn = Mid$(sReturn, 2, Len(sReturn) - 2)
        End If
    End If

    ExtractQuotes = sReturn

End Function



