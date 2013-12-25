VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CApplicationSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit



' Types
' =====

Private Type tSetting           ' Single setting contained in file
    st_Name As String
    st_Value As String
End Type

Private Type tBookmarkList      ' A list of bookmarks for a file
    bl_Filename As String
    bl_Bookmarks As String
End Type


Private Type tSettingsFile      ' Application settings file
    sf_Settings() As tSetting
    sf_BookmarkLists() As tBookmarkList
End Type



' Private Variables
' =================

Private m_oSettingsFile As tSettingsFile
Private m_oSnippets As CObjectCollection
Private m_sFilename As String


' Public Methods
' ==============


Public Sub InitializeSettings(ByVal Filename As String)

    ReDim m_oSettingsFile.sf_Settings(0 To 10)
    
    ' Initialize important application settings/values
    m_oSettingsFile.sf_Settings(0).st_Name = "MainWindow_X"
    m_oSettingsFile.sf_Settings(1).st_Name = "MainWindow_Y"
    m_oSettingsFile.sf_Settings(2).st_Name = "MainWindow_Width"
    m_oSettingsFile.sf_Settings(3).st_Name = "MainWindow_Height"
    m_oSettingsFile.sf_Settings(4).st_Name = "MainWindow_State"
    m_oSettingsFile.sf_Settings(5).st_Name = "Path_Duke3D"
    m_oSettingsFile.sf_Settings(6).st_Name = "Path_Cons"
    m_oSettingsFile.sf_Settings(7).st_Name = "Path_Maps"
    m_oSettingsFile.sf_Settings(8).st_Name = "Path_Sounds"
    m_oSettingsFile.sf_Settings(9).st_Name = "Path_Art"
    
    m_oSettingsFile.sf_Settings(10).st_Name = "Recent_Count"
    m_oSettingsFile.sf_Settings(10).st_Value = "4"

    ReDim m_oSettingsFile.sf_BookmarkLists(0)

    SaveFile Filename

End Sub

Public Function ReadSetting(ByVal SettingName As String) As String

    On Error GoTo ProcedureError
    
    ' Retrieve a specific application setting
    Dim i As Integer
    For i = LBound(m_oSettingsFile.sf_Settings) To UBound(m_oSettingsFile.sf_Settings)
        If UCase(m_oSettingsFile.sf_Settings(i).st_Name) = UCase(SettingName) Then
            ReadSetting = m_oSettingsFile.sf_Settings(i).st_Value
            Exit Function
        End If
    Next

    Exit Function
    
ProcedureError:
    If err.Number = 9 Then
    Else
        MsgBox err.Number & vbNewLine & err.Description
    End If

End Function


Public Sub WriteSetting(ByVal SettingName As String, ByVal Value As String)

    On Error GoTo ProcedureError
    
    ' Save a specific application setting
    
    Dim i As Integer
    Dim iIndex As Integer
    
    iIndex = -1
    
    For i = LBound(m_oSettingsFile.sf_Settings) To UBound(m_oSettingsFile.sf_Settings)
        If UCase(m_oSettingsFile.sf_Settings(i).st_Name) = UCase(SettingName) Then
            iIndex = i
            Exit For
        End If
    Next
    
    If iIndex = -1 Then
        ReDim Preserve m_oSettingsFile.sf_Settings(LBound(m_oSettingsFile.sf_Settings) To UBound(m_oSettingsFile.sf_Settings) + 1)
        m_oSettingsFile.sf_Settings(UBound(m_oSettingsFile.sf_Settings)).st_Name = SettingName
        iIndex = UBound(m_oSettingsFile.sf_Settings)
    End If

    m_oSettingsFile.sf_Settings(iIndex).st_Value = Value

    Exit Sub
    
ProcedureError:
    If err.Number = 9 Then
        ReDim m_oSettingsFile.sf_Settings(0)
        WriteSetting SettingName, Value
        Exit Sub
    Else
        MsgBox err.Number & vbNewLine & err.Description
    End If

End Sub


Public Sub LoadFile(ByVal Filename As String)

    ' Load contents of a CFG file into memory
    Dim lHnd As Long
    lHnd = FreeFile()
    Open Filename For Binary As #lHnd
    Get #lHnd, , m_oSettingsFile
    Close #lHnd
    
    m_sFilename = Filename

End Sub


Public Sub SaveFile(ByVal Filename As String)

    ' Save contents of SettingsFile object to CFG file
    Dim lHnd As Long
    lHnd = FreeFile()
    Open Filename For Binary As #lHnd
    Put #lHnd, , m_oSettingsFile
    Close #lHnd

End Sub





Public Function ReadBookmarks(ByVal Filename As String) As String
    
    Dim i As Integer
    
    For i = LBound(m_oSettingsFile.sf_BookmarkLists) To UBound(m_oSettingsFile.sf_BookmarkLists)
        With m_oSettingsFile.sf_BookmarkLists(i)
            If UCase(.bl_Filename) = UCase(Filename) Then
                ReadBookmarks = .bl_Bookmarks
            End If
        End With
    Next

End Function


Public Sub WriteBookmarks(ByVal Filename As String, ByVal Values As String)

    Dim i As Integer
    Dim iIndex As Integer
    
    iIndex = -1
    
    ' Search for an occurance of the specified filename in the bookmarklist
    For i = LBound(m_oSettingsFile.sf_BookmarkLists) To UBound(m_oSettingsFile.sf_BookmarkLists)
        With m_oSettingsFile.sf_BookmarkLists(i)
            If UCase(.bl_Filename) = UCase(Filename) Then
                iIndex = i
                Exit For
            End If
        End With
    Next

    ' If necessary, add a new bookmarklist entry for the specified filename
    With m_oSettingsFile
        If iIndex = -1 Then
            ReDim Preserve .sf_BookmarkLists(LBound(.sf_BookmarkLists) To UBound(.sf_BookmarkLists) + 1)
            iIndex = UBound(.sf_BookmarkLists)
            .sf_BookmarkLists(iIndex).bl_Filename = Filename
        End If
    End With
        
    ' Set values string
    With m_oSettingsFile.sf_BookmarkLists(iIndex)
        .bl_Bookmarks = Values
    End With

End Sub


Public Sub ChangeBookmarks(ByVal OldFilename As String, ByVal NewFilename As String)

    Dim i As Integer
    Dim iIndex As Integer
    
    iIndex = -1
    
    ' Search for an occurance of the specified filename in the bookmarklist
    For i = LBound(m_oSettingsFile.sf_BookmarkLists) To UBound(m_oSettingsFile.sf_BookmarkLists)
        With m_oSettingsFile.sf_BookmarkLists(i)
            If UCase(.bl_Filename) = UCase(OldFilename) Then
                iIndex = i
                Exit For
            End If
        End With
    Next

    ' If found, change the filename setting over to the new filename
    If iIndex > -1 Then
        m_oSettingsFile.sf_BookmarkLists(iIndex).bl_Filename = NewFilename
    End If

End Sub


Public Sub WriteSnippets()
    
    Dim sSetting As String
    
    Dim oSnippet As CSnippet
    
    For Each oSnippet In m_oSnippets
        sSetting = sSetting & oSnippet.Filename & ","
    Next

    If Right$(Trim$(sSetting), 1) = "," Then
        sSetting = Mid$(sSetting, 1, Len(sSetting) - 1)
    End If
    
    Settings.WriteSetting "Snippets_List", sSetting

End Sub


Public Sub ReadSnippets()

    Dim sFiles() As String
    Dim sSetting As String
    Dim oSnippet As CSnippet
    Dim i As Integer
    
    m_oSnippets.Clear
    
    sSetting = Settings.ReadSetting("Snippets_List")
    
    If Len(Trim$(sSetting)) > 0 Then
    
        sFiles = Split(sSetting, ",")
        For i = LBound(sFiles) To UBound(sFiles)
            If IsFileExistant(sFiles(i)) Then
                ' Create new snippet object
                Set oSnippet = New CSnippet
                oSnippet.LoadFile sFiles(i)
                
                ' Add to collection
                m_oSnippets.Add oSnippet
            End If
        Next
    
    End If

End Sub



' Public Properties
' =================


Public Property Get Snippets() As CObjectCollection
    Set Snippets = m_oSnippets
End Property


Public Property Get Filename() As String
    Filename = m_sFilename
End Property

Public Property Let Filename(ByVal NewValue As String)
    m_sFilename = NewValue
End Property

Private Sub Class_Initialize()
    Set m_oSnippets = New CObjectCollection
End Sub


