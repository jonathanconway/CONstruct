Attribute VB_Name = "ModSettings"
'Option Explicit
'
'
'' Types
'' =====
'
'' Application setting (stored as collection in .CFG file)
'Public Type tSetting
'    st_Name As String
'    st_Value As String
'End Type
'
'' Application settings file (stored in a .CFG file)
'Public Type tSettingsFile
'    sf_Settings() As tSetting
'End Type
'
'
'
'' Private Variables
'' =================
'
'
'Private m_oSettingsFile As tSettingsFile
'
'
'
'
'
'Public Sub GenerateSettings()
'
'    ReDim m_oSettingsFile.sf_Settings(0 To 9)
'
'    m_oSettingsFile.sf_Settings(0).st_Name = "MainWindow_X"
'    m_oSettingsFile.sf_Settings(1).st_Name = "MainWindow_Y"
'    m_oSettingsFile.sf_Settings(2).st_Name = "MainWindow_Width"
'    m_oSettingsFile.sf_Settings(3).st_Name = "MainWindow_Height"
'    m_oSettingsFile.sf_Settings(4).st_Name = "MainWindow_State"
'
'    m_oSettingsFile.sf_Settings(5).st_Name = "Path_Duke3D"
'    m_oSettingsFile.sf_Settings(6).st_Name = "Path_Cons"
'    m_oSettingsFile.sf_Settings(7).st_Name = "Path_Maps"
'    m_oSettingsFile.sf_Settings(8).st_Name = "Path_Sounds"
'    m_oSettingsFile.sf_Settings(9).st_Name = "Path_Art"
'
'    SaveSettings
'
'End Sub
'
'
'' Settings Stuff
'' --------------
'
'Public Sub LoadSettings()
'
'    ' Load contents of a CFG file into memory
'    Dim lHnd As Long
'    lHnd = FreeFile()
'    Open App.Path & "\CONstruct.cfg" For Binary As #lHnd
'    Get #lHnd, , m_oSettingsFile
'    Close #lHnd
'
'End Sub
'
'Public Sub SaveSettings()
'
'    ' Save contents of SettingsFile object to CFG file
'    Dim lHnd As Long
'    lHnd = FreeFile()
'    Open App.Path & "\CONstruct.cfg" For Binary As #lHnd
'    Put #lHnd, , m_oSettingsFile
'    Close #lHnd
'
'End Sub
'
'Public Function ReadSetting(ByVal SettingName As String) As String
'
'    On Error GoTo ProcedureError
'
'    ' Retrieve a specific application setting
'    Dim i As Integer
'    For i = LBound(m_oSettingsFile.sf_Settings) To UBound(m_oSettingsFile.sf_Settings)
'        If UCase(m_oSettingsFile.sf_Settings(i).st_Name) = UCase(SettingName) Then
'            Settings.ReadSetting = m_oSettingsFile.sf_Settings(i).st_Value
'        End If
'    Next
'
'    Exit Function
'
'ProcedureError:
'    If err.Number = 9 Then
'    Else
'        MsgBox err.Number & vbNewLine & err.Description
'    End If
'End Function
'
'Public Sub WriteSetting(ByVal SettingName As String, ByVal NewValue As String)
'
'    ' Save a specific application setting
'
'    Dim i As Integer
'    Dim iIndex As Integer
'
'    iIndex = -1
'
'    For i = LBound(m_oSettingsFile.sf_Settings) To UBound(m_oSettingsFile.sf_Settings)
'        If UCase(m_oSettingsFile.sf_Settings(i).st_Name) = UCase(SettingName) Then
'            iIndex = i
'            Exit For
'        End If
'    Next
'
'    If iIndex = -1 Then
'        ReDim Preserve m_oSettingsFile.sf_Settings(LBound(m_oSettingsFile.sf_Settings) To UBound(m_oSettingsFile.sf_Settings) + 1)
'        m_oSettingsFile.sf_Settings(UBound(m_oSettingsFile.sf_Settings)).st_Name = SettingName
'        iIndex = UBound(m_oSettingsFile.sf_Settings)
'    End If
'
'    m_oSettingsFile.sf_Settings(iIndex).st_Value = NewValue
'
'End Sub
'
'
