Attribute VB_Name = "ModArt"
' =============================================================================
' Module Name:      ModArt
' Module Type:      Code Module
' Description:      Functionality for loading/manipulating Duke Nukem ART
'                   files. Code was taken from BastART by Marijn Kentie.
' Author(s):        Marijn "Patchboy" Kentie (modified by Jonathan A. Conway)
' -----------------------------------------------------------------------------
' Log:
'                   04 05 07 : Created ModArt using modified code from BastART
' =============================================================================





Type RGBColor
R As Byte
G As Byte
b As Byte
Reserved As Byte
End Type


Public PalColor(255) As RGBColor





' Types
' =====

Type PicType
    Pixels() As Byte
End Type

Type PropType
    AnimType As Byte
    OffsetX As Byte
    OffsetY As Byte
    AnimSpeed As Byte
End Type


Type TileType
    XSize As Integer
    YSize As Integer
    Properties As PropType
    PicData As PicType
    Changed As Boolean
End Type


Type tpArtFile
    Filename As String
    Version As Long
    NumTiles As Long
    LocalStart As Long
    LocalEnd As Long
    ArtNum As Byte
    Tiles() As TileType
End Type



' Public Variables
' ================

Public ArtFile As tpArtFile


' Public Methods
' ==============

Public Sub OpenArt(File As String)
    'On Error GoTo err
    Dim strTemp As String
    Dim lnTemp As Long
    Dim tByte As Byte
    Dim i As Integer
    Dim o As Integer
    Dim p As Integer
    
    
    Dim HeaderEnd As Long
    Dim XSizeEnd As Long
    Dim YSizeEnd As Long
    Dim PropsEnd As Long
    Dim lImageData As Long
    
    'StopAnim
    
    'Open an ART file.
    Open File For Binary As #1
    
    
    'Skip all of this if only a single tile needs to be loaded.
    With ArtFile
    
    
    CurrTile = 0
    OffsetX = 0
    OffsetY = 0
    DoEvents
    
    'The art file's number is the last three characters of the 8. in the filename.
    .ArtNum = Val(Right(Left(Right(File, Len(File) - i), 8), 3))
    
    o = -1
    
    'Check where the last slash is in the filename, and store that pos in I.
    Do Until o = 0
    If o = -1 Then o = 0
    i = o
    o = InStr(o + 1, File, "\")
    
    Loop
    
    
    
    'ART Header processing.
    Get (1), , .Version 'Get a long containing the version number.
    Get (1), , lnTemp 'Skip the (outdated) file NumTiles value.
    Get (1), , .LocalStart 'Get the start tile.
    Get (1), , .LocalEnd 'End tile.
    .NumTiles = .LocalEnd - .LocalStart + 1
    
    ReDim .Tiles(.NumTiles)
    
    
    'Calculate the data offsets.
    HeaderEnd = LenB(.Version) + LenB(lnTemp) + LenB(.LocalStart) + LenB(.LocalEnd) + 1
    XSizeEnd = HeaderEnd + 2 * .NumTiles '2 byte chunks.
    YSizeEnd = XSizeEnd + 2 * .NumTiles '2 byte chunks.
    PropsEnd = YSizeEnd + 4 * .NumTiles '4 byte chunks.
    
    End With
    'Read the data for each tile at a calculated distance from the header.
    
    
    For i = 1 To ArtFile.NumTiles
    
    With ArtFile.Tiles(i)
    
    Get #1, HeaderEnd + (i - 1) * 2, .XSize
    Get #1, XSizeEnd + (i - 1) * 2, .YSize
    Get #1, YSizeEnd + (i - 1) * 4, .Properties.AnimType
    Get #1, , .Properties.OffsetX
    Get #1, , .Properties.OffsetY
    'MsgBox .Properties.OffsetX & " " & .Properties.OffsetY
    Get #1, , .Properties.AnimSpeed
    
    
    ReDim .PicData.Pixels(.XSize, .YSize)
    
    For o = 1 To .XSize
    For p = 1 To .YSize
    Get #1, PropsEnd + lImageData, .PicData.Pixels(o, p)
    lImageData = lImageData + 1
    Next p
    Next o
    
    End With
    
    
    Next i
    
    
    Close #1
    
    ArtFile.Filename = File
    
    
    'Form2.VScroll1.Value = 0
    'Form1.Caption = "bastART - " & File
    'Form1.mnuSave.Enabled = True
    'Form1.mnuTile.Enabled = True
    'Form1.mnuSaveAs.Enabled = True
    
    CurrTile = 1
    'CenterTile
    'FillBrowser
    'RenderTile (CurrTile)
    Exit Sub
err:
    If FreeFile > 1 Then Close #1
    MsgBox "Error opening ART file.", vbOKOnly + vbExclamation, "bastART"
    'NewArt
End Sub

'Public Sub SaveArt(File As String, Optional JustSave As Boolean)
'    'Write an ART file.
'
'    If JustSave = True Then
'    'If not saving as, but just saving, always overwrite the old file.
'    Kill File
'    Else
'    If Overwrite(File) = False Then Exit Sub
'    End If
'
'    Form1.Label1.Caption = "Writing " & File & "..."
'    DoEvents
'
'    Dim i As Integer
'    Dim o As Integer
'    Dim p As Integer
'
'    Dim HeaderEnd As Long
'    Dim XSizeEnd As Long
'    Dim YSizeEnd As Long
'    Dim PropsEnd As Long
'    Dim lImageData As Long
'    With ArtFile
'    'Calculate the data offsets.
'    HeaderEnd = LenB(.Version) + 4 + LenB(.LocalStart) + LenB(.LocalEnd) + 1
'    XSizeEnd = HeaderEnd + 2 * .NumTiles '2 byte chunks.
'    YSizeEnd = XSizeEnd + 2 * .NumTiles '2 byte chunks.
'    PropsEnd = YSizeEnd + 4 * .NumTiles '4 byte chunks.
'
'    'Write the header.
'    Open File For Binary As #1
'    Put #1, , .Version
'    Put #1, , .NumTiles
'    Put #1, , .LocalStart
'    Put #1, , .LocalEnd
'    End With
'    For i = 1 To ArtFile.NumTiles
'
'    'Write the data at the calculated offsets.
'    With ArtFile.Tiles(i)
'
'    Put #1, HeaderEnd + (i - 1) * 2, .XSize
'    Put #1, XSizeEnd + (i - 1) * 2, .YSize
'    Put #1, YSizeEnd + (i - 1) * 4, .Properties.AnimType
'
'
'    Put #1, YSizeEnd + (i - 1) * 4 + 1, .Properties.OffsetX
'
'
'    Put #1, YSizeEnd + (i - 1) * 4 + 2, .Properties.OffsetY
'
'    Put #1, YSizeEnd + (i - 1) * 4 + 3, .Properties.AnimSpeed
'
'    For o = 1 To .XSize
'    For p = 1 To .YSize
'    Put #1, PropsEnd + lImageData, .PicData.Pixels(o, p)
'    lImageData = lImageData + 1
'    Next p
'    Next o
'
'    End With
'
'    Next i
'    Close #1
'    Form1.Label1.Caption = "Done."
'
'    'So the name updates when saving as.
'    Form1.mnuSave.Enabled = True
'    ArtFile.Filename = File
'    Form1.Caption = "bastART - " & File
'End Sub

'Public Sub NewArt()
'    Form1.mnuSave.Enabled = False
'    Form1.mnuSaveAs.Enabled = True
'    Form1.Picture1.Cls
'    ArtFile.LocalStart = 0
'    ArtFile.LocalEnd = 255
'    ArtFile.NumTiles = 256
'    ArtFile.Version = 1
'    ReDim ArtFile.Tiles(256)
'    Form1.mnuTile.Enabled = True
'    CurrTile = 1
'    FillBrowser
'    RenderTile (CurrTile)
'End Sub







' ****************




Public Sub LoadPalette(File As String)
'On Error GoTo err
'Read the palette file to find out the RGP values of the 256 color-values the ART files use.

Dim i As Integer
Dim Y As Byte
Dim X As Byte
Dim o As Integer
Dim p As Integer
Dim eb As Byte 'Placeholder empty byte
Y = 0



Open File For Binary As #1

'Skip a .pal file's header.
If UCase(Right(File, 3)) = "PAL" Then
For i = 1 To 24
Get #1, , eb
Next i
End If

For i = 0 To 255

Get #1, , PalColor(i).R
Get #1, , PalColor(i).G
Get #1, , PalColor(i).b

If UCase(Right(File, 3)) = "PAL" Then
Get #1, , eb 'If a Win palette, read an empty byte.
Else
'Multiply each color by 4 to range from 0-255 instead of 0-63 (only Build pals).

PalColor(i).R = PalColor(i).R * 4
PalColor(i).G = PalColor(i).G * 4
PalColor(i).b = PalColor(i).b * 4
End If



'Preview the palette in a picturebox.
If i = 128 Then Y = 1: X = 0 'Start on a new line at half the colors
For o = 0 To 3
For p = 0 To 3
'Form1.Picture2.PSet ((X * 4 + o) * Screen.TwipsPerPixelX, (Y * 4 + p) * Screen.TwipsPerPixelY), RGB(PalColor(i).R, PalColor(i).G, PalColor(i).B)
Next p
Next o
X = X + 1

Next i

Close #1

'If CurrTile > 0 Then FillBrowser: RenderTile (CurrTile)
'Form1.mnuOpen.Enabled = True
'Form1.mnuexppal.Enabled = True

'Form1.mnuNew.Enabled = True

'Form1.Label2.BackColor = RGB(PalColor(0).R, PalColor(0).G, PalColor(0).B)
'Form1.Label3.BackColor = RGB(PalColor(255).R, PalColor(255).G, PalColor(255).B)



Exit Sub
err:
If FreeFile > 1 Then Close #1
MsgBox "Error opening palette.", vbOKOnly + vbExclamation, "bastART"
End Sub

'Public Sub SavePal(File As String)
'
'
'Dim Resp As Integer
'
'If Overwrite(File) = False Then Exit Sub
'
'Open File For Binary As #1
'Dim b As Byte
'Dim zb As Byte
'Dim i As Integer
'zb = 0
'If Right(File, 3) = "PAL" Then
''Export the palette data to a standard Windows palette.
'
''Write PAL header.
'Put #1, , "RIFF"
'b = 14
'Put #1, , b
'b = 4
'Put #1, , b
'Put #1, , zb
'Put #1, , zb
'Put #1, , "PAL data"
'b = 8
'Put #1, , b
'b = 4
'Put #1, , b
'Put #1, , zb
'Put #1, , zb
'Put #1, , zb
'b = 3
'Put #1, , b
'Put #1, , zb
'b = 1
'Put #1, , b
'
''PAL data.
'
'For i = 0 To 255
'Put #1, , PalColor(i).R
'Put #1, , PalColor(i).G
'Put #1, , PalColor(i).b
'Put #1, , zb
'Next i
'
''Write PAL footer.
'For i = 1 To 4
'Put #1, , zb
'Next i
'
'Else
''Save a standard Build .PAL file.
'For i = 0 To 255
'b = PalColor(i).R / 4
'Put #1, , b
'b = PalColor(i).G / 4
'Put #1, , b
'b = PalColor(i).b / 4
'Put #1, , b
'Next i
'End If
'
'
'Close #1
'
'
'
'End Sub

