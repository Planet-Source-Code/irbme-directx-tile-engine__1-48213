Attribute VB_Name = "modEngine"
Option Explicit

#Const DebugMode = False

Private Declare Function GetTickCount Lib "kernel32" () As Long
Private Declare Function ShowCursor Lib "user32" (ByVal bShow As Long) As Long
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)


#If DebugMode Then
    Public Const Walk_Speed As Long = 2
#Else
    Public Const Walk_Speed As Long = 1
#End If

Public Const SpriteWidth As Long = 32
Public Const SpriteHeight As Long = 64


Public Type udtPoint
    X As Long
    Y As Long
End Type


Public Type udtSurface
    DXSurface           As DirectDrawSurface7   'The direct draw surface
    Width               As Long                 'The width
    Height              As Long                 'The height
    TransparentColor    As Long                 'The transparent color (If Transparent)
    Transparent         As Boolean              'Whether or not it is transparent
End Type


Public Type udtSingleTile
    GraphicIndex        As Integer              'The tile index from the tileset.
    Walkable            As Boolean              'Whether or not this tile is walkable

    HasPortal           As Boolean
    PortalX             As Long
    PortalY             As Long
    PortalMapName       As String
End Type

'Example tile set indices:
'   _ _ _ _ _ _ _ _ _ _ _ _ _
'  | 1  2  3  4  5  6  7  8  |
'  | 9  10 11 12 13 14 15 16 |
'  | 17 18 19 20 21 22 23 24 |
'  | 25 26 27 28 29 30 31 32 |
'   ¯ ¯ ¯ ¯ ¯ ¯ ¯ ¯ ¯ ¯ ¯ ¯ ¯

Public Type udtTileSet
    TilesX              As Long                 'Number of tiles along the way
    TilesY              As Long                 'Number of tiles down the way
    TileWidth           As Long                 'Tile width
    TileHeight          As Long                 'Tile height
    Surface             As udtSurface           'The suface descriptor
End Type


Public Type udtMap
    TilesX              As Long                 'Number of tiles along the way
    TilesY              As Long                 'Number of tiles down the way
    StartX              As Long                 'The starting X position (in tile numbers)
    StartY              As Long                 'The starting Y position (in tile numbers)
    OffsetX             As Long                 'The current offset along the way (see note)
    OffsetY             As Long                 'The current offset down the way (see note)
    TileSet             As udtTileSet           'The tile set information
    Tiles()             As udtSingleTile        'The actual tile array
End Type

'Note about OffsetX and OffsetY:

'The whole map will not fit on the screen. For this reason, it can be shifted along the X or Y axis
'OffsetX and OffsetY store how much they are shifted.
'If both are 0, the map is drawn starting from the top left tile.
'As the map shifts, OffsetX and OffsetY grow negative

'eg. Map Width = 10 tiles. Screen holds 7
'    Map Height = 10 tiles. Screen holds 6

'(-3,-4)
'    _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _
'   |                                   |
'   |     MAP                           | OffsetY = -4
'   |                                   |
'   |      (0,0) _ _ _ _ _ _ _ _ _ _ _ _|
'   |           |                       |
'   |           |                       |
'   |           |                       |
'   |           |                       | Visible Tiles Y = 6
'   |           |        SCREEN         |
'   |           |                       |
'   |           |                       |
'   |_ _ _ _ _ _|_ _ _ _ _ _ _ _ _ _ _ _|
'                                     (7,6)
'
'   OffsetX = -3   Visible Tiles X = 7

'(Offsets are not in tiles but are in pixels:
'  OffsetX = NumberTilesX * TileWidth
'  OffsetY = NumberTilesY * TileHeight

Public Type udtPlayer
    AnimationOffsetX As Long        'The X position of the animation
    AnimationOffsetY As Long        'The Y position of the animation
    PositionX        As Long        'The X position on the screen
    PositionY        As Long        'The Y position on the screen
    GoingUp          As Boolean     'Are they going up?
    GoingDown        As Boolean     'Are they going down?
    GoingLeft        As Boolean     'Are they going left?
    GoingRight       As Boolean     'Are they going right?
    DXSurface        As udtSurface  'The surface descriptor
End Type



Private DirectX              As DirectX7                'Main DirextX object
Private DirectDraw           As DirectDraw7             'Main DirectDraw object

Private PrimarySurface       As DirectDrawSurface7      'Screen surface
Private PrimaryDescriptor    As DDSURFACEDESC2          'Screen descriptor
Private Backbuffer           As DirectDrawSurface7      'Backbuffer surface
Private BackBufferDescriptor As DDSURFACEDESC2          'Backbuffer descriptor

Private Handle               As Long                    'Parent Handle (Form.hWnd)

Public ScreenWidth           As Long                    'Screen width in pixels
Public ScreenHeight          As Long                    'Screen height in pixels

Private LastTickCount        As Long                    'The last recorded tick count
Private CurrentFrameCount    As Long                    'Current frame counter
Private LastFrameCount       As Long                    'Last frame counter

Private GammaController      As DirectDrawGammaControl  'Gamma controller
Private GammaRamp            As DDGAMMARAMP             'Current gamma ramp
Private OriginalRamp         As DDGAMMARAMP             'Original gamma ramp

Private GammaRedVal          As Integer                 'Amount of red in gamma
Private GammaGreenVal        As Integer                 'Amount of green in gamma
Private GammaBlueVal         As Integer                 'Amount of blue in gamma

Public GammaIsSupported      As Boolean                 'Is gamma supported

Public Map                   As udtMap                  'The map
Public Player                As udtPlayer               'The players Player

Private LastPortalEntered    As udtPoint
Private AllowPortalTravel    As Boolean


Public Function InitializeDirectX(hWnd As Long, Width As Long, Height As Long, BitDepth As Long) As Boolean
  On Error GoTo ErrHandler
  
  Dim DisplayModesEnum  As DirectDrawEnumModes
  Dim SurfaceDescriptor As DDSURFACEDESC2
  Dim Caps              As DDSCAPS2
  Dim hwCaps            As DDCAPS
  Dim helCaps           As DDCAPS
  Dim nCount            As Long
  Dim i                 As Long
    
    'We are already initialized
    If Handle <> 0 Then Exit Function
    
    'Create a new instance of DirectX and DirectDraw
    Set DirectX = New DirectX7
    Set DirectDraw = DirectX.DirectDrawCreate(vbNullString)
    
    'Full screen exclusive mode
    DirectDraw.SetCooperativeLevel hWnd, DDSCL_FULLSCREEN Or DDSCL_ALLOWMODEX Or DDSCL_EXCLUSIVE
    
    'Create new display mode enumeration object
    Set DisplayModesEnum = DirectDraw.GetDisplayModesEnum(0, PrimaryDescriptor)
    nCount = DisplayModesEnum.GetCount
    
    'Redimension the array of display modes to the display mode count
    ReDim DisplayModes(1 To nCount) As DDSURFACEDESC2
    
    'Loop through each element of the array and add the display mode to it.
    For i = 1 To nCount
        DisplayModesEnum.GetItem i, DisplayModes(i)
    Next
    
    'If the requested display mode is not supported return false
    If Not DisplayModeIsSupported(Width, Height, BitDepth) Then
        InitializeDirectX = False
        Exit Function
    End If
    
    'The display mode is supported so set it
    DirectDraw.SetDisplayMode Width, Height, BitDepth, 0, DDSDM_DEFAULT
    
    'Fill out the primary surface descriptor
    PrimaryDescriptor.lFlags = DDSD_CAPS Or DDSD_BACKBUFFERCOUNT
    PrimaryDescriptor.ddsCaps.lCaps = DDSCAPS_PRIMARYSURFACE Or DDSCAPS_FLIP Or DDSCAPS_COMPLEX
    PrimaryDescriptor.lBackBufferCount = 1
    
    'Create the primary surface and get it's descriptor
    Set PrimarySurface = DirectDraw.CreateSurface(PrimaryDescriptor)
    PrimarySurface.GetSurfaceDesc PrimaryDescriptor

    Caps.lCaps = DDSCAPS_BACKBUFFER
    
    'Create a back buffer surface and get it's descriptor
    Set Backbuffer = PrimarySurface.GetAttachedSurface(Caps)
    Backbuffer.GetSurfaceDesc BackBufferDescriptor
    
    'The background of text is transparent
    Backbuffer.SetFontTransparency True
    
    'Is gamma supported
    DirectDraw.GetCaps hwCaps, helCaps
    GammaIsSupported = hwCaps.lCaps2 And DDCAPS2_PRIMARYGAMMA <> 0

    'Get a gamma controller
    Set GammaController = PrimarySurface.GetDirectDrawGammaControl
    
    'Get the original gamma ramp
    GammaController.GetGammaRamp DDSGR_DEFAULT, OriginalRamp

    GammaRedVal = 0
    GammaGreenVal = 0
    GammaBlueVal = 0

    'Hide the cursor
    While ShowCursor(0) >= 0: Wend
    
    'Store some variables
    ScreenWidth = Width
    ScreenHeight = Height
    Handle = hWnd
    
    LastPortalEntered.X = -1
    LastPortalEntered.Y = -1
    AllowPortalTravel = True
    
    'Return true
    InitializeDirectX = True
    
Exit Function
ErrHandler:
End Function



Public Function LoadMap(FileName As String) As Boolean
  
  On Error GoTo ErrHandler
  
  Dim FileNum   As Integer
  Dim TempByte  As Byte
  Dim TempInt   As Integer
  Dim X         As Long
  Dim Y         As Long

  Dim NumberOfPortals As Long
  Dim PortalX As Long
  Dim PortalY As Long
  Dim PortalDestX As Long
  Dim PortalDestY As Long
  Dim PortalMapName As String * 16
  
    FileNum = FreeFile

    Open FileName For Binary Access Read Lock Write As #FileNum
        
        'Get map width and height
        Get #FileNum, , TempInt
        Map.TilesX = CLng(TempInt)
        Get #FileNum, , TempInt
        Map.TilesY = CLng(TempInt)
        
        'Get the starting position
        Get #FileNum, , TempInt
        Map.StartX = CLng(TempInt)
        Get #FileNum, , TempInt
        Map.StartY = CLng(TempInt)
        
        'Get the tileset information
        Get #FileNum, , TempByte
        Map.TileSet.TilesX = CLng(TempByte)
        Get #FileNum, , TempByte
        Map.TileSet.TilesY = CLng(TempByte)
        Get #FileNum, , TempByte
        Map.TileSet.TileWidth = CLng(TempByte)
        Get #FileNum, , TempByte
        Map.TileSet.TileHeight = CLng(TempByte)
       
        ReDim Map.Tiles(1 To Map.TilesX, 1 To Map.TilesY)
        
        'Read the tile information
        For X = 1 To Map.TilesX
            For Y = 1 To Map.TilesY
                Get #FileNum, , TempByte
                Map.Tiles(X, Y).GraphicIndex = CInt(TempByte)
                Get #FileNum, , TempByte
                Map.Tiles(X, Y).Walkable = IIf(TempByte = 1, True, False)
            Next
        Next
        
        'Portal information
        Get #FileNum, , TempInt
        NumberOfPortals = CLng(TempInt)
        
        For X = 1 To NumberOfPortals
            Get #FileNum, , TempInt
            PortalX = CLng(TempInt)

            Get #FileNum, , TempInt
            PortalY = CLng(TempInt)

            Get #FileNum, , TempInt
            PortalDestX = CLng(TempInt)

            Get #FileNum, , TempInt
            PortalDestY = CLng(TempInt)

            Get #FileNum, , PortalMapName
            PortalMapName = Replace$(PortalMapName, Chr$(0), vbNullString)

            Map.Tiles(PortalX, PortalY).HasPortal = True
            Map.Tiles(PortalX, PortalY).PortalX = PortalDestX
            Map.Tiles(PortalX, PortalY).PortalY = PortalDestY
            Map.Tiles(PortalX, PortalY).PortalMapName = PortalMapName
        Next
        
    Close #FileNum
    
    PutPlayerOnTile Map.StartX, Map.StartY
        
    LoadMap = True

Exit Function
ErrHandler:
End Function




Public Function LoadSurface(BitmapFile As String, Optional Width As Long = 0, Optional Height As Long = 0, Optional Transparent As Boolean = False) As udtSurface
  On Error GoTo ErrHandler
      
  Dim DXSurfaceDescriptor   As DDSURFACEDESC2
  Dim TransparencyKey       As DDCOLORKEY
  Dim TempRect              As RECT
  Dim Pixel                 As Long
  
  
    'If the width and height are given then add them to the structure
    If Width <> 0 And Height <> 0 Then
        DXSurfaceDescriptor.lFlags = DDSD_CAPS Or DDSD_HEIGHT Or DDSD_WIDTH
        DXSurfaceDescriptor.lWidth = Width
        DXSurfaceDescriptor.lHeight = Height
    Else
        DXSurfaceDescriptor.lFlags = DDSD_CAPS
    End If

    'Create a new surface from a file
    DXSurfaceDescriptor.ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN
    Set LoadSurface.DXSurface = DirectDraw.CreateSurfaceFromFile(BitmapFile, DXSurfaceDescriptor)
    
    'If we need a color key
    If Transparent Then
        TempRect = MakeRect(0, 0, DXSurfaceDescriptor.lWidth, DXSurfaceDescriptor.lHeight)
        
        'Get the pixel in the upper left corner and use it as the transparent color
        LoadSurface.DXSurface.Lock TempRect, DXSurfaceDescriptor, DDLOCK_NOSYSLOCK, Handle
        Pixel = LoadSurface.DXSurface.GetLockedPixel(0, 0)
        LoadSurface.DXSurface.Unlock TempRect
    
        'Create a color key
        TransparencyKey.low = Pixel
        TransparencyKey.high = Pixel
    
        'Apply it to the surface
        LoadSurface.DXSurface.SetColorKey DDCKEY_SRCBLT, TransparencyKey
    End If
    
    'Store some information about the surface in the structure
    LoadSurface.Width = DXSurfaceDescriptor.lWidth
    LoadSurface.Height = DXSurfaceDescriptor.lHeight
    LoadSurface.Transparent = Transparent
    LoadSurface.TransparentColor = Pixel
    
Exit Function
ErrHandler:
End Function



Public Sub CheckExlusiveMode()
  On Error GoTo ErrHandler

  Dim TestCoopRes As Long
    
    'Test the cooperative level
    TestCoopRes = DirectDraw.TestCooperativeLevel
    
    'If it returns not ok then go into a loop until we can reset the device
    If TestCoopRes <> DD_OK Then
        Do Until DirectDraw.TestCooperativeLevel = DD_OK
            DoEvents
        Loop
        
        'Reset the device and restore the surfaces
        DirectDraw.RestoreAllSurfaces
    End If
    
Exit Sub
ErrHandler:
End Sub



Public Sub SetFont(Bold As Boolean, Italic As Boolean, Underlined As Boolean, Size As Long, Name As String)
  On Error GoTo ErrHandler
  
  Dim NewFont As New StdFont
    
    'Fill out the stdFont structure
    With NewFont
        .Bold = Bold
        .Italic = Italic
        .Underline = Underlined
        .Name = Name
        .Size = Size
    End With
    
    'Set the back buffer's font
    Backbuffer.SetFont NewFont
    
Exit Sub
ErrHandler:
End Sub



Public Sub SetLineProperties(ForeColor As Long, LineWidth As Long)
  On Error GoTo ErrHandler
  
    'Set the new forecolor and line width
    Backbuffer.SetForeColor ForeColor
    Backbuffer.setDrawWidth LineWidth
    
Exit Sub
ErrHandler:
End Sub


Public Sub ClearScreen(Color As Long)
  On Error GoTo ErrHandler

  Dim ScreenRect As RECT
    
    'Get the window rect and color fill it
    DirectX.GetWindowRect Handle, ScreenRect
    Backbuffer.BltColorFill ScreenRect, Color

Exit Sub
ErrHandler:
End Sub



Public Sub DrawBox(X1 As Long, Y1 As Long, X2 As Long, Y2 As Long)
  On Error GoTo ErrHandler

    'Draw a box from (X1,Y1) to (X2,Y2)
    Backbuffer.DrawBox X1, Y1, X2, Y2

Exit Sub
ErrHandler:
End Sub



Public Sub DrawCircle(CentreX As Long, CentreY As Long, Radius As Long)
  On Error GoTo ErrHandler

    'Draw a circle with centre (CentreX,CentreY)and radius, Radius
    Backbuffer.DrawCircle CentreX, CentreY, Radius

Exit Sub
ErrHandler:
End Sub



Public Sub DrawEllipse(X1 As Long, Y1 As Long, X2 As Long, Y2 As Long)
  On Error GoTo ErrHandler

    'Draw an ellipse from (X1,Y1) to (X2,Y2)
    Backbuffer.DrawEllipse X1, Y1, X2, Y2

Exit Sub
ErrHandler:
End Sub



Public Sub DrawLine(X1 As Long, Y1 As Long, X2 As Long, Y2 As Long)
  On Error GoTo ErrHandler

    'Draw a line from (X1,Y1) to (X2,Y2)
    Backbuffer.DrawLine X1, Y1, X2, Y2

Exit Sub
ErrHandler:
End Sub



Public Sub DrawRoundedBox(X1 As Long, Y1 As Long, X2 As Long, Y2 As Long, RoundWidth As Long, RoundHeight As Long)
  On Error GoTo ErrHandler

    'Draw an rounded box from (X1,Y1) to (X2,Y2)
    Backbuffer.DrawRoundedBox X1, Y1, X2, Y2, RoundWidth, RoundHeight

Exit Sub
ErrHandler:
End Sub



Public Sub DrawText(X As Long, Y As Long, Text As String)
  On Error GoTo ErrHandler
    
    'Draw Text at position (X,Y)
    Backbuffer.DrawText X, Y, Text, False
    
Exit Sub
ErrHandler:
End Sub



Public Sub PresentScene()
  On Error GoTo ErrHandler
    
    'Flip the back buffer onto the screen
    PrimarySurface.Flip Nothing, DDFLIP_WAIT

Exit Sub
ErrHandler:
End Sub



Public Function UpdateFrameRate() As Long
  On Error GoTo ErrHandler
    
    'If this is the first time, the last tick count will be 0 and we have to get it
    If LastTickCount = 0 Then
        LastTickCount = GetTickCount
    End If
    
    'If a second has passed, update the new frame rate
    If GetTickCount - LastTickCount >= 1000 Then
        LastTickCount = GetTickCount
        LastFrameCount = CurrentFrameCount
        CurrentFrameCount = 0
    End If
    
    'Increment the frame count and return the frame count from the last second
    UpdateFrameRate = LastFrameCount
    CurrentFrameCount = CurrentFrameCount + 1

Exit Function
ErrHandler:
End Function



Public Function MoveMap(OffsetX As Long, OffsetY As Long) As Boolean
  On Error GoTo ErrHandler
  
    MoveMap = True
    
    'Ensure the map stays in the screen.
    'Returns true if the map was moved or false if it can't move (at the edge)
    
    'Check the right edge
    If ScreenWidth - OffsetX > Map.TilesX * Map.TileSet.TileWidth Then
        OffsetX = ScreenWidth - Map.TilesX * Map.TileSet.TileWidth
        MoveMap = False
    'Check the bottom edge
    ElseIf ScreenHeight - OffsetY > Map.TilesY * Map.TileSet.TileHeight Then
        OffsetY = ScreenHeight - Map.TilesY * Map.TileSet.TileHeight
        MoveMap = False
    'Check the left edge
    ElseIf OffsetX > 0 Then
        OffsetX = 0
        MoveMap = False
    'Check the right edge
    ElseIf OffsetY > 0 Then
        OffsetY = 0
        MoveMap = False
    End If
    
Exit Function
ErrHandler:
End Function



Public Sub MovePlayer()
  On Error GoTo ErrHandler
  
  Dim TempPosX As Long
  Dim TempPosY As Long
  Dim TempMapX As Long
  Dim TempMapY As Long
  
  Dim NewPosition As udtPoint
  
  Dim pt As udtPoint
    
    'Store the initial positions in temporary variables
    TempPosX = Player.PositionX
    TempPosY = Player.PositionY
    TempMapX = Map.OffsetX
    TempMapY = Map.OffsetY
    
    'X Axis
    
    If Player.PositionX < ScreenWidth \ 2 - 2 Or Player.PositionX > ScreenWidth \ 2 + 2 Then
        'if the player is not in the middle of the map then move the player
        If Player.GoingLeft Then Player.PositionX = Player.PositionX - Walk_Speed
        If Player.GoingRight Then Player.PositionX = Player.PositionX + Walk_Speed
    Else
        'Try to scroll the map
        If Player.GoingLeft Then Map.OffsetX = Map.OffsetX + Walk_Speed
        If Player.GoingRight Then Map.OffsetX = Map.OffsetX - Walk_Speed
        
        'If the map is at the edge then move the player
        If Not MoveMap(Map.OffsetX, Map.OffsetY) Then
            If Player.GoingLeft Then Player.PositionX = Player.PositionX - Walk_Speed
            If Player.GoingRight Then Player.PositionX = Player.PositionX + Walk_Speed
        End If
    End If

    'Y Axis

    If Player.PositionY < ScreenHeight \ 2 - 2 Or Player.PositionY > ScreenHeight \ 2 + 2 Then
        'if the player is not in the middle of the map then move the player
        If Player.GoingUp Then Player.PositionY = Player.PositionY - Walk_Speed
        If Player.GoingDown Then Player.PositionY = Player.PositionY + Walk_Speed
    Else
        'Try to scroll the map
        If Player.GoingUp Then Map.OffsetY = Map.OffsetY + Walk_Speed
        If Player.GoingDown Then Map.OffsetY = Map.OffsetY - Walk_Speed
        
        'If the map is at the edge then move the player
        If Not MoveMap(Map.OffsetX, Map.OffsetY) Then
            If Player.GoingUp Then Player.PositionY = Player.PositionY - Walk_Speed
            If Player.GoingDown Then Player.PositionY = Player.PositionY + Walk_Speed
        End If
    End If
    
    'If the player is not at the edge of the map or on a walkable tile then don't restore the
    'original positions. The new ones are fine.
    If Player.PositionX >= 0 And Player.PositionX + SpriteWidth <= ScreenWidth And Player.PositionY >= 0 And Player.PositionY + SpriteHeight \ 2 <= ScreenHeight Then

        If Not TileIsWalkable(Player.PositionX, Player.PositionY) Then
            'We need to restore the original positions
            Player.PositionX = TempPosX
            Player.PositionY = TempPosY
            Map.OffsetX = TempMapX
            Map.OffsetY = TempMapY
            
            Exit Sub
        End If
    Else
        'We need to restore the original positions
        Player.PositionX = TempPosX
        Player.PositionY = TempPosY
        Map.OffsetX = TempMapX
        Map.OffsetY = TempMapY
            
        Exit Sub
    End If
    
    pt = TileFromPixels(Player.PositionX + SpriteWidth \ 2, Player.PositionY + SpriteHeight \ 4)

    If pt.X <> LastPortalEntered.X Or pt.Y <> LastPortalEntered.Y Then
        LastPortalEntered.X = -1
        LastPortalEntered.Y = -1
        AllowPortalTravel = True
    End If

    If HasPortal(Player.PositionX, Player.PositionY) And AllowPortalTravel Then
        FadeOut 1
        
        #If DebugMode Then
            ClearScreen vbBlack
        #End If
        
        NewPosition.X = Map.Tiles(pt.X, pt.Y).PortalX
        NewPosition.Y = Map.Tiles(pt.X, pt.Y).PortalY
        
        If Map.Tiles(pt.X, pt.Y).PortalMapName <> vbNullString Then
            LoadMap App.Path & "\" & Replace$(Map.Tiles(pt.X, pt.Y).PortalMapName, Chr$(0), vbNullString)
        End If
        
        If NewPosition.X >= 0 And NewPosition.Y >= 0 Then
            PutPlayerOnTile NewPosition.X, NewPosition.Y
            LastPortalEntered.X = NewPosition.X
            LastPortalEntered.Y = NewPosition.Y
        Else
            LastPortalEntered.X = Map.StartX
            LastPortalEntered.Y = Map.StartY
        End If
        
        AllowPortalTravel = False
        
        DrawTiles
        DrawPlayer
        modEngine.DrawText 10, 10, "Frame rate: " & modEngine.UpdateFrameRate & " FPS"
        PresentScene
        DoEvents
    
        FadeIn 1
    End If
    
Exit Sub
ErrHandler:
End Sub



Public Sub DrawTiles()
  On Error GoTo ErrHandler
  
  Dim X As Long, Y As Long
  Dim OffsetX As Long, OffsetY As Long
  Dim SourcePoint As udtPoint
    
    #If DebugMode Then
        ClearScreen vbBlack
    #End If
    
    'Loop through all the tiles along the x axis
    For X = 1 To Map.TilesX
        
        'Only draw the row if some of them are visible
        OffsetX = (X - 1) * Map.TileSet.TileWidth + Map.OffsetX
        If OffsetX + Map.TileSet.TileWidth >= 0 And OffsetX - Map.TileSet.TileWidth <= ScreenWidth Then
            
            'Loop through all the tiles along the y axis
            For Y = 1 To Map.TilesY
            
                'Only draw the tiles in this column which are visible
                OffsetY = (Y - 1) * Map.TileSet.TileHeight + Map.OffsetY
                If OffsetY + Map.TileSet.TileHeight >= 0 And OffsetY - Map.TileSet.TileHeight <= ScreenHeight Then
                    'Get the position and draw it
                    SourcePoint = TileSet_OffsetFromIndex(Map.Tiles(X, Y).GraphicIndex)
                    
                    #If DebugMode Then
                        If Map.Tiles(X, Y).Walkable Then
                            If Map.Tiles(X, Y).HasPortal Then
                                SetLineProperties vbMagenta, 1
                            Else
                                SetLineProperties vbGreen, 1
                            End If
                        Else
                            SetLineProperties vbRed, 1
                        End If
                        
                        DrawBox OffsetX, OffsetY, OffsetX + Map.TileSet.TileWidth, OffsetY + Map.TileSet.TileHeight
                    #Else
                        BltSurface Map.TileSet.Surface, OffsetX, OffsetY, SourcePoint.X, SourcePoint.Y, Map.TileSet.TileWidth, Map.TileSet.TileHeight
                    #End If
                End If
            Next Y
        End If
    Next X
    
Exit Sub
ErrHandler:
End Sub



Public Sub DrawPlayer()
    BltSurface Player.DXSurface, Player.PositionX, Player.PositionY - SpriteHeight \ 2, Player.AnimationOffsetX, Player.AnimationOffsetY, SpriteWidth, SpriteHeight
End Sub



Public Sub BltSurface(Surface As udtSurface, ByVal DestX As Long, ByVal DestY As Long, Optional SourceX As Long = 0, Optional SourceY As Long = 0, Optional SourceWidth As Long = 0, Optional SourceHeight As Long = 0)
  On Error GoTo ErrHandler:
  
  Dim lngWidth  As Long
  Dim lngHeight As Long
  Dim lngX      As Long
  Dim lngY      As Long
    
    'If the width or height are not given then use the whole surface from the given point
    If SourceWidth = 0 Or SourceHeight = 0 Then
        SourceWidth = Surface.Width - SourceX
        SourceHeight = Surface.Height - SourceY
    End If
    
    'If the destination is off the screen then don't draw it
    If DestX >= ScreenWidth Or DestY >= ScreenHeight Or DestX + SourceWidth <= 0 Or DestY + SourceHeight <= 0 Then
        Exit Sub
    End If
    
    'If the surface goes off the screen to the left or the right then we have to compute the
    'x position and the width of the rect that is still on the screen.
    If DestX < 0 Then
        lngX = SourceX - DestX
        lngWidth = SourceWidth + DestX
    ElseIf DestX + SourceWidth > ScreenWidth Then
        lngX = SourceX
        lngWidth = ScreenWidth - DestX
    Else
        lngWidth = SourceWidth
        lngX = SourceX
    End If

    'If the surface goes off the screen to the top or the bottom then we have to compute the
    'y position and the height of the rect that is still on the screen.
    If DestY < 0 Then
        lngY = SourceY - DestY
        lngHeight = SourceHeight + DestY
    ElseIf DestY + SourceHeight > ScreenHeight Then
        lngY = SourceY
        lngHeight = ScreenHeight - DestY
    Else
        lngHeight = SourceHeight
        lngY = SourceY
    End If

    'Correct the destination co-ordinate incase it was off the screen
    If DestX < 0 Then DestX = 0
    If DestY < 0 Then DestY = 0

    'And finally blit to the back buffer
    Backbuffer.BltFast DestX, DestY, Surface.DXSurface, MakeRect(lngX, lngY, lngWidth, lngHeight), IIf(Surface.Transparent, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY, DDBLTFAST_WAIT)

Exit Sub
ErrHandler:
End Sub



Public Sub SetGamma(intRed As Integer, intGreen As Integer, intBlue As Integer)
  On Error GoTo ErrHandler
  
  Dim i As Integer

    'Alter the gamma ramp to the percent given by comparing to original state
    '-> A value of zero for intRed, intGreen, or intBlue will result in the gamma level being set back to the original levels
    '-> Anything above zero will fade towards full colour
    '-> anything below zero will fade towards no colour

    For i = 0 To 255
        Select Case intRed
            'Negative (convert to unsigned)
            Case Is < 0
                GammaRamp.red(i) = UnsignedToSignedValue(SignedToUnSignedValue(OriginalRamp.red(i)) * (100 - Abs(intRed)) / 100)
            'Zero (original gamma ramp)
            Case Is = 0
                 GammaRamp.red(i) = OriginalRamp.red(i)
            'Positive
            Case Is > 0
                GammaRamp.red(i) = UnsignedToSignedValue(65535 - ((65535 - SignedToUnSignedValue(OriginalRamp.red(i))) * (100 - intRed) / 100))
        End Select
        
        Select Case intGreen
            'Negative (convert to unsigned)
            Case Is < 0
                GammaRamp.green(i) = UnsignedToSignedValue(SignedToUnSignedValue(OriginalRamp.green(i)) * (100 - Abs(intGreen)) / 100)
            'Zero (original gamma ramp)
            Case Is = 0
                 GammaRamp.green(i) = OriginalRamp.green(i)
            'Positive
            Case Is > 0
                GammaRamp.green(i) = UnsignedToSignedValue(65535 - ((65535 - SignedToUnSignedValue(OriginalRamp.green(i))) * (100 - intGreen) / 100))
        End Select

        Select Case intBlue
            'Negative (convert to unsigned)
            Case Is < 0
                GammaRamp.blue(i) = UnsignedToSignedValue(SignedToUnSignedValue(OriginalRamp.blue(i)) * (100 - Abs(intBlue)) / 100)
            'Zero (original gamma ramp)
            Case Is = 0
                 GammaRamp.blue(i) = OriginalRamp.blue(i)
            'Poaitive
            Case Is > 0
                GammaRamp.blue(i) = UnsignedToSignedValue(65535 - ((65535 - SignedToUnSignedValue(OriginalRamp.blue(i))) * (100 - intBlue) / 100))
        End Select

    Next

    GammaRedVal = intRed
    GammaGreenVal = intGreen
    GammaBlueVal = intBlue

    GammaController.SetGammaRamp DDSGR_DEFAULT, GammaRamp

Exit Sub
ErrHandler:
End Sub



Public Sub FadeIn(DelayTime As Long)

  Dim i As Integer

    If Not GammaIsSupported Then Exit Sub
    
    'Fade from no color up to full color
    For i = -99 To 0
        Sleep DelayTime
        SetGamma i, i, i
    Next
    
End Sub



Public Sub FadeOut(DelayTime As Long)

  Dim i As Integer

    If Not GammaIsSupported Then Exit Sub
    
    'Fade from full color to no color
    For i = 0 To -99 Step -1
        Sleep DelayTime
        SetGamma i, i, i
    Next

End Sub


Public Sub CleanUpDirectX()
  On Error Resume Next

    If Handle > 0 Then
        'Display the cursor
        While ShowCursor(1) <= 0: Wend
        
        'Restore the display mode and cooperative level
        DirectDraw.RestoreDisplayMode
        DirectDraw.SetCooperativeLevel Handle, DDSCL_NORMAL
        
        'Destroy all objects
        Set Backbuffer = Nothing
        Set PrimarySurface = Nothing
        Set DirectDraw = Nothing
        Set DirectX = Nothing
        
        Handle = 0
    End If

End Sub
