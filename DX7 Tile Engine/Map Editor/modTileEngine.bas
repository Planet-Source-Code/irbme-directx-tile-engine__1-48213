Attribute VB_Name = "modTileEngine"
Option Explicit

#Const DebugMode = False

Private Declare Function GetTickCount Lib "kernel32" () As Long
Private Declare Function ShowCursor Lib "user32" (ByVal bShow As Long) As Long
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Public Declare Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer


Public Const Walk_Speed As Long = 1
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
    Portal_FadeInOut    As Boolean
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

Public Map                   As udtMap                  'The map



Public Function InitializeDirectX(hWnd As Long, Width As Long, Height As Long) As Boolean
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
    DirectDraw.SetCooperativeLevel hWnd, DDSCL_NORMAL
    
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
    
    Set Clipper = DirectDraw.CreateClipper(0)
    Clipper.SetHWnd hWnd
    PrimarySurface.SetClipper Clipper
    
    'Store some variables
    ScreenWidth = Width
    ScreenHeight = Height
    Handle = hWnd
    
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



Public Sub ClearScreen(Color As Long)
  On Error GoTo ErrHandler

  Dim ScreenRect As RECT
    
    'Get the window rect and color fill it
    DirectX.GetWindowRect Handle, ScreenRect
    Backbuffer.BltColorFill ScreenRect, Color

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
        
        'Only draw the column if some of them are visible
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



