Attribute VB_Name = "modMacros"
Option Explicit


Public DisplayModes() As DDSURFACEDESC2



Public Function TileIsWalkable(lngX As Long, lngY As Long) As Boolean

  Dim pt1 As udtPoint
  Dim pt2 As udtPoint
  Dim pt3 As udtPoint
  Dim pt4 As udtPoint
    
    'First calculate which tiles the player is in by calculating the 4 points as shown:
    
    '(pt1) _ _ (pt2)
    '     |   |
    '     |   |
    '     |_ _|
    '(pt3)     (pt4)
    
    'Then get the corresponding tile co-ordinate from these 4 points.
    
    pt1 = TileFromPixels(lngX, lngY)
    pt2 = TileFromPixels(lngX + SpriteWidth - 4, lngY)
    pt3 = TileFromPixels(lngX, lngY + SpriteHeight \ 2 - 4)
    pt4 = TileFromPixels(lngX + SpriteWidth - 4, lngY + SpriteHeight \ 2 - 4)
    
    'The rect joining the 4 points could occupy either 2 or 4 tiles.
    'We check all 4 points incase they are in 4 seperate tiles.
    
    'In order to return true, all (2 or 4) tiles must be walkable
    TileIsWalkable = Map.Tiles(pt1.X, pt1.Y).Walkable And _
                     Map.Tiles(pt2.X, pt2.Y).Walkable And _
                     Map.Tiles(pt3.X, pt3.Y).Walkable And _
                     Map.Tiles(pt4.X, pt4.Y).Walkable

End Function


Public Function HasPortal(lngX As Long, lngY As Long) As Boolean

  Dim pt As udtPoint

    pt = TileFromPixels(lngX + SpriteWidth \ 2, lngY + SpriteHeight \ 4)
    HasPortal = Map.Tiles(pt.X, pt.Y).HasPortal

End Function


Public Sub PutPlayerOnTile(TileX As Long, TileY As Long)

    'Start the Player in the centre of the screen
    Player.PositionX = ScreenWidth \ 2
    Player.PositionY = ScreenHeight \ 2
    
    'Work out the map offsets so that the starting tile is in the centre of the screen,
    'where the Player now is.
    Map.OffsetX = Player.PositionX - (TileX - 1) * Map.TileSet.TileWidth
    Map.OffsetY = Player.PositionY - (TileY - 1) * Map.TileSet.TileHeight
    
    'Sometimes this situation can occur:
    
'    _ _ _ _ _ _ _ _ _ _ _ _ _ _
'   |                           |
'   |     MAP                   |
'   |                           |
'   |            _ _ _ _ _ _ _ _|_ _ _ _
'   |           |               |       |
'   |           |               |       |
'   |           |               |       |
'   |           |             S |       |  S = Starting Square
'   |_ _ _ _ _ _|_ _ _ _ _ _ _ _|       |
'               |                       |
'               |        SCREEN         |
'               |_ _ _ _ _ _ _ _ _ _ _ _|
'
' Where placing the starting tile in the middle causes the map to be drawn too far left (or right or up or down)
    
    
    'Ensure the map isn't drawn too far right
    If Map.OffsetX > 0 Then
         Player.PositionX = (TileX - 1) * Map.TileSet.TileWidth
         Map.OffsetX = 0
    End If

    'Ensure the map isn't drawn too far left
    If Map.TilesX * Map.TileSet.TileWidth - (ScreenWidth - Map.OffsetX) < 0 Then
        Map.OffsetX = ScreenWidth - Map.TilesX * Map.TileSet.TileWidth
        Player.PositionX = (TileX - 1) * Map.TileSet.TileWidth + Map.OffsetX
    End If
    
    'Ensure the map isn't drawn too far down
    If Map.OffsetY > 0 Then
         Player.PositionY = (TileY - 1) * Map.TileSet.TileHeight
         Map.OffsetY = 0
    End If
    
    'Ensure the map isn't drawn too far up
    If Map.TilesY * Map.TileSet.TileHeight - (ScreenHeight - Map.OffsetY) < 0 Then
        Map.OffsetY = ScreenHeight - Map.TilesY * Map.TileSet.TileHeight
        Player.PositionY = (TileY - 1) * Map.TileSet.TileHeight + Map.OffsetY
    End If


End Sub


Public Function DisplayModeIsSupported(Width As Long, Height As Long, BitDepth As Long) As Boolean

  On Error GoTo ErrHandler

  Dim i As Long
    
    'Loop through each display mode
    For i = 1 To UBound(DisplayModes)
        
        'If this mode matches the one being queried then it is supported - no need to continue
        If DisplayModes(i).lWidth = Width And DisplayModes(i).lHeight = Height And DisplayModes(i).ddpfPixelFormat.lRGBBitCount = BitDepth Then
            DisplayModeIsSupported = True
            Exit Function
        End If
    Next i
    
'If we get here then either some error occured or the mode is not supported.
ErrHandler:
    DisplayModeIsSupported = False
End Function



Public Function UnsignedToSignedValue(lngValue As Long) As Integer
    UnsignedToSignedValue = CInt(lngValue - IIf(lngValue <= 32767, 0, 65535))
End Function



Public Function SignedToUnSignedValue(intValue As Integer) As Long
    SignedToUnSignedValue = CLng(intValue + IIf(intValue >= 0, 0, 65535))
End Function



Public Function TileFromPixels(X As Long, Y As Long) As udtPoint
    TileFromPixels.X = (X - Map.OffsetX) \ Map.TileSet.TileWidth + 1
    TileFromPixels.Y = (Y - Map.OffsetY) \ Map.TileSet.TileHeight + 1
    
    If TileFromPixels.X > Map.TilesX Then TileFromPixels.X = Map.TilesX
    If TileFromPixels.Y > Map.TilesY Then TileFromPixels.Y = Map.TilesY
End Function



Public Function CoordinatesFromPixels(PixelsX As Long, PixelsY As Long) As udtPoint
    CoordinatesFromPixels.X = (PixelsX \ Map.TileSet.TileWidth) + 1
    CoordinatesFromPixels.Y = (PixelsY \ Map.TileSet.TileHeight) + 1
End Function



Public Function PixelsFromCoordinate(CoordinateX As Long, CoordinateY As Long) As udtPoint
    PixelsFromCoordinate.X = (CoordinateX - 1) * Map.TileSet.TileWidth
    PixelsFromCoordinate.Y = (CoordinateY - 1) * Map.TileSet.TileHeight
End Function



Public Function TileSet_OffsetFromIndex(Index As Integer) As udtPoint
    TileSet_OffsetFromIndex = Tileset_CoordinateFromIndex(Index)
    TileSet_OffsetFromIndex = PixelsFromCoordinate(TileSet_OffsetFromIndex.X, TileSet_OffsetFromIndex.Y)
End Function



Public Function Tileset_CoordinateFromIndex(Index As Integer) As udtPoint

  Dim OffsetX As Integer
  Dim OffsetY As Integer
    
    OffsetY = (Index \ Map.TileSet.TilesX) + 1
    OffsetX = Index Mod Map.TileSet.TilesX
    
    If OffsetX = 0 Then
        OffsetX = Map.TileSet.TilesX
        OffsetY = OffsetY - 1
    End If
    
    Tileset_CoordinateFromIndex.X = OffsetX
    Tileset_CoordinateFromIndex.Y = OffsetY
    
End Function



Public Function Tileset_IndexFromCoordinate(OffsetX As Long, OffsetY As Long) As Integer
    
    OffsetX = OffsetX \ Map.TileSet.TileWidth
    OffsetY = OffsetY \ Map.TileSet.TileHeight
    
    Tileset_IndexFromCoordinate = (OffsetY * Map.TileSet.TilesX) + (OffsetX + 1)
    
End Function



Public Function MakePoint(X As Long, Y As Long) As udtPoint
    MakePoint.X = X
    MakePoint.Y = Y
End Function



Public Function MakeRect(X As Long, Y As Long, Width As Long, Height As Long) As RECT
    With MakeRect
        .Left = X
        .Top = Y
        .Right = X + Width
        .Bottom = Y + Height
    End With
End Function


