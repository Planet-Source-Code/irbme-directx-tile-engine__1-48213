Attribute VB_Name = "modGlobals"
Option Explicit

Public Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Public Declare Function TransparentBlt Lib "msimg32.dll" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal crTransparent As Long) As Boolean
Public Declare Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer


Public Const TILESET_TILESX As Long = 9
Public Const TILESET_TILESY As Long = 16

Public Type udtPoint
    X As Integer
    Y As Integer
End Type


Public Type udtSurface
    DXSurface           As Long
    Width               As Long
    Height              As Long
    TransparentColor    As Long
    Transparent         As Boolean
End Type


Public Type udtSingleTile
    GraphicIndex        As Integer              'The tile index from the tileset.
    Walkable            As Boolean              'Whether or not this tile is walkable

    HasPortal           As Boolean
    PortalX             As Long
    PortalY             As Long
    PortalMapName       As String
End Type


Public Type udtTileSet
    TilesX              As Long
    TilesY              As Long
    TileWidth           As Long
    TileHeight          As Long
    Surface             As udtSurface
End Type


Public Type udtMap
    TilesX              As Long
    TilesY              As Long
    StartX              As Long
    StartY              As Long
    TileSet             As udtTileSet
    Tiles()             As udtSingleTile
End Type


Public Map              As udtMap
Public bDrawing         As Boolean

Public CurrentIndex     As Integer
Public CurrentTile      As udtPoint


Public Function LoadMap(FileName As String)

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

        Get #FileNum, , TempInt
        Map.StartX = CLng(TempInt)
        Get #FileNum, , TempInt
        Map.StartY = CLng(TempInt)

        Get #FileNum, , TempByte
        Map.TileSet.TilesX = CLng(TempByte)
        Get #FileNum, , TempByte
        Map.TileSet.TilesY = CLng(TempByte)
        Get #FileNum, , TempByte
        Map.TileSet.TileWidth = CLng(TempByte)
        Get #FileNum, , TempByte
        Map.TileSet.TileHeight = CLng(TempByte)

        frmMap.NewMap Map.TilesX, Map.TilesY, Map.TileSet.TileWidth, Map.TileSet.TileHeight

        'Read the tile index
        For X = 1 To Map.TilesX
            For Y = 1 To Map.TilesY
                Get #FileNum, , TempByte
                Map.Tiles(X, Y).GraphicIndex = CInt(TempByte)
                Get #FileNum, , TempByte
                Map.Tiles(X, Y).Walkable = IIf(TempByte = 1, True, False)
            Next
        Next
        
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
End Function


Public Function SaveMap(FileName As String)
  
  Dim FileNum   As Integer
  Dim TempByte  As Byte
  Dim TempInt   As Integer
  Dim X         As Long
  Dim Y         As Long
  
  Dim NumberOfPortals As Long
  Dim PortalMapName As String * 16
  
    FileNum = FreeFile

    Open FileName For Binary Access Write As #FileNum
        
        'Get map width and height
        TempInt = CInt(Map.TilesX)
        Put #FileNum, , TempInt
        TempInt = CInt(Map.TilesY)
        Put #FileNum, , TempInt
        
        TempInt = CInt(Map.StartX)
        Put #FileNum, , TempInt
        TempInt = CInt(Map.StartY)
        Put #FileNum, , TempInt
        
        TempByte = CByte(Map.TileSet.TilesX)
        Put #FileNum, , TempByte
        TempByte = CByte(Map.TileSet.TilesY)
        Put #FileNum, , TempByte
        TempByte = CByte(Map.TileSet.TileWidth)
        Put #FileNum, , TempByte
        TempByte = CByte(Map.TileSet.TileHeight)
        Put #FileNum, , TempByte

        'Read the tile index
        For X = 1 To Map.TilesX
            For Y = 1 To Map.TilesY
                
                If Map.Tiles(X, Y).HasPortal Then
                    NumberOfPortals = NumberOfPortals + 1
                End If

                TempByte = CByte(Map.Tiles(X, Y).GraphicIndex)
                Put #FileNum, , TempByte
                TempByte = CByte(IIf(Map.Tiles(X, Y).Walkable, 1, 0))
                Put #FileNum, , TempByte
            Next
        Next
        
        TempInt = CInt(NumberOfPortals)
        Put #FileNum, , TempInt
        
        For X = 1 To Map.TilesX
            For Y = 1 To Map.TilesY
            
                If Map.Tiles(X, Y).HasPortal Then
                    TempInt = CInt(X)
                    Put #FileNum, , TempInt
                    
                    TempInt = CInt(Y)
                    Put #FileNum, , TempInt
                        
                    TempInt = CInt(Map.Tiles(X, Y).PortalX)
                    Put #FileNum, , TempInt
                
                    TempInt = CInt(Map.Tiles(X, Y).PortalY)
                    Put #FileNum, , TempInt
                    
                    PortalMapName = Map.Tiles(X, Y).PortalMapName & String(16 - Len(Map.Tiles(X, Y).PortalMapName), Chr$(0))
                    Put #FileNum, , PortalMapName
                End If
            Next Y
        Next X
        
    Close #FileNum

End Function


Public Function CoordinatesFromPixels(PixelsX As Integer, PixelsY As Integer) As udtPoint
    CoordinatesFromPixels.X = (PixelsX \ Map.TileSet.TileWidth) + 1
    CoordinatesFromPixels.Y = (PixelsY \ Map.TileSet.TileHeight) + 1
End Function


Public Function PixelsFromCoordinate(CoordinateX As Integer, CoordinateY As Integer) As udtPoint
    PixelsFromCoordinate.X = (CoordinateX - 1) * Map.TileSet.TileWidth
    PixelsFromCoordinate.Y = (CoordinateY - 1) * Map.TileSet.TileHeight
End Function


Public Function Tileset_CoordinateFromIndex(TileSet As udtTileSet, Index As Integer) As udtPoint

  Dim OffsetX As Integer
  Dim OffsetY As Integer
    
    OffsetY = (Index \ TileSet.TilesX) + 1
    OffsetX = Index Mod TileSet.TilesX
    
    If OffsetX = 0 Then
        OffsetX = TileSet.TilesX
        OffsetY = OffsetY - 1
    End If
    
    Tileset_CoordinateFromIndex.X = OffsetX
    Tileset_CoordinateFromIndex.Y = OffsetY
    
End Function


Public Function Tileset_IndexFromCoordinate(TileSet As udtTileSet, OffsetX As Integer, OffsetY As Integer) As Integer
    
    OffsetX = OffsetX \ Map.TileSet.TileWidth
    OffsetY = OffsetY \ Map.TileSet.TileHeight
    
    Tileset_IndexFromCoordinate = (OffsetY * TileSet.TilesX) + (OffsetX + 1)
    
End Function
