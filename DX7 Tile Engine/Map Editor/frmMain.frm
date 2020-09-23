VERSION 5.00
Begin VB.Form frmMap 
   Caption         =   "Map"
   ClientHeight    =   4665
   ClientLeft      =   525
   ClientTop       =   825
   ClientWidth     =   4995
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   4665
   ScaleWidth      =   4995
   WindowState     =   2  'Maximized
   Begin VB.VScrollBar VScroll1 
      Height          =   3585
      Left            =   4200
      Max             =   10
      Min             =   1
      TabIndex        =   3
      Top             =   210
      Value           =   1
      Width           =   225
   End
   Begin VB.HScrollBar HScroll1 
      Height          =   225
      Left            =   210
      Max             =   10
      Min             =   1
      TabIndex        =   2
      Top             =   3885
      Value           =   1
      Width           =   3900
   End
   Begin VB.PictureBox picMain 
      BackColor       =   &H8000000C&
      Height          =   3585
      Left            =   210
      ScaleHeight     =   235
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   256
      TabIndex        =   0
      Top             =   210
      Width           =   3900
      Begin VB.PictureBox picTiles 
         BackColor       =   &H80000006&
         BorderStyle     =   0  'None
         DrawWidth       =   2
         Height          =   960
         Left            =   0
         ScaleHeight     =   64
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   92
         TabIndex        =   1
         Top             =   0
         Width           =   1380
      End
   End
End
Attribute VB_Name = "frmMap"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Public Sub NewMap(TilesX As Long, TilesY As Long, TileWidth As Long, TileHeight As Long)

  Dim X As Long, Y As Long

    Erase Map.Tiles

    ReDim Map.Tiles(1 To TilesX, 1 To TilesY)
    ReDim Map.Tiles(1 To TilesX, 1 To TilesY)

    For X = 1 To TilesX
        For Y = 1 To TilesY
            Map.Tiles(X, Y).Walkable = True
        Next Y
    Next X

    With Map
        .TilesX = TilesX
        .TilesY = TilesY
        .TileSet.TileWidth = TileWidth
        .TileSet.TileHeight = TileHeight

        .TileSet.TilesX = TILESET_TILESX
        .TileSet.TilesY = TILESET_TILESY
    End With
    
    picTiles.Top = 0
    picTiles.Left = 0
    
    picTiles.Width = Map.TilesX * Map.TileSet.TileWidth
    picTiles.Height = Map.TilesY * Map.TileSet.TileHeight
    
    BitBlt picTiles.hdc, 0, 0, picTiles.ScaleWidth, picTiles.ScaleHeight, 0, 0, 0, vbBlackness
    
    picTiles_MouseUp 2, 0, 0, 0
    
    If picMain.ScaleWidth \ Map.TileSet.TileWidth < Map.TilesX Then
        HScroll1.Max = Map.TilesX - (picMain.ScaleWidth \ Map.TileSet.TileWidth) + 1
    End If
    
    If picMain.ScaleHeight \ Map.TileSet.TileHeight < Map.TilesY Then
        VScroll1.Max = Map.TilesY - (picMain.ScaleHeight \ Map.TileSet.TileHeight)
    End If
    
End Sub



Private Sub Form_Resize()
    
    If Me.WindowState = vbMinimized Or MDIMain.WindowState = vbMinimized Then Exit Sub

    picMain.Top = 0
    picMain.Left = 0
    picMain.Width = Me.Width - 360
    picMain.Height = Me.Height - 765
    
    VScroll1.Top = 0
    VScroll1.Left = Me.Width - 360
    VScroll1.Height = Me.Height - 765
    
    HScroll1.Left = 0
    HScroll1.Top = Me.Height - 765
    HScroll1.Width = Me.Width - 360
    
    If picMain.ScaleWidth \ Map.TileSet.TileWidth < Map.TilesX Then
        HScroll1.Max = Map.TilesX - (picMain.ScaleWidth \ Map.TileSet.TileWidth) + 1
    End If
    
    If picMain.ScaleHeight \ Map.TileSet.TileHeight < Map.TilesY Then
        VScroll1.Max = Map.TilesY - (picMain.ScaleHeight \ Map.TileSet.TileHeight)
    End If
    
End Sub


Private Sub HScroll1_Change()
    
    If picTiles.ScaleWidth > picMain.ScaleWidth Then
        picTiles.Left = -((HScroll1.Value - 1) * Map.TileSet.TileWidth)
    End If
    
End Sub


Private Sub VScroll1_Change()
    
    If picTiles.ScaleHeight > picMain.ScaleHeight Then
        picTiles.Top = -((VScroll1.Value - 1) * Map.TileSet.TileHeight)
    End If
    
End Sub


Private Sub picTiles_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    bDrawing = True
    picTiles_MouseMove Button, Shift, X, Y
End Sub


Private Sub picTiles_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
  Dim Offset As udtPoint
    
    If Button <> 1 And Shift = 0 Then
        Offset.X = X \ Map.TileSet.TileWidth
        Offset.Y = Y \ Map.TileSet.TileHeight

        CurrentTile.X = Offset.X + 1
        CurrentTile.Y = Offset.Y + 1
        
        If CurrentTile.X > UBound(Map.Tiles, 1) Or CurrentTile.Y > UBound(Map.Tiles, 2) Then Exit Sub
        
        MDIMain.chkHasPortal.Value = IIf(Map.Tiles(CurrentTile.X, CurrentTile.Y).HasPortal, vbChecked, vbUnchecked)
        MDIMain.txtPortalX.Text = Map.Tiles(Offset.X + 1, Offset.Y + 1).PortalX
        MDIMain.txtPortalY.Text = Map.Tiles(Offset.X + 1, Offset.Y + 1).PortalY
        MDIMain.txtPortalMapName.Text = Map.Tiles(Offset.X + 1, Offset.Y + 1).PortalMapName
        
        MDIMain.lblCoordinates.Caption = CurrentTile.X & " x " & CurrentTile.Y
        
        picTiles_Paint
    ElseIf GetKeyState(vbKeyControl) < 0 Then
        Offset.X = X \ Map.TileSet.TileWidth
        Offset.Y = Y \ Map.TileSet.TileHeight
        
        Map.StartX = Offset.X + 1
        Map.StartY = Offset.Y + 1
        Map.Tiles(Offset.X + 1, Offset.Y + 1).Walkable = True
        
        picTiles_Paint
    End If
    
    bDrawing = False
End Sub


Private Sub picTiles_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

  Dim Offset As udtPoint

    If bDrawing Then

        Offset.X = X \ Map.TileSet.TileWidth
        Offset.Y = Y \ Map.TileSet.TileHeight
        
        If Offset.X + 1 > Map.TilesX Or Offset.Y + 1 > Map.TilesY Then Exit Sub
        If Offset.X < 0 Or Offset.Y < 0 Then Exit Sub

        If Shift = 0 And Button = 1 Then
            Map.Tiles(Offset.X + 1, Offset.Y + 1).GraphicIndex = CurrentIndex
        ElseIf Shift = 1 And (Map.StartX <> Offset.X + 1 Or Map.StartY <> Offset.Y + 1) Then
            Map.Tiles(Offset.X + 1, Offset.Y + 1).Walkable = IIf(Button = 1, True, False)
        End If
        
        picTiles_Paint
    End If
    
End Sub


Private Sub picTiles_Paint()

  Dim X As Long, Y As Long
  Dim Offset As udtPoint

    For X = 0 To Map.TilesX - 1
        For Y = 0 To Map.TilesY - 1

            Offset = modGlobals.Tileset_CoordinateFromIndex(Map.TileSet, Map.Tiles(X + 1, Y + 1).GraphicIndex)
            Offset = modGlobals.PixelsFromCoordinate(Offset.X, Offset.Y)

            If Map.Tiles(X + 1, Y + 1).GraphicIndex <> 0 Then
                BitBlt picTiles.hdc, X * Map.TileSet.TileWidth, Y * Map.TileSet.TileHeight, Map.TileSet.TileWidth, Map.TileSet.TileHeight, MDIMain.TileSet.hdc, Offset.X, Offset.Y, vbSrcCopy
            End If

            If Map.StartX = X + 1 And Map.StartY = Y + 1 Then
                TransparentBlt picTiles.hdc, X * Map.TileSet.TileWidth, Y * Map.TileSet.TileHeight, Map.TileSet.TileWidth, Map.TileSet.TileHeight, MDIMain.picStart.hdc, 0, 0, Map.TileSet.TileWidth, Map.TileSet.TileHeight, vbBlue
            ElseIf Not Map.Tiles(X + 1, Y + 1).Walkable And MDIMain.chkShowNonWalkable.Value = vbChecked Then
                TransparentBlt picTiles.hdc, X * Map.TileSet.TileWidth, Y * Map.TileSet.TileHeight, Map.TileSet.TileWidth, Map.TileSet.TileHeight, MDIMain.picUnwalkable.hdc, 0, 0, Map.TileSet.TileWidth, Map.TileSet.TileHeight, vbBlue
            End If
            
            If Map.Tiles(X + 1, Y + 1).HasPortal Then
                picTiles.ForeColor = vbMagenta
            
                Offset.X = X * Map.TileSet.TileWidth
                Offset.Y = Y * Map.TileSet.TileHeight
                
                picTiles.Line (Offset.X, Offset.Y)-(Offset.X + Map.TileSet.TileWidth, Offset.Y)
                picTiles.Line (Offset.X, Offset.Y)-(Offset.X, Offset.Y + Map.TileSet.TileHeight)
            
                picTiles.Line (Offset.X, Offset.Y + Map.TileSet.TileHeight)-(Offset.X + Map.TileSet.TileWidth, Offset.Y + Map.TileSet.TileHeight)
                picTiles.Line (Offset.X + Map.TileSet.TileWidth, Offset.Y)-(Offset.X + Map.TileSet.TileWidth, Offset.Y + Map.TileSet.TileHeight)
            End If
 
        Next Y
    Next X
    
    picTiles.ForeColor = vbRed
    
    Offset.X = (CurrentTile.X - 1) * Map.TileSet.TileWidth
    Offset.Y = (CurrentTile.Y - 1) * Map.TileSet.TileHeight
    
    picTiles.Line (Offset.X, Offset.Y)-(Offset.X + Map.TileSet.TileWidth, Offset.Y)
    picTiles.Line (Offset.X, Offset.Y)-(Offset.X, Offset.Y + Map.TileSet.TileHeight)

    picTiles.Line (Offset.X, Offset.Y + Map.TileSet.TileHeight)-(Offset.X + Map.TileSet.TileWidth, Offset.Y + Map.TileSet.TileHeight)
    picTiles.Line (Offset.X + Map.TileSet.TileWidth, Offset.Y)-(Offset.X + Map.TileSet.TileWidth, Offset.Y + Map.TileSet.TileHeight)

End Sub
