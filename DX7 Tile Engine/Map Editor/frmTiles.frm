VERSION 5.00
Begin VB.Form frmTiles 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Tileset"
   ClientHeight    =   3540
   ClientLeft      =   11985
   ClientTop       =   5415
   ClientWidth     =   2700
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3540
   ScaleWidth      =   2700
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox picObject 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   480
      Left            =   2100
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   5
      Top             =   2310
      Width           =   480
   End
   Begin VB.PictureBox picObject2 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   480
      Left            =   2100
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   4
      Top             =   2940
      Width           =   480
   End
   Begin VB.PictureBox picObjects 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   480
      Left            =   0
      Picture         =   "frmTiles.frx":0000
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   128
      TabIndex        =   3
      Top             =   2940
      Width           =   1920
      Begin VB.Shape Selector2 
         BorderColor     =   &H000000FF&
         BorderWidth     =   3
         Height          =   480
         Left            =   0
         Top             =   0
         Width           =   480
      End
   End
   Begin VB.PictureBox picTile2 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   480
      Left            =   2100
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   2
      Top             =   735
      Width           =   480
   End
   Begin VB.PictureBox TileSet 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   2880
      Left            =   0
      Picture         =   "frmTiles.frx":3042
      ScaleHeight     =   192
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   128
      TabIndex        =   1
      Top             =   0
      Width           =   1920
      Begin VB.Shape Selector 
         BorderColor     =   &H000000FF&
         BorderWidth     =   3
         Height          =   480
         Left            =   0
         Top             =   0
         Width           =   480
      End
   End
   Begin VB.PictureBox picTile 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   480
      Left            =   2100
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   0
      Top             =   105
      Width           =   480
   End
End
Attribute VB_Name = "frmTiles"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub Form_Load()
    TileSet_MouseUp 1, 0, 0, 0
    TileSet_MouseUp 2, 0, 0, 0
    picObjects_MouseUp 1, 0, 0, 0
    picObjects_MouseUp 2, 0, 0, 0
End Sub


Private Sub picObjects_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

  Dim Offset As udtPoint

    If Button = 1 Then
        CurrentObjectIndex1 = modGlobals.Tileset_IndexFromCoordinate(Map.Layer2TileSet, CInt(X), CInt(Y))
        Offset = modGlobals.Tileset_CoordinateFromIndex(Map.Layer1TileSet, CurrentObjectIndex1)
        Offset = modGlobals.PixelsFromCoordinate(Offset.X, Offset.Y)
        
        TransparentBlt picObject.hdc, 0, 0, Map.TileWidth, Map.TileHeight, picObjects.hdc, Offset.X, Offset.Y, Map.TileWidth, Map.TileHeight, Map.TransparentColor
        picObject.Refresh
    Else
        CurrentObjectIndex2 = modGlobals.Tileset_IndexFromCoordinate(Map.Layer2TileSet, CInt(X), CInt(Y))
        Offset = modGlobals.Tileset_CoordinateFromIndex(Map.Layer2TileSet, CurrentObjectIndex2)
        Offset = modGlobals.PixelsFromCoordinate(Offset.X, Offset.Y)
        
        TransparentBlt picObject2.hdc, 0, 0, Map.TileWidth, Map.TileHeight, picObjects.hdc, Offset.X, Offset.Y, Map.TileWidth, Map.TileHeight, Map.TransparentColor
        picObject2.Refresh
    End If

    Selector2.Left = Offset.X
    Selector2.Top = Offset.Y

End Sub


Private Sub TileSet_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
  Dim Offset As udtPoint

    If Button = 1 Then
        CurrentIndex1 = modGlobals.Tileset_IndexFromCoordinate(Map.Layer1TileSet, CInt(X), CInt(Y))
        Offset = modGlobals.Tileset_CoordinateFromIndex(Map.Layer1TileSet, CurrentIndex1)
        Offset = modGlobals.PixelsFromCoordinate(Offset.X, Offset.Y)
        
        BitBlt picTile.hdc, 0, 0, Map.TileWidth, Map.TileHeight, TileSet.hdc, Offset.X, Offset.Y, vbSrcCopy
        picTile.Refresh
    Else
        CurrentIndex2 = modGlobals.Tileset_IndexFromCoordinate(Map.Layer1TileSet, CInt(X), CInt(Y))
        Offset = modGlobals.Tileset_CoordinateFromIndex(Map.Layer1TileSet, CurrentIndex2)
        Offset = modGlobals.PixelsFromCoordinate(Offset.X, Offset.Y)
        
        BitBlt picTile2.hdc, 0, 0, Map.TileWidth, Map.TileHeight, TileSet.hdc, Offset.X, Offset.Y, vbSrcCopy
        picTile2.Refresh
    End If
    
    Selector.Left = Offset.X
    Selector.Top = Offset.Y
    
End Sub

