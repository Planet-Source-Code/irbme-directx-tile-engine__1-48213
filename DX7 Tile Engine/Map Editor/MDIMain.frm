VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.MDIForm MDIMain 
   BackColor       =   &H8000000C&
   Caption         =   "Tile Editor"
   ClientHeight    =   10710
   ClientLeft      =   105
   ClientTop       =   750
   ClientWidth     =   11400
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.PictureBox Picture1 
      Align           =   4  'Align Right
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   10710
      Left            =   7035
      ScaleHeight     =   10680
      ScaleWidth      =   4335
      TabIndex        =   0
      Top             =   0
      Width           =   4370
      Begin VB.CheckBox chkHasPortal 
         Caption         =   "Contains Portal"
         Height          =   225
         Left            =   105
         TabIndex        =   11
         Top             =   8295
         Width           =   1380
      End
      Begin VB.TextBox txtPortalX 
         Height          =   285
         Left            =   2100
         MaxLength       =   4
         TabIndex        =   10
         Text            =   "0"
         Top             =   8610
         Width           =   435
      End
      Begin VB.TextBox txtPortalY 
         Height          =   285
         Left            =   2625
         MaxLength       =   4
         TabIndex        =   9
         Text            =   "0"
         Top             =   8610
         Width           =   435
      End
      Begin VB.TextBox txtPortalMapName 
         Height          =   285
         Left            =   2100
         MaxLength       =   16
         TabIndex        =   8
         Top             =   8925
         Width           =   1275
      End
      Begin VB.CommandButton cmdApply 
         Caption         =   "Apply"
         Height          =   645
         Left            =   3465
         TabIndex        =   7
         Top             =   8610
         Width           =   750
      End
      Begin VB.PictureBox picStart 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   480
         Left            =   105
         Picture         =   "MDIMain.frx":0000
         ScaleHeight     =   32
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   32
         TabIndex        =   6
         Top             =   9660
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.CheckBox chkShowNonWalkable 
         Caption         =   "Show Non Walkable Symbol"
         Height          =   195
         Left            =   1575
         TabIndex        =   5
         Top             =   8295
         Value           =   1  'Checked
         Width           =   2325
      End
      Begin VB.PictureBox picUnwalkable 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   480
         Left            =   735
         Picture         =   "MDIMain.frx":0C42
         ScaleHeight     =   32
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   32
         TabIndex        =   3
         Top             =   9660
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.PictureBox picTile 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   480
         Left            =   735
         ScaleHeight     =   30
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   30
         TabIndex        =   2
         Top             =   7770
         Width           =   480
      End
      Begin VB.PictureBox TileSet 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   7680
         Left            =   0
         OLEDropMode     =   1  'Manual
         Picture         =   "MDIMain.frx":1884
         ScaleHeight     =   512
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   288
         TabIndex        =   1
         Top             =   0
         Width           =   4320
         Begin VB.Shape Selector 
            BorderColor     =   &H000000FF&
            BorderWidth     =   3
            Height          =   480
            Left            =   0
            Top             =   0
            Width           =   480
         End
      End
      Begin VB.Label Label4 
         Caption         =   "Destination Tile:"
         Height          =   225
         Left            =   105
         TabIndex        =   14
         Top             =   8610
         Width           =   1485
      End
      Begin VB.Label Label3 
         Caption         =   "Destination Map Name:"
         Height          =   225
         Left            =   105
         TabIndex        =   13
         Top             =   8925
         Width           =   1695
      End
      Begin VB.Label lblCoordinates 
         Caption         =   "Label5"
         Height          =   225
         Left            =   1365
         TabIndex        =   12
         Top             =   7875
         Width           =   3270
      End
      Begin VB.Label Label2 
         Caption         =   "Tile:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   105
         TabIndex        =   4
         Top             =   7875
         Width           =   1275
      End
   End
   Begin MSComDlg.CommonDialog cd 
      Left            =   4620
      Top             =   5670
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Menu mnuFule 
      Caption         =   "File"
      Begin VB.Menu mnuNew 
         Caption         =   "New Map"
      End
      Begin VB.Menu mnuLoad 
         Caption         =   "Load Map"
      End
      Begin VB.Menu mnuSave 
         Caption         =   "Save Map"
      End
      Begin VB.Menu Sep 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu mnuMap 
      Caption         =   "Map"
      Begin VB.Menu mnuFill 
         Caption         =   "Fill With Current Tile"
      End
      Begin VB.Menu mnuRandomFillWalk 
         Caption         =   "Random Fill (Walkable)"
      End
      Begin VB.Menu mnuRandomFillUnwalk 
         Caption         =   "Random Fill (Unwalkable)"
      End
      Begin VB.Menu Sep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuBorderWalk 
         Caption         =   "Make Border (Walkable)"
      End
      Begin VB.Menu mnuBorderUnwalk 
         Caption         =   "Make Border (Unwalkable)"
      End
   End
   Begin VB.Menu mnuHlp 
      Caption         =   "Help"
      Begin VB.Menu mnuHelp 
         Caption         =   "Help"
      End
   End
End
Attribute VB_Name = "MDIMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Const SW_SHOWNORMAL = 1


Private Sub chkShowNonWalkable_Click()
    frmMap.picTiles.Refresh
End Sub
    

Private Sub cmdApply_Click()
    
    Map.Tiles(CurrentTile.X, CurrentTile.Y).HasPortal = IIf(chkHasPortal.Value = vbChecked, True, False)
    
    If Map.Tiles(CurrentTile.X, CurrentTile.Y).HasPortal Then
        Map.Tiles(CurrentTile.X, CurrentTile.Y).PortalX = CLng(txtPortalX.Text)
        Map.Tiles(CurrentTile.X, CurrentTile.Y).PortalY = CLng(txtPortalY.Text)
        Map.Tiles(CurrentTile.X, CurrentTile.Y).PortalMapName = txtPortalMapName.Text & String(16 - Len(txtPortalMapName.Text), Chr$(0))
    End If
    
End Sub


Private Sub MDIForm_Load()

    frmMap.NewMap 10, 10, 32, 32
    frmMap.picTiles.ForeColor = vbRed
    Load frmMap
    
    TileSet_MouseUp 1, 0, 0, 0
    TileSet_MouseUp 2, 0, 0, 0
    
    mnuFill_Click

End Sub


Private Sub mnuBorderWalk_Click()

  Dim X As Long
  Dim Y As Long
  
    For Y = 1 To Map.TilesY
        Map.Tiles(1, Y).Walkable = True
        Map.Tiles(1, Y).GraphicIndex = CurrentIndex
    Next Y
    
    For X = 1 To Map.TilesX
        Map.Tiles(X, 1).Walkable = True
        Map.Tiles(X, 1).GraphicIndex = CurrentIndex
    Next X
    
    For Y = 1 To Map.TilesY
        Map.Tiles(Map.TilesX, Y).Walkable = True
        Map.Tiles(Map.TilesX, Y).GraphicIndex = CurrentIndex
    Next Y
    
    For X = 1 To Map.TilesX
        Map.Tiles(X, Map.TilesY).Walkable = True
        Map.Tiles(X, Map.TilesY).GraphicIndex = CurrentIndex
    Next X
    
    frmMap.picTiles.Refresh
    
End Sub


Private Sub mnuBorderUnwalk_Click()

  Dim X As Long
  Dim Y As Long
  
    For Y = 1 To Map.TilesY
        Map.Tiles(1, Y).Walkable = False
        Map.Tiles(1, Y).GraphicIndex = CurrentIndex
    Next Y
    
    For X = 1 To Map.TilesX
        Map.Tiles(X, 1).Walkable = False
        Map.Tiles(X, 1).GraphicIndex = CurrentIndex
    Next X
    
    For Y = 1 To Map.TilesY
        Map.Tiles(Map.TilesX, Y).Walkable = False
        Map.Tiles(Map.TilesX, Y).GraphicIndex = CurrentIndex
    Next Y
    
    For X = 1 To Map.TilesX
        Map.Tiles(X, Map.TilesY).Walkable = False
        Map.Tiles(X, Map.TilesY).GraphicIndex = CurrentIndex
    Next X
    
    frmMap.picTiles.Refresh

End Sub



Private Sub mnuExit_Click()
    Unload Me
End Sub


Private Sub mnuFill_Click()
  
  Dim X As Long, Y As Long
  
    For X = 1 To Map.TilesX
        For Y = 1 To Map.TilesY
            Map.Tiles(X, Y).GraphicIndex = CurrentIndex
        Next Y
    Next X
    
    frmMap.picTiles.Refresh
  
End Sub


Private Sub mnuHelp_Click()
    ShellExecute Me.hwnd, vbNullString, App.Path & "\Help.txt", vbNullString, "C:\", SW_SHOWNORMAL
End Sub


Private Sub mnuNew_Click()

  Dim strInput    As String
  Dim bOK         As Boolean
  
    Do While Not bOK
        strInput = InputBox("Enter the number of tiles across the way", "Tile Dimensions")
        
        If StrPtr(strInput) = 0 Then
            Exit Sub
        End If
        
        If IsNumeric(strInput) Then
            bOK = True
            Map.TilesX = CInt(strInput)
        End If
    Loop
    
    bOK = False
    
    Do While Not bOK
        strInput = InputBox("Enter the number of tiles down the way", "Tile Dimensions")
        
        If IsNumeric(strInput) Then
            bOK = True
            Map.TilesY = CInt(strInput)
        End If
    Loop
    
    frmMap.NewMap Map.TilesX, Map.TilesY, 32, 32
    mnuFill_Click
    
End Sub


Private Sub mnuRandomFillWalk_Click()

  Dim strInput    As String
  Dim bOK         As Boolean
  Dim intNumber   As Long
  Dim i As Long
  Dim intRand1 As Long
  Dim intRand2 As Long
  
    Do While Not bOK
        strInput = InputBox("Enter the number tiles to add. Warning: If the map is already filled with this tile or you enter a larger number than there are squares not filled with this tile then continuing will crash the editor.", "Tile Dimensions")
        
        If StrPtr(strInput) = 0 Then
            Exit Sub
        End If
        
        If IsNumeric(strInput) Then
            bOK = True
            intNumber = CLng(strInput)
        End If
    Loop
    
    For i = 1 To intNumber
        
        intRand1 = Int(Rnd * Map.TilesX) + 1
        intRand2 = Int(Rnd * Map.TilesY) + 1
    
        If Map.Tiles(intRand1, intRand2).GraphicIndex = CurrentIndex Then
            i = i - 1
        Else
           Map.Tiles(intRand1, intRand2).GraphicIndex = CurrentIndex
           Map.Tiles(intRand1, intRand2).Walkable = True
        End If
    Next i
    
    frmMap.picTiles.Refresh
    
End Sub


Private Sub mnuRandomFillUnwalk_Click()

  Dim strInput    As String
  Dim bOK         As Boolean
  Dim intNumber   As Long
  Dim i As Long
  Dim intRand1 As Long
  Dim intRand2 As Long
  
    Do While Not bOK
        strInput = InputBox("Enter the number tiles to add. Warning: If the map is already filled with this tile or you enter a larger number than there are squares not filled with this tile then continuing will crash the editor.", "Tile Dimensions")
        
        If StrPtr(strInput) = 0 Then
            Exit Sub
        End If
        
        If IsNumeric(strInput) Then
            bOK = True
            intNumber = CLng(strInput)
        End If
    Loop
    
    For i = 1 To intNumber
        
        intRand1 = Int(Rnd * Map.TilesX) + 1
        intRand2 = Int(Rnd * Map.TilesY) + 1
    
        If Map.Tiles(intRand1, intRand2).GraphicIndex = CurrentIndex Then
            i = i - 1
        Else
           Map.Tiles(intRand1, intRand2).GraphicIndex = CurrentIndex
           Map.Tiles(intRand1, intRand2).Walkable = False
        End If
    Next i
    
    frmMap.picTiles.Refresh
    
End Sub


Private Sub mnuSave_Click()

    cd.InitDir = App.Path
    cd.Flags = cdlOFNHideReadOnly Or cdlOFNPathMustExist Or cdlOFNOverwritePrompt
    cd.DefaultExt = ".map"
    cd.Filter = "Map files | *.map"

    cd.ShowSave
    If cd.FileName <> vbNullString Then SaveMap (cd.FileName)
End Sub

Private Sub mnuLoad_Click()

    cd.InitDir = App.Path
    cd.Flags = cdlOFNHideReadOnly Or cdlOFNPathMustExist
    cd.DefaultExt = ".map"
    cd.Filter = "Map files | *.map"
    
    cd.ShowOpen
    
    If cd.FileName <> vbNullString Then
        LoadMap (cd.FileName)
        frmMap.picTiles.Refresh
    End If
    
End Sub


Private Sub TileSet_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
  Dim Offset As udtPoint

    CurrentIndex = modGlobals.Tileset_IndexFromCoordinate(Map.TileSet, CInt(X), CInt(Y))
    Offset = modGlobals.Tileset_CoordinateFromIndex(Map.TileSet, CurrentIndex)
    Offset = modGlobals.PixelsFromCoordinate(Offset.X, Offset.Y)
    
    BitBlt picTile.hdc, 0, 0, Map.TileSet.TileWidth, Map.TileSet.TileHeight, TileSet.hdc, Offset.X, Offset.Y, vbSrcCopy
    picTile.Refresh

    
    Selector.Left = Offset.X
    Selector.Top = Offset.Y
    
End Sub

