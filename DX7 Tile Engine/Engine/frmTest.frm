VERSION 5.00
Begin VB.Form frmTest 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   3090
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "frmTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
Private Declare Function GetTickCount Lib "kernel32" () As Long


Private LastTick    As Long
Private bRunning    As Boolean
Private bMoving     As Boolean


Private Sub form_click()
    bRunning = False
End Sub


Private Sub Form_Load()

  Dim PlayerPt As udtPoint
    
    Me.Show
    Me.Refresh
    
    '--- Initialize -------
    bRunning = modEngine.InitializeDirectX(Me.hWnd, 640, 480, 32)
    modEngine.SetFont False, False, False, 12, "Arial"
    modEngine.SetLineProperties vbRed, 2
    
    Player.DXSurface = modEngine.LoadSurface(App.Path & "\Player.Bmp", , , True)
    modEngine.LoadMap (App.Path & "\World.Map")
    Map.TileSet.Surface = modEngine.LoadSurface(App.Path & "\Tileset.Bmp")

    LastTick = GetTickCount
    '-----------------------
    
    '--- Fade in -----------
    modEngine.DrawTiles
    modEngine.DrawPlayer
    modEngine.PresentScene
    
    FadeIn 1
    DoEvents
    '-----------------------
    
    '--- Main Loop ---------
    While bRunning
        modEngine.CheckExlusiveMode
        CheckKeys
        modEngine.DrawTiles
        modEngine.DrawPlayer
        
        'Draw framerate
        modEngine.DrawText 10, 10, "Frame rate: " & modEngine.UpdateFrameRate & " FPS"
        
        'Draw current tile position
        PlayerPt = TileFromPixels(Player.PositionX + SpriteWidth \ 2, Player.PositionY + SpriteHeight \ 4)
        modEngine.DrawText 10, 30, "Position: " & PlayerPt.X & ":" & PlayerPt.Y
        
        'Flip
        modEngine.PresentScene
        DoEvents
    Wend
    '-----------------------
    
    '--- Cleanup ----------
    modEngine.FadeOut 1
    modEngine.CleanUpDirectX
    Me.Hide

    Unload Me
    '-----------------------

End Sub



Private Sub CheckKeys()
    
    bMoving = False
    
    'Left
    If GetAsyncKeyState(vbKeyLeft) <> 0 Then
        Player.GoingLeft = True: Player.GoingRight = False
        bMoving = True
    'Right
    ElseIf GetAsyncKeyState(vbKeyRight) <> 0 Then
        Player.GoingLeft = False: Player.GoingRight = True
        bMoving = True
    'Neither
    Else
        Player.GoingLeft = False: Player.GoingRight = False
    End If
    
    'Up
    If GetAsyncKeyState(vbKeyUp) <> 0 Then
        Player.GoingUp = True: Player.GoingDown = False
        bMoving = True
    'Down
    ElseIf GetAsyncKeyState(vbKeyDown) <> 0 Then
        Player.GoingUp = False: Player.GoingDown = True
        bMoving = True
    'Neither
    Else
        Player.GoingUp = False: Player.GoingDown = False
    End If
    
    'Escape
    If GetAsyncKeyState(vbKeyEscape) <> 0 Then
        bRunning = False
        Exit Sub
    End If
    
    'Update animations
    With Player
        If .GoingUp And .GoingRight Then
            .AnimationOffsetY = 1 * SpriteHeight
        ElseIf .GoingDown And .GoingRight Then
            .AnimationOffsetY = 3 * SpriteHeight
        ElseIf .GoingDown And .GoingLeft Then
            .AnimationOffsetY = 5 * SpriteHeight
        ElseIf .GoingUp And .GoingLeft Then
            .AnimationOffsetY = 7 * SpriteHeight
        ElseIf .GoingUp Then
            .AnimationOffsetY = 0 * SpriteHeight
        ElseIf .GoingRight Then
            .AnimationOffsetY = 2 * SpriteHeight
        ElseIf .GoingDown Then
            .AnimationOffsetY = 4 * SpriteHeight
        ElseIf .GoingLeft Then
            .AnimationOffsetY = 6 * SpriteHeight
        End If
    End With
    
    'Limit animation speed
    If bMoving Then
        modEngine.MovePlayer

        If GetTickCount - LastTick > 100 Then
            LastTick = GetTickCount
            Player.AnimationOffsetX = Player.AnimationOffsetX + SpriteWidth
            If Player.AnimationOffsetX > SpriteWidth * 3 Then Player.AnimationOffsetX = 0
        End If
    End If

End Sub
