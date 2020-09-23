VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PanelDePonCls"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit 'I upgraded to a class so I could have multiple instances and events
'New switching effect originally contributed by dafhi
'Garbage < 0, =0 empty, =1 cursor, > 1 tile
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long

Public Enum ComboOP
    None = -1
    AddOne = 0
    Enable
    Disable
    Clear
End Enum

Public Enum GarbageAction
    NoOp = 0
    destroy = 1
    Drop = 2
End Enum

Public Enum PanelAction
    Moveup
    MoveDown
    MoveLeft
    MoveRight
    MoveTiles
    MoveBoard
End Enum

Public Enum GameMode
    Normal = 0 'Normal go on for infinite mode
    Puzzle = 1 'Get rid of all the blocks with a certain number of moves
    Edit = 2 'Create a custom puzzle
    Lineclear = 3 'Get all the blocks below a certain line
    Timelimit = 4 'clear as many blocks as you can before time runs out
End Enum

Private Type Tile '6*4 bytes + 1 bool = 24 and 1/8 bytes per tile
    Color As Long
    X As Long
    Y As Long
    Clean As Boolean
    Z As Long
    
    'Garbage block dimensions
    width As Long
    height As Long
End Type

Private Type Combo
    value As Long
    Used As Boolean
End Type

'Used to keep two remote classes synchronized
Public Event PDPGameOver() 'Occurs when a game is ended by lack of moves or tiles reaching the top
Public Event PDPGameWon() 'the game was won!
Public Event NewRow(Row As String) 'a new row was created, send it to the remote client
Public Event CursorMoved(X As Long, Y As Long) 'the user moved the cursor
Public Event GridCleared() 'the grid was empties
Public Event ActionPerformed(Index As PanelAction) 'a user action was performed
Public Event MapLoaded() 'a new map was loaded/created
Public Event OffsetIncreased() 'the board offset was increased
Public Event OffsetsDecremented() 'all the little offsets were decreased
Public Event RaiseLevel() 'increase the speed that the board goes upwards
Public Event KillTiles(count As Long) 'Used to increment the score on normal clears
Public Event TopChanged(height As Long) 'Used to change the music when the top reached a certain level
Public Event TopRaised() 'Used when the bar is raised to synch the AI
Public Event ComboMade(Hits As Long) 'Used when combos are made
Public Event ChainInProgress(Hits As Long) 'Used while combos are being made
Public Event ScoreChanged(Score As Long) 'used when the score changes
Public Event GarbageCreated(X As Long, Y As Long, width As Long, height As Long)
Public Event GarbageSubstituted(X As Long, Y As Long, Color As Long)

Public Event MovesChainged(Moves As Long)
Public Event LinesChainged(Lines As Long)

Private Const Adjacent As Long = 2, pi As Long = 3.14159265358979, two_pi As Long = 2 * pi, convert As Double = pi / 180 'Constants
Public Destination As PictureBox, Source As PictureBox, Display As PictureBox  'Objects
Public GridWidth As Long, GridHeight As Long, TileSize As Long, Colors As Long, Ghost As Long, Clean As Boolean 'Graphic variables
Public TileCount As Long, Offset As Long, CursorX As Long, CursorY As Long, BoardHeight As Long 'Dimension variables
Public GameOver As Boolean, GameWon As Boolean, Score As Long, Lines As Long, Moves As Long, Currtime As Long, Speed As Long 'Game Variables
Public Stopper As Long, Mode As GameMode, IncreaseSpeed As Boolean, Level As Long, doBG As Boolean, PixelSpeed As Long, doGarbage As Long, GColors As Long, Locked As Boolean
Private Grid() As Tile, UndoList() As String, UndoCount As Long, Combos() As Combo 'Game board (row, column)
Private ThreeDEffect() As Long, BoardWidth As Long, ThreeDMode As Boolean, OffsetX As Long

Private Sub TransBLT(SrcHDC As Long, xSrc As Long, ySrc As Long, MaskHDC As Long, ByVal Xmsk As Long, ByVal Ymsk As Long, width As Long, height As Long, DestHDC As Long, X As Long, Y As Long, Optional UseSrcAsMask As Boolean)
    Const SRCPAINT = &HEE0086
    If UseSrcAsMask Then Xmsk = xSrc: Ymsk = ySrc
    BitBlt DestHDC, X, Y, width, height, MaskHDC, Xmsk, Ymsk, SRCPAINT
    BitBlt DestHDC, X, Y, width, height, SrcHDC, xSrc, ySrc, vbSrcAnd
End Sub

'Initialize all startup values
Public Sub InitPanelDePon(DEST As PictureBox, SRC As PictureBox, Disp As PictureBox, Optional width As Long = 6, Optional height As Long = 10, Optional Size As Long = -1, Optional Tiles As Long = -1, Optional Remains As Long = 32, Optional MaxMoves As Long, Optional GameSpeed As Long = 1, Optional GarbageColors As Long)
    If Size = -1 Then Size = SRC.height / 3
    If Tiles = -1 Then Tiles = width * height * (25 / 60)
    Set Destination = DEST
    Set Source = SRC
    Set Display = Disp
    TileSize = Size
    Colors = Source.width \ TileSize
    GColors = GarbageColors
    doBG = True
    Destination.FillStyle = vbSolid
    Destination.FillColor = Destination.BackColor
    GridWidth = width
    GridHeight = height
    BoardHeight = GridHeight * Size
    BoardWidth = GridWidth * Size
    TileCount = Tiles
    Ghost = Remains
    Speed = 1
    If TileSize \ GameSpeed = TileSize / GameSpeed Then Speed = GameSpeed 'make sure its a multiple of tilesize
    Moves = MaxMoves
    PixelSpeed = 2
    ReDim Grid(0 To height, 1 To width)
    ReDim ThreeDEffect(0 To BoardWidth - 1)
    Seed3D
    Destination.Move Destination.Left, Destination.top, Size * width, Size * height
    Display.Move Display.Left, Display.top, Destination.width, Destination.height
    Reset
End Sub

Public Sub INIT3DMODE()
    If Not ThreeDMode Then
        ThreeDMode = True
        GridWidth = GridWidth * 2
        ReDim Grid(0 To GridHeight, 1 To GridWidth)
        CursorX = GridWidth / 2
        ResetCombos
    End If
End Sub

'Reset Level variables to default
Public Sub Reset()
    If ThreeDMode Then
        ThreeDMode = False
        OffsetX = 0
        InitPanelDePon Destination, Source, Display, GridWidth / 2, GridHeight, TileSize, TileCount, Ghost, Moves, Speed, GColors
        Exit Sub
    End If
    PurgeUndo
    CursorX = 1
    CursorY = 1
    Offset = 0
    Score = 0
    Currtime = 0
    Level = 0
    IncreaseSpeed = False
    Mode = Normal
    doGarbage = 0
    GameOver = True
    GameWon = False
    Clean = False
    RaiseEvent GridCleared
    ResetCombos
    RaiseEvent ScoreChanged(0)
    RaiseEvent MovesChainged(0)
    RaiseEvent LinesChainged(0)
End Sub

'Purge combocounters
Public Sub ResetCombos()
    ReDim Combos(1 To GridWidth)
End Sub

'Purge all undos
Public Sub PurgeUndo()
    UndoCount = 0
    ReDim UndoList(0)
End Sub

'Basic keyboard interpretor
Public Sub Keydown(KeyCode As Integer)
    Select Case KeyCode
        Case 38, 104: Action Moveup
        Case 40, 98: Action MoveDown
        Case 37, 100: Action MoveLeft
        Case 39, 102: Action MoveRight
        Case 13, 32, 101: Action MoveTiles
        Case 9, 20, 33, 34, 96, 107: Action MoveBoard
        Case 8, 46, 85, 90: If Mode = Puzzle Then PullUndo
    End Select
End Sub

'Basic command interpretor
Public Sub Action(Index As PanelAction)
    If GameOver And Mode <> Edit Then Exit Sub
    RaiseEvent ActionPerformed(Index)
    Select Case Index
        Case Moveup: MoveCursor CursorX, CursorY + 1
        Case MoveDown: MoveCursor CursorX, CursorY - 1
        Case MoveLeft: MoveCursor CursorX - 1, CursorY
        Case MoveRight: MoveCursor CursorX + 1, CursorY
        Case MoveBoard: ShiftUp
        Case MoveTiles
            If Mode = Edit Then
                EditTile CursorX, CursorY
            Else
                SwitchTiles CursorX, CursorY
            End If
    End Select
    DrawScreen
End Sub

'Edit a tile (user)
Public Sub EditTile(X As Long, Y As Long, Optional Direction As Long = 1)
    With Grid(Y, X)
        .Color = .Color + Direction
        Select Case .Color
            Case 1: .Color = .Color + 1
            Case Is >= Colors: .Color = 0
            Case Is < 0: .Color = Colors - 1
        End Select
        .Clean = False
        DrawScreen 'Might as well do it automatically
    End With
End Sub

'Convert a map to a string
Public Function SaveCustom() As String
    Dim tempstr As String, temp As Long, temp2 As Long, temp3 As Long
    tempstr = GridWidth & "," & GridHeight & "," & Mode & "," & Moves & "," & Lines & "," & Score & "," & PixelSpeed & vbNewLine
    For temp = 1 To GridWidth
        For temp2 = 0 To GridHeight
            temp3 = Grid(temp2, temp).Color
            If temp3 = 1 Then temp3 = 0
            tempstr = tempstr & temp3 & IIf(temp2 < GridHeight, ",", Empty)
        Next
        tempstr = tempstr & IIf(temp < GridWidth, vbNewLine, Empty)
    Next
    SaveCustom = tempstr
End Function

'Push the current map onto the undo list
Public Sub PushUndo()
    UndoCount = UndoCount + 1
    ReDim Preserve UndoList(UndoCount)
    UndoList(UndoCount - 1) = SaveCustom
End Sub

'Pull an undo from the undo list
Public Sub PullUndo()
    If UndoCount > 0 Then
        LoadCustom UndoList(UndoCount - 1)
        UndoCount = UndoCount - 1
        ReDim Preserve UndoList(UndoCount)
    End If
End Sub

'Count how many tiles there are on the map
Public Function CountTiles() As Long
    Dim temp As Long, temp2 As Long, temp3 As Long
    For temp = 1 To GridWidth
        For temp2 = 1 To GridHeight
            If Grid(temp2, temp).Color > 1 Then temp3 = temp3 + 1
        Next
    Next
    CountTiles = temp3
End Function

'Convert a string to a map, returns true if successful
Public Function LoadCustom(MAP As String) As Boolean
    On Error Resume Next
    Dim tempstr() As String, tempstr2() As String, temp As Long, temp2 As Long
    tempstr = Split(MAP, vbNewLine)
    GameOver = False
    tempstr2 = Split(tempstr(0), ",")
    GridWidth = Val(tempstr2(0))
    GridHeight = Val(tempstr2(1))
    Mode = Val(tempstr2(2))
    Moves = Val(tempstr2(3))
    Lines = Val(tempstr2(4))
    Score = Val(tempstr2(5))
    PixelSpeed = Val(tempstr2(6))
    ReDim Grid(0 To GridHeight, 1 To GridWidth)
    
    For temp = 1 To GridWidth
        tempstr2 = Split(tempstr(temp), ",")
        For temp2 = 0 To GridHeight
            With Grid(temp2, temp)
                .Color = Val(tempstr2(temp2))
                .Clean = False
            End With
        Next
    Next
    
    'ADD SOLUTION LOADING
    
    Clean = False
    LoadCustom = True
    RaiseEvent MapLoaded
    RaiseEvent MovesChainged(Moves)
    RaiseEvent LinesChainged(Lines)
    DrawScreen
End Function

'Draws the entire grid and both cursors
Public Sub DrawScreen()
    Dim temp As Long, temp2 As Long
    
    If ThreeDMode Then
        Destination.Cls
        DrawBackground
        DrawScreen3D
        Clean = True
    Else
        If Not doBG Then ClearArea
        If Not Clean Or doBG Then Destination.Cls
        If doBG Then DrawBackground
        For temp = 1 To GridWidth
            For temp2 = 0 To GridHeight
                DrawGrid temp, temp2
            Next
        Next
        DrawCursor CursorX, CursorY
        If Mode <> Edit Then
            DrawCursor CursorX + 1, CursorY
            If Offset > 0 Then Shade -Offset, BoardHeight - Offset, Destination.width, Offset
            If Mode = Lineclear And Lines > 0 Then
                temp2 = GetLine
                Destination.Line (0, temp2)-(Destination.width, temp2), vbRed, B
            End If
        End If
        Clean = Not doBG
    End If
    'BitBlt Display, 0, 0, BoardWidth, BoardHeight, Destination, 0, 0, vbSrcCopy
    Set Display.Picture = Destination.Image
    Display.Refresh
End Sub

'Get the Y coord of the line
Public Function GetLine() As Long
    GetLine = Destination.height - Lines * TileSize - Offset
End Function

'Draws a cursor on to destination
Private Sub DrawCursor(X As Long, Y As Long)
    Dim temp As Long
    temp = BoardHeight - ((Y * TileSize)) - Offset
    If ThreeDMode Then temp = BoardHeight - ((Y * TileSize) + Offset) + ThreeDEffect(BoardWidth / 2)
    DrawTile 0, (X - 1) * TileSize, temp
End Sub

'Draws a tile on the grid
Private Sub DrawGrid(X As Long, Y As Long)
    With Grid(Y, X)
        If .Clean = False Or Clean = False Then
            If (.X <> 0 Or .Y <> 0) And .Color = 0 Then Exit Sub
            If .Color = 0 And doBG Then Exit Sub
            If .Color = 1 And isEven(.height) Then
                DrawEffect .width, (X - 1) * TileSize + .X, BoardHeight - (Y * TileSize) - .Y - Offset
            Else
                DrawTile .Color, (X - 1) * TileSize + .X, BoardHeight - (Y * TileSize) - .Y - Offset, True, .width, .height
                .Clean = True
            End If
        End If
    End With
End Sub

Private Sub DrawEffect(ByVal Color As Long, X As Long, Y As Long)
    If Color < 1 Then Color = 1
    BitBlt Destination.hDC, X, Y, TileSize, TileSize, Source.hDC, Color * TileSize, TileSize, vbSrcCopy
End Sub

Public Function isEven(Number As Long) As Boolean
    isEven = (Number And 1) = 0
End Function

'Used for mouse control of the cursor
Public Sub MouseMove(ByVal X As Single, ByVal Y As Single)
    Static OLDX As Long, hasINIT As Boolean
    Dim temp As Long, newX As Long
    If Not GameOver And TileSize > 0 Then
        If ThreeDMode Then
            If Not hasINIT Then
                OLDX = X
                hasINIT = True
            End If
            temp = X - OLDX
            newX = CursorX
            If temp > TileSize / 2 Then newX = newX + 1: OLDX = X
            If temp < -TileSize / 2 Then newX = newX - 1: OLDX = X
            MoveCursor newX, GridHeight - ((Y + Offset) \ TileSize)
        Else
            MoveCursor X \ TileSize + 1, GridHeight - ((Y + Offset) \ TileSize)
        End If
    End If
End Sub

'If the new position is not the same as the old one, dirty the tile under the current cursor position, then move the cursor
Public Sub MoveCursor(X As Long, Y As Long)
    If X < 1 And Not ThreeDMode Then X = 1
    If Mode = Edit Then
        If X > GridWidth Then X = GridWidth
    Else
        If ThreeDMode = True Then
            If Not (X = GridWidth And CursorX = 1) Then
                X = GetCol(X - CursorX)
            End If
            'If X <= 1 Then X = GridWidth
            'If X >= GridWidth - 1 Then X = 1
        Else
            If X > GridWidth - 1 Then X = GridWidth - 1
        End If
    End If
    If Y < 1 Then Y = 1
    If Y > GridHeight Then Y = GridHeight
    
    If X <> CursorX Or Y <> CursorY Then
        Grid(CursorY, CursorX).Clean = False

        If Mode <> Edit And Not ThreeDMode Then Grid(CursorY, CursorX + 1).Clean = False
        DrawGrid CursorX, CursorY
        If Mode <> Edit And Not ThreeDMode Then DrawGrid CursorX + 1, CursorY
        
        If ThreeDMode And X <> CursorX Then
            If X <= 1 And CursorX <> X + 1 Then
                CursorX = GridWidth
                OffsetX = -TileSize
            Else
                If X = CursorX + 1 Then OffsetX = -TileSize
                If X = CursorX - 1 Then OffsetX = TileSize
                If OffsetX = 0 And X = GridWidth Then OffsetX = TileSize
                'If X = 1 And CursorX = GridWidth Then OffsetX = -TileSize
                'If X = GridWidth And CursorX = 1 Then OffsetX = TileSize
            End If
        End If
        
        CursorX = X
        CursorY = Y
        RaiseEvent CursorMoved(X, Y)
    End If
End Sub

'Place random tiles
Public Sub InitGrid(Optional doReset As Boolean = True)
    Dim temp As Long
    If doReset Then Reset
    GameOver = False
    ClearGrid
    For temp = 1 To GridWidth
        ClearTile temp, 0, True
    Next
    For temp = 1 To TileCount
        PlaceRandomTile
    Next
    RaiseEvent MapLoaded
    DrawScreen
End Sub

'Clear all tiles
Public Sub ClearGrid()
    Dim temp As Long, temp2 As Long
    For temp = 1 To GridWidth
        For temp2 = 0 To GridHeight
            ClearTile temp, temp2
        Next
    Next
    RaiseEvent GridCleared
End Sub

'Shift the board up a given amount of pixels. If the offset matches the tilesize than shift everything up one row
'In line mode, the time is increased
Public Sub IncreaseOffset(Optional Pixels As Long = 1)
    If GameOver Or Locked Then Exit Sub
    RaiseEvent OffsetIncreased
    If Mode = Puzzle Then
        Currtime = Currtime + 1
    Else
        If Mode = Timelimit Then
            Currtime = Currtime - 1
            If Currtime = 0 Then
                RaiseEvent PDPGameOver
                GameOver = True
            End If
        End If
        
        If Stopper <= 0 Then
            Stopper = 0
            Offset = Offset + Pixels
            If Offset < TileSize Then
                Clean = False
            Else
                ShiftUp
                MoveCursor CursorX, CursorY + 1
                Offset = 0
            End If
        Else
            Stopper = Stopper - Pixels
        End If
    End If
End Sub

'Clears the screen for pausing
Public Sub Clearscreen()
    Clean = False
    Destination.Cls
    Set Display.Picture = LoadPicture(Empty)
    Display.Refresh
End Sub

'Shift each of the tiles that are in movement
Public Function DecrementOffsets(Optional Pixels As Long = -1) As Boolean
    If GameOver Then Exit Function
    If Pixels = -1 Then Pixels = PixelSpeed
    If Pixels = 0 Then Pixels = 2
    Dim temp As Long, temp2 As Long, temp3 As Long, topwaschanged As Boolean, Y As Long, Tiles As Long, DecrementedOffsets As Boolean, maxcombos As Long
    DisableCombos
    Locked = False
    maxcombos = 0
    If OffsetX <> 0 Then
        Clean = False
        If OffsetX < 0 Then OffsetX = OffsetX + Abs(Pixels)
        If OffsetX > 0 Then OffsetX = OffsetX - Abs(Pixels)
    End If
    For temp = 1 To GridWidth
        For temp2 = 1 To GridHeight 'the buffer row cant be moved by the user and thus shouldnt have any offsets
            With Grid(temp2, temp)
                If .X <> 0 Or .Y <> 0 Or .Z <> 0 Then
                    DecrementedOffsets = True
                    If .Color = 0 Then
                        DropColumn temp, temp2 + 1
                        If .X < 0 Then .X = .X + Pixels
                        If .X > 0 Then .X = .X - Pixels
                        If .Y > 0 Then .Y = .Y - Pixels
                        .Clean = False
                    Else
                        If .Y > 0 And .X = 0 Then
                            AddCombo temp, Enable
                            .Y = .Y - Pixels
                            If .Y <= 0 Then
                                .Y = 0
                                If temp2 < GridHeight Then
                                    If Grid(temp2 + 1, temp).Color = 0 Then Grid(temp2 + 1, temp).Clean = False
                                End If
                                Tiles = Tiles + CheckTile(temp, temp2)
                            End If
                        End If
                        
                        'New tile switching effect
                        If .X <> 0 Then
                            Y = BoardHeight - (temp2 * TileSize) - Offset
                            If .X < 0 Then 'moving right
                                .X = .X + Pixels
                                If temp > 1 Then
                                    If doBG And Grid(temp2, temp - 1).Color > 1 Then .Y = Y - DegreesToXY(((temp - 1) * TileSize), Y, (360 / TileSize) * Abs(.X), TileSize, TileSize / 3, False)
                                Else
                                    If ThreeDMode And Grid(temp2, GridWidth).Color > 1 Then .Y = Y - DegreesToXY(((temp - 1) * TileSize), Y, (360 / TileSize) * Abs(.X), TileSize, TileSize / 3, False)
                                End If
                            End If
                            If .X > 0 Then 'moving left
                                .X = .X - Pixels
                                If .X = 0 Then
                                    .Y = 0
                                Else
                                    If temp < GridWidth Then
                                        If doBG And Grid(temp2, temp + 1).Color > 1 Then .Y = DegreesToXY((temp) * TileSize, Y, 360 - ((360 / TileSize) * .X), TileSize, TileSize / 3, False) - Y
                                    Else
                                        If ThreeDMode And Grid(temp2, 1).Color > 1 Then .Y = DegreesToXY((temp) * TileSize, Y, 360 - ((360 / TileSize) * .X), TileSize, TileSize / 3, False) - Y
                                    End If
                                End If
                            End If
                            If .X = 0 Then
                                Tiles = Tiles + CheckTile(temp, temp2)
                            End If
                        End If
                        
                        If .Z <> 0 Then Locked = True
                        If .Z > 0 Then
                            .Z = .Z - 1
                            If .height > 0 Then .height = .height - 1
                            If .Z = 0 Then
                                .Color = 0
                                temp3 = temp3 + 1
                                topwaschanged = True
                                DropColumn temp, temp2 + 1
                            End If
                            AddCombo temp, Enable 'Combos(temp).Used = True
                        End If
                        
                        If .Z < 0 Then
                            .Z = .Z + 1
                            'If Z = <<>> Then ClearTile temp, temp2, True
                            If .Z = 0 Then Tiles = Tiles + CheckTile(temp, temp2) 'DropTile temp, temp2
                        End If
                        
                        .Clean = False
                    End If
                End If
            End With
        Next
        If AddCombo(temp, None) = False Then ' Combos(temp).Used = False Then
            If Combos(temp).value > 1 Then
                Score = Score + ChainScore(Combos(temp).value)
                RaiseEvent ComboMade(Combos(temp).value)
                RaiseEvent ScoreChanged(Score)
            End If
            AddCombo temp, Clear 'Combos(temp).value = 0
        End If
        If Combos(temp).value > maxcombos Then maxcombos = Combos(temp).value
    Next
    If topwaschanged Then
        CheckLineWon
        RaiseEvent TopChanged(top)
    End If
    RaiseEvent OffsetsDecremented
    If Tiles > 0 Then
        Score = Score + ComboScore(Tiles)
        RaiseEvent KillTiles(Tiles)
        RaiseEvent ScoreChanged(Score)
    End If
    DecrementOffsets = DecrementedOffsets
    If Not DecrementedOffsets And Mode = Puzzle Then
        If CountTiles = 0 Then
            GameWon = True
            GameOver = True
            RaiseEvent PDPGameWon
        Else
            If Moves <= 0 Then
                GameOver = True
                RaiseEvent PDPGameOver
            End If
        End If
    End If
    If maxcombos > 1 Then RaiseEvent ChainInProgress(maxcombos)
End Function

'Shift all tiles up one row. If the top row is filled already, game over. Create random tiles in the buffer/bottom row
Public Sub ShiftUp()
    Dim temp As Long, temp2 As Long, tempstr As String, Tiles As Long
    For temp = 1 To GridWidth
        If Bottom(temp) = GridHeight Then
            GameOver = True
            RaiseEvent PDPGameOver
            Exit Sub
        Else
            If Offset < TileSize Then Clean = False
            For temp2 = GridHeight - 1 To 0 Step -1
                Grid(temp2 + 1, temp) = Grid(temp2, temp)
            Next
            Tiles = Tiles + CheckTile(temp, 1)
            ClearTile temp, 0, True
            tempstr = tempstr & IIf(Len(tempstr) > 0, ",", Empty) & Grid(0, temp).Color
        End If
    Next
    RaiseEvent TopChanged(top)
    RaiseEvent TopRaised
    RaiseEvent NewRow(tempstr)
    Lines = Lines + 1
    RaiseEvent LinesChainged(Lines)
    If Mode = Lineclear Then
        'Lines = Lines - 1
        CheckLineWon
    Else
        'Lines = Lines + 1
        IncreaseSpeed = Getlevel(Level) = Lines
        If IncreaseSpeed Then
            RaiseEvent RaiseLevel
            Level = Level + 1
        End If
    End If
    If doGarbage > 0 Then CreateRandomGarbage
    Offset = 0
    If Tiles > 0 Then
        RaiseEvent KillTiles(Tiles)
        Score = Score + ComboScore(Tiles)
        RaiseEvent ScoreChanged(Score)
    End If
End Sub

'checks if the user won line mode
Public Sub CheckLineWon()
    If Mode = Lineclear Then
        If Lines >= 0 Then
            If top <= Lines Then
                GameWon = True
                GameOver = True
                RaiseEvent PDPGameWon
            End If
        End If
    End If
End Sub

'Used for speed enhancing
Public Function Getlevel(Optional Level As Long) As Long
    Getlevel = 2 ^ Level 'Took the easy way out and did exponential growth
End Function

'Place a random tile
Private Sub PlaceRandomTile()
    Dim X As Long, Y As Long, Color As Long, temp As Boolean
    Do Until temp
        X = Rnd * (GridWidth - 1) + 1
        Y = Bottom(X)
        Color = RandomTile
        If Y <= GridHeight \ 2 Then
            Grid(Y, X).Color = Color
            temp = Not IsAScore(X, Y, Color)
            If Not temp Then Grid(Y, X).Color = 0
        End If
    Loop
    'MsgBox "Creating a " & Color2String(Color) & " tile at " & X & ", " & Y
    ClearTile X, Y 'Clear it out just in case
    Grid(Y, X).Color = Color
    DrawScreen
End Sub

Private Function Color2String(Color As Long) As String
    Dim tempstr() As String
    tempstr = Split("Cursor,Placeholder,Black,Green,Blue,Red", ",")
    Color2String = tempstr(Color)
End Function

'Switch a tile with the one beside it IF they arent moving
Public Sub SwitchTiles(X As Long, Y As Long)
    Dim temp As Tile, temp2 As Long
    temp2 = X + 1
    If temp2 > GridWidth Then temp2 = 1
    If Grid(Y, X).X = 0 And Grid(Y, X).Y = 0 And Grid(Y, temp2).X = 0 And Grid(Y, temp2).Y = 0 Then
        If Grid(Y, X).Color = 0 And Grid(Y, temp2).Color = 0 Then Exit Sub
        If IsGarbage(X, Y) Or IsGarbage(temp2, Y) Then Exit Sub
        If Mode = Puzzle Then
            PushUndo
            Moves = Moves - 1
        Else
            If Grid(Y, X).Color = 1 Or Grid(Y, temp2).Color = 1 Then Exit Sub
            Moves = Moves + 1
        End If
        RaiseEvent MovesChainged(Moves)
        temp = Grid(Y, X)
        Grid(Y, X) = Grid(Y, temp2)
        Grid(Y, temp2) = temp
        Grid(Y, X).Clean = False
        Grid(Y, X).X = TileSize
        Grid(Y, temp2).Clean = False
        Grid(Y, temp2).X = -TileSize
    End If
End Sub

'If the tile is not the lowest or moving already it can go, drop it down one row, and set its offset along Y to 32 so it gradually comes down
Public Sub DropTile(X As Long, Y As Long)
    If Bottom(X, Y) < Y And (Grid(Y, X).Color < 0 Or Grid(Y, X).Color > 1) And Grid(Y, X).X = 0 Then
        If Grid(Y, X).Color < 0 Then If Not CanDropGarbage(X, Y) Then Exit Sub 'The the tile is a garbage block, and cant fall, exit this sub
        If Grid(Y - 1, X).Color <> 0 Or (IsGarbage(X, Y - 1) And Not IsGarbage(X, Y)) Then Exit Sub 'if the tile underneath is not blank, or (is occupied by garbage, that is not part of itself [if the source is garbage]) then exit sub
        If Grid(Y, X).Color > 1 And Grid(Y, X).Z <> 0 Then Exit Sub
        Grid(Y - 1, X) = Grid(Y, X)
        ClearTile X, Y
        Grid(Y - 1, X).Y = TileSize
        Grid(Y - 1, X).Clean = False
    End If
End Sub

'Clears a tiles properties at x,y. If Newtile = true, a random tile is created
Private Function ClearTile(X As Long, Y As Long, Optional Newtile As Boolean) As Long
    With Grid(Y, X)
        .Color = 0
        If Newtile Then '
            .Color = RandomTile
            Do Until CountX(X, Y, .Color) < Adjacent
                .Color = RandomTile
            Loop
            ClearTile = .Color
        End If
        .X = 0
        .Y = 0
        .Z = 0
        .width = 0
        .height = 0
        .Clean = False
    End With
End Function

'Finds the first empty spot in a column (0 is used as a buffer)
Public Function Bottom(X As Long, Optional ByVal Y As Long) As Long
    Dim temp As Long
    'pregarbage
    'For temp = 1 To GridHeight
    '    If Grid(temp, X).Color < 2 Or temp = GridHeight Then
    '        Bottom = temp
    '        Exit For
    '    End If
    'Next
    
    'Post garbage
    If Y = 0 Then Y = GridHeight
    If IsGarbage(X, Y) Then Y = Y - (GarbageTile(X, Y).height) + 1
    For temp = Y - 1 To 0 Step -1
        If IsGarbage(X, temp) Or Grid(temp, X).Color > 1 Or temp = 0 Then
            Bottom = temp + 1
            Exit For
        End If
    Next
End Function
    
'Draws a tile from source into destination. If 0 is ignored (you dont want to draw a cursor) and index = 0 a blank square is drawn
Private Sub DrawTile(Index As Long, X As Long, Y As Long, Optional ignorezero As Boolean, Optional width As Long, Optional height As Long)
    If Index = 0 And ignorezero Then
        DrawSquare X, Y
    Else
        If Index < 0 Then
            DrawGarbage X, Y, width, height, , Index
        Else
            If Index = 0 Then
                TransBLT Source.hDC, 0, 0, Source.hDC, 0, TileSize, TileSize, TileSize, Destination.hDC, X, Y
            Else
                BitBlt Destination.hDC, X, Y, TileSize, TileSize, Source.hDC, Index * TileSize, 0, vbSrcCopy
            End If
        End If
    End If
End Sub

'Draws a blank square
Private Sub DrawSquare(X As Long, Y As Long)
    Destination.Line (X, Y)-(X + TileSize - 1, Y + TileSize - 1), Destination.BackColor, B
End Sub

'Draws a line
Private Sub DrawLine(X1 As Long, Y1 As Long, X2 As Long, Y2 As Long, Optional Color As OLE_COLOR = vbBlack)
    Destination.Line (X1, Y1)-(X2, Y2), Color
End Sub

'Dither/shade
Private Sub Shade(X As Long, Y As Long, width As Long, height As Long)
    Dim temp As Long, X1 As Long, Y1 As Long, X2 As Long, Y2 As Long
    For temp = X - height To X + width + height Step 2 'Middle
        'drawaline temp, Y, Width, Height
        X1 = temp
        Y1 = Y
        
        If X1 < X Then
            X1 = X
            Y1 = Y1 + (X - X1)
        End If
        
        X2 = X1 + height
        Y2 = Y1 + height
        
        If Y2 > Y + height Then
            X2 = X2 - (Y2 - (Y + height - 1))
            Y2 = Y + height - 1
        End If
        
        DrawLine X1, Y1, X2, Y2
    Next
End Sub

'Chooses a random tile color (0 is the cursor, 1 is the placeholder)
Public Function RandomTile() As Long
    Randomize Timer
    RandomTile = (Rnd * (Colors - 3)) + 2
End Function

'Drop a tiles above a certain tile
Public Sub DropColumn(X As Long, ByVal Y As Long)
    For Y = Y To GridHeight
        If Grid(Y, X).Color > 1 Then DropTile X, Y
        IsGarbage X, Y, Drop
    Next
End Sub

'Erase the area under moving tiles
Public Sub ClearArea()
    Dim temp As Long, temp2 As Long
    For temp = 1 To GridWidth
        For temp2 = 1 To GridHeight
            With Grid(temp2, temp)
                If (.X <> 0 Or .Y <> 0) And .Color > 1 Then
                    DrawSquare (temp - 1) * TileSize, BoardHeight - (temp2 * TileSize) - Offset
                    'Clear garbage from moving blocks beside blank ones
                    If temp < GridWidth Then If .X > 0 And Grid(temp2, temp + 1).Color = 0 Then DrawSquare (temp) * TileSize, BoardHeight - (temp2 * TileSize) - Offset
                    If temp > 1 Then If .X < 0 And Grid(temp2, temp - 1).Color = 0 Then DrawSquare (temp - 2) * TileSize, BoardHeight - (temp2 * TileSize) - Offset
                    If temp2 < GridHeight Then If .Y > 0 And Grid(temp2 + 1, temp).Color = 0 Then DrawSquare (temp - 1) * TileSize, BoardHeight - ((temp2 + 1) * TileSize) - Offset
                End If
            End With
        Next
    Next
End Sub

'If the tile hasnt reached the bottom yet, drop it, otherwise check for a scoring move
Public Function CheckTile(X As Long, Y As Long) As Long
    Dim pLeft As Long, pRight As Long, pTop As Long, pBottom As Long, Color As Long, HOR As Boolean, VER As Boolean, Left2 As Long
    Dim temp2 As Long
    Color = Grid(Y, X).Color
    
    If (Color = 0) Or (Bottom(X, Y) < Y) Then
        DropColumn X, Y ' DropTile temp, temp2
    Else
        pLeft = StartX(X, Y, Color)
        pRight = EndX(X, Y, Color)
        HOR = pRight - pLeft + 1 > Adjacent
        
        pTop = StartY(X, Y, Color)
        pBottom = EndY(X, Y, Color)
        VER = pBottom - pTop + 1 > Adjacent
        
        
        If ThreeDMode And (pLeft = 1 Or pRight = GridWidth) Then
            temp2 = GhostCheck(Y, Color)
        Else
            If HOR Then temp2 = GhostX(pLeft, Y, pRight)
        End If
                
        If VER Then temp2 = temp2 + GhostY(X, pTop, pBottom)
        CheckTile = temp2
    End If
    
    Select Case Mode
        Case Lineclear: GameWon = CountTiles = 0 And Lines = 0
        Case Puzzle: GameWon = CountTiles = 0
    End Select
End Function

Private Function GhostCheck(Y As Long, Color As Long, Optional doGhost As Boolean = True) As Long
        Dim LS As Long, RS As Long, temp As Long, temp2 As Long
        LS = CheckXLeft(Y, Color, True)
        RS = CheckXLeft(Y, Color)
        If LS + RS > Adjacent Then
            GhostCheck = LS + RS
            If doGhost Then
                temp2 = Ghost
                For temp = GridWidth - RS + 1 To GridWidth
                    AddCombo temp, AddOne
                    GhostTile temp, Y, temp2
                    temp2 = temp2 + Ghost
                Next
                For temp = 1 To LS
                    AddCombo temp, AddOne
                    GhostTile temp, Y, temp2
                    temp2 = temp2 + Ghost
                Next
            End If
        End If
End Function

Private Function CheckXLeft(Y As Long, Color As Long, Optional RS As Boolean) As Long
    Dim temp As Long, temp2 As Long
    Dim Start As Long, Finish As Long, Inc As Long
    Start = 1
    Finish = 1
    Inc = 1
    If RS Then
        Finish = GridWidth
    Else
        Start = GridWidth
        Inc = -1
    End If
    For temp = Start To Finish Step Inc
        If Color <> Grid(Y, temp).Color Then
            CheckXLeft = temp2
            Exit Function
        End If
        temp2 = temp2 + 1
    Next
End Function

'Ghosts a horizontal Score
Private Function GhostX(X As Long, Y As Long, X2 As Long) As Long
    Dim temp As Long, temp2 As Long
    If Not X = 1 And ThreeDMode Then
        For temp = X + 1 To X2 'Attempt to root out random tile kidnapping
            If Grid(Y, temp).Color <> Grid(Y, X).Color And Grid(Y, temp).Color <> 1 And 1 <> Grid(Y, X).Color Then Exit Function
        Next
    End If
    temp2 = Ghost
    temp = X2 - X + 1
    GhostX = temp
    If doScore(temp) Or ThreeDMode Then
        For temp = X To X2
            AddCombo temp, AddOne
            GhostTile temp, Y, temp2
            temp2 = temp2 + Ghost
        Next
    End If
End Function

'Ghosts a vertical Score
Private Function GhostY(X As Long, Y As Long, Y2 As Long) As Long
    Dim temp As Long, temp2 As Long
    For temp = Y + 1 To Y2 'Attempt to root out random tile kidnapping
        If Grid(temp, X).Color <> Grid(Y, X).Color And Grid(temp, X).Color <> 1 And 1 <> Grid(Y, X).Color Then Exit Function
    Next
    AddCombo X, AddOne
    temp = Y2 - Y + 1
    GhostY = temp
    temp2 = Ghost
    If doScore(temp) Then
        For temp = Y To Y2
            GhostTile X, temp, temp2
            temp2 = temp2 + Ghost
        Next
    End If
End Function

'Auto do score, remove less than 3 accidental clears
Public Function doScore(Tiles As Long) As Boolean
    doScore = Tiles > Adjacent
End Function

'Ghosts a single tile
Private Sub GhostTile(X As Long, Y As Long, Optional GhostTime As Long)
    If X > 0 And X <= GridWidth And Y >= 0 And Y <= GridHeight Then 'AHAHAH, stupidly put gridwidth in the Y checked
        With Grid(Y, X)
            If .Color > 1 Then .width = .Color
            .height = Ghost / 2
            
            .Color = 1
            .Clean = False
            .Z = GhostTime
            
            'Destroy adjacent garbage blocks
            IsGarbage X + 1, Y, destroy
            IsGarbage X - 1, Y, destroy
            IsGarbage X, Y + 1, destroy
            IsGarbage X, Y - 1, destroy
        End With
    End If
End Sub

'Returns the highest block
Public Function top() As Long
    Dim temp As Long, temp2 As Long, temp3 As Long
    For temp = 1 To GridWidth
        temp2 = Bottom(temp) - 1
        If temp2 > temp3 Then temp3 = temp2
    Next
    top = temp3
End Function

'Find the start of concurrent tiles in a row
Private Function StartX(X As Long, Y As Long, Optional Color As Long = -1) As Long
    Dim temp As Long
    If Color = -1 Then Color = Grid(Y, X).Color
    For temp = X - 1 To 1 Step -1
        If Grid(Y, temp).Color <> Color Then
            StartX = temp + 1
            Exit Function
        End If
    Next
    StartX = 1
End Function

'Find the start of concurrent tiles in a column
Private Function StartY(X As Long, Y As Long, Optional Color As Long = -1) As Long
    Dim temp As Long
    If Color = -1 Then Color = Grid(Y, X).Color
    For temp = Y - 1 To 1 Step -1
        If Grid(temp, X).Color <> Color Then
            StartY = temp + 1
            Exit Function
        End If
    Next
    StartY = 1
End Function

'Find the end of concurrent tiles in a row
Private Function EndX(X As Long, Y As Long, Optional Color As Long = -1) As Long
    Dim temp As Long
    If Color = -1 Then Color = Grid(Y, X).Color
    For temp = X + 1 To GridWidth
        If Grid(Y, temp).Color <> Color Then
            EndX = temp - 1
            Exit Function
        End If
    Next
    EndX = GridWidth
End Function

'Find the end of concurrent tiles in a column
Private Function EndY(X As Long, Y As Long, Optional Color As Long = -1) As Long
    Dim temp As Long
    If Color = -1 Then Color = Grid(Y, X).Color
    For temp = Y + 1 To GridHeight
        If Grid(temp, X).Color <> Color Then
            EndY = temp - 1
            Exit Function
        End If
    Next
    EndY = GridHeight
End Function

'Find the number of concurrent tiles in a row
Private Function CountX(X As Long, Y As Long, Optional Color As Long = -1) As Long
    Dim temp As Long, temp2 As Long
    temp = StartX(X, Y, Color)
    temp2 = EndX(temp, Y, Color)
    If ThreeDMode And temp = 1 Then
        CountX = GhostCheck(Y, Color, False)
    Else
        CountX = temp2 - temp + 1
    End If
End Function

'Find the number of concurrent tiles in a column
Private Function CountY(X As Long, Y As Long, Optional Color As Long = -1) As Long
    Dim temp As Long, temp2 As Long
    temp = StartY(X, Y, Color)
    temp2 = EndY(X, temp, Color)
    CountY = temp2 - temp + 1
End Function

'Check if the resulting tile placement will result in a score
Public Function IsAScore(X As Long, Y As Long, Color As Long) As Boolean
    IsAScore = Color > 0 And (CountX(X, Y, Color) > Adjacent Or CountY(X, Y, Color) > Adjacent)
End Function

'Load a file and return its contents as a string
Public Function LoadWholeFile(Filename As String) As String
    On Error Resume Next
    If FileLen(Filename) = 0 Then Exit Function
    Dim temp As Long, tempstr As String, tempstr2 As String
    temp = FreeFile
    If Dir(Filename) <> Filename Then
        Open Filename For Input As temp
            Do Until EOF(temp)
                Line Input #temp, tempstr
                If Len(tempstr) = 0 Then
                    tempstr2 = tempstr2 & vbNewLine
                Else
                    tempstr2 = tempstr2 & tempstr & IIf(Len(tempstr) > 0, vbNewLine, Empty)
                End If
            Loop
            LoadWholeFile = tempstr2
        Close temp
    End If
End Function

'Save a string as a file
Public Function SaveWholeFile(Filename As String, Optional text As String) As Boolean
    On Error Resume Next
    Dim tempfile As Long
    tempfile = FreeFile
    Open Filename For Output As tempfile
        Print #tempfile, text
    Close tempfile
    SaveWholeFile = True
End Function

'Drop all Tiles on the screen (perform after a load)
Public Sub DropAll()
    Dim temp As Long
    For temp = 1 To GridWidth
        DropColumn temp, 1
    Next
End Sub

'Loads a remote row into the buffer
Public Sub LoadRow(Row As String)
    Dim tempstr() As String, temp As Long
    tempstr = Split(Row, ",")
    For temp = 0 To UBound(tempstr)
        ClearTile temp + 1, 0
        Grid(0, temp + 1).Color = Val(tempstr(temp))
    Next
End Sub

'Used for the AI to read the board
Public Function GetColor(X As Long, Y As Long) As Long
    Dim temp As Long
    temp = Grid(Y, X).Color
    If temp = 0 Then If IsGarbage(X, Y) Then temp = -1 'doesnt reallymatter which garbage color does it?
    GetColor = temp
End Function

'Draws a background from the offset
Public Sub DrawBackground()
    Dim temp As Long, temp2 As Long
    With Destination
        For temp = Offset - TileSize To .width Step TileSize
            For temp2 = Offset - TileSize To .height Step TileSize
                BitBlt Destination.hDC, temp, temp2, TileSize, TileSize, Source.hDC, 0, TileSize * 2, vbSrcCopy
            Next
        Next
    End With
End Sub

'Draws a garbage tile
Public Sub DrawGarbage(X As Long, Y As Long, width As Long, height As Long, Optional Size As Long = 6, Optional ByVal Color As Long = -1)
    Dim leftside As Long, top As Long
    top = TileSize * 2
    leftside = Abs(Color) * TileSize
    
    BitBlt Destination.hDC, X, Y, Size, Size, Source.hDC, leftside, top, vbSrcCopy  'Upper left corner
    BitBlt Destination.hDC, X, Y + (height * TileSize) - Size - 1, Size, Size, Source.hDC, leftside, top + TileSize - Size, vbSrcCopy 'upper left corner
    BitBlt Destination.hDC, X + (width * TileSize) - Size - 1, Y, Size, Size, Source.hDC, leftside + TileSize - Size, top, vbSrcCopy  'Upper right corner
    BitBlt Destination.hDC, X + (width * TileSize) - Size - 1, Y + (height * TileSize) - Size - 1, Size, Size, Source.hDC, leftside + TileSize - Size, top + TileSize - Size, vbSrcCopy 'lower right corner
    
    TileBoth Source.hDC, leftside + Size, top + Size, TileSize - Size * 2, TileSize - Size * 2, Destination.hDC, X + Size, Y + Size, width * TileSize - Size * 2, height * TileSize - Size * 2      'center center
    
    TileHorizontal Source.hDC, leftside + Size, top, TileSize - Size * 2, Size, Destination.hDC, X + Size, Y, width * TileSize - Size * 2   'upper middle
    TileHorizontal Source.hDC, leftside + Size, top + TileSize - Size, TileSize - Size * 2, Size, Destination.hDC, X + Size, Y + (height * TileSize) - 1 - Size, width * TileSize - Size * 2  'lower middle
    
    TileVertical Source.hDC, leftside, top + Size, Size, TileSize - Size * 2, Destination.hDC, X, Y + Size, height * TileSize - Size * 2 'left middle
    TileVertical Source.hDC, leftside + TileSize - Size, top + Size, Size, TileSize - Size * 2, Destination.hDC, X + (width * TileSize) - 1 - Size, Y + Size, height * TileSize - Size * 2 'right middle
End Sub

'tile horizontal
Public Sub TileHorizontal(SRC As Long, xSrc As Long, ySrc As Long, WidthSrc As Long, HeightSrc As Long, DEST As Long, Xdest As Long, Ydest As Long, WidthDest As Long)
    Dim temp As Long
    For temp = Xdest To Xdest + WidthDest - 1 Step WidthSrc
        If temp + WidthSrc - 1 <= Xdest + WidthDest - 1 Then
            BitBlt DEST, temp, Ydest, WidthSrc, HeightSrc, SRC, xSrc, ySrc, vbSrcCopy
        Else
            BitBlt DEST, temp, Ydest, (Xdest + WidthDest - 1) - temp, HeightSrc, SRC, xSrc, ySrc, vbSrcCopy
        End If
    Next
End Sub

'tile Vertical
Public Sub TileVertical(SRC As Long, xSrc As Long, ySrc As Long, WidthSrc As Long, HeightSrc As Long, DEST As Long, Xdest As Long, Ydest As Long, HeightDest As Long)
    Dim temp As Long
    For temp = Ydest To Ydest + HeightDest - 1 Step HeightSrc
        If temp + HeightSrc - 1 <= Ydest + HeightDest - 1 Then
            BitBlt DEST, Xdest, temp, WidthSrc, HeightSrc, SRC, xSrc, ySrc, vbSrcCopy
        Else
            BitBlt DEST, Xdest, temp, WidthSrc, (Ydest + HeightDest - 1) - temp, SRC, xSrc, ySrc, vbSrcCopy
        End If
    Next
End Sub

'tile horizontal and vertical
Public Sub TileBoth(SRC As Long, xSrc As Long, ySrc As Long, WidthSrc As Long, HeightSrc As Long, DEST As Long, Xdest As Long, Ydest As Long, WidthDest As Long, HeightDest As Long)
    Dim temp As Long
    For temp = Ydest To Ydest + HeightDest - 1 Step HeightSrc
        If temp + HeightSrc - 1 <= Ydest + HeightDest - 1 Then
            TileHorizontal SRC, xSrc, ySrc, WidthSrc, HeightSrc, DEST, Xdest, temp, WidthDest
        Else
            TileHorizontal SRC, xSrc, ySrc, WidthSrc, (Ydest + HeightDest - 1) - temp, DEST, Xdest, temp, WidthDest
        End If
    Next
End Sub

'Create Garbage blocks
Public Sub CreateGarbage(X As Long, Y As Long, width As Long, height As Long, Optional Color = -1)
    If X < 1 Then X = 1
    With Grid(Y, X)
        .Color = Color
        .height = height
        .width = width
        .Clean = False
    End With
    RaiseEvent GarbageCreated(X, Y, width, height)
    DropTile X, Y
End Sub

'Is the space occupied by garbage?
Public Function IsGarbage(X As Long, Y As Long, Optional Action As GarbageAction) As Boolean
    Dim temp As Long, temp2 As Long
    If X > GridWidth Or Y > GridHeight Or X < 0 Or Y < 0 Then Exit Function
    For temp = 1 To X
        For temp2 = GridHeight To Y Step -1
            With Grid(temp2, temp)
                If .Color < 0 Then
                    If temp + .width - 1 >= X And temp2 - (.height - 1) <= Y Then
                        IsGarbage = True
                        Select Case Action
                            Case destroy: DestroyGarbage temp, temp2
                            Case Drop: DropTile temp, temp2
                        End Select
                        Exit Function
                    End If
                End If
            End With
        Next
    Next
End Function

Private Function GarbageTile(X As Long, Y As Long) As Tile
Dim temp As Long, temp2 As Long
    If X > GridWidth Or Y > GridHeight Or X < 0 Or Y < 0 Then Exit Function
    For temp = 1 To X
        For temp2 = GridHeight To Y Step -1
            With Grid(temp2, temp)
                If .Color < 0 Then
                    If temp + .width - 1 >= X And temp2 - (.height - 1) <= Y Then
                        GarbageTile = Grid(Y, X)
                        Exit Function
                    End If
                End If
            End With
        Next
    Next
End Function

'Destroys garbage blocks
Public Sub DestroyGarbage(X As Long, Y As Long)
    Dim temp As Long, temp2 As Long, temp3 As Long, temp4 As Long, temptile As Tile
    temptile = Grid(Y, X)
    temp4 = 1
    temp3 = -Ghost * 3 'TEST VALUE FOR NOW
    For temp = X To X + temptile.width - 1
            'temp3 = Ghost * temp4
            For temp2 = (Y - temptile.height) + 1 To Y
                'If temp <> X And temp2 <> Y Then
                With Grid(temp2, temp)
                    ClearTile temp, temp2, True
                    '.Color = 1
                    .Z = temp3
                    RaiseEvent GarbageSubstituted(temp, temp2, .Color)
                End With
                'temp3 = temp3 + Ghost
                'End If
            Next
            temp4 = temp4 + 1
            'DropColumn temp, (Y - temptile.height) + 1
    Next
    'Stopper = temp3
End Sub

'check every tile underneath a garbage block to see if it can be dropped one level
Public Function CanDropGarbage(X As Long, Y As Long) As Boolean
    Dim temp As Long
    With Grid(Y, X)
        If .Color < 0 Then
            CanDropGarbage = True
            For temp = X To X + Grid(Y, X).width - 1
                If Grid(Y - .height, temp).Color <> 0 Then
                    CanDropGarbage = False
                    Exit Function
                End If
            Next
        End If
    End With
End Function

'used to find the Y coordinate in the new switching effect
Public Function DegreesToXY(CenterX As Long, CenterY As Long, degree As Double, radiusX As Long, radiusY As Long, Optional IsX As Boolean = True) As Long
    If IsX Then DegreesToXY = CenterX - (Sin(-degree * convert) * radiusX) Else DegreesToXY = CenterY - (Sin((90 + (degree)) * convert) * radiusY)
End Function

'Returns the points for a combo
Public Function ComboScore(Tiles As Long) As Long
    Select Case Tiles
        Case 3, 4, 5: ComboScore = Tiles - 2
        Case 6, 7, 8, 9: ComboScore = Tiles - 1
        Case 10: ComboScore = 10
        Case Else: ComboScore = 14 + (Tiles - 11) * 3
    End Select
End Function

'returns the points for a chain
Public Function ChainScore(Hits As Long) As Long
    Select Case Hits
        Case 0, 1: ChainScore = 0
        Case 2: ChainScore = 5
        Case 3: ChainScore = 8
        Case 4: ChainScore = 15
        Case 5, 6, 7: ChainScore = (Hits - 2) * 10
        Case 8, 9, 10, 11, 12: ChainScore = 70 + (Hits - 8) * 2
        Case Else: ChainScore = 180
    End Select
End Function

Public Function StartCombo(X As Long) As Long
    Dim temp As Long
    If Combos(X).value = 0 Then
        StartCombo = X
    Else
        For temp = X - 1 To 1 Step -1
            If Combos(temp).value <> Combos(X).value Then
                StartCombo = temp + 1
                Exit Function
            End If
        Next
        StartCombo = 1
    End If
End Function

Public Function EndCombo(X As Long) As Long
    Dim temp As Long
    If Combos(X).value = 0 Then
        EndCombo = X
    Else
        For temp = X + 1 To GridWidth
            If Combos(temp).value <> Combos(X).value Then
                EndCombo = temp - 1
                Exit Function
            End If
        Next
        EndCombo = GridWidth
    End If
End Function

Private Sub DisableCombos()
    Dim temp As Long
    For temp = 1 To GridWidth
        Combos(temp).Used = False
    Next
End Sub

Public Function AddCombo(X As Long, Optional Operation As ComboOP) As Boolean
    Dim temp As Long, tempbl As Boolean
    For temp = StartCombo(X) To EndCombo(X)
        Select Case Operation
            Case AddOne: Combos(temp).value = Combos(temp).value + 1
            Case Clear: Combos(temp).value = 0
            Case Enable: Combos(temp).Used = True
            Case Disable: Combos(temp).Used = False
        End Select
        tempbl = Combos(temp).Used Or tempbl
    Next
    AddCombo = tempbl
End Function

'How the chain system works:
'The chain detection and quantification process resides within the decrementoffsets sub
'at the beginning of the sub, the combo counter for each column is disabled
'during the loops, if a tile is in motion or destruction, its columns combo counter is enabled
'if scores are made, the combo counters in the columns of the score are incremented by 1
'after the loop, if the combo counter is still enabled, a chain is registered as in progress
'otherwise its registered as complete, passed as an event to the user, then the combo counters are cleared and disabled
'now since a combo can take place along multiple columns, i had to devise a method to treat multiple columns as one
'combo counters with the same number if > 0 are treated as one column/combo
'however, if you make 2 simple scores one on top of the other, itll regsiter incorrectly as a chain, but I dont mind that much
'cause the method to fix that would be a lot more complex

Public Sub CreateRandomGarbage(Optional width As Long, Optional height As Long)
    Dim temp As Long
    Randomize Timer
    temp = -(Rnd * (GColors - 1) + 1)
    If doGarbage > 0 Or height > 0 Then
        If height = 0 Then
            If Lines Mod doGarbage = 0 Then
                Randomize Timer
                width = Int(Rnd * (GridWidth / 2) + (GridWidth / 2))
                height = Int(Rnd * 3) + 1
                CreateGarbage (GridWidth - width) / 2, GridHeight, width, height, temp
            End If
        Else
            CreateGarbage (GridWidth - width) / 2, GridHeight, width, height, temp
        End If
    End If
End Sub

Public Sub AutoCreateRandomGarbage(Optional isChain As Boolean, Optional ByVal count As Long)
    Dim temp As Long, width As Long, height As Long
    If GameOver Or GameWon Then Exit Sub
    If isChain Then
        If count < 2 Then Exit Sub
        If count > 13 Then count = 13
        height = count - 1
        width = GridWidth
    Else 'is a normal combo
        If count < 4 Then Exit Sub
        height = 1
        For temp = 1 To count
            width = width + 1
            If width > GridWidth Then
                height = height + 1
                width = 2
            End If
        Next
    End If
    CreateRandomGarbage width, height
End Sub

Public Function DegreesToY3D(degree As Double) As Long
    Const convert As Double = pi / 180
    DegreesToY3D = -(Sin((90 + (degree)) * convert) * (TileSize / 2))
End Function

Public Sub Seed3D()
    Dim temp As Long, temp2 As Long
    For temp = 0 To BoardWidth / 2
        ThreeDEffect(temp) = DegreesToY3D(99 + (temp * (BoardWidth / 160)))
        If ThreeDEffect(temp) > temp2 Then temp2 = ThreeDEffect(temp)
        If ThreeDEffect(temp) < temp2 Then ThreeDEffect(temp) = temp2
    Next
    For temp = BoardWidth / 2 + 1 To BoardWidth - 1
        ThreeDEffect(temp) = ThreeDEffect(BoardWidth / 2 - (temp - BoardWidth / 2) + 1)
    Next
End Sub

Public Sub DrawTile3D(ByVal Index As Long, X As Long, Y As Long, Optional Back As Boolean, Optional Buffer As Boolean, Optional Effect As Boolean)
'BitBlt Destination.hDC, X, Y, TileSize, TileSize, Source.hDC, Index * TileSize, 0, vbSrcCopy
Dim temp As Long
Index = Index * TileSize
For temp = 0 To TileSize - 1
    If X + temp >= 0 And X + temp < BoardWidth Then
    If Back Then
        If isEven(temp) Then
            BitBlt Destination.hDC, X + temp, Y - ThreeDEffect(X + temp), 1, TileSize, Source.hDC, Index + temp, IIf(Effect, TileSize, 0), vbSrcCopy
        Else
            DrawLine X + temp, Y - ThreeDEffect(X + temp) + 1, X + temp, Y - ThreeDEffect(X + temp) + TileSize + 1, vbBlack
        End If
    Else
        If isEven(temp) Or Not Buffer Then
            BitBlt Destination.hDC, X + temp, Y + ThreeDEffect(X + temp) - 1, 1, TileSize, Source.hDC, Index + temp, IIf(Effect, TileSize, 0), vbSrcCopy
        Else
            DrawLine X + temp, Y + ThreeDEffect(X + temp) - 1, X + temp, Y + ThreeDEffect(X + temp) - 1 + TileSize, vbBlack
        End If
    End If
    End If
Next
End Sub

Public Sub DrawScreen3D()
    Dim temp As Long, X As Long, Half As Long
    Half = GridWidth / 4
    X = 7
    For temp = -Half * 3 - 1 To -Half
        Draw3DColumn GetCol(temp), True, X
        X = X - 1
    Next

    X = 0
    For temp = -Half - 1 To Half
        Draw3DColumn GetCol(temp), False, X
        X = X + 1
    Next
    
    DrawCursor Half, CursorY
    DrawCursor Half + 1, CursorY
End Sub

Public Sub Draw3DColumn(Column As Long, Back As Boolean, X As Long)
    Dim Y As Long, temp As Long, X2 As Long
    temp = OffsetX
    If Back Then temp = temp * -1
    For Y = 0 To GridHeight
        With Grid(Y, Column)
            X2 = .X
            If Not Back Then X2 = X2 * -1
            If .Color > 0 Then
                If .Color = 1 And isEven(.height) Then
                    DrawTile3D .width, (X - 1) * TileSize + X2 + temp, BoardHeight - (Y * TileSize) - .Y - Offset, Back, Y = 0, True
                Else
                    DrawTile3D .Color, (X - 1) * TileSize + X2 + temp, BoardHeight - (Y * TileSize) - .Y - Offset, Back, Y = 0
                End If
                .Clean = True
            End If
        End With
    Next
End Sub

Public Function GetCol(ByVal RelativeToCursor As Long) As Long
    RelativeToCursor = CursorX - RelativeToCursor
    Do Until RelativeToCursor > 0
        RelativeToCursor = RelativeToCursor + GridWidth
    Loop
    Do Until RelativeToCursor <= GridWidth
        RelativeToCursor = RelativeToCursor - GridWidth
    Loop
    GetCol = RelativeToCursor
End Function
