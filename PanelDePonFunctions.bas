Attribute VB_Name = "PanelDePonFunctions"
Option Explicit

Public Enum GameMode
    Normal = 0
    Puzzle = 1
    edit = 2
    lineclear = 3
End Enum

Public Enum PanelAction
    MoveUp
    MoveDown
    MoveLeft
    MoveRight
    MoveTiles
    MoveBoard
End Enum

Private Type Tile
    Color As Long
    X As Long
    Y As Long
    Clean As Boolean
    Z As Long
End Type

Private Const Adjacent As Long = 2 'Constants
Private Destination As PictureBox, Source As PictureBox, Mask As PictureBox, Display As PictureBox 'Objects
Public GridWidth As Long, GridHeight As Long, TileSize As Long, Colors As Long, Ghost As Long, Clean As Boolean 'Graphic variables
Public TileCount As Long, Offset As Long, CursorX As Long, CursorY As Long, BoardHeight As Long 'Dimension variables
Public GameOver As Boolean, GameWon As Boolean, Score As Long, Lines As Long, Moves As Long, Currtime As Long, Speed As Long 'Game Variables
Public Stopper As Long, Mode As GameMode, IncreaseSpeed As Boolean, Level As Long
Public Grid() As Tile, UndoList() As String, UndoCount As Long 'Game board (row, column)

'Initialize all startup values
Public Sub InitPanelDePon(Dest As PictureBox, Src As PictureBox, Buffer As PictureBox, Msk As PictureBox, Disp As PictureBox, Optional Width As Long = 6, Optional Height As Long = 10, Optional Size As Long = -1, Optional Tiles As Long = -1, Optional Remains As Long = 1024, Optional MaxMoves As Long, Optional GameSpeed As Long = 1)
    If Size = -1 Then Size = Src.Height
    If Tiles = -1 Then Tiles = Width * Height * (25 / 60)
    Set Destination = Dest
    Set Source = Src
    Set Mask = Msk
    Set Display = Disp
    Destination.FillStyle = vbSolid
    Destination.FillColor = Destination.BackColor
    MakeMask Source.hDC, Buffer.hDC, Mask.hDC, 0, 0, Source.Width, Source.Height, vbWhite
    GridWidth = Width
    GridHeight = Height
    BoardHeight = GridHeight * Size
    TileSize = Size
    TileCount = Tiles
    Ghost = Remains
    Speed = 1
    If TileSize \ GameSpeed = TileSize / GameSpeed Then Speed = GameSpeed 'make sure its a multiple of tilesize
    Moves = MaxMoves
    Colors = Source.Width \ Source.Height
    ReDim Grid(0 To Height, 1 To Width)
    Destination.Move Destination.Left, Destination.Top, Size * Width, Size * Height
    Display.Move Display.Left, Display.Top, Destination.Width, Destination.Height
    Reset
End Sub

'Reset Level variables to default
Public Sub Reset()
    PurgeUndo
    CursorX = 1
    CursorY = 1
    Offset = 0
    Score = 0
    Currtime = 0
    Level = 0
    IncreaseSpeed = False
    Mode = Normal
    GameOver = True
    GameWon = False
    Clean = False
End Sub

'Purge all undos
Public Sub PurgeUndo()
    UndoCount = 0
    ReDim UndoList(0)
End Sub

'Basic keyboard interpretor
Public Sub Keydown(KeyCode As Integer)
    Select Case KeyCode
        Case 38, 104: Action MoveUp
        Case 40, 98: Action MoveDown
        Case 37, 100: Action MoveLeft
        Case 39, 102: Action MoveRight
        Case 13, 32, 101: Action MoveTiles
        Case 9, 16, 18, 20, 33, 34, 96, 107: Action MoveBoard
        Case 8, 46, 85, 90: If Mode = Puzzle Then PullUndo
    End Select
End Sub

'Basic command interpretor
Public Sub Action(Index As PanelAction)
    If GameOver And Mode <> edit Then Exit Sub
    Select Case Index
        Case MoveUp: MoveCursor CursorX, CursorY + 1
        Case MoveDown: MoveCursor CursorX, CursorY - 1
        Case MoveLeft: MoveCursor CursorX - 1, CursorY
        Case MoveRight: MoveCursor CursorX + 1, CursorY
        Case MoveBoard: ShiftUp
        Case MoveTiles
            If Mode = edit Then
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
    tempstr = GridWidth & "," & GridHeight & "," & Mode & "," & Moves & "," & Lines & "," & Score & vbNewLine
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
    If UndoCount > 1 Then
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
Public Function LoadCustom(Map As String) As Boolean
    On Error Resume Next
    Dim tempstr() As String, tempstr2() As String, temp As Long, temp2 As Long
    tempstr = Split(Map, vbNewLine)
    GameOver = False
    tempstr2 = Split(tempstr(0), ",")
    GridWidth = Val(tempstr2(0))
    GridHeight = Val(tempstr2(1))
    Mode = Val(tempstr2(2))
    Moves = Val(tempstr2(3))
    Lines = Val(tempstr2(4))
    Score = Val(tempstr2(5))
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
    Clean = False
    LoadCustom = True
End Function

'Draws the entire grid and both cursors
Public Sub DrawScreen()
    Dim temp As Long, temp2 As Long
    If Not Clean Then Destination.Cls
    ClearArea
    For temp = 1 To GridWidth
        For temp2 = 0 To GridHeight
            DrawGrid temp, temp2
        Next
    Next
    DrawCursor CursorX, CursorY
    If Mode <> edit Then
        DrawCursor CursorX + 1, CursorY
        If Offset > 0 Then Shade -Offset, BoardHeight - Offset, Destination.Width, Offset
    End If
    Clean = True
    Set Display.Picture = Destination.Image
    Display.Refresh
End Sub

'Draws a cursor on to destination
Public Sub DrawCursor(X As Long, Y As Long)
    DrawTile 0, (X - 1) * TileSize, BoardHeight - (Y * TileSize) - Offset
End Sub

'Draws a tile on the grid
Public Sub DrawGrid(X As Long, Y As Long)
    With Grid(Y, X)
        If .Clean = False Or Clean = False Then
            If (.X <> 0 Or .Y <> 0) And .Color = 0 Then Exit Sub
            DrawTile .Color, (X - 1) * TileSize + .X, BoardHeight - (Y * TileSize) - .Y - Offset, True
            .Clean = True
        End If
    End With
End Sub

'Used for mouse control of the cursor
Public Sub MouseMove(ByVal X As Single, ByVal Y As Single)
    If Not GameOver And TileSize > 0 Then MoveCursor X \ TileSize + 1, GridHeight - ((Y + Offset) \ TileSize)
End Sub

'If the new position is not the same as the old one, dirty the tile under the current cursor position, then move the cursor
Public Sub MoveCursor(X As Long, Y As Long)
    If X < 1 Then X = 1
    If Mode = edit Then
        If X > GridWidth Then X = GridWidth
    Else
        If X > GridWidth - 1 Then X = GridWidth - 1
    End If
    If Y < 1 Then Y = 1
    If Y > GridHeight Then Y = GridHeight
    
    If X <> CursorX Or Y <> CursorY Then
        Grid(CursorY, CursorX).Clean = False
        If Mode <> edit Then Grid(CursorY, CursorX + 1).Clean = False
        DrawGrid CursorX, CursorY
        If Mode <> edit Then DrawGrid CursorX + 1, CursorY
        CursorX = X
        CursorY = Y
    End If
End Sub

'Place random tiles
Public Sub InitGrid()
    Dim temp As Long
    Reset
    GameOver = False
    ClearGrid
    For temp = 1 To GridWidth
        ClearTile temp, 0, True
    Next
    For temp = 1 To TileCount
        PlaceRandomTile
    Next
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
End Sub

'Shift the board up a given amount of pixels. If the offset matches the tilesize than shift everything up one row
'In line mode, the time is increased
Public Sub IncreaseOffset(Optional Pixels As Long = 1)
    If GameOver Then Exit Sub
    If Mode = lineclear Then
        Currtime = Currtime + 1
    Else
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
Public Sub DecrementOffsets(Optional Pixels As Long = 2)
    If GameOver Then Exit Sub
    Dim temp As Long, temp2 As Long
    For temp = 1 To GridWidth
        For temp2 = 1 To GridHeight 'the buffer row cant be moved by the user and thus shouldnt have any offsets
            With Grid(temp2, temp)
                If .X <> 0 Or .Y <> 0 Or .Z <> 0 Then
                    If .Color = 0 Then
                        DropColumn temp, temp2 + 1
                        .X = 0
                        .Y = 0
                        .Clean = False
                    Else
                        If .X < 0 Then .X = .X + Pixels
                        If .X > 0 Then .X = .X - Pixels
                        If .Y > 0 Then
                            .Y = .Y - Pixels
                            If .Y <= 0 Then
                                .Y = 0
                                If temp2 < GridHeight Then
                                    If Grid(temp2 + 1, temp).Color = 0 Then Grid(temp2 + 1, temp).Clean = False
                                End If
                            End If
                        End If
                        If .Z > 0 Then
                            .Z = .Z - 1
                            If .Z = 0 Then .Color = 0
                            DropColumn temp, temp2 + 1
                        Else
                            If .X = 0 And .Y = 0 Then
                                'MsgBox "Checking " & temp & ", " & temp2
                                CheckTile temp, temp2
                            End If
                        End If
                        .Clean = False
                    End If
                End If
            End With
        Next
    Next
End Sub

'Shift all tiles up one row. If the top row is filled already, game over. Create random tiles in the buffer/bottom row
Public Sub ShiftUp()
    Dim temp As Long, temp2 As Long
    For temp = 1 To GridWidth
        If Bottom(temp) = GridHeight Then
            GameOver = True
            Exit Sub
        Else
            If Offset < TileSize Then Clean = False
            For temp2 = GridHeight - 1 To 0 Step -1
                Grid(temp2 + 1, temp) = Grid(temp2, temp)
            Next
            CheckTile temp, 1
            ClearTile temp, 0, True
        End If
    Next
    If Mode = lineclear Then
        Lines = Lines - 1
    Else
        Lines = Lines + 1
        IncreaseSpeed = Getlevel(Level) = Lines
    End If
    Offset = 0
End Sub

'Used for speed enhancing
Public Function Getlevel(Optional Level As Long) As Long
    Getlevel = 2 ^ Level 'Took the easy way out and did exponential growth
End Function

'Place a random tile
Public Sub PlaceRandomTile()
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
    Dim temp As Tile
    If Grid(Y, X).X = 0 And Grid(Y, X).Y = 0 And Grid(Y, X + 1).X = 0 And Grid(Y, X + 1).Y = 0 Then
        If Mode = Puzzle Then
            PushUndo
            Moves = Moves - 1
            If Moves = 0 Then GameOver = True
        Else
            Moves = Moves + 1
        End If
        temp = Grid(Y, X)
        Grid(Y, X) = Grid(Y, X + 1)
        Grid(Y, X + 1) = temp
        Grid(Y, X).Clean = False
        Grid(Y, X).X = TileSize
        Grid(Y, X + 1).Clean = False
        Grid(Y, X + 1).X = -TileSize
    End If
End Sub

'If the tile is not the lowest it can go, drop it down one row, and set its offset along Y to 32 so it gradually comes down
Public Sub DropTile(X As Long, Y As Long)
    If Bottom(X) < Y Then
        Grid(Y - 1, X) = Grid(Y, X)
        ClearTile X, Y
        Grid(Y - 1, X).Y = TileSize
        Grid(Y - 1, X).Clean = False
    End If
End Sub

'Clears a tiles properties at x,y. If Newtile = true, a random tile is created
Public Sub ClearTile(X As Long, Y As Long, Optional Newtile As Boolean)
    With Grid(Y, X)
        .Color = 0
        If Newtile Then '
            .Color = RandomTile
            Do Until CountX(X, Y, .Color) < Adjacent
                .Color = RandomTile
            Loop
        End If
        .X = 0
        .Y = 0
        .Z = 0
        .Clean = False
    End With
End Sub

'Finds the first empty spot in a column (0 is used as a buffer)
Public Function Bottom(X As Long) As Long
    Dim temp As Long
    For temp = 1 To GridHeight
        If Grid(temp, X).Color < 2 Then
            Bottom = temp
            Exit Function
        End If
    Next
End Function
    
'Draws a tile from source into destination. If 0 is ignored (you dont want to draw a cursor) and index = 0 a blank square is drawn
Public Sub DrawTile(Index As Long, X As Long, Y As Long, Optional ignorezero As Boolean)
    Dim temp As Long
    If Index = 0 And ignorezero Then
        DrawSquare X, Y
    Else
        temp = Index * TileSize
        TransBLT Source.hDC, temp, 0, Mask.hDC, temp, 0, TileSize, TileSize, Destination.hDC, X, Y
    End If
End Sub

'Draws a blank square
Public Sub DrawSquare(X As Long, Y As Long)
    Destination.Line (X, Y)-(X + TileSize - 1, Y + TileSize - 1), Destination.BackColor, B
End Sub

'Draws a line
Public Sub DrawLine(X1 As Long, Y1 As Long, X2 As Long, Y2 As Long, Optional Color As OLE_COLOR = vbBlack)
    Destination.Line (X1, Y1)-(X2, Y2), Color
End Sub

'Dither/shade
Public Sub Shade(X As Long, Y As Long, Width As Long, Height As Long)
    Dim temp As Long, X1 As Long, Y1 As Long, X2 As Long, Y2 As Long
    For temp = X - Height To X + Width + Height Step 2 'Middle
        'drawaline temp, Y, Width, Height
        X1 = temp
        Y1 = Y
        
        If X1 < X Then
            X1 = X
            Y1 = Y1 + (X - X1)
        End If
        
        X2 = X1 + Height
        Y2 = Y1 + Height
        
        If Y2 > Y + Height Then
            X2 = X2 - (Y2 - (Y + Height - 1))
            Y2 = Y + Height - 1
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
        DropTile X, Y
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
Public Sub CheckTile(X As Long, Y As Long)
    Dim pLeft As Long, pRight As Long, pTop As Long, pBottom As Long, Color As Long, HOR As Boolean, VER As Boolean
    Color = Grid(Y, X).Color
    
    If Color = 0 Or Bottom(X) < Y Then
        DropColumn X, Y ' DropTile temp, temp2
    Else
        pLeft = StartX(X, Y, Color)
        pRight = EndX(X, Y, Color)
        HOR = pRight - pLeft + 1 > Adjacent
        
        pTop = StartY(X, Y, Color)
        pBottom = EndY(X, Y, Color)
        VER = pBottom - pTop + 1 > Adjacent
        
        If HOR Then GhostX pLeft, Y, pRight
        If VER Then GhostY X, pTop, pBottom
    End If
    
    Select Case Mode
        Case lineclear: GameWon = CountTiles = 0 And Lines = 0
        Case Puzzle: GameWon = CountTiles = 0
    End Select
End Sub

'Ghosts a horizontal Score
Public Sub GhostX(X As Long, Y As Long, X2 As Long)
    Dim temp As Long
    For temp = X To X2
        GhostTile temp, Y
    Next
End Sub

'Ghosts a vertical Score
Public Sub GhostY(X As Long, Y As Long, Y2 As Long)
    Dim temp As Long
    For temp = Y To Y2
        GhostTile X, temp
    Next
End Sub

'Ghosts a single tile
Public Sub GhostTile(X As Long, Y As Long)
    If X > 0 And X <= GridWidth And Y >= 0 And Y <= GridWidth Then
        Grid(Y, X).Color = 1
        Grid(Y, X).Clean = False
        Grid(Y, X).Z = Ghost
    End If
End Sub

'Returns the highest block
Public Function Top() As Long
    Dim temp As Long, temp2 As Long, temp3 As Long
    For temp = 1 To GridWidth
        temp2 = Bottom(temp) - 1
        If temp2 > temp3 Then temp3 = temp2
    Next
    Top = temp3
End Function

'Find the start of concurrent tiles in a row
Public Function StartX(X As Long, Y As Long, Optional Color As Long = -1) As Long
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
Public Function StartY(X As Long, Y As Long, Optional Color As Long = -1) As Long
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
Public Function EndX(X As Long, Y As Long, Optional Color As Long = -1) As Long
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
Public Function EndY(X As Long, Y As Long, Optional Color As Long = -1) As Long
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
Public Function CountX(X As Long, Y As Long, Optional Color As Long = -1) As Long
    Dim temp As Long, temp2 As Long
    temp = StartX(X, Y, Color)
    temp2 = EndX(temp, Y, Color)
    CountX = temp2 - temp + 1
End Function

'Find the number of concurrent tiles in a column
Public Function CountY(X As Long, Y As Long, Optional Color As Long = -1) As Long
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
Public Function SaveWholeFile(Filename As String, Optional Text As String) As Boolean
    On Error Resume Next
    Dim tempfile As Long
    tempfile = FreeFile
    Open Filename For Output As tempfile
        Print #tempfile, Text
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
