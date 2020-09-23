VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PanelDePonAI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit 'My pathetic attempt to make a decent AI
'While simultaneously taking time into affect as to not do everything instantaneously
'having distinct seperately disablable strategies as to perform on multiple levels

Public Event NoActions() 'Called when the AI has nothing planned (therefore plan something new)

Private Type PDPMove
    Action As PanelAction
    Quantity As Long 'saves more resources then pushing one action many times
End Type

Private Const StrategyCount As Long = 1, Adjacent As Long = 2 'Constants
Private StrategyList(0 To StrategyCount) As Boolean
'Public Enum Strategies 'Put more complex/advanced strategies at the top to be executed first
    Private Const ThreeHorizontal As Long = 0
    Private Const ThreeVertical As Long = 1
'End Enum 'for some reason VB keeps asking for a constant expression, enums are constants!!!

Private MoveList() As PDPMove, MoveBottom As Long, MoveTop As Long, MoveCount As Long
Public WithEvents PANEL As PanelDePonCls
Attribute PANEL.VB_VarHelpID = -1

'INITialize the AI
Public Sub INIT(PDP As PanelDePonCls)
    Dim temp As Long
    Set PANEL = PDP
    For temp = 0 To StrategyCount
        StrategyList(temp) = True
    Next
    ClearActions
End Sub

'Auto AI strategy Level
Public Sub AILevel(ByVal Level As Long)
    Dim temp As Long
    If Level > StrategyCount Then Level = StrategyCount
    For temp = 2 To StrategyCount
        StrategyList(temp) = temp <= Level
    Next
End Sub

'Clears the Action stack
Public Sub ClearActions()
    MoveCount = 0
    MoveBottom = 0
    MoveTop = 0
    ReDim MoveList(0)
    RaiseEvent NoActions
    CalculateAction
End Sub

'Push an action series on to the stack
Public Sub PushAction(Action As PanelAction, Optional Quantity As Long = 1)
    MoveTop = MoveTop + 1
    MoveCount = MoveCount + 1
    ReDim Preserve MoveList(MoveBottom To MoveTop)
    With MoveList(MoveTop - 1)
        .Action = Action
        .Quantity = Quantity
    End With
End Sub

'Pull an action off the stack and decrease its quantity
'if the quantity is 0, remove the action series from the stack and redimension it
'if there is nothing on the stack, reset it to zero position
Public Sub PullAction()
    If MoveCount = 0 Then CalculateAction
    If MoveCount = 0 Or PANEL.GridWidth = 0 Then Exit Sub
    
    PANEL.Action MoveList(MoveBottom).Action 'perform action
    MoveList(MoveBottom).Quantity = MoveList(MoveBottom).Quantity - 1 'decrement quantity
    If MoveList(MoveBottom).Quantity = 0 Then 'remove it
        MoveBottom = MoveBottom + 1
        MoveCount = MoveCount - 1
        If MoveCount = 0 Then
            ClearActions
        Else
            ReDim Preserve MoveList(0 To MoveTop)
        End If
    End If
End Sub

'Move the cursor to a set position and switch the tiles if specified
Public Sub MoveCursor(X As Long, Y As Long, Optional Switch As Boolean)
    With PANEL
        If .CursorX < X Then PushAction MoveRight, X - .CursorX
        If .CursorX > X Then PushAction MoveLeft, .CursorX - X
        If .CursorY < Y Then PushAction Moveup, Y - .CursorY
        If .CursorY > Y Then PushAction MoveDown, .CursorY - Y
        If Switch Then PushAction MoveTiles
    End With
End Sub

'Switch a tile repeatedly along X axis
Public Sub MoveTile(ByVal X As Long, Y As Long, Direction As Long)
    Dim temp As Long
    MoveCursor X, Y, False
    If X = PANEL.GridWidth Or Direction > 0 Then PushAction MoveTiles
    For temp = 1 To Abs(Direction)
        PushAction IIf(Direction < 0, MoveLeft, MoveRight)
        PushAction MoveTiles
    Next
End Sub

'AI engine to detect whether or not to perform strategies
Public Sub CalculateAction()
    Dim temp As Long
    For temp = StrategyCount To 0 Step -1
        If StrategyList(temp) Then If DoStrategy(temp) Then Exit Sub
    Next
    PANEL.Action MoveBoard
End Sub

'Checks if a row has a color
Public Function YhasColor(Y As Long, Color As Long) As Boolean
    Dim temp As Long
    With PANEL
        For temp = 1 To .GridWidth
            If .GetColor(temp, Y) = Color Then
                YhasColor = True
                Exit Function
            End If
        Next
    End With
End Function

'Checks if a column has a color
Public Function XhasColor(X As Long, Color As Long) As Boolean
    Dim temp As Long
    With PANEL
        For temp = 1 To .GridHeight
            If .GetColor(X, temp) = Color Then
                XhasColor = True
                Exit Function
            End If
        Next
    End With
End Function

'Checks if a tile can be moved across
Public Function XCanMove(X As Long, Y As Long, X2 As Long) As Boolean
    Dim temp As Long
    If Y = 1 Then
        XCanMove = True
    Else
        With PANEL
            If X2 > X Then
                For temp = X To X2
                    If .GetColor(temp, Y - 1) < 2 Then Exit Function
                Next
                XCanMove = True
            Else
                For temp = X2 To X Step -1
                    If .GetColor(temp, Y - 1) < 2 Then Exit Function
                Next
                XCanMove = True
            End If
        End With
    End If
End Function

'Gets a tile that can move across
Public Function XFindTile(X As Long, Y As Long, Color As Long) As Long
    Dim temp As Long
    With PANEL
        If .GetColor(X, Y) = Color Then
            XFindTile = X
        Else
            For temp = 1 To .GridWidth
                If .GetColor(temp, Y) = Color Then
                    If XCanMove(temp, Y, X) Then
                        XFindTile = temp
                        Exit For
                    End If
                End If
            Next
        End If
    End With
End Function

'Counts the amount of one color on a line
Public Function CountTiles(Y As Long, Color As Long) As Long
    Dim temp As Long, temp2 As Long
    With PANEL
        For temp = 1 To .GridWidth
            If .GetColor(temp, Y) = Color Then temp2 = temp2 + 1
        Next
        CountTiles = temp2
    End With
End Function

'Gets the first tile along X
Public Function XFindFirstTile(X As Long, Y As Long, Color As Long) As Long
    Dim temp As Long
    With PANEL
        For temp = X + 1 To .GridWidth
            If .GetColor(temp, Y) = Color Then
                XFindFirstTile = temp
                Exit For
            End If
        Next
    End With
End Function

'Calculates the distance automatically
Public Sub MoveTileTo(X As Long, Y As Long, X2 As Long)
    If X <> X2 Then MoveTile X, Y, X2 - X
End Sub

'Tries to move any tile
Public Function XMoveAnyTile(X As Long, Y As Long, Color As Long) As Boolean
    Dim temp As Long
    temp = XFindTile(X, Y, Color)
    If temp > 0 Then
        MoveTileTo temp, Y, X
        XMoveAnyTile = True
    End If
End Function

'Performs (or tries to) the AI
Public Function DoStrategy(Index As Long) As Boolean
    Dim temp As Long, temp2 As Long, temp3 As Long, temp4 As Long, temp5 As Long
    With PANEL
        Select Case Index 'a dash represents any number of tiles in between (0 to gridwidth -1)
            
            Case ThreeVertical 'make three in a column
                'if three tiles exit in seperate consecutive rows that can be put together, try and do so
                '1 holds Y, 2 holds color, 3 holds X
                For temp = 1 To .GridHeight - 2 'check each row (except the top ones, they get checked later)
                    For temp2 = 2 To .Colors 'check each color
                        If CountTiles(temp, temp2) > 0 Then 'if this row has the color
                            If CountTiles(temp + 1, temp2) > 0 Then 'and the next row
                                If CountTiles(temp + 2, temp2) > 0 Then 'and the last row
                                    For temp3 = 1 To .GridWidth 'check each tile in the top row
                                        If .GetColor(temp3, temp + 2) = temp2 Then 'if its the right color
                                            temp4 = XFindTile(temp3, temp + 1, temp2)
                                            temp5 = XFindTile(temp3, temp, temp2)
                                            If temp4 > 0 And temp5 > 0 Then 'and you can move a tile from underneath it, and one underneath that one
                                                MoveTileTo temp4, temp + 1, temp3
                                                MoveTileTo temp5, temp, temp3
                                                DoStrategy = True 'Complete
                                            End If
                                        End If
                                    Next
                                End If
                            End If
                        End If
                    Next
                 Next
                 
            Case ThreeHorizontal 'X-X-X to XXX make three in a row
                'if three tiles of one color are in a single row that can be moved together, try and do so
                '1 holds Y, 2 holds color, 3 hold X coord of first tile, 4 holds current X,5 holds last X
                For temp = 1 To .GridHeight 'check each row
                    For temp2 = 2 To .Colors 'check each color
                        If CountTiles(temp, temp2) > Adjacent Then 'if the amount of tiles of one color in a row are greater than 2
                            'Move first tile
                            temp3 = XFindFirstTile(0, temp, temp2) 'get the first one
                            temp5 = temp3 + 1
                            'Move second tile
                            temp4 = XFindFirstTile(temp3, temp, temp2)
                            If temp4 > temp5 Then MoveTile temp4, temp, temp5 - temp4
                            temp5 = temp5 + 1
                            'Move third tile
                            temp4 = XFindFirstTile(temp4, temp, temp2)
                            If temp4 > temp5 Then MoveTile temp4, temp, temp5 - temp4
                            'Complete
                            DoStrategy = True
                            Exit Function
                        End If
                    Next
                Next
                
                
        End Select
    End With
End Function

'The game has reacted, change desision
Private Sub PANEL_KillTiles(count As Long)
    ClearActions
End Sub
