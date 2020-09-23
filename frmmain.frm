VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmmain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Panel de Pon - Microsoft Visual Basic [run]"
   ClientHeight    =   8130
   ClientLeft      =   150
   ClientTop       =   765
   ClientWidth     =   7035
   Icon            =   "frmmain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   Picture         =   "frmmain.frx":0442
   ScaleHeight     =   542
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   469
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox lstmoves 
      Columns         =   5
      Height          =   255
      ItemData        =   "frmmain.frx":37FB
      Left            =   1695
      List            =   "frmmain.frx":37FD
      TabIndex        =   12
      Top             =   7800
      Visible         =   0   'False
      Width           =   2265
   End
   Begin VB.CommandButton cmdmain 
      Caption         =   "6"
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   8.25
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3705
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   7545
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.PictureBox picmain 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   4800
      Index           =   4
      Left            =   1080
      ScaleHeight     =   320
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   192
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   840
      Width           =   2880
      Begin VB.Image imgwon 
         Height          =   2325
         Left            =   0
         Picture         =   "frmmain.frx":37FF
         Top             =   1200
         Visible         =   0   'False
         Width           =   2835
      End
      Begin VB.Image imgmain 
         Height          =   2730
         Left            =   -120
         Picture         =   "frmmain.frx":5628
         Top             =   960
         Visible         =   0   'False
         Width           =   3015
      End
   End
   Begin VB.PictureBox picmain 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   480
      Index           =   6
      Left            =   1080
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   192
      TabIndex        =   3
      Tag             =   "Player 2 Buffer"
      Top             =   840
      Visible         =   0   'False
      Width           =   2880
   End
   Begin VB.PictureBox picmain 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   4800
      Index           =   5
      Left            =   4080
      ScaleHeight     =   320
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   192
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   840
      Width           =   2880
      Begin VB.Image imgplayer2 
         Height          =   2730
         Left            =   -120
         Picture         =   "frmmain.frx":7C17
         Top             =   960
         Visible         =   0   'False
         Width           =   3015
      End
   End
   Begin VB.PictureBox picmain 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   1440
      Index           =   1
      Left            =   1080
      Picture         =   "frmmain.frx":A206
      ScaleHeight     =   96
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   192
      TabIndex        =   1
      Top             =   840
      Visible         =   0   'False
      Width           =   2880
   End
   Begin VB.PictureBox picmain 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   480
      Index           =   0
      Left            =   1080
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   192
      TabIndex        =   0
      Tag             =   "Player 1 Buffer"
      Top             =   960
      Visible         =   0   'False
      Width           =   2880
   End
   Begin MSWinsockLib.Winsock wskmain 
      Left            =   450
      Top             =   5190
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      RemoteHost      =   "5151"
      LocalPort       =   5151
   End
   Begin VB.Timer TimerSmall 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   15
      Top             =   3480
   End
   Begin VB.Timer TimerBig 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   0
      Top             =   3480
   End
   Begin VB.Timer TimerAI 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   0
      Top             =   3480
   End
   Begin VB.Label lblbutton 
      BackStyle       =   0  'Transparent
      Height          =   375
      Left            =   960
      TabIndex        =   19
      Top             =   5640
      Width           =   6135
   End
   Begin VB.Label lblscore 
      BackColor       =   &H00FFFFFF&
      Caption         =   "0"
      Height          =   210
      Index           =   11
      Left            =   4740
      TabIndex        =   18
      Top             =   7815
      Width           =   2055
   End
   Begin VB.Label lblscore 
      BackColor       =   &H00FFFFFF&
      Caption         =   "0"
      Height          =   210
      Index           =   10
      Left            =   1725
      TabIndex        =   17
      Top             =   7815
      Width           =   2055
   End
   Begin VB.Label lblscore 
      BackColor       =   &H00FFFFFF&
      Caption         =   "0"
      Height          =   210
      Index           =   9
      Left            =   4740
      TabIndex        =   16
      Top             =   7575
      Width           =   2055
   End
   Begin VB.Label lblscore 
      BackColor       =   &H00FFFFFF&
      Caption         =   "0"
      Height          =   210
      Index           =   8
      Left            =   1725
      TabIndex        =   15
      Top             =   7575
      Width           =   2055
   End
   Begin VB.Label lblscore 
      BackColor       =   &H00FFFFFF&
      Caption         =   "0"
      Height          =   210
      Index           =   7
      Left            =   4740
      TabIndex        =   14
      Top             =   7335
      Width           =   2055
   End
   Begin VB.Label lblscore 
      BackColor       =   &H00FFFFFF&
      Caption         =   "0"
      Height          =   210
      Index           =   6
      Left            =   1725
      TabIndex        =   13
      Top             =   7335
      Width           =   2055
   End
   Begin VB.Label lblscore 
      BackColor       =   &H00FFFFFF&
      Caption         =   "0"
      Height          =   210
      Index           =   5
      Left            =   4740
      TabIndex        =   10
      Top             =   7095
      Width           =   2055
   End
   Begin VB.Label lblscore 
      BackColor       =   &H00FFFFFF&
      Caption         =   "0"
      Height          =   210
      Index           =   4
      Left            =   1725
      TabIndex        =   9
      Top             =   7095
      Width           =   2055
   End
   Begin VB.Label lblscore 
      BackColor       =   &H00FFFFFF&
      Caption         =   "0"
      Height          =   210
      Index           =   3
      Left            =   4740
      TabIndex        =   8
      Top             =   6855
      Width           =   2055
   End
   Begin VB.Label lblscore 
      BackColor       =   &H00FFFFFF&
      Caption         =   "0"
      Height          =   210
      Index           =   2
      Left            =   1725
      TabIndex        =   7
      Top             =   6855
      Width           =   2055
   End
   Begin VB.Label lblscore 
      BackColor       =   &H00FFFFFF&
      Caption         =   "0"
      Height          =   210
      Index           =   1
      Left            =   4740
      TabIndex        =   6
      Top             =   6615
      Width           =   2055
   End
   Begin VB.Label lblscore 
      BackColor       =   &H00FFFFFF&
      Caption         =   "0"
      Height          =   210
      Index           =   0
      Left            =   1725
      TabIndex        =   5
      Top             =   6615
      Width           =   2055
   End
   Begin VB.Image imgAI 
      Height          =   1095
      Left            =   0
      Top             =   6360
      Width           =   975
   End
   Begin VB.Menu mnufile 
      Caption         =   "&File"
      Begin VB.Menu mnustory 
         Caption         =   "&Story"
      End
      Begin VB.Menu mnumarathon 
         Caption         =   "&Marathon"
      End
      Begin VB.Menu mnulineclear 
         Caption         =   "&Line Clear"
      End
      Begin VB.Menu mnutimeclear 
         Caption         =   "&Time Clear"
      End
      Begin VB.Menu mnugarbage 
         Caption         =   "&Garbage"
      End
      Begin VB.Menu mnucustom 
         Caption         =   "&Custom"
         Begin VB.Menu mnunew 
            Caption         =   "&New"
         End
         Begin VB.Menu mnuload 
            Caption         =   "&Load"
         End
         Begin VB.Menu mnusave 
            Caption         =   "&Save"
         End
         Begin VB.Menu mnuplay 
            Caption         =   "&Play"
         End
      End
      Begin VB.Menu mnufilesep 
         Caption         =   "-"
      End
      Begin VB.Menu mnuonline 
         Caption         =   "&MultiPlayer"
         Begin VB.Menu mnuai 
            Caption         =   "&Artificial Intelligence"
         End
         Begin VB.Menu mnuonlinesep 
            Caption         =   "-"
         End
         Begin VB.Menu mnulisten 
            Caption         =   "&Listen"
         End
         Begin VB.Menu mnulocalport 
            Caption         =   "&Set Local Port"
         End
         Begin VB.Menu mnuonlinesep1 
            Caption         =   "-"
         End
         Begin VB.Menu mnuconnect 
            Caption         =   "&Connect"
         End
         Begin VB.Menu mnuremoteport 
            Caption         =   "&Set Remote Port"
         End
         Begin VB.Menu mnuonlinesep2 
            Caption         =   "-"
         End
         Begin VB.Menu mnuhostexe 
            Caption         =   "&Host the EXE"
            Enabled         =   0   'False
         End
         Begin VB.Menu mnudisconnect 
            Caption         =   "&Disconnect/Disable"
         End
      End
      Begin VB.Menu mnufilesep3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuthreed 
         Caption         =   "&3D Mode"
      End
      Begin VB.Menu mnufilesep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnureset 
         Caption         =   "&Reset"
      End
      Begin VB.Menu mnuexit 
         Caption         =   "&Exit"
      End
   End
   Begin VB.Menu mnuoption 
      Caption         =   "&Options"
   End
   Begin VB.Menu mnupause 
      Caption         =   "&Pause"
   End
End
Attribute VB_Name = "frmmain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public WithEvents PANEL As PanelDePonCls, WithEvents PANEL2 As PanelDePonCls, WithEvents AI As PanelDePonAI
Attribute PANEL.VB_VarHelpID = -1
Attribute PANEL2.VB_VarHelpID = -1
Attribute AI.VB_VarHelpID = -1
Public host As Boolean

Public Sub Update(Index As Integer, value As Long)
    lblscore(Index).Caption = CStr(value)
    lblscore(Index).ForeColor = vbRed
End Sub

'Used for new front end
Private Function isinregion(Left, top, width, height, X, Y) As Boolean
    isinregion = X >= Left And X <= Left + width And Y >= top And Y <= top + height
End Function

Private Sub cmdmain_Click()
    lstmoves.Visible = Not lstmoves.Visible
End Sub

'make the bg interactive
Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If isinregion(0, 0, 102, 26, X, Y) Then PopupMenu mnucustom, , X, Y
    If isinregion(103, 0, 25, 26, X, Y) Then mnuload_Click
    If isinregion(131, 0, 25, 26, X, Y) Then mnusave_Click
End Sub

'Save position
Private Sub Form_Unload(Cancel As Integer)
    SaveSetting "Panel de Pon", "Main", "Left", Left
    SaveSetting "Panel de Pon", "Main", "Top", top
End Sub

Private Sub lblbutton_Click()
    If height = Maximized Then height = Minimized Else height = Maximized
End Sub

Private Sub lstmoves_Click()
    lstmoves.Visible = False
End Sub

'Activate the AI
Private Sub mnuai_Click()
    StoryLevel = 0
    If PANEL.Mode <> Edit And PANEL.Mode <> Puzzle Then
        mnudisconnect_Click
        TimerAI.Enabled = True
        width = IsOnline
    Else
        MsgBox "Artificial Intelligence is not available in Custom Puzzle Creation or Play modes"
    End If
End Sub

'Connect to a remote IP
Private Sub mnuconnect_Click()
    host = False
    wskmain.Close
    wskmain.Connect InputBox("Please enter an IP", , "127.0.0.1")
End Sub

'Disconnect/stop listening/ai
Private Sub mnudisconnect_Click()
    wskmain.Close
    host = False
    mnufile.Visible = True
    width = NotOnline
    mnuoption.Visible = True
    TimerAI.Enabled = False
    PANEL.Reset
    PANEL.ClearGrid
End Sub

'Initialize classes, PdP classes, pictureboxes, AI, load settings
Private Sub Form_Load()
    Dim temp As Long
    width = NotOnline
    Set PANEL = New PanelDePonCls
    Set PANEL2 = New PanelDePonCls
    PANEL.InitPanelDePon picmain(0), picmain(1), picmain(4), , , 32, , 48, , , 2
    PANEL2.InitPanelDePon picmain(6), picmain(1), picmain(5), , , 32, , 48, , , 2

    Left = Val(GetSetting("Panel de Pon", "Main", "Left", Left))
    top = Val(GetSetting("Panel de Pon", "Main", "Top", top))
    
    Set AI = New PanelDePonAI
    AI.INIT PANEL2
    Set AIPIC = imgAI
    
    For temp = 1 To 5
        lstmoves.AddItem CStr(temp)
    Next
    height = Maximized
End Sub

'Exit the program
Private Sub mnuexit_Click()
    If TimerSmall.Enabled Then mnupause_Click
    Unload Me
    End
End Sub

Private Sub mnugarbage_Click()
    mnumarathon_Click
    PANEL.doGarbage = 3
End Sub

'Activate line clear mode
Private Sub mnulineclear_Click()
    mnumarathon_Click
    PANEL.Lines = -5 'arbitrary value for now
    PANEL.Mode = Lineclear
    Synch
    If isConnected(wskmain) And host Then SOCKSend wskmain, "MODE=" & Timelimit & ";LINES=" & PANEL.Lines
End Sub

'Start listening
Private Sub mnulisten_Click()
    On Error Resume Next
    wskmain.Listen
    host = True
    MsgBox wskmain.LocalIP & " (" & wskmain.LocalHostName & ") is listening on port " & wskmain.LocalPort, vbInformation
End Sub

'Load a file from common dialog
Private Sub mnuload_Click() 'dafhi reformatted this code
    Dim temp As String
    InitOpen "Panel de Pon Custom Maps (*.pon)" & Chr(0) & "*.pon", "Load a Panel de Pon custom map file"
    temp = Open_File(Me.hWnd)
    If LoadTmp(temp) Then mnuplay_Click Else cmdmain.Visible = True
End Sub

'Allows other subs to load a file with the error checking
Private Function LoadTmp(temp As String) As Boolean 'dafhi reformatted this code
    If Len(temp) > 0 Then
        temp = PANEL.LoadWholeFile(temp)
        If Len(temp) > 0 Then
            mnunew_Click
            If Not PANEL.LoadCustom(temp) Then
                MsgBox "The file was invalid", vbCritical, "Unable to load"
            Else
                PANEL.Clean = False
                PANEL.DrawScreen
                LoadTmp = True
            End If
        End If
    End If
End Function

'Set the local port (for listening)
Private Sub mnulocalport_Click()
    wskmain.LocalPort = Val(InputBox("Please set the local port", , wskmain.LocalPort))
End Sub

'Activate marathon mode
Private Sub mnumarathon_Click()
    INITPdP
    PANEL.InitGrid
    If host Then PANEL2.LoadCustom PANEL.SaveCustom
    NORMALAnyMusic
    SYNCHTimer
End Sub

Public Sub INITPdP()
    PANEL.Reset
    imgmain.Visible = False
    imgwon.Visible = False
    TimerSmall.Enabled = True
    TimerBig.Enabled = True
    picmain(4).Enabled = True
    cmdmain.Visible = False
    PANEL.GameOver = False
End Sub

'Create a new custom puzzle
Private Sub mnunew_Click()
    INITPdP
    TimerSmall.Enabled = False
    TimerBig.Enabled = False
    With PANEL
        .ClearGrid
        .Reset
        .GameOver = False
        .Mode = Edit
        .DrawScreen
    End With
    cmdmain.Visible = True
    picmain(4).Enabled = True
End Sub

'Bring up the options form
Private Sub mnuoption_Click()
    If TimerSmall.Enabled Then mnupause_Click
    frmsec.ChangeDetails
    frmsec.Show vbModal, Me
End Sub

'Pause the game
Private Sub mnupause_Click()
    If Not PANEL.GameOver And PANEL.GridWidth > 0 Then
        If isConnected(wskmain) Then SOCKSend wskmain, "PAUSE"
        TimerSmall.Enabled = Not TimerSmall.Enabled
        TimerBig.Enabled = TimerSmall.Enabled
        picmain(4).Enabled = TimerSmall.Enabled
        mnupause.Caption = IIf(TimerSmall.Enabled, "Pause", "Resume")
        If Not TimerSmall.Enabled Then PANEL.Clearscreen
        SOCKSend wskmain, "PAUSE"
    End If
End Sub

'Bring the system out of editing mode, and activate gravity
Private Sub mnuplay_Click()
    PANEL.Mode = Puzzle
    PANEL.DropAll
    TimerSmall.Enabled = True
    picmain(4).Enabled = True
    NORMALAnyMusic
    cmdmain.Visible = False
End Sub

'Set the remote port (for connecting)
Private Sub mnuremoteport_Click()
    wskmain.RemotePort = Val(InputBox("Please set the remote port", , wskmain.RemotePort))
End Sub

Private Sub mnureset_Click()
mnudisconnect_Click
End Sub

'Save a custom puzzle
Private Sub mnusave_Click()
    Dim temp As String
    If PANEL.Mode = Edit Then
        InitSave "Panel de Pon Custom Maps (*.pon)" & Chr(0) & "*.pon", "Save a Panel de Pon custom map file"
        temp = Save_File(Me.hWnd, "pon")
        If Len(temp) > 0 Then
            If Not PANEL.SaveWholeFile(temp, PANEL.SaveCustom) Then MsgBox "An unknown error has occured", vbCritical, "Unable to save"
        End If
    Else
        MsgBox "You aren't even in edit mode!", vbCritical, "Unable to save"
    End If
End Sub

Private Sub mnustory_Click()
    MsgBox "Story mode is a Proof-of-concept right now, I have to make actual artwork for this, so right now it's just a demonstration"
    NewLevel
End Sub

Private Sub mnuthreed_Click()
    INITPdP
    PANEL.INIT3DMODE
    PANEL.InitGrid False
End Sub

'Activate time limit mode
Private Sub mnutimeclear_Click()
    mnumarathon_Click
    PANEL.Currtime = 120000 / TimerBig.Interval '2 minutes divided by the occurance of the big timer
    PANEL.Mode = Timelimit
    Synch
    If isConnected(wskmain) And host Then SOCKSend wskmain, "MODE=" & Timelimit & ";TIME=" & PANEL.Currtime
End Sub

'synch panel2 with 1
Public Sub Synch()
    PANEL2.Currtime = PANEL.Currtime
    PANEL2.Mode = PANEL.Mode
    PANEL2.Lines = PANEL.Lines
    PANEL2.Moves = PANEL.Moves
End Sub

'An action was performed by the user, send it to the remote client
Private Sub PANEL_ActionPerformed(Index As PanelAction)
    If isConnected(wskmain) Then SOCKSend wskmain, "ACTION=" & Index
End Sub

Private Sub PANEL_ComboMade(Hits As Long)
    PANEL2.AutoCreateRandomGarbage True, Hits
    If Val(lblscore(4)) < Hits Then Update 4, Hits
    lblscore(4).ForeColor = vbRed
End Sub

'The cursor was moved by the user, send it to the remote client
Private Sub PANEL_CursorMoved(X As Long, Y As Long)
    PANEL.DrawScreen
    If isConnected(wskmain) Then SOCKSend wskmain, "X=" & X & ";Y=" & Y
End Sub

'The grid was cleared, send it to the remote client
Private Sub PANEL_GridCleared()
    If isConnected(wskmain) And host Then SOCKSend wskmain, "GRIDCLEARED"
End Sub

Private Sub PANEL_KillTiles(count As Long)
    PANEL2.AutoCreateRandomGarbage False, count
    If Val(lblscore(2)) < count Then Update 2, count
End Sub

Private Sub PANEL_LinesChainged(Lines As Long)
    Update 6, Lines
End Sub

'A new map was created/loaded, send it to the remote client
Private Sub PANEL_MapLoaded()
    Dim tempstr() As String, temp As Long
    imgmain.Visible = False
    imgwon.Visible = False
    For temp = 0 To lblscore.UBound
        lblscore(temp).Caption = "0"
        lblscore(temp).ForeColor = vbBlack
    Next
    If isConnected(wskmain) And host Then 'SOCKSend wskmain, "MAP=" & PANEL.SaveCustom
        tempstr = Split(PANEL.SaveCustom, vbNewLine)
        For temp = 0 To UBound(tempstr)
            SOCKSend wskmain, "MAP=" & tempstr(temp)
        Next
        SOCKSend wskmain, "MAPCOMPLETE"
    Else
        If TimerAI.Enabled Then PANEL2.LoadCustom PANEL.SaveCustom
    End If
End Sub

Public Sub SYNCHTimer()
    SOCKSend wskmain, "TIMERSMALL=" & TimerSmall.Interval & ";TIMERBIG=" & TimerBig.Interval & ";ENABLE"
End Sub

Private Sub PANEL_MovesChainged(Moves As Long)
    Update 8, Moves
End Sub

'A new row was created, send it to the remote client
Private Sub PANEL_NewRow(Row As String)
    If isConnected(wskmain) Then SOCKSend wskmain, "ROW=" & Row
End Sub

'THIS IS NO LONGER HANDLED BY THE NETWORK CODE
'The board/global offset was increased, send it to the remote client
'Private Sub PANEL_OffsetIncreased()
    'If isConnected(wskmain) Then SOCKSend wskmain, "OFFSETINCREASED"
'End Sub
'The incremental offsets were decremented, send it to the remote client
'Private Sub PANEL_OffsetsDecremented()
    'If isConnected(wskmain) Then SOCKSend wskmain, "OFFSETDECREASED"
'End Sub

'The user lost, send it to the remote client
Private Sub PANEL_PDPGameOver()
    StoryLevel = 0
    If isConnected(wskmain) Then SOCKSend wskmain, "GAMEOVER"
    imgmain.Visible = True
    TimerBig.Enabled = False
    TimerSmall.Enabled = False
    picmain(4).Enabled = False
    If TimerAI.Enabled And StoryLevel > 0 Then STORYAIWIN
End Sub

'The user won, send it to the remote client
Private Sub PANEL_PDPGameWon()
    If isConnected(wskmain) Then SOCKSend wskmain, "GAMEWON"
    imgmain.Visible = False
    picmain(4).Enabled = False
    TimerBig.Enabled = False
    TimerSmall.Enabled = False
    If TimerAI.Enabled Then
        If StoryLevel > 0 Then STORYAILOSE
    Else
        imgwon.Visible = True
    End If
End Sub

Private Sub PANEL_RaiseLevel()
    Dim temp As Long
    If Not isConnected(wskmain) Then
        If PANEL.IncreaseSpeed Then
            temp = TimerBig.Interval * 0.9
            If temp > 0 And temp < TimerBig.Interval Then TimerBig.Interval = temp
            SYNCHTimer
        End If
    End If
    PANEL.IncreaseSpeed = False
End Sub

Private Sub PANEL_ScoreChanged(Score As Long)
    Update 0, Score * 10
End Sub

Private Sub PANEL_TopChanged(height As Long)
    If StoryLevel > 0 Then AUTOMusicEmotion
End Sub

Private Sub PANEL2_ComboMade(Hits As Long)
    PANEL.AutoCreateRandomGarbage True, Hits
    If Val(lblscore(5)) < Hits Then Update 5, Hits
End Sub

Private Sub PANEL2_KillTiles(count As Long)
    PANEL.AutoCreateRandomGarbage False, count
    If Val(lblscore(3)) < count Then Update 3, count
End Sub

Private Sub PANEL2_MapLoaded()
    imgplayer2.Visible = False
End Sub

Private Sub PANEL2_PDPGameOver()
    imgplayer2.Visible = True
    PANEL.GameWon = True
    PANEL_PDPGameWon
    TimerAI.Enabled = False
End Sub

Private Sub PANEL2_ScoreChanged(Score As Long)
    Update 1, Score * 10
End Sub

Private Sub PANEL2_TopChanged(height As Long)
    If StoryLevel > 0 Then
        AUTOMusicEmotion
        AUTOAIEmotion
    End If
End Sub

'The user clicked on the board
Private Sub picmain_Click(Index As Integer)
    PANEL.Action MoveTiles
End Sub

'The user pressed a key
Private Sub picmain_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If Index = 4 Then PANEL.Keydown KeyCode
End Sub

'The user moved the mouse
Private Sub picmain_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Index = 4 Then PANEL.MouseMove X, Y
End Sub

'Perform an AI action
Private Sub TimerAI_Timer()
    If TimerBig.Enabled Then AI.PullAction
End Sub

'Increment the board/global offset
Private Sub TimerBig_Timer()
    Const Increment As Long = 32
    Dim temp As Integer
    PANEL.IncreaseOffset PANEL.Speed
    
    If TimerAI.Enabled Or isConnected(wskmain) Then PANEL2.IncreaseOffset PANEL2.Speed
    
    For temp = 0 To lblscore.UBound
        'alphablend from red
        Select Case lblscore(temp).ForeColor
            Case 0 'GNDN
            Case Is < Increment: lblscore(temp).ForeColor = vbBlack
            Case Else: lblscore(temp).ForeColor = lblscore(temp).ForeColor - Increment
        End Select
        'If lblscore(temp).ForeColor = vbRed Then lblscore(temp).ForeColor = vbBlack
    Next
End Sub

'Decrement the incremental offsets
Private Sub TimerSmall_Timer()
    PANEL.DecrementOffsets
    PANEL.DrawScreen
    If TimerAI.Enabled Or isConnected(wskmain) Then
        PANEL2.DecrementOffsets
        PANEL2.DrawScreen
    End If
End Sub

'User was disconnected
Private Sub wskmain_Close()
    mnudisconnect_Click
    width = NotOnline
End Sub

'User has connected to a remote host
Private Sub wskmain_Connect()
    MsgBox "Connection has been made a host!"
    mnufile.Visible = False
    mnuoption.Visible = False
    picmain(4).Enabled = True
    width = IsOnline
End Sub

'A remote client is requesting to connect
Private Sub wskmain_ConnectionRequest(ByVal requestID As Long)
    wskmain.Close
    wskmain.Accept requestID
    width = IsOnline
    MsgBox "Connection has been made with a client!"
    If wskmain.LocalPort <> 80 Then SOCKSend wskmain, "VERSION=" & App.Major & "." & App.Minor & "." & App.Revision
End Sub

'Checks if the winsock is connected
Public Function isConnected(sock As Winsock) As Boolean
    isConnected = sock.State = sckConnected
End Function

'Data has arrived, parse it
Private Sub wskmain_DataArrival(ByVal bytesTotal As Long)
    Dim temp As String
    wskmain.GetData temp, vbString
    ProcessCommand temp
End Sub

'Parse data commands in string form
Public Sub ProcessCommand(ByVal temp As String)
    Const VERSIONFAIL As String = "The connected client's version does not match yours. Both of you must upgrade to the newest or same version"
    Static MAP As String, GARBAGEX As Long, GARBAGEY As Long, GARBAGEW As Long, GARBAGEH As Long
    Dim temp2 As String
    Do Until InStr(temp, ";") = Len(temp) Or InStr(temp, ";") = 0
        temp2 = Left(temp, InStr(temp, ";"))
        ProcessCommand temp2
        temp = Right(temp, Len(temp) - InStr(temp, ";"))
    Loop
    temp = Replace(temp, ";", Empty)
    If InStr(temp, "=") > 0 Then
        temp2 = Right(temp, Len(temp) - InStr(temp, "="))
        temp = Left(temp, InStr(temp, "=") - 1)
    End If
    With PANEL2
        Select Case UCase(temp)
            'has no parameters
            Case "PAUSE"
            Case "ENABLE"
                TimerBig.Enabled = True
                TimerSmall.Enabled = True
            Case "GAMEWON": .GameWon = True
            Case "GRIDCLEARED"
            Case "GAMEOVER": .GameOver = True
            
            'THESE ARE NO LONGER HANDLED BY THE NETWORK CODE
            'Case "OFFSETINCREASED"
            '    .IncreaseOffset
            '    If Not host Then TimerBig_Timer
            'Case "OFFSETDECREASED"
            '    .DecrementOffsets
            '    If Not host Then TimerSmall_Timer
            '    .DrawScreen
                
            'has a single parameter
            Case "MODE"
                PANEL.Mode = Val(temp2)
                PANEL2.Mode = PANEL.Mode
            Case "LINES"
                PANEL.Lines = Val(temp2)
                PANEL2.Lines = PANEL.Lines
            Case "TIME"
                PANEL.Currtime = Val(temp2)
                PANEL2.Currtime = PANEL.Currtime
            Case "ROW"
                .LoadRow temp2
                .DrawScreen
            Case "ACTION": .Action Val(temp2) ': PANEL.DrawScreen
            Case "VERSION"
                If temp2 <> App.Major & "." & App.Minor & "." & App.Revision Then
                    SOCKSend wskmain, "VERSIONFAIL"
                    MsgBox VERSIONFAIL
                    wskmain.Close
                End If
            Case "VERSIONFAIL"
                MsgBox VERSIONFAIL
            Case "TIMERBIG"
                TimerBig.Interval = Val(temp2)
            Case "TIMERSMALL"
                TimerSmall.Interval = Val(temp2)
            Case "PAUSE"
                TimerSmall.Enabled = Not TimerSmall.Enabled
                TimerBig.Enabled = Not TimerBig.Enabled
            
            'has multiple parameters spread across multiple commands
            Case "MAP"
                MAP = MAP & temp2 & vbNewLine
            Case "X" 'Y always follows after the X, so only draw when Y is received
                .CursorX = Val(temp2)
            Case "GARBAGEX" 'X coord of a garbage block
                GARBAGEX = Val(temp2)
            Case "GARBAGEY" 'Y coord of a garbage block
                GARBAGEY = Val(temp2)
            Case "GARBAGEW" 'width of a garbage block
                GARBAGEW = Val(temp2)
            Case "GARBAGEH" 'height of a garbage block
                GARBAGEH = Val(temp2)
                
            'Puts multiple parameters into one command
            Case "Y"
                .MoveCursor .CursorX, Val(temp2)
                .DrawScreen
            Case "GARBAGE1" 'send bgarbage to player1
                .CreateGarbage GARBAGEX, GARBAGEY, GARBAGEW, GARBAGEH
                .DrawScreen
            Case "GARBAGE2" 'send bgarbage to player1
                PANEL2.CreateGarbage GARBAGEX, GARBAGEY, GARBAGEW, GARBAGEH
                PANEL2.DrawScreen
            Case "MAPCOMPLETE" 'full map received, load it into both panels
                MAP = Left(MAP, Len(MAP) - Len(vbNewLine))
                .LoadCustom MAP
                PANEL.LoadCustom MAP
                MAP = Empty
                .GameOver = False
                PANEL2.GameOver = False
                
        End Select
    End With
    DoEvents
End Sub
