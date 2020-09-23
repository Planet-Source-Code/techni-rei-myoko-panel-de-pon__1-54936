VERSION 5.00
Begin VB.Form frmsec 
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   " Options / About / Help"
   ClientHeight    =   4245
   ClientLeft      =   60
   ClientTop       =   330
   ClientWidth     =   9615
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   283
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   641
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdok 
      Caption         =   "&Ok"
      Height          =   375
      Left            =   1320
      TabIndex        =   36
      Top             =   3840
      Width           =   1215
   End
   Begin VB.PictureBox picmain 
      HasDC           =   0   'False
      Height          =   330
      Index           =   6
      Left            =   120
      ScaleHeight     =   18
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   237
      TabIndex        =   32
      Top             =   3150
      Width           =   3615
      Begin VB.HScrollBar hscrmain 
         Height          =   255
         Index           =   6
         LargeChange     =   10
         Left            =   1440
         Max             =   128
         TabIndex        =   33
         Top             =   0
         Value           =   16
         Width           =   1215
      End
      Begin VB.Label lblcaption 
         Alignment       =   2  'Center
         Caption         =   "Ghost duration:"
         Height          =   255
         Index           =   6
         Left            =   120
         TabIndex        =   35
         Top             =   30
         Width           =   1215
      End
      Begin VB.Label lblvalue 
         Alignment       =   2  'Center
         Caption         =   "16"
         Height          =   255
         Index           =   6
         Left            =   2760
         TabIndex        =   34
         Top             =   30
         Width           =   735
      End
   End
   Begin VB.PictureBox picmain 
      HasDC           =   0   'False
      Height          =   330
      Index           =   5
      Left            =   120
      ScaleHeight     =   18
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   237
      TabIndex        =   28
      Top             =   3480
      Width           =   3615
      Begin VB.HScrollBar hscrmain 
         Height          =   255
         Index           =   5
         LargeChange     =   2
         Left            =   1440
         Max             =   6
         Min             =   1
         TabIndex        =   29
         Top             =   0
         Value           =   2
         Width           =   1215
      End
      Begin VB.Label lblcaption 
         Alignment       =   2  'Center
         Caption         =   "Pixel increment all:"
         Height          =   255
         Index           =   5
         Left            =   0
         TabIndex        =   31
         Top             =   30
         Width           =   1455
      End
      Begin VB.Label lblvalue 
         Alignment       =   2  'Center
         Caption         =   "2"
         Height          =   255
         Index           =   5
         Left            =   2760
         TabIndex        =   30
         Top             =   30
         Width           =   735
      End
   End
   Begin VB.PictureBox picmain 
      HasDC           =   0   'False
      Height          =   330
      Index           =   4
      Left            =   120
      ScaleHeight     =   18
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   237
      TabIndex        =   24
      Top             =   2280
      Width           =   3615
      Begin VB.HScrollBar hscrmain 
         Height          =   255
         Index           =   4
         LargeChange     =   10
         Left            =   1440
         Max             =   2
         Min             =   2
         TabIndex        =   25
         Top             =   0
         Value           =   2
         Width           =   1215
      End
      Begin VB.Label lblvalue 
         Alignment       =   2  'Center
         Caption         =   "2"
         Height          =   255
         Index           =   4
         Left            =   2760
         TabIndex        =   27
         Top             =   30
         Width           =   735
      End
      Begin VB.Label lblcaption 
         Alignment       =   2  'Center
         Caption         =   "Strategies:"
         Height          =   255
         Index           =   4
         Left            =   120
         TabIndex        =   26
         Top             =   30
         Width           =   1215
      End
   End
   Begin VB.CommandButton cmdmain 
      Caption         =   "2"
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   8.25
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3975
      Left            =   3840
      TabIndex        =   23
      Top             =   120
      Width           =   255
   End
   Begin VB.PictureBox picmain 
      HasDC           =   0   'False
      Height          =   330
      Index           =   3
      Left            =   120
      ScaleHeight     =   18
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   237
      TabIndex        =   18
      Top             =   1920
      Width           =   3615
      Begin VB.HScrollBar hscrmain 
         Height          =   255
         Index           =   3
         LargeChange     =   10
         Left            =   1440
         Max             =   320
         Min             =   1
         TabIndex        =   19
         Top             =   0
         Value           =   100
         Width           =   1215
      End
      Begin VB.Label lblcaption 
         Alignment       =   2  'Center
         Caption         =   "Speed:"
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   21
         Top             =   30
         Width           =   1215
      End
      Begin VB.Label lblvalue 
         Alignment       =   2  'Center
         Caption         =   "1000 MS"
         Height          =   255
         Index           =   3
         Left            =   2760
         TabIndex        =   20
         Top             =   30
         Width           =   735
      End
   End
   Begin VB.CheckBox chkmain 
      Caption         =   "Scrolling Background and curved animation"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   16
      Top             =   2880
      Width           =   3615
   End
   Begin VB.PictureBox picmain 
      HasDC           =   0   'False
      Height          =   330
      Index           =   2
      Left            =   120
      ScaleHeight     =   18
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   237
      TabIndex        =   12
      Top             =   360
      Width           =   3615
      Begin VB.HScrollBar hscrmain 
         Height          =   255
         Index           =   2
         LargeChange     =   2
         Left            =   1440
         Max             =   20
         Min             =   1
         TabIndex        =   13
         Top             =   0
         Value           =   4
         Width           =   1215
      End
      Begin VB.Label lblcaption 
         Alignment       =   2  'Center
         Caption         =   "Maximum moves:"
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   15
         Top             =   30
         Width           =   1215
      End
      Begin VB.Label lblvalue 
         Alignment       =   2  'Center
         Caption         =   "4"
         Height          =   255
         Index           =   2
         Left            =   2760
         TabIndex        =   14
         Top             =   30
         Width           =   735
      End
   End
   Begin VB.PictureBox picmain 
      HasDC           =   0   'False
      Height          =   330
      Index           =   1
      Left            =   120
      ScaleHeight     =   18
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   237
      TabIndex        =   8
      Top             =   1320
      Width           =   3615
      Begin VB.HScrollBar hscrmain 
         Height          =   255
         Index           =   1
         LargeChange     =   2
         Left            =   1440
         Max             =   320
         Min             =   1
         TabIndex        =   9
         Top             =   0
         Value           =   50
         Width           =   1215
      End
      Begin VB.Label lblcaption 
         Alignment       =   2  'Center
         Caption         =   "Time:"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   11
         Top             =   30
         Width           =   1215
      End
      Begin VB.Label lblvalue 
         Alignment       =   2  'Center
         Caption         =   "500 MS"
         Height          =   255
         Index           =   1
         Left            =   2760
         TabIndex        =   10
         Top             =   30
         Width           =   735
      End
   End
   Begin VB.PictureBox picmain 
      HasDC           =   0   'False
      Height          =   330
      Index           =   0
      Left            =   120
      ScaleHeight     =   18
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   237
      TabIndex        =   4
      Top             =   960
      Width           =   3615
      Begin VB.HScrollBar hscrmain 
         Height          =   255
         Index           =   0
         LargeChange     =   2
         Left            =   1440
         Max             =   6
         Min             =   1
         TabIndex        =   6
         Top             =   0
         Value           =   4
         Width           =   1215
      End
      Begin VB.Label lblvalue 
         Alignment       =   2  'Center
         Caption         =   "4"
         Height          =   255
         Index           =   0
         Left            =   2760
         TabIndex        =   7
         Top             =   30
         Width           =   735
      End
      Begin VB.Label lblcaption 
         Alignment       =   2  'Center
         Caption         =   "Pixel increment up:"
         Height          =   255
         Index           =   0
         Left            =   0
         TabIndex        =   5
         Top             =   30
         Width           =   1455
      End
   End
   Begin VB.ListBox lstmain 
      Height          =   1230
      ItemData        =   "frmsec.frx":0000
      Left            =   4080
      List            =   "frmsec.frx":0016
      TabIndex        =   1
      Top             =   120
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.TextBox txtmain 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3975
      Left            =   4080
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Text            =   "frmsec.frx":004C
      Top             =   120
      Width           =   5415
   End
   Begin VB.Label lblmain 
      Caption         =   "Artificial Intelligence:"
      Height          =   255
      Index           =   3
      Left            =   120
      TabIndex        =   22
      Top             =   1680
      Width           =   3615
   End
   Begin VB.Label lblmain 
      Caption         =   "Visual/Audio Effects:"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   17
      Top             =   2640
      Width           =   3615
   End
   Begin VB.Label lblmain 
      Caption         =   "Marathon and Line Mode:"
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   3
      Top             =   720
      Width           =   3615
   End
   Begin VB.Label lblmain 
      Caption         =   "Puzzle Mode:"
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   3615
   End
End
Attribute VB_Name = "frmsec"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim tempstr() As String 'holds the help file seperated into sections

'Load settings
Public Sub ChangeDetails()
    On Error Resume Next
    Me.Enabled = False
    If hscrmain(0).Value <> frmmain.PANEL.Speed Then hscrmain(0).Value = frmmain.PANEL.Speed
    If hscrmain(1).Value <> frmmain.TimerBig.Interval \ 10 Then hscrmain(1).Value = frmmain.TimerBig.Interval \ 10
    If hscrmain(2).Value <> frmmain.TimerAI.Interval \ 10 Then hscrmain(2).Value = frmmain.TimerAI.Interval \ 10
    If frmmain.PANEL.Mode = Edit And frmmain.PANEL.Moves > 0 And frmmain.PANEL.Moves < hscrmain(2).Max Then hscrmain(2).Value = frmmain.PANEL.Moves
    chkmain(0).Value = IIf(frmmain.PANEL.doBG, vbChecked, vbUnchecked)
    hscrmain(5).Value = frmmain.PANEL.PixelSpeed
    hscrmain(6).Value = frmmain.PANEL.Ghost
    Me.Enabled = True
End Sub

'Check box options
Private Sub chkmain_Click(Index As Integer)
    Select Case Index
        Case 0 'Scrolling background
            frmmain.PANEL.doBG = chkmain(0).Value = vbChecked
            frmmain.PANEL2.doBG = frmmain.PANEL.doBG
    End Select
End Sub

'Hide/show the list box
Private Sub cmdmain_Click()
    lstmain.Visible = Not lstmain.Visible
End Sub

Private Sub cmdok_Click()
    Me.Hide
End Sub

'Load the pipe delimeted help file
Private Sub Form_Load()
    On Error Resume Next
    tempstr = Split(frmmain.PANEL.LoadWholeFile(Replace(App.PATH & "\panel.txt", "\\", "\")), "|")
End Sub

'Resize the text box
Private Sub Form_Resize()
    If ScaleWidth > 280 Then txtmain.width = ScaleWidth - 280
    If ScaleHeight > 15 Then txtmain.height = ScaleHeight - 15
    cmdmain.height = txtmain.height
End Sub

'Vertical scroll bar options
Private Sub hscrmain_Change(Index As Integer)
    If Not frmmain.PANEL.Mode = Edit And (Not frmmain.PANEL.GameOver Or Not Me.Enabled) Then Exit Sub
    lblvalue(Index).Caption = hscrmain(Index).Value
    Select Case Index
        Case 0 'Pixel
            frmmain.PANEL.Speed = hscrmain(Index).Value
        Case 1 'Time
            frmmain.TimerBig.Interval = hscrmain(Index).Value * 10
            lblvalue(Index).Caption = frmmain.TimerBig.Interval & " MS"
        Case 2 'Moves
            frmmain.PANEL.Moves = hscrmain(Index).Value
            'frmmain.PANEL_MovesChainged  hscrmain(index).value
        Case 3 'AI
            frmmain.TimerAI.Interval = hscrmain(Index).Value * 10
            lblvalue(Index).Caption = frmmain.TimerAI.Interval & " MS"
        Case 4 'AI intelligence level
        Case 5 'Pixel increment
            frmmain.PANEL.PixelSpeed = hscrmain(Index).Value
            frmmain.PANEL2.PixelSpeed = hscrmain(Index).Value
        Case 6 'Ghost duration
            frmmain.PANEL.Ghost = hscrmain(Index).Value
            frmmain.PANEL2.Ghost = hscrmain(Index).Value
    End Select
End Sub

'Load a section from the help file (loaded previously)
Private Sub lstmain_Click()
    On Error Resume Next
    If lstmain.ListIndex > -1 Then
        txtmain = tempstr(lstmain.ListIndex)
        lstmain.Visible = False
    End If
End Sub

