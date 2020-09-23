Attribute VB_Name = "StoryModeFunctions"
Option Explicit 'This will be used to retreive story data
'http://www.geocities.com/jpop_ayumi/Ayumi_Gallery_Site.html
Public StoryLevel As Long, MusicEmotion As Boolean, AIEmotion As Boolean, AIPIC As Image
Public AICLEAN As Boolean, MUSICCLEAN As Boolean
Public Const IsOnline = 7125, NotOnline = 4140, Maximized = 8835, Minimized = 6675

'HARD CODED TO FRMMAIN (but I dont care right now)
'STORYEmotion is hardcoded to the image box on frmmain
'NewLevel is hardcoded to the timers, AI, pictures, Panels, on and frmmain itself
'STORYLoadBackground is hardcoded to use frmmain's picture

Public Sub NORMALAnyMusic()
    Randomize Timer
    StoryLevel = 0
    MUSICCLEAN = False
    STORYRefreshMusic
End Sub

Public Function PATH(Optional Filename As String = "Story.INI") As String
    PATH = chkpath(App.PATH, Filename)
End Function

Public Function STORYLevels() As Long
    STORYLevels = Val(getvalue(PATH, "Main", "Levels", "0"))
End Function

Public Function STORYProperty(Prop As String, Optional Default As String) As String
    STORYProperty = getvalue(PATH, CStr(StoryLevel), Prop, Default)
End Function

Public Function STORYEmotion(Optional Player As Boolean = True, Optional Tolerance As Double = 0.7) As Boolean 'Player (true=player, false=ai) Emotion (true=Angry, False=Happy)
    If Player = True Then 'HARD CODED TO FRMMAIN (but I dont care right now)
        STORYEmotion = frmmain.PANEL.Top >= frmmain.PANEL.GridHeight * Tolerance
    Else
        STORYEmotion = frmmain.PANEL2.Top >= frmmain.PANEL2.GridHeight * Tolerance
    End If
End Function

Public Function STORYMusicEmotion() As Boolean 'Emotion (true=Angry, False=Happy)
    STORYMusicEmotion = STORYEmotion(True) Or STORYEmotion(False)
End Function

Public Function NewLevel(Optional Level As Long = 1)
    StoryLevel = Level
    AICLEAN = False
    MUSICCLEAN = False
    STORYRefreshAI
    STORYRefreshMusic
    STORYLoadBackground
    
    'HARD CODED TO FRMMAIN (but I dont care right now)
    With frmmain
        .PANEL.TileCount = Val(STORYProperty("Fill", "25"))
        .PANEL.InitGrid
        
        .AI.AILevel Val(STORYProperty("Strategy", "2"))
        .TimerBig.Interval = Val(STORYProperty("Speed", "1000"))
        .TimerAI.Interval = Val(STORYProperty("AISpeed", "500"))

        .imgplayer2.Visible = False

        .TimerSmall.Enabled = True
        .TimerBig.Enabled = True
        .TimerAI.Enabled = True
        .picmain(4).Enabled = True
        .width = IsOnline
    End With
End Function

Public Sub STORYRefreshAI(Optional Emotion As Boolean)
    If AIEmotion <> Emotion Or Not AICLEAN Then
        AICLEAN = True
        AIEmotion = Emotion
        AIPIC.Picture = LoadPicture(PATH(STORYProperty(STORYEmotion2Text(Emotion) & " Picture")))
    End If
End Sub
Public Sub STORYRefreshMusic(Optional Emotion As Boolean)
    If AIEmotion <> MusicEmotion Or Not MUSICCLEAN Then
        MUSICCLEAN = True
        MusicEmotion = Emotion
        'PLAYSONG PATH(STORYProperty(STORYEmotion2Text(Emotion) & " Music"))
    End If
End Sub

Public Function STORYEmotion2Text(Emotion As Boolean) As String
    STORYEmotion2Text = IIf(Emotion, "Angry", "Happy")
End Function

Public Sub AUTOAIEmotion()
    If StoryLevel > 0 Then STORYRefreshAI STORYEmotion(False)
End Sub

Public Sub AUTOMusicEmotion()
    If StoryLevel > 0 Then STORYRefreshMusic STORYMusicEmotion
End Sub

Public Sub STORYAILOSE()
    MsgBox STORYProperty("Lose", "Congratulations")
    StopGame True
End Sub
Public Sub STORYAIWIN()
    MsgBox STORYProperty("Win", "I look forward to our next match")
    StopGame
End Sub

Public Sub StopGame(Optional didwin As Boolean)
    With frmmain
        .TimerAI.Enabled = False
        .TimerBig.Enabled = False
        .PANEL.GameOver = True
        .PANEL2.GameOver = True
    End With
    If didwin Then
        StoryLevel = StoryLevel + 1
        If StoryLevel <= STORYLevels Then
            NewLevel StoryLevel
        Else
            MsgBox "You have beaten story mode"
        End If
    Else
        StoryLevel = 0
    End If
End Sub

Public Sub STORYLoadBackground()
    Dim tempstr As String
    tempstr = PATH(STORYProperty("Default Picture"))
    If FileExists(tempstr) Then 'HARD CODED TO FRMMAIN (but I dont care right now)
        frmmain.Picture = LoadPicture(tempstr)
    End If
End Sub

'INIfile access
Private Function issection(value As String) As Boolean
    If Left(value, 1) = "[" And Right(value, 1) = "]" And stripsection(value) <> Empty Then issection = True Else issection = False
End Function

Private Function isvalue(value As String) As Boolean
    If issection(value) = False And InStr(value, "=") > 0 Then isvalue = True Else isvalue = False
End Function

Private Function stripsection(section As String) As String
    stripsection = Mid(section, 2, Len(section) - 2)
End Function

Private Function stripvalue(value As String) As String
    stripvalue = Right(value, Len(value) - InStr(value, "="))
End Function

Private Function stripname(value As String) As String
    stripname = Left(value, InStr(value, "=") - 1)
End Function

Private Function iscomment(value As String) As Boolean
    If Left(value, 1) = "#" Or Left(value, 1) = "'" Then iscomment = True Else iscomment = False
End Function

Public Function getvalue(Filename As String, ByVal section As String, value As String, Optional Default As String = Empty) As String
    On Error Resume Next
    section = LCase(section)
    getvalue = Default
    Dim tempfile As Long, found As Boolean, temp As String, currentsection As String
    If FileExists(Filename) = True Then
        tempfile = FreeFile
        Open Filename For Input As #tempfile
            Do Until EOF(tempfile) Or found = True
                Line Input #tempfile, temp
                If iscomment(temp) = False Then
                    If issection(temp) = True Then
                        currentsection = LCase(stripsection(temp))
                    Else
                        If currentsection = section And isvalue(temp) Then
                            If LCase(stripname(temp)) = LCase(value) Then
                                getvalue = stripvalue(temp)
                                found = True
                            End If
                    End If
                    End If
                End If
            Loop
        Close #tempfile
    End If
End Function

'File access functions
Public Function chkpath(ByVal basehref As String, ByVal url As String) As String
'Debug.Print basehref & " " & URL
Const goback As String = "..\"
Const slash As String = "\"
Dim spoth As Long
If Left(url, 1) = slash Then url = Right(url, Len(url) - 1)
If Right(basehref, 1) = slash And Len(basehref) > 3 Then basehref = Left(basehref, Len(basehref) - 1)
If LCase(url) <> LCase(basehref) And url <> Empty And basehref <> Empty Then
If url Like "?:*" Then 'is absolute
    chkpath = url
Else
    If containsword(url, goback) Then 'is relative
        If containsword(Right(basehref, Len(basehref) - 3), slash) Then
            For spoth = 1 To countwords(url, goback)
                If countwords(basehref, slash) > 0 Then
                    url = Right(url, Len(url) - Len(goback))
                    basehref = Left(basehref, InStrRev(basehref, slash) - 1)
                Else
                    url = Replace(url, goback, "")
                End If
            Next
        Else
            chkpath = Left(basehref, 3)
            Exit Function
            url = Replace(url, goback, "")
        End If
        If Right(basehref, 1) <> slash Then chkpath = basehref & slash & url Else chkpath = basehref & url
    Else 'is additive
        If Right(basehref, 1) <> slash Then chkpath = basehref & slash & url Else chkpath = basehref & url
    End If
End If
End If
End Function

Private Function containsword(text As String, word As String) As Boolean
    containsword = InStr(1, text, word, vbTextCompare) > 0
End Function

Private Function countwords(text As String, word As String) As Long
    Dim temp As Long, count As Long
    temp = InStr(1, text, word, vbTextCompare)
    Do Until temp = 0
        count = count + 1
        temp = InStr(temp + 1, text, word, vbTextCompare)
    Loop
    countwords = count
End Function

Public Function FileExists(Filename As String) As Boolean
On Error Resume Next 'Checks to see if a file exists
Dim temp As Long
temp = GetAttr(Filename)
FileExists = temp > 0
End Function

'Used to make PdP pretend to be a webserver. Code borrowed from Nebula
Public Function SendFile(Filename As String, sock As Winsock, Optional closewhendone As Boolean = True) As Boolean
    On Error GoTo err:
    Filename = Replace(Filename, "/", "\")
    Filename = Replace(Filename, "%20", " ")
    Dim tempfile As Long, filebin As String, filesize As Long, sentsize As Long, temp As Long, issending As Boolean
    Const buffer = 1024
    filesize = FileLen(Filename)
    tempfile = FreeFile
    Open Filename For Binary As #tempfile 'open filename
    filebin = Space(buffer) 'create 1024 byte buffer
    issending = True
    Do Until issending = False
        Get #tempfile, , filebin
        sentsize = sentsize + Len(filebin)
        sock.Tag = Empty
        If sentsize > filesize Then
            temp = sentsize - filesize
            SOCKSend sock, Left(filebin, Len(filebin) - temp)
            issending = False
        Else
            SOCKSend sock, filebin
        End If
        Do Until sock.Tag <> Empty
            DoEvents
        Loop
        If sock.Tag = "False" Then
            closewhendone = True
            GoTo err
        End If
        DoEvents
    Loop
    Close tempfile
    SendFile = True
err:
    If closewhendone Then sock.Close
End Function

'Send data if the winsock is connected
Public Sub SOCKSend(sock As Winsock, ByVal text As String, Optional block As Long = 1024)
    If sock.State = sckConnected Then sock.SendData text & ";"
End Sub
