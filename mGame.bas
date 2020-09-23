Attribute VB_Name = "mGame"
'----main game control
Public IntroRunning As Boolean
Public NumLevels As Integer
Public IntroCounter As Long
Public TextAnimCounter As Long
Public MainMenuRunning As Boolean
Public PauseMenuRunning As Boolean
Public ScoresRunning As Boolean
Public CreditsRunning As Boolean
Public GameRunning As Boolean
Public ResetGame As Boolean
Public ReStartGame As Boolean
Public ResetCounter As Long
Public GameOver As Boolean
Public Score As Long
Public ExitGame As Boolean
Public ExitGameCounter As Long
Public DoneLevel As Boolean
Dim NewBackMade As Boolean
Dim StartPauseMark As Single
Dim EndPauseMark As Single
Public DoorMove As Integer
Dim DispHealthR As Integer
Dim DispShieldR As Integer
Dim DispBombNo As Integer
Public ShowStatusCounter As Integer
Public CurShip As Integer


'----level control
Public Type ObjectLauncher
    Type As Integer     'pre-loaded types to appear in level
    LaunchTime As Single  'pre-loaded time markers for when enemies appear
    ObjectPath As Integer
    IsHazard As Boolean
    NumInGroup As Integer
    PowerUpNo As Integer
End Type

Public Type LevelData
    LevelStart As Single        'marks start time of level
    LevelTimeMark As Single     'checks time since start of level
    NumObjects As Integer
    LevelObject(30) As ObjectLauncher
    BackGrd As String
    ShowStars As Boolean
End Type

Public Level(9) As LevelData
Public CurrentLevel As Integer
Public CurrentObject As Integer
Public ShowLevelPrompt As Boolean

'----graphic effects control
Public BackDC As Long
Public ScreenGamma As New clsGamma
Public GammaValue As Integer
Public FadeIn As Boolean
Public FadeOut As Boolean
Public LocalLight As New clsLighting

'----Windows API graphics functions
Public Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, _
ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, _
ByVal dwRop As Long) As Long

Public Declare Function StretchBlt Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, _
ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, _
ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, _
ByVal dwRop As Long) As Long

'----Windows API keyboard input status function
Public Declare Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer

'----screen font definitions
Public MenuFnt As New StdFont
Public ScoreFnt As New StdFont
Public CreditsFnt As New StdFont
Public StatusFnt As New StdFont

'----turn cursor on/off (will not be implemented until debugging is done)
Public Declare Function ShowCursor Lib "user32" (ByVal bShow As Long) As Long

'----game music and sound control
Public Music As clsMusic
Public GameSpeed As Long
Public GameSounds As New clsDirectSound
Dim recDisplay As RECT

Public Sub Main()
Dim ind As Integer
    
    Randomize
    GetGameSettings         'extract FPS, Sound, Volume settings
    GetLevelData
    GetScores               'extract scores from scores.txt
    SetFrameRate GameSpeed  'set FPS for game
    SetupDDrawScreen        'initialize DirectX7, DirectDraw
    InitSoundsAndMusic      'initalize DirectShow and DirectSound, load sounds and music
    SetupStars              'initialize arrays with random star data
    SetupShip               'setup initial player ship data
    CurShip = 0
    SetupOptions            'load the settings for the options screen
    InitScreenEffects       'load initial settings for lights, screen fonts, gamma
    InitPowerUps
    CreateBackGrdSurf "BACKMENU"
    mDInput.Initialize frmGame
    
    InitEnemyTypes          'load all the enemy data for the game into array
    SetupTitleSequence      'setup initial settings for the intro
    
    IntroRunning = True

    Do While IntroRunning
        mDInput.RefreshMouseState
        PrepareBack
        DoFades
        ShowExplosions
        DoTitleSequence
        primary.Flip Nothing, DDFLIP_NOVSYNC
        UpdateFPS
        DoEvents
    Loop
    
    Do While GameRunning
        mDInput.RefreshMouseState
        If Music.Position = Music.Duration Then
            Music.Position = 0
            Music.Play
        End If
        PrepareBack
        ShowBack
        If MainMenuRunning Then
            DoMainMenu
        ElseIf PauseMenuRunning Then
            DoPauseMenu
        ElseIf ScoresRunning Then
            DoScores
        ElseIf CreditsRunning Then
            DoCredits
        ElseIf OptionsRunning Then
            DoOptions
        Else
            ShowStars
            ShowHazard
            ShowEnemies
            ShowShip
            ShowPowerUps
            ShowExplosions
            UpdateLevel
            If ResetGame Then
                DoResetSequence
            Else
                ShowStatus
            End If
        End If
        DrawSidePanels
        If GammaValue > 0 Then
            ScreenGamma.UpdateGamma GammaValue, GammaValue, GammaValue
            GammaValue = GammaValue - 5
        End If
        primary.Flip Nothing, DDFLIP_NOVSYNC
        UpdateFPS
        DoEvents
    Loop
    
    Do While ExitGame
        DoExitGame
    Loop

End Sub

Public Sub UpdateLevel()
Dim i As Integer
Dim Interval As Single

    With Level(CurrentLevel)
        .LevelTimeMark = VBA.Timer
'        If .LevelTimeMark <= .LevelStart Then          'this tests if user has played
'            .LevelTimeMark = .LevelTimeMark + 86400    'past midnight - adds a day
'        End If
        Interval = .LevelTimeMark - .LevelStart
        If CurrentObject < .NumObjects Then
            If Interval >= .LevelObject(CurrentObject).LaunchTime Then
                If .LevelObject(CurrentObject).IsHazard Then
                    InitHazard .LevelObject(CurrentObject).Type
                Else
                    InitEnemyGroup .LevelObject(CurrentObject).Type, .LevelObject(CurrentObject).ObjectPath, .LevelObject(CurrentObject).NumInGroup, .LevelObject(CurrentObject).PowerUpNo
                End If
                CurrentObject = CurrentObject + 1
            End If
        Else
            DoneLevel = True
            For i = 0 To 4
                If EnGrp(i).Active Then
                    DoneLevel = False
                End If
            Next i
            For i = 0 To 8
                If Explode(i).DoExplosion Then
                    DoneLevel = False
                End If
            Next i
            If DoneLevel Then
                If StartPauseMark = 0 Then
                    StartPauseMark = VBA.Timer
                Else
                    EndPauseMark = VBA.Timer
                    If EndPauseMark > StartPauseMark + 3 Then
                        If CurrentLevel < NumLevels - 1 Then
                            ResetEnemies
                            ResetHazard
                            ResetPowerUps
                            CurrentLevel = CurrentLevel + 1
                            ResetGame = True
                            ShowLevelPrompt = True
                            Level(CurrentLevel).LevelStart = VBA.Timer
                            CurrentObject = 0
                            If NewBackMade = False Then
                                CreateBackGrdSurf Level(CurrentLevel).BackGrd
                                NewBackMade = True
                            End If
                        End If
                    End If
                End If
            End If
        End If
    End With
    
End Sub

Public Function IsKeyDown(AsciiKeyCode As Byte) As Boolean

    If GetKeyState(AsciiKeyCode) < -125 Then IsKeyDown = True

End Function

Public Sub InitScreenEffects()

    LocalLight.CreateLight 32, 1, 0, 1, 1, 0, 255 '     1   explosion light
    LocalLight.CreateLight 64, 1, 1, 1, 1, 0, 100 '     2   intro light
    LocalLight.CreateLight 8, 1, 1, 1, 0.5, 0, 250  '   3   shooting light
    LocalLight.CreateLight 32, 1, 1, 1, 1, 0, 50 '      4   powerup light
    LocalLight.CreateLight 8, 1, 1, 1, 0.5, 0, 100  '   5   shooting light trail 1
    LocalLight.CreateLight 8, 1, 1, 1, 0.5, 0, 50  '    6   shooting light trail 2
    LocalLight.CreateLight 8, 1, 0, 0.2, 1, 0, 100  '   7   enemy shot
    LocalLight.CreateLight 64, 1, 1, 1, 1, 0, 80  '     8   cursor light
    LocalLight.CreateLight 32, 1, 0.5, 1, 1, 0, 200  '  9   shield light
    LocalLight.CreateLight 16, 1, 1, 0.5, 0.5, 30, 255  '   10  big gun shot light
    LocalLight.CreateLight 2, 1, 1, 0.5, 0.75, 0, 255  '         11  spacestation light
    
    ScoreFnt.Size = 20
    ScoreFnt.Name = "Bauhaus 93"
    MenuFnt.Size = 16
    MenuFnt.Name = "Bauhaus 93"
    CreditsFnt.Size = 14
    CreditsFnt.Name = "Times New Roman"
    StatusFnt.Name = "Arial Bold"
    StatusFnt.Size = 14
    

    ScreenGamma.CreateGammaRamp
    GammaValue = -100
    ScreenGamma.UpdateGamma GammaValue, GammaValue, GammaValue

    FadeIn = True

End Sub

Public Sub DoFades()

    If FadeIn = True And GammaValue < 0 Then
        GammaValue = GammaValue + 1
        ScreenGamma.UpdateGamma GammaValue, GammaValue, GammaValue
    Else
        FadeIn = False
    End If
    If FadeOut = True And GammaValue > -100 Then
        GammaValue = GammaValue - 1
        ScreenGamma.UpdateGamma GammaValue, GammaValue, GammaValue
    Else
        FadeOut = False
    End If

End Sub

Public Sub InitSoundsAndMusic()

    Set Music = New clsMusic
    Music.Window = frmGame.hWnd
    
    Music.FileName = App.Path & "\sounds\intmusic.mp3"
    Music.Volume = 0 - (5000 - MusicVolume * 50)
    Music.Balance = 0
    Music.Speed = 1
    Music.Position = 0

    GameSounds.Init_sound frmGame.hWnd
    GameSounds.Load_Sound App.Path & "\sounds\explode.wav"
    GameSounds.Load_Sound App.Path & "\sounds\gun1.wav"
    GameSounds.Load_Sound App.Path & "\sounds\doorsound.wav"
    GameSounds.Load_Sound App.Path & "\sounds\doorclose.wav"
    GameSounds.Load_Sound App.Path & "\sounds\shipexplode.wav"
    GameSounds.Load_Sound App.Path & "\sounds\bombshoot.wav"
    GameSounds.Load_Sound App.Path & "\sounds\bombexplode.wav"
    GameSounds.SetDXVolume SoundVolume
    
End Sub

Public Sub DoExitGame()

    If SplashCounter < 200 Then
        SplashCounter = SplashCounter + 1
        If SplashCounter > 100 Then
            FadeOut = True
        End If
        DoFades
        ScrGame.Left = 0: ScrGame.Right = 1024
        backbuffer.BltColorFill ScrGame, 0
        recDisplay.Right = 215: recDisplay.Bottom = 170
        backbuffer.BltFast GameCtr - recDisplay.Right / 2, 384 - recDisplay.Bottom / 2, ddsSplash, recDisplay, DDBLTFAST_SRCCOLORKEY Or DDBLTFAST_WAIT
        recDisplay.Right = 400: recDisplay.Bottom = 50
        backbuffer.BltFast GameCtr - recDisplay.Right / 2, 500, ddsThanks, recDisplay, DDBLTFAST_SRCCOLORKEY Or DDBLTFAST_WAIT
        primary.Flip Nothing, DDFLIP_NOVSYNC
        UpdateFPS
        Exit Sub
    Else
        ExitGame = False
        EndIt
    End If

End Sub

Public Sub EndIt()
    
    'This line restores you back to your default (windows) resolution.
    Call dd.RestoreDisplayMode
    
    'This tells windows/directX that we no longer want exclusive access
    'to the graphics features/directdraw
    Call dd.SetCooperativeLevel(frmGame.hWnd, DDSCL_NORMAL)
    mDInput.Terminate
    
    SetAllToNothing
    
    'turn the Windows cursor back on
'    ShowCursor 1
    
    
    'Stop the program:
    End

End Sub

Public Sub SetAllToNothing()
Dim i As Integer

    SetSurfsToNothing
    
    Music.StopPlaying
    Set Music = Nothing
    Set GameSounds = Nothing
    
    Set LocalLight = Nothing
    Set DX = Nothing
    Set frmGame = Nothing
    
    Set ScoreFnt = Nothing
    Set MenuFnt = Nothing
    Set CreditsFnt = Nothing
    
    
End Sub


Public Sub DoResetSequence()
Dim strText As String, TextLeft As Long, TextColour As Long
Dim j As Integer
Dim BlankRec As RECT

    Player.OnScreen = False
    If DoneLevel Then
        BlankRec.Left = 150
        BlankRec.Right = 874
        BlankRec.Top = ResetCounter * 1.875
        BlankRec.Bottom = 768 - ResetCounter * 1.875
        backbuffer.BltColorFill BlankRec, 100
    End If
    
    
    If ResetCounter / 12 - Int(ResetCounter / 12) = 0 Then
        recDisplay.Left = 0: recDisplay.Right = 60: recDisplay.Bottom = 60
        backbuffer.BltFast Player.Left, Player.Top, ddsShip(CurShip), recDisplay, DDBLTFAST_SRCCOLORKEY Or DDBLTFAST_WAIT
    End If
    
    If ResetCounter = 0 Then
        ResetPlayerShots
        SetupShip
        ResetCounter = ResetCounter + 1
    ElseIf ResetCounter < 200 Then
        ResetCounter = ResetCounter + 1
        backbuffer.SetFont MenuFnt
        If ResetCounter > 50 Then
            If Player.NumLives > 0 Then
                TextLeft = GameCtr - 35
                TextColour = RGB(0, 255, 0)
                If ResetCounter < 150 Then
                    If ShowLevelPrompt Then
                        strText = "LEVEL " & CurrentLevel + 1
                    Else
                        strText = "READY?"
                    End If
                Else
                    strText = "   GO!"
                    ShowLevelPrompt = False
                End If
            Else
                TextLeft = GameCtr - 60
                TextColour = vbRed
                strText = "GAME OVER"
            End If
            recDisplay.Left = 0: recDisplay.Top = 0
            recDisplay.Right = 200: recDisplay.Bottom = 40
            backbuffer.BltFast GameCtr - 100, ScrGame.Bottom / 2 - 20, ddsMenuBack, recDisplay, DDBLTFAST_SRCCOLORKEY Or DDBLTFAST_WAIT
            backbuffer.SetForeColor TextColour
            backbuffer.DrawText TextLeft, ScrGame.Bottom / 2 - 13, strText, False
        End If
    Else
        If GameOver Then
            ResetEnemies
            InsertNewScore
            DoorMove = 2
            GameSounds.play_snd 2
            SetupMainMenuItems
            ScoresRunning = True
            SetupBackButton
            GameOver = False
        End If
        DoneLevel = False
        NewBackMade = False
        Player.OnScreen = True
        Player.Hit = 0
        ResetGame = False
        ResetCounter = 0
        gintMouseX = 512
        gintMouseY = 700
    End If

    ShowStatus
    ShowDoors

End Sub

Public Sub ShowStatus()
Dim k As Integer, strScore As String
    
    recDisplay.Right = 30: recDisplay.Bottom = 30
    For k = 1 To Player.NumLives
        backbuffer.BltFast ScrGame.Left + k * 37 - 30, 10, ddsSmShip(CurShip), recDisplay, DDBLTFAST_SRCCOLORKEY Or DDBLTFAST_WAIT
    Next k
    
    If Score < 10 Then strScore = "0000" & Int(Score)
    If Score >= 10 Then strScore = "000" & Int(Score)
    If Score >= 100 Then strScore = "00" & Int(Score)
    If Score >= 1000 Then strScore = "0" & Int(Score)
    If Score >= 10000 Then strScore = Int(Score)
    
    backbuffer.SetFont ScoreFnt
    
    backbuffer.SetForeColor RGB(130, 130, 0)
    backbuffer.DrawText GameCtr + 20, 12, "Level " & CurrentLevel + 1, False
    backbuffer.SetForeColor RGB(200, 200, 0)
    backbuffer.DrawText GameCtr + 23, 10, "Level " & CurrentLevel + 1, False
    
    backbuffer.SetForeColor RGB(130, 130, 0)
    backbuffer.DrawText GameCtr - 103, 12, strScore, False
    backbuffer.SetForeColor RGB(200, 200, 0)
    backbuffer.DrawText GameCtr - 100, 10, strScore, False
    

    recDisplay.Right = 200: recDisplay.Bottom = 65
    If ShowStatusCounter < 70 Then
        ShowStatusCounter = ShowStatusCounter + 1
        If ShowStatusCounter <= 65 Then
            recDisplay.Top = 65 - ShowStatusCounter
            backbuffer.BltFast GameCtr + 155, 0, ddsHBarBack, recDisplay, DDBLTFAST_SRCCOLORKEY Or DDBLTFAST_WAIT
        Else
            recDisplay.Top = 0
            backbuffer.BltFast GameCtr + 155, ShowStatusCounter - 65, ddsHBarBack, recDisplay, DDBLTFAST_SRCCOLORKEY Or DDBLTFAST_WAIT
        End If
        Exit Sub
    End If
    
    backbuffer.BltFast GameCtr + 155, 5, ddsHBarBack, recDisplay, DDBLTFAST_SRCCOLORKEY Or DDBLTFAST_WAIT
    
    If Player.Hit < 7 Then
        If DispHealthR > (6 - Player.Hit) / 6 * 100 Then
            DispHealthR = DispHealthR - 1
        ElseIf DispHealthR < (6 - Player.Hit) / 6 * 100 Then
            DispHealthR = DispHealthR + 1
        End If
        backbuffer.SetForeColor RGB(0, 200, 200)
        backbuffer.SetFillColor RGB(0, 100, 100)
        backbuffer.DrawBox GameCtr + 240, 17, (GameCtr + 240) + DispHealthR, 26
    End If

    If Player.ShieldLife > 0 Then
        If DispShieldR > (Player.ShieldLife / 10) * 100 Then
            DispShieldR = DispShieldR - 1
        ElseIf DispShieldR < (Player.ShieldLife / 10) * 100 Then
            DispShieldR = DispShieldR + 1
        End If
        backbuffer.SetForeColor RGB(200, 0, 150)
        backbuffer.SetFillColor RGB(100, 0, 50)
        backbuffer.DrawBox GameCtr + 240, 33, (GameCtr + 240) + DispShieldR, 42
    End If
    
    If Player.GotBombs And BombsLeft > 0 Then
        backbuffer.SetForeColor RGB(0, 200, 150)
        backbuffer.SetFillColor RGB(0, 100, 50)
        For k = 0 To BombsLeft - 1
            backbuffer.DrawBox GameCtr + 240 + k * 28, 49, (GameCtr + 257) + k * 28, 58
        Next k
    End If
    
End Sub

Public Sub GetLevelData()
On Error GoTo CantLoad
Dim fs, f, ts
Dim i As Integer, j As Integer
Dim strLine As String, LineLabel As String
Dim LevelNumber As Integer
Dim ObjectNumber As Integer
   
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set f = fs.getfile(App.Path & "\levels.txt")
'    Reading = 1, Writing = 2, Appending = 8
'    TristateUseDefault = -2, TristateTrue = -1, TristateFalse = 0
    Set ts = f.OpenAsTextStream(1, -2)

    strLine = ts.Readline
    strLine = Right(strLine, Len(strLine) - 10)
    NumLevels = Val(strLine)
    strLine = ts.Readline
    
    For i = 1 To NumLevels
        strLine = ts.Readline
        LineLabel = Left(strLine, 7)
        If LineLabel = "LEVELNO" Then
            LevelNumber = Val(Right(strLine, Len(strLine) - 8)) - 1
            strLine = ts.Readline
            LineLabel = Left(strLine, 7)
            If LineLabel = "BACKGRD" Then
                Level(LevelNumber).BackGrd = Right(strLine, Len(strLine) - 8)
            End If
            strLine = ts.Readline
            LineLabel = Left(strLine, 7)
            If LineLabel = "SHSTARS" Then
                Level(LevelNumber).ShowStars = IIf(Right(strLine, Len(strLine) - 8) = "F", False, True)
            End If
            strLine = ts.Readline
            LineLabel = Left(strLine, 7)
            If LineLabel = "NUMBOBJ" Then
                Level(LevelNumber).NumObjects = Val(Right(strLine, Len(strLine) - 8))
                strLine = ts.Readline
                For j = 0 To Level(LevelNumber).NumObjects - 1
                    strLine = ts.Readline
                    LineLabel = Left(strLine, 7)
                    If LineLabel = "OBJDATA" Then
                        ObjectNumber = Val(Mid(strLine, 9, 2)) - 1
                        With Level(LevelNumber)
                            .LevelObject(ObjectNumber).Type = Val(Mid(strLine, 12, 2)) - 1
                            .LevelObject(ObjectNumber).LaunchTime = Val(Mid(strLine, 15, 3))
                            .LevelObject(ObjectNumber).ObjectPath = Val(Mid(strLine, 19, 2)) - 1
                            .LevelObject(ObjectNumber).NumInGroup = Val(Mid(strLine, 22, 2))
                            .LevelObject(ObjectNumber).IsHazard = IIf(Mid(strLine, 25, 1) = "F", False, True)
                            .LevelObject(ObjectNumber).PowerUpNo = Val(Mid(strLine, 27, 2))
                        End With
                    End If
                Next j
            End If
        End If
        strLine = ts.Readline
    Next i
    
    ts.Close
    Set fs = Nothing
    Exit Sub
    
CantLoad:
    MsgBox "Error loading level data."
    End

End Sub

Private Sub GetGameSettings()
On Error GoTo CantLoad
Dim fs, f, ts
Dim strLine As String
   
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set f = fs.getfile(App.Path & "\settings.txt")
'    Reading = 1, Writing = 2, Appending = 8
'    TristateUseDefault = -2, TristateTrue = -1, TristateFalse = 0
    Set ts = f.OpenAsTextStream(1, -2)

    strLine = ts.Readline
    strLine = Right(strLine, Len(strLine) - 4)
    GameSpeed = Val(strLine)
    strLine = ts.Readline
    strLine = Right(strLine, Len(strLine) - 9)
    SoundVolume = Val(strLine)
    strLine = ts.Readline
    strLine = Right(strLine, Len(strLine) - 9)
    MusicVolume = Val(strLine)

    ts.Close
    Set fs = Nothing
    Exit Sub


CantLoad:
    MsgBox "Error loading info from settings file."
    End
    
End Sub
