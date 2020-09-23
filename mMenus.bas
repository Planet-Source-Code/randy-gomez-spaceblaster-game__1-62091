Attribute VB_Name = "mMenus"
Public Type objMenuItem
    Left As Single
    Top As Single
    Text As String
    TextLeft As Single
    TextTop As Single
    Hilite As Long
End Type

Public MainMenuItem(4) As objMenuItem
Public PauseMenuItem(1) As objMenuItem
Public BackButton As objMenuItem
Public ActiveItem As Integer
Public FirstTime As Boolean

Dim recDisplay As RECT

Public MenuShowCounter As Long

Public Sub SetupMainMenuItems()
Dim i As Integer
    
    ScrGame.Left = 150
    ScrGame.Right = 874
    
    MenuShowCounter = 0
    ActiveItem = 5
    
    For i = 0 To UBound(MainMenuItem)
        With MainMenuItem(i)
            .Hilite = 0
            .Left = 410
            .Top = 200
            .TextLeft = .Left + 40
            .TextTop = .Top + 5
        End With
    Next i
    
    MainMenuItem(0).Text = "START GAME"
    MainMenuItem(1).Text = "VIEW SCORES"
    MainMenuItem(2).Text = "    OPTIONS"
    MainMenuItem(3).Text = "    CREDITS"
    MainMenuItem(4).Text = "        EXIT"
    
End Sub

Public Sub DoMainMenu()
Dim i As Integer
    
'    If FirstTime Then
        FadeIn = True
        DoFades
'    End If
    
    ShowDoors
    recDisplay.Right = 400: recDisplay.Bottom = 60
    backbuffer.BltFast GameCtr - 200, 100, ddsTitle, recDisplay, DDBLTFAST_SRCCOLORKEY Or DDBLTFAST_WAIT
    
    recDisplay.Right = 200: recDisplay.Bottom = 40
    If MenuShowCounter < 80 Then
       MenuShowCounter = MenuShowCounter + 1
        For i = 0 To UBound(MainMenuItem)
            With MainMenuItem(i)
                .Top = MainMenuItem(0).Top + i * MenuShowCounter
                .TextTop = .Top + 6
                backbuffer.BltFast .Left, .Top, ddsMenuBack, recDisplay, DDBLTFAST_SRCCOLORKEY Or DDBLTFAST_WAIT
                backbuffer.SetForeColor 0
                backbuffer.SetFont MenuFnt
                backbuffer.DrawText .TextLeft, .TextTop, .Text, False
            End With
        Next i
        If MenuShowCounter = 75 Then
            If DoorMove = 2 Then
                GameSounds.stop_snd 2
                GameSounds.play_snd 3
            End If
        End If
    Else
        FirstTime = False
        DoorMove = 0
        For i = 0 To UBound(MainMenuItem)
            If (gintMouseX >= MainMenuItem(0).Left And gintMouseX < MainMenuItem(0).Left + 200) And _
                (gintMouseY >= MainMenuItem(i).Top And gintMouseY < MainMenuItem(i).Top + 40) Then
                If gblnLMouseButton Then
                    ProcessMainMenuItem ActiveItem
                End If
                ActiveItem = i
                MainMenuItem(i).Hilite = 256
            End If
            With MainMenuItem(i)
                backbuffer.BltFast .Left, .Top, ddsMenuBack, recDisplay, DDBLTFAST_SRCCOLORKEY Or DDBLTFAST_WAIT
                If ActiveItem = i Then
                    If .Hilite > 0 Then
                        .Hilite = .Hilite - 16
                    End If
                    backbuffer.SetForeColor RGB(255, 255, .Hilite)
                Else
                    If .Hilite > 0 Then
                        .Hilite = .Hilite - 16
                    End If
                    backbuffer.SetForeColor RGB(.Hilite, .Hilite, 0)
                End If
                backbuffer.SetFont MenuFnt
                backbuffer.DrawText .TextLeft, .TextTop, .Text, False
            End With
        Next i
    End If

    DisplayCursor

End Sub

Public Sub ProcessMainMenuItem(GetItem As Integer)

    MainMenuRunning = False
    
    Select Case GetItem
        Case 0
            CreateBackGrdSurf Level(0).BackGrd
            Level(0).LevelStart = VBA.Timer
            CurrentLevel = 0
            CurrentObject = 0
            DoneLevel = False
            ResetGame = True
            ResetHazard
            ResetEnemies
            ResetPowerUps
            ResetExplosions
            ResetPlayerShots
            Score = 0
            Player.NumLives = 5
            DoorMove = 1
            ShowStatusCounter = 0
            GameSounds.play_snd 2
            Music.StopPlaying
            Music.FileName = App.Path & "\Sounds\gamemusic.mid"
            Music.Play
        Case 1
            SetupBackButton
            ScoreShowCounter = 0
            ScoresRunning = True
        Case 2
            SetupBackButton
            OptionsRunning = True
        Case 3
            SetupBackButton
            CreditsRunning = True
        Case 4
            MainMenuRunning = False
            SplashCounter = 0
            GameRunning = False
            GammaValue = 0
            ExitGameCounter = 100
            ExitGame = True
    End Select

End Sub

Public Sub SetupPauseMenuItems()
Dim k As Integer

    ActiveItem = 0
    MenuShowCounter = 0

    For k = 0 To 1
        With PauseMenuItem(k)
            .Hilite = 0
            .Left = 410
            .Top = 200
            .TextLeft = .Left + 45
            .TextTop = .Top + 7
        End With
    Next k
    
    PauseMenuItem(0).Text = "CONTINUE"
    PauseMenuItem(1).Text = "QUIT GAME"

End Sub

Public Sub DoPauseMenu()
Dim i As Integer

    recDisplay.Top = 0: recDisplay.Left = 0
    recDisplay.Right = 400: recDisplay.Bottom = 60
    backbuffer.BltFast GameCtr - 200, 100, ddsTitle, recDisplay, DDBLTFAST_SRCCOLORKEY Or DDBLTFAST_WAIT
    
    recDisplay.Right = 200: recDisplay.Bottom = 40
    If MenuShowCounter < 40 Then
       MenuShowCounter = MenuShowCounter + 1
        For i = 0 To UBound(PauseMenuItem)
            With PauseMenuItem(i)
                .Top = PauseMenuItem(0).Top + i * MenuShowCounter * 2
                .TextTop = .Top + 6
                backbuffer.BltFast .Left, .Top, ddsMenuBack, recDisplay, DDBLTFAST_SRCCOLORKEY Or DDBLTFAST_WAIT
                backbuffer.SetForeColor 0
                backbuffer.SetFont MenuFnt
                backbuffer.DrawText .TextLeft, .TextTop, .Text, False
            End With
        Next i
    Else
        For i = 0 To UBound(PauseMenuItem)
            If (gintMouseX >= PauseMenuItem(0).Left And gintMouseX < PauseMenuItem(0).Left + 200) And _
                (gintMouseY >= PauseMenuItem(i).Top And gintMouseY < PauseMenuItem(i).Top + 40) Then
                ActiveItem = i
                PauseMenuItem(i).Hilite = 256
                If gblnLMouseButton Then
                    ProcessPauseMenuItem ActiveItem
                End If
            End If
    
            With PauseMenuItem(i)
                backbuffer.BltFast .Left, .Top, ddsMenuBack, recDisplay, DDBLTFAST_SRCCOLORKEY Or DDBLTFAST_WAIT
                If ActiveItem = i Then
                    If .Hilite > 0 Then
                        .Hilite = .Hilite - 16
                    End If
                    backbuffer.SetForeColor RGB(255, 255, .Hilite)
                Else
                    If .Hilite > 0 Then
                        .Hilite = .Hilite - 16
                    End If
                    backbuffer.SetForeColor RGB(.Hilite, .Hilite, 0)
                End If
                backbuffer.SetFont MenuFnt
                backbuffer.DrawText .TextLeft, .TextTop, .Text, False
            End With
        Next i
    End If

    DisplayCursor

End Sub

Public Sub ProcessPauseMenuItem(GetItem As Integer)

    Select Case GetItem
        Case 0
            MainMenuRunning = False
            gintMouseX = 512
            gintMouseY = 700
        Case 1
            ResetEnemies
            MainMenuRunning = True
            InsertNewScore
            SetupMainMenuItems
            DoorMove = 2
            GameSounds.play_snd 2
            Music.StopPlaying
            Music.FileName = App.Path & "\Sounds\menumusic.mid"
            Music.Play
    End Select
    PauseMenuRunning = False

End Sub

Public Sub DoCredits()
Dim fs, f, ts
Dim i As Integer
Dim strText(16) As String
Dim TextPos As Long
   
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set f = fs.getfile(App.Path & "\Credits.txt")
'    Reading = 1, Writing = 2, Appending = 8
'    TristateUseDefault = -2, TristateTrue = -1, TristateFalse = 0
    Set ts = f.OpenAsTextStream(1, -2)

    For i = 0 To UBound(strText)
        strText(i) = ts.Readline
    Next i

    ts.Close
    Set fs = Nothing

    ShowDoors
    recDisplay.Top = 0: recDisplay.Left = 0
    recDisplay.Right = 400: recDisplay.Bottom = 60
    backbuffer.BltFast GameCtr - 200, 100, ddsTitle, recDisplay, DDBLTFAST_SRCCOLORKEY Or DDBLTFAST_WAIT

    backbuffer.SetFillColor RGB(0, 0, 50)
    backbuffer.SetForeColor RGB(100, 100, 255)
    backbuffer.DrawBox 180, 180, 850, 600

    TextPos = 180
    backbuffer.SetForeColor RGB(100, 255, 255)
    
    backbuffer.SetFont CreditsFnt
    For i = 0 To UBound(strText)
        If strText(i) <> "-" Then
            TextPos = TextPos + 20
            backbuffer.DrawText GameCtr - 300, TextPos, strText(i), False
        Else
            TextPos = TextPos + 30
        End If
    Next i

    DoBackButton
    
    DisplayCursor

End Sub

Public Sub SetupBackButton()

    With BackButton
        .Hilite = 0
        .Left = 410
        .Text = "BACK"
        .TextLeft = .Left + 70
        .Top = 630
        .TextTop = .Top + 7
    End With

End Sub

Public Sub DoBackButton()
On Error GoTo DoBackErr

    If gintMouseX >= BackButton.Left And gintMouseX < BackButton.Left + 200 And _
        gintMouseY >= BackButton.Top And gintMouseY < BackButton.Top + 40 Then
        BackButton.Hilite = 256
        If gblnLMouseButton Then
            CreditsRunning = False
            ScoresRunning = False
            If OptionsRunning Then
                WriteOptionValuesToFile
                OptionsRunning = False
            End If
            SetupMainMenuItems
            MainMenuRunning = True
            ActiveItem = 6
        End If
    Else
        If BackButton.Hilite > 0 Then
            BackButton.Hilite = BackButton.Hilite - 16
        End If
    End If
    
    recDisplay.Right = 200: recDisplay.Bottom = 40
    backbuffer.BltFast BackButton.Left, BackButton.Top, ddsMenuBack, recDisplay, DDBLTFAST_SRCCOLORKEY Or DDBLTFAST_WAIT
    backbuffer.SetForeColor RGB(255, 255, BackButton.Hilite)
    backbuffer.SetFont MenuFnt
    backbuffer.DrawText BackButton.TextLeft, BackButton.TextTop, BackButton.Text, False

    Exit Sub
    
DoBackErr:
    MsgBox "Error in DoBackButton procedure"
    EndIt

End Sub
