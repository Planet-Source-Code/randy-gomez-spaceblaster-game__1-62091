Attribute VB_Name = "mIntro"
Private Type IntroShotObject
    Fired As Boolean
    Left As Single
    Top As Single
    Mov As Single
End Type

Private Type IntroShipObject
    Counter As Long
    Top As Single
    Left As Single
    Angle As Single
    ImgNo As Integer
    FireCounter As Integer
    Shot(5) As IntroShotObject
    Trails(10) As TrailObject
End Type

Dim IntroShip As IntroShipObject

Dim recIntroPanelL As RECT
Dim recIntroPanelR As RECT

Public SplashCounter As Integer
Dim StationImgCounter As Integer
Dim recDisplay As RECT
Dim recRG As RECT

Dim TitleBackTop As Single

Public Sub SetupTitleSequence()

'    recIntroShip.Right = 400
    IntroShip.Top = ScrGame.Bottom
    IntroShip.Left = GameCtr - 200
    recIntroPanelL.Right = 150
    recIntroPanelL.Bottom = 768
    recIntroPanelR.Right = 150
    recIntroPanelR.Bottom = 768
    recRG.Right = 200
    recRG.Bottom = 40

End Sub

Public Sub DoTitleSequence()
Dim j As Integer
Dim TitleDC As Long


    If SplashCounter < 500 Then
        SplashCounter = SplashCounter + 1
        If SplashCounter > 300 Then
            FadeOut = True
        End If
        recDisplay.Right = 215: recDisplay.Bottom = 170
        backbuffer.BltFast GameCtr - recDisplay.Right / 2, 384 - recDisplay.Bottom / 2, ddsSplash, recDisplay, DDBLTFAST_SRCCOLORKEY Or DDBLTFAST_WAIT
        Exit Sub
    ElseIf SplashCounter = 500 Then
        FadeIn = True
        Music.Play
        SplashCounter = 501
    End If

    recDisplay.Right = 1024: recDisplay.Bottom = 468
    If IntroCounter < 1150 Then
        recDisplay.Bottom = Int(IntroCounter / 4)
    End If
    backbuffer.BltFast 0, 768 - Int(IntroCounter / 4), ddsEarth, recDisplay, DDBLTFAST_SRCCOLORKEY Or DDBLTFAST_WAIT
    DoIntroStars
    IntroCounter = IntroCounter + 1
    If IntroCounter < 1100 Then
        If IntroCounter > 560 Then
            With IntroShip
                .Counter = .Counter + 1
                If .Counter <= 48 Then
                    .Top = .Top - 4
                    If .Top + 195 > ScrGame.Bottom Then
                        recIntroShip.Bottom = ScrGame.Bottom - .Top
                    Else
                        recIntroShip.Bottom = 195
                    End If
                ElseIf .Counter = 115 Then
                    GameSounds.play_snd 1, True
                ElseIf .Counter > 115 And .Counter < 125 Then
                    backbuffer.SetForeColor vbCyan
                    backbuffer.DrawLine GameCtr - 100, ScrGame.Bottom - 110, GameCtr - 5, ScrGame.Bottom - 330
                    backbuffer.DrawLine GameCtr - 102, ScrGame.Bottom - 112, GameCtr - 5, ScrGame.Bottom - 330
                    backbuffer.DrawLine GameCtr - 104, ScrGame.Bottom - 114, GameCtr - 5, ScrGame.Bottom - 330
                
                    backbuffer.DrawLine GameCtr + 100, ScrGame.Bottom - 110, GameCtr + 5, ScrGame.Bottom - 330
                    backbuffer.DrawLine GameCtr + 102, ScrGame.Bottom - 108, GameCtr + 5, ScrGame.Bottom - 330
                ElseIf .Counter > 160 Then
                    .Top = .Top + 4
                    recIntroShip.Bottom = ScrGame.Bottom - .Top
                End If

                backbuffer.BltFast .Left, .Top, ddsIntroShip, recIntroShip, DDBLTFAST_SRCCOLORKEY Or DDBLTFAST_WAIT
            
            End With
        End If
        If IntroCounter < 680 Then
            StationImgCounter = StationImgCounter + 1
            recDisplay.Top = 0: recDisplay.Bottom = 300
            recDisplay.Left = 0: recDisplay.Right = 300
            If IntroCounter <= 300 Then
                recDisplay.Right = recDisplay.Left + IntroCounter
            End If
            backbuffer.BltFast GameCtr + 512 - IntroCounter, ScrGame.Bottom / 2 - 170, ddsStation, recDisplay, DDBLTFAST_SRCCOLORKEY Or DDBLTFAST_WAIT
            If StationImgCounter <= 40 Then
                LocalLight.ShowLight 11, (GameCtr + 512 - IntroCounter) + 160, (ScrGame.Bottom / 2 - 170) + 39
                LocalLight.ShowLight 11, (GameCtr + 512 - IntroCounter) + 39, (ScrGame.Bottom / 2 - 170) + 130
                LocalLight.ShowLight 11, (GameCtr + 512 - IntroCounter) + 280, (ScrGame.Bottom / 2 - 170) + 29
            ElseIf StationImgCounter = 80 Then
                StationImgCounter = 0
            End If
        ElseIf IntroCounter = 680 Then
            GameSounds.play_snd 0, True
            LocalLight.ShowLight 2, GameCtr - 20, ScrGame.Bottom - 360
            StartExplosion GameCtr - 20, ScrGame.Bottom - 360, 1
            LocalLight.ShowLight 2, GameCtr + 30, ScrGame.Bottom - 300
            StartExplosion GameCtr + 30, ScrGame.Bottom - 300, 1
        ElseIf IntroCounter = 682 Then
            LocalLight.ShowLight 2, GameCtr + 50, ScrGame.Bottom - 420
            StartExplosion GameCtr + 50, ScrGame.Bottom - 420, 1
        ElseIf IntroCounter = 684 Then
            LocalLight.ShowLight 2, GameCtr - 50, ScrGame.Bottom - 370
            StartExplosion GameCtr - 50, ScrGame.Bottom - 370, 1
        ElseIf IntroCounter > 684 Then
            BackDC = backbuffer.GetDC
            TitleDC = ddsTitle.GetDC
            StretchBlt BackDC, GameCtr - (IntroCounter - 400) / 2, ScrGame.Bottom / 2 + 70 - (IntroCounter - 640) * 1.25, _
                IntroCounter - 380, 20 + (IntroCounter - 480) / 10, _
                    TitleDC, 0, 0, 400, 60, vbSrcPaint
            ddsTitle.ReleaseDC TitleDC
            backbuffer.ReleaseDC BackDC
            If IntroCounter > 830 And IntroCounter < 1060 Then
                If recRG.Top < 360 Then
                    If IntroCounter / 2 - Int(IntroCounter / 2) = 0 Then
                        recRG.Top = recRG.Top + 40
                        recRG.Bottom = recRG.Top + 40
                    End If
                End If
                backbuffer.BltFast GameCtr - 100, ScrGame.Bottom / 2 - 40, ddsRG, recRG, DDBLTFAST_SRCCOLORKEY Or DDBLTFAST_WAIT
            ElseIf IntroCounter > 1060 And IntroCounter < 1080 Then
                If recRG.Top > 0 Then
                    If IntroCounter / 2 - Int(IntroCounter / 2) = 0 Then
                        recRG.Top = recRG.Top - 40
                        recRG.Bottom = recRG.Top + 40
                    End If
                End If
                backbuffer.BltFast GameCtr - 100, ScrGame.Bottom / 2 - 40, ddsRG, recRG, DDBLTFAST_SRCCOLORKEY Or DDBLTFAST_WAIT
            End If
        End If
    ElseIf IntroCounter > 1100 And IntroCounter < 1200 Then
        FadeOut = True
        DoFades
    ElseIf IntroCounter = 1200 Then
        FadeOut = False
    ElseIf IntroCounter = 1202 Then
        FirstTime = True
        SetupMainMenuItems
        IntroRunning = False
        GameRunning = True
        MainMenuRunning = True
        ResetGame = True
        ScrGame.Left = 150
        ScrGame.Right = 874
        Music.StopPlaying
        Music.FileName = App.Path & "\Sounds\menumusic.mid"
        Music.Volume = 0 - (5000 - MusicVolume * 50)
        Music.Play
        Set ddsIntroShip = Nothing
        Set ddsStation = Nothing
        Set ddsEarth = Nothing
    End If

End Sub


Private Sub DoIntroTrails()
Dim j As Integer
Dim TrailsDC As Long

        With IntroShip
            For j = 0 To 10
                If .Trails(j).LifeTime > 0 Then
                    .Trails(j).LifeTime = .Trails(j).LifeTime - 1
                    .Trails(j).X = .Trails(j).X + .Trails(j).XMov
                    .Trails(j).Y = .Trails(j).Y + .Trails(j).YMov
                    If .Trails(j).LifeTime / 2 - Int(.Trails(j).LifeTime / 2) = 0 Then
                        .Trails(j).picno = .Trails(j).picno + 10
                    End If
                    BackDC = backbuffer.GetDC
                    TrailsDC = ddsRocTrails.GetDC
                    BitBlt BackDC, .Trails(j).X + 5, .Trails(j).Y + Int(Rnd * 3), 10, 10, TrailsDC, .Trails(j).picno, 0, vbSrcPaint
                    BitBlt BackDC, .Trails(j).X - 17, .Trails(j).Y + Int(Rnd * 3), 10, 10, TrailsDC, .Trails(j).picno, 0, vbSrcPaint
                    ddsRocTrails.ReleaseDC TrailsDC
                    backbuffer.ReleaseDC BackDC
                Else
                    .Trails(j).X = .Left + (Int(Rnd - 0.5) + 40)
                    .Trails(j).Y = .Top + 75
                    .Trails(j).LifeTime = Int(Rnd * 16)
                    .Trails(j).XMov = Round(Rnd - 0.5, 2)
                    .Trails(j).YMov = Round(Rnd * 2, 2)
                    .Trails(j).picno = 0
                End If
            Next j
        End With
        
End Sub

Public Sub DoIntroStars()
Dim i As Integer
Dim ReturnVal As Long
Dim recDisplay As RECT

       
        For i = 1 To 50
            If i < 11 Then
                FasterStar(i).Y = FasterStar(i).Y - 0.25
                recDisplay.Right = 4: recDisplay.Bottom = 4
                ReturnVal = backbuffer.BltFast(FasterStar(i).X, FasterStar(i).Y, ddsFStar, recDisplay, DDBLTFAST_SRCCOLORKEY Or DDBLTFAST_WAIT)
                If FasterStar(i).Y < 0 Then
                    FasterStar(i).X = Rnd * (ScrGame.Right - ScrGame.Left) + ScrGame.Left
                    FasterStar(i).Y = ScrGame.Bottom
                    FasterStar(i).Move = Rnd * 1.25 + 0.75
                End If
            End If
            SlowStar(i).Y = SlowStar(i).Y - 0.25
            recDisplay.Right = 1: recDisplay.Bottom = 1
            ReturnVal = backbuffer.BltFast(SlowStar(i).X, SlowStar(i).Y, ddsSStar, recDisplay, DDBLTFAST_SRCCOLORKEY Or DDBLTFAST_WAIT)
            
            If SlowStar(i).Y < 0 Then
                SlowStar(i).X = Rnd * (ScrGame.Right - ScrGame.Left) + ScrGame.Left
                SlowStar(i).Y = ScrGame.Bottom
                SlowStar(i).Move = Rnd + 0.5
            End If
        Next i

End Sub
