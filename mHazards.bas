Attribute VB_Name = "mHazards"
Public Type HazardObject
    Top As Single
    Firing As Boolean
    FireTicker As Integer
    OnScreen As Boolean
End Type

Public Type AsteroidObject
    OnScreen As Boolean
    Left As Single
    Top As Single
    AnimFrames As Integer
    FrameNo As Integer
    FrameCounter As Integer
    HitLimit As Integer
    Hit As Integer
    Width As Integer
    Height As Integer
    SurfNo As Integer
End Type

Public Type BigGunObject
    Left As Single
    Hit As Integer
    Done As Boolean
    CenterX As Single
    CenterY As Single
    Shot(8) As AmmoObject
    CollRad As Integer
End Type

Public Type BigGunMotion
    OnScreen As Boolean
    Top As Single
    FrameNo As Integer
    FrameCounter As Integer
End Type

Public BigGunLeft As BigGunObject
Public BigGunRight As BigGunObject
Public BigGunMove As BigGunMotion

Public LaserBeam As HazardObject
Public Barrier As HazardObject
Public BallShooter(3) As EnemyObject
Public Asteroid(30) As AsteroidObject
Public blnAstOn As Boolean
Dim recDisplay As RECT

Public Sub ShowHazard()

        If LaserBeam.OnScreen Then
            ShowLaserBeam
        ElseIf blnAstOn Then
            ShowAsteroids
        ElseIf Barrier.OnScreen Then
            ShowBarrier
        ElseIf BigGunMove.OnScreen Then
            ShowBigGuns
        End If

End Sub

Private Sub ShowLaserBeam()
Dim i As Integer
Dim ChkTime As Single
Dim OuterColor As Long
Dim InnerColor As Long
Dim ShowTop As Single
    
        With LaserBeam
            .Top = .Top + 1
            .FireTicker = .FireTicker + 1
            If .FireTicker = 220 Then
                .Firing = True
            ElseIf .FireTicker = 240 Then
                .Firing = False
                .FireTicker = 0
            End If
            
            recDisplay.Top = 0
            recDisplay.Bottom = 19
            If .Top < 0 Then
                ShowTop = 0
                recDisplay.Top = .Top * -1
            ElseIf .Top + 19 > ScrGame.Bottom Then
                ShowTop = .Top
                recDisplay.Bottom = ScrGame.Bottom - .Top
            Else
                ShowTop = .Top
            End If

            If .Top > ScrGame.Bottom Then .OnScreen = False
            
            If .Firing Then
                InnerColor = RGB(255, 100 - Abs(.FireTicker - 226) * 7, 255 - Abs(.FireTicker - 226) * 17)
                OuterColor = RGB(255 - Abs(.FireTicker - 226) * 17, 0, 100 - Abs(.FireTicker - 226) * 7)
                backbuffer.SetForeColor OuterColor
                backbuffer.DrawLine 165, .Top + 7, 860, .Top + 7
                backbuffer.DrawLine 165, .Top + 8, 860, .Top + 8
                backbuffer.SetForeColor InnerColor
                backbuffer.DrawLine 165, .Top + 9, 860, .Top + 9
                backbuffer.DrawLine 165, .Top + 10, 860, .Top + 10
                TestForLaserBeamHit
            End If
            
            recDisplay.Left = 0
            recDisplay.Right = 19
            backbuffer.BltFast 150, ShowTop, ddsLaserCannon, recDisplay, DDBLTFAST_SRCCOLORKEY Or DDBLTFAST_WAIT
            recDisplay.Left = 19
            recDisplay.Right = 38
            backbuffer.BltFast 855, ShowTop, ddsLaserCannon, recDisplay, DDBLTFAST_SRCCOLORKEY Or DDBLTFAST_WAIT
        End With

End Sub

Private Sub ShowAsteroids()
Dim i As Integer
    
    recDisplay.Bottom = 40
    For i = 0 To UBound(Asteroid)
        With Asteroid(i)
            If .OnScreen Then
                .Top = .Top + 1
                .FrameCounter = .FrameCounter + 1
                If .FrameCounter = 2 Then
                    If .FrameNo < .AnimFrames - 1 Then
                        .FrameNo = .FrameNo + 1
                    Else
                        .FrameNo = 0
                    End If
                    .FrameCounter = 0
                End If
                recDisplay.Left = .FrameNo * .Width
                recDisplay.Right = recDisplay.Left + .Width
                TestShotHitAsteroid i
                TestAsteroidHitShip i
                If .Hit >= .HitLimit Then
                    GameSounds.play_snd 0, True
                    .OnScreen = False
                    StartExplosion .Left + .Width / 2, .Top + .Height / 2, 2
                End If
                If .Top > ScrGame.Bottom Then .OnScreen = False
                If .SurfNo = 0 Then
                    backbuffer.BltFast .Left, .Top, ddsAsteroid(0), recDisplay, DDBLTFAST_SRCCOLORKEY Or DDBLTFAST_WAIT
                ElseIf .SurfNo = 1 Then
                    backbuffer.BltFast .Left, .Top, ddsAsteroid(1), recDisplay, DDBLTFAST_SRCCOLORKEY Or DDBLTFAST_WAIT
                End If
            End If
        End With
    Next i

    blnAstOn = False
    For i = 0 To UBound(Asteroid)
        If Asteroid(i).OnScreen Then
            blnAstOn = True
            Exit For
        End If
    Next i

End Sub

Private Sub ShowBarrier()
Dim ShowTop As Single
Dim k As Integer, a As Integer

    Barrier.Top = Barrier.Top + 1
    If Barrier.Top <= 0 Then
        ShowTop = 0
        recBarrier.Top = Barrier.Top * -1
    ElseIf Barrier.Top + 85 > ScrGame.Bottom Then
        ShowTop = Barrier.Top
        recBarrier.Bottom = ScrGame.Bottom - Barrier.Top - 0.5
    Else
        ShowTop = Barrier.Top
        recBarrier.Top = 0
        recBarrier.Bottom = 85
    End If
    backbuffer.BltFast 150, ShowTop, ddsBarrier, recBarrier, DDBLTFAST_SRCCOLORKEY Or DDBLTFAST_WAIT
    For k = 0 To 3
        If BallShooter(k).OnScreen Then
            BallShooter(k).Top = BallShooter(k).Top + 1
            BallShooter(k).AnimCounter = BallShooter(k).AnimCounter + 1
            If BallShooter(k).AnimCounter = 3 Then
                If BallShooter(k).FrameNo < 19 Then
                    BallShooter(k).FrameNo = BallShooter(k).FrameNo + 1
                    BallShooter(k).ImgRECT.Left = BallShooter(k).ImgRECT.Left + 25
                    BallShooter(k).ImgRECT.Right = BallShooter(k).ImgRECT.Left + 25
                Else
                    BallShooter(k).FrameNo = 0
                    BallShooter(k).ImgRECT.Left = 0
                    BallShooter(k).ImgRECT.Right = 25
                End If
                BallShooter(k).AnimCounter = 0
            End If
            If BallShooter(k).Top <= 0 Then
                ShowTop = 0
                BallShooter(k).ImgRECT.Top = BallShooter(k).Top * -1
            ElseIf BallShooter(k).Top + 25 > ScrGame.Bottom Then
                ShowTop = BallShooter(k).Top
                BallShooter(k).ImgRECT.Bottom = ScrGame.Bottom - BallShooter(k).Top
            Else
                ShowTop = BallShooter(k).Top
                BallShooter(k).ImgRECT.Top = 0
                BallShooter(k).ImgRECT.Bottom = 25
            End If
            backbuffer.BltFast BallShooter(k).Left, ShowTop, ddsBallShooter, BallShooter(k).ImgRECT, DDBLTFAST_SRCCOLORKEY Or DDBLTFAST_WAIT
            BallShooter(k).FireTicker = BallShooter(k).FireTicker + 1
            If BallShooter(k).FireTicker = 50 Then
                For a = 0 To 6
                    If BallShooter(k).Ammo(a).Fired = False Then
                        BallShooter(k).Ammo(a).Fired = True
                        BallShooter(k).AimAngle = Int(Rnd * 3) + 1
                        BallShooter(k).Ammo(a).xSpan = 8
                        If BallShooter(k).AimAngle = 1 Then
                            BallShooter(k).Ammo(a).XMove = 4
                            BallShooter(k).Ammo(a).xCtr = BallShooter(k).Left + 25
                            BallShooter(k).Ammo(a).Y = BallShooter(k).Top + 30
                        ElseIf BallShooter(k).AimAngle = 2 Then
                            BallShooter(k).Ammo(a).XMove = -4
                            BallShooter(k).Ammo(a).xCtr = BallShooter(k).Left
                            BallShooter(k).Ammo(a).Y = BallShooter(k).Top + 30
                        ElseIf BallShooter(k).AimAngle = 3 Then
                            BallShooter(k).Ammo(a).XMove = 0
                            BallShooter(k).Ammo(a).xCtr = BallShooter(k).Left + 12
                            BallShooter(k).Ammo(a).Y = BallShooter(k).Top + 35
                        End If
                        BallShooter(k).Ammo(a).YMove = 6
                        Exit For
                    End If
                Next a
                BallShooter(k).FireTicker = 0
            End If
            ShowBallShooterFiring BallShooter(k)
            TestforShotHitShip BallShooter(k)
            TestForShotHitBallShooter BallShooter(k)
            TestForBallShooterHitShip BallShooter(k)
            If BallShooter(k).HitCount = 5 Then
                BallShooter(k).OnScreen = False
                GameSounds.play_snd 0, True
                StartExplosion BallShooter(k).Left + 12, BallShooter(k).Top + 12, 1
            End If
        End If
    Next k
    If Barrier.Top > ScrGame.Bottom Then
        Barrier.OnScreen = False
        For k = 0 To 3
            BallShooter(k).OnScreen = False
        Next k
    End If

End Sub

Private Sub ShowBallShooterFiring(GetEnemy As EnemyObject)
Dim k As Integer

    If GetEnemy.Firing Then
        With GetEnemy
            For k = 0 To 6
                If .Ammo(k).Fired = True Then
                    .Ammo(k).Y = .Ammo(k).Y + .Ammo(k).YMove
                    .Ammo(k).xCtr = .Ammo(k).xCtr + .Ammo(k).XMove
                    LocalLight.ShowLight 6, .Ammo(k).xCtr - .Ammo(k).XMove * 3, .Ammo(k).Y - .Ammo(k).YMove * 3
                    LocalLight.ShowLight 6, .Ammo(k).xCtr - .Ammo(k).XMove * 2, .Ammo(k).Y - .Ammo(k).YMove * 2
                    LocalLight.ShowLight 5, .Ammo(k).xCtr - .Ammo(k).XMove, .Ammo(k).Y - .Ammo(k).YMove
                    LocalLight.ShowLight 3, .Ammo(k).xCtr, .Ammo(k).Y
                    If .Ammo(k).Y > ScrGame.Bottom Or .Ammo(k).xCtr > 874 Or .Ammo(k).xCtr < 150 Then
                        .Ammo(k).Fired = False
                    End If
                End If
            Next k
        End With
    End If

End Sub
Public Sub ResetHazard()

    If LaserBeam.OnScreen Then
        LaserBeam.OnScreen = False
    ElseIf blnAstOn Then
        Dim i As Integer
        For i = 0 To UBound(Asteroid)
            Asteroid(i).OnScreen = False
        Next i
        blnAstOn = False
    ElseIf Barrier.OnScreen Then
        Barrier.OnScreen = False
    ElseIf BigGunMove.OnScreen Then
        BigGunMove.OnScreen = False
    End If

End Sub

Public Sub InitHazard(GetHzNo As Integer)

    If GetHzNo = 0 Then
        LaserBeam.Top = -19
        LaserBeam.Firing = False
        LaserBeam.FireTicker = 0
        LaserBeam.OnScreen = True
    ElseIf GetHzNo = 1 Then
        recAsteroid(0).Right = 40
        recAsteroid(1).Right = 40
        blnAstOn = True
        Dim i As Integer
        For i = 0 To UBound(Asteroid)
            With Asteroid(i)
                If i / 2 - Int(i / 2) = 0 Then
                    .SurfNo = 0
                Else
                    .SurfNo = 1
                End If
                .OnScreen = True
                .AnimFrames = 15
                .FrameCounter = 0
                .FrameNo = Int(Rnd * 14)
                .Height = 40
                .Width = 40
                .Hit = 0
                .HitLimit = 8
                .Left = GameCtr + ((i - 15) * Int(Rnd * 4 + 20)) - .Width / 2
                .Top = -400 + Int(Rnd * 350)
            End With
        Next i
    ElseIf GetHzNo = 2 Then
        Barrier.Top = -85
        recBarrier.Top = 0
        recBarrier.Bottom = 85
        Barrier.OnScreen = True
        BallShooter(0).Left = 225
        BallShooter(0).FrameNo = 0
        BallShooter(0).FireTicker = 0
        BallShooter(1).Left = 362
        BallShooter(1).FrameNo = 5
        BallShooter(1).FireTicker = 30
        BallShooter(2).Left = 635
        BallShooter(2).FrameNo = 10
        BallShooter(2).FireTicker = 15
        BallShooter(3).Left = 774
        BallShooter(3).FrameNo = 15
        BallShooter(3).FireTicker = 45
        Dim j As Integer
        For k = 0 To 3
            BallShooter(k).CanBeHit = True
            BallShooter(k).AnimCounter = 0
            BallShooter(k).ImgRECT.Top = 0
            BallShooter(k).ImgRECT.Right = 25
            BallShooter(k).ImgRECT.Bottom = 25
            BallShooter(k).AimAngle = 0
            BallShooter(k).OnScreen = True
            BallShooter(k).Firing = True
            BallShooter(k).Top = -53
            BallShooter(k).HitCount = 0
            For j = 0 To 6
                BallShooter(k).Ammo(j).Fired = False
            Next j
        Next k
    ElseIf GetHzNo = 3 Then
        BigGunMove.Top = -176
        BigGunMove.OnScreen = True
        BigGunMove.FrameNo = 0
        
        BigGunLeft.Left = 282
        BigGunLeft.CenterX = 48
        BigGunLeft.CenterY = 43
        BigGunLeft.Hit = 0
        BigGunLeft.Done = False
        BigGunLeft.CollRad = 45
        
        BigGunRight.Left = 620
        BigGunRight.CenterX = 72
        BigGunRight.CenterY = 43
        BigGunRight.Hit = 0
        BigGunRight.Done = False
        BigGunRight.CollRad = 45
        
        For i = 0 To 8
            BigGunLeft.Shot(i).Fired = False
            BigGunRight.Shot(i).Fired = False
        Next i
        
        recGunStation.Top = 0
        recGunStation.Bottom = 166
        recBigGun.Top = 0
        recBigGun.Bottom = 120
    End If

End Sub

Private Sub ShowBigGuns()
Dim ShowTop As Single, ShowBackTop As Single
Dim i As Integer

        If BigGunMove.Top < ScrGame.Bottom Then
            BigGunMove.Top = BigGunMove.Top + 1
            If BigGunMove.Top <= 0 Then
                ShowBackTop = 0
                recGunStation.Top = BigGunMove.Top * -1 + 0.5
                If BigGunMove.Top <= -157 Then
                    ShowTop = 0
                    recBigGun.Top = 120
                ElseIf BigGunMove.Top < -37 Then
                    ShowTop = 0
                    recBigGun.Top = 120 - (157 + BigGunMove.Top)
                Else
                    ShowTop = BigGunMove.Top + 37
                    recBigGun.Top = 0
                End If
            ElseIf BigGunMove.Top + 157 > ScrGame.Bottom Then
                ShowBackTop = BigGunMove.Top
                recGunStation.Bottom = ScrGame.Bottom - BigGunMove.Top - 0.5
                ShowTop = ShowBackTop + 37
                recBigGun.Bottom = (ScrGame.Bottom - BigGunMove.Top) - 37.5
            ElseIf BigGunMove.Top + 166 > ScrGame.Bottom Then
                ShowBackTop = BigGunMove.Top
                recGunStation.Bottom = ScrGame.Bottom - BigGunMove.Top - 0.5
                ShowTop = ShowBackTop + 37
            Else
                ShowBackTop = BigGunMove.Top
                ShowTop = ShowBackTop + 37
                recBigGun.Top = 0
                recBigGun.Bottom = 120
            End If
            BigGunMove.FrameCounter = BigGunMove.FrameCounter + 1
            If BigGunMove.FrameCounter = 5 Then
                BigGunMove.FrameNo = BigGunMove.FrameNo + 1
                If BigGunMove.FrameNo = 20 Then BigGunMove.FrameNo = 0
                BigGunMove.FrameCounter = 0
            End If
            
            recBigGun.Left = BigGunMove.FrameNo * 120
            recBigGun.Right = recBigGun.Left + 120
            
            backbuffer.BltFast 150, ShowBackTop, ddsGunStation, recGunStation, DDBLTFAST_SRCCOLORKEY Or DDBLTFAST_WAIT
            If Not BigGunLeft.Done Then
                backbuffer.BltFast BigGunLeft.Left, ShowTop, ddsBigGunLeft, recBigGun, DDBLTFAST_SRCCOLORKEY Or DDBLTFAST_WAIT
            End If
            If Not BigGunRight.Done Then
                backbuffer.BltFast BigGunRight.Left, ShowTop, ddsBigGunRight, recBigGun, DDBLTFAST_SRCCOLORKEY Or DDBLTFAST_WAIT
            End If
            
            TestForShotHitBigGun BigGunLeft
            If Not BigGunLeft.Done And BigGunLeft.Hit = 8 Then
                BigGunLeft.Done = True
                GameSounds.play_snd 4, True
                StartExplosion BigGunLeft.Left + BigGunLeft.CenterX, BigGunMove.Top + BigGunLeft.CenterY + 37, 4
            End If
            
            TestForShotHitBigGun BigGunRight
            If Not BigGunRight.Done And BigGunRight.Hit = 8 Then
                BigGunRight.Done = True
                GameSounds.play_snd 4, True
                StartExplosion BigGunRight.Left + BigGunRight.CenterX, BigGunMove.Top + BigGunRight.CenterY + 37, 4
            End If
        
        Else
            BigGunMove.OnScreen = False
        End If
        
        ShowBigGunShots
        
        If BigGunMove.FrameNo = 0 And BigGunMove.FrameCounter = 0 Then
            If Not BigGunLeft.Done Then
                For i = 0 To 8
                    With BigGunLeft.Shot(i)
                        If .Fired = False Then
                            .Fired = True
                            .xLeft = BigGunLeft.Left + 18
                            .Y = BigGunMove.Top + 132
                            .XMove = 0
                            .YMove = 15
                            With BigGunLeft.Shot(i + 1)
                                .Fired = True
                                .xLeft = BigGunLeft.Left + 70
                                .Y = BigGunMove.Top + 132
                                .XMove = 0
                                .YMove = 15
                            End With
                            Exit For
                        End If
                    End With
                Next i
            End If
                    
            If Not BigGunRight.Done Then
                For i = 0 To 8
                    With BigGunRight.Shot(i)
                        If .Fired = False Then
                            .Fired = True
                            .xLeft = BigGunRight.Left + 40
                            .Y = BigGunMove.Top + 132
                            .XMove = 0
                            .YMove = 15
                            With BigGunRight.Shot(i + 1)
                                .Fired = True
                                .xLeft = BigGunRight.Left + 92
                                .Y = BigGunMove.Top + 132
                                .XMove = 0
                                .YMove = 15
                            End With
                            Exit For
                        End If
                    End With
                Next i
            End If
        
        ElseIf BigGunMove.FrameNo = 9 And BigGunMove.FrameCounter = 0 Then
            If Not BigGunLeft.Done Then
                For i = 0 To 8
                    With BigGunLeft.Shot(i)
                        If .Fired = False Then
                            .Fired = True
                            .xLeft = BigGunLeft.Left + 90
                            .Y = BigGunMove.Top + 135
                            .XMove = 15
                            .YMove = 11
                            With BigGunLeft.Shot(i + 1)
                                .Fired = True
                                .xLeft = BigGunLeft.Left + 105
                                .Y = BigGunMove.Top + 88
                                .XMove = 15
                                .YMove = 11
                            End With
                            Exit For
                        End If
                    End With
                Next i
            End If
                    
            If Not BigGunRight.Done Then
                For i = 0 To 8
                    With BigGunRight.Shot(i)
                        If .Fired = False Then
                            .Fired = True
                            .xLeft = BigGunRight.Left - 7
                            .Y = BigGunMove.Top + 85
                            .XMove = -15
                            .YMove = 11
                            With BigGunRight.Shot(i + 1)
                                .Fired = True
                                .xLeft = BigGunRight.Left + 20
                                .Y = BigGunMove.Top + 130
                                .XMove = -15
                                .YMove = 11
                            End With
                            Exit For
                        End If
                    End With
                Next i
            End If
                
        End If

End Sub

Private Sub ShowBigGunShots()
Dim i As Integer

For i = 0 To 8
    If Not BigGunLeft.Done Then
        With BigGunLeft.Shot(i)
            If .Fired Then
                .xLeft = .xLeft + .XMove
                .Y = .Y + .YMove
                If .Y > ScrGame.Bottom Then .Fired = False
                LocalLight.ShowLight 10, .xLeft + 8, .Y + 8
                LocalLight.ShowLight 3, .xLeft + 5 + (Int(Rnd * 2) - 1), .Y + 5 + (Int(Rnd * 2) - 1)
                TestForBigGunShotHitShip BigGunLeft.Shot(i)
            End If
        End With
    End If
    
    If Not BigGunRight.Done Then
        With BigGunRight.Shot(i)
            If .Fired Then
                .xLeft = .xLeft + .XMove
                .Y = .Y + .YMove
                If .Y > ScrGame.Bottom Then .Fired = False
                LocalLight.ShowLight 10, .xLeft + 8, .Y + 8
                LocalLight.ShowLight 3, .xLeft + 5 + (Int(Rnd * 2) - 1), .Y + 5 + (Int(Rnd * 2) - 1)
                TestForBigGunShotHitShip BigGunRight.Shot(i)
            End If
        End With
    End If
Next i

End Sub


Public Sub TestForShotHitBigGun(GetGun As BigGunObject)
'collision detection using the circular method
Dim j As Integer
Dim GunCtrX As Single, GunCtrY As Single
Dim DistL As Double, DistR As Double, DistC As Double

    GunCtrX = GetGun.Left + GetGun.CenterX
    GunCtrY = BigGunMove.Top + GetGun.CenterY + 37

    With Player
        For j = 0 To 15
            If .shipAmmo(j).Fired Then
                If Not GetGun.Done And BigGunMove.OnScreen Then
                    DistL = Sqr((.shipAmmo(j).xLeft - GunCtrX) ^ 2 + (.shipAmmo(j).Y - GunCtrY) ^ 2)
                    DistR = Sqr((.shipAmmo(j).xRight - GunCtrX) ^ 2 + (.shipAmmo(j).Y - GunCtrY) ^ 2)
                    If DistL <= GetGun.CollRad Or DistR <= GetGun.CollRad Then
                       LocalLight.ShowLight 2, GunCtrX, GunCtrY
                       GetGun.Hit = GetGun.Hit + Player.LaserPower
                       .shipAmmo(j).Fired = False
                    End If
                End If
            End If
        Next j
    End With

End Sub
