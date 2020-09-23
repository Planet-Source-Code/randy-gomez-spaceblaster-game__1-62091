Attribute VB_Name = "mBosses"
Public BossExplodeCounter As Integer

Public Sub DoBossFiring(GetEGno As Integer, GetRecLeft As Long, GetRecTop As Long)
Dim k As Integer
Dim recDisplay As RECT

    recDisplay.Left = GetRecLeft
    recDisplay.Top = GetRecTop
    recDisplay.Right = recDisplay.Left + 10
    recDisplay.Bottom = recDisplay.Top + 10
    
    With EnGrp(GetEGno)
        If .Enemy(0).Firing Then
            .Enemy(0).FireTicker = .Enemy(0).FireTicker + 1
            If .Enemy(0).FireTicker = EnemyType(.TypeNo).FirePause Then
                For k = 0 To 6
                    If .Enemy(0).Ammo(k).Fired = False Then
                        .Enemy(0).Ammo(k).Fired = True
                        .Enemy(0).Ammo(k).xSpan = 5
                        .Enemy(0).Ammo(k).YMove = 20
                        .Enemy(0).Ammo(k).XMove = 0
                        .Enemy(0).Ammo(k).xCtr = .Enemy(0).Left + (EnemyType(.TypeNo).Width / 2)
                        .Enemy(0).Ammo(k).Y = .Enemy(0).Top + EnemyType(.TypeNo).Height + 5
                        Exit For
                    End If
                Next k
                .Enemy(0).FireTicker = 0
            End If
        End If

        With .Enemy(0)
            For k = 0 To 6
                If .Ammo(k).Fired = True Then
                    backbuffer.BltFast .Ammo(k).xCtr + 46, .Ammo(k).Y - 23, ddsEnShot, recDisplay, DDBLTFAST_SRCCOLORKEY Or DDBLTFAST_WAIT
                    TestforBossShotHitShip GetEGno, k, .Ammo(k).xCtr + 51, .Ammo(k).Y - 23
                        
                    LocalLight.ShowLight 3, .Ammo(k).xCtr + 18 + 5, .Ammo(k).Y
                    TestforBossShotHitShip GetEGno, k, .Ammo(k).xCtr + 23, .Ammo(k).Y
                        
                    LocalLight.ShowLight 3, .Ammo(k).xCtr - 28 + 5, .Ammo(k).Y
                    TestforBossShotHitShip GetEGno, k, .Ammo(k).xCtr - 23, .Ammo(k).Y
                        
                    backbuffer.BltFast .Ammo(k).xCtr - 56, .Ammo(k).Y - 23, ddsEnShot, recDisplay, DDBLTFAST_SRCCOLORKEY Or DDBLTFAST_WAIT
                    TestforBossShotHitShip GetEGno, k, .Ammo(k).xCtr - 51, .Ammo(k).Y - 23
                        
                    .Ammo(k).Y = .Ammo(k).Y + .Ammo(k).YMove
                    If .Ammo(k).Y > ScrGame.Bottom Or .Ammo(k).xCtr > 874 Or .Ammo(k).xCtr < 150 Then
                       .Ammo(k).Fired = False
                    End If
                End If
            Next k
        End With

    End With


End Sub

Public Sub TestforBossShotHitShip(GrpNo As Integer, AmmoNo As Integer, AmmoX As Single, AmmoY As Single)
'collision detection using the circular method
Dim Dist As Double
Dim PlayerCtrX As Single, PlayerCtrY As Single
Dim i As Integer

    If Player.OnScreen And Player.Hit < 6 Then
        PlayerCtrX = Player.Left + Player.Width / 2
        PlayerCtrY = Player.Top + Player.Height / 2
    Else
        Exit Sub
    End If

    For i = 0 To 6
        If EnGrp(GrpNo).Enemy(0).Ammo(AmmoNo).Fired Then
            Dist = Sqr((PlayerCtrX - (AmmoX) + 5) ^ 2 + (PlayerCtrY - (AmmoY + 5)) ^ 2)
            If Dist <= Player.CollRad Then
               EnGrp(GrpNo).Enemy(0).Ammo(AmmoNo).Fired = False
               If Player.ShieldLife = 0 Then
                   Player.Hit = Player.Hit + 1
                   LocalLight.ShowLight 2, PlayerCtrX, PlayerCtrY
               Else
                   Player.ShieldLife = Player.ShieldLife - 1
               End If
            End If
        End If
    Next i

End Sub

Public Sub DoBossExplosion(GetEGno As Integer, GetENo As Integer, GetShowTop As Single)

    BossExplodeCounter = BossExplodeCounter + 1

With EnGrp(GetEGno)
    If .Enemy(GetENo).OnScreen Then
        If BossExplodeCounter < 100 Then
            If BossExplodeCounter = 1 Then
                GameSounds.play_snd 0
                StartExplosion .Enemy(j).Left + EnemyType(.TypeNo).Width * 0.5, .Enemy(j).Top + EnemyType(.TypeNo).Height * 0.5, 3
            ElseIf BossExplodeCounter = 22 Then
                GameSounds.play_snd 0, True
                StartExplosion .Enemy(j).Left + EnemyType(.TypeNo).Width * 0.33, .Enemy(j).Top + EnemyType(.TypeNo).Height * 0.33, 3
            ElseIf BossExplodeCounter = 36 Then
                GameSounds.play_snd 0, True
                StartExplosion .Enemy(j).Left + EnemyType(.TypeNo).Width * 0.8, .Enemy(j).Top, 3
            ElseIf BossExplodeCounter = 54 Then
                GameSounds.play_snd 0, True
                StartExplosion .Enemy(j).Left + EnemyType(.TypeNo).Width * 0.2, .Enemy(j).Top + EnemyType(.TypeNo).Height * 0.66, 3
            ElseIf BossExplodeCounter = 65 Then
                GameSounds.play_snd 0, True
                StartExplosion .Enemy(j).Left + EnemyType(.TypeNo).Width * 0.75, .Enemy(j).Top + EnemyType(.TypeNo).Height * 0.5, 3
            End If
        Else
            .Enemy(GetENo).OnScreen = False
            GammaValue = 50
            GameSounds.play_snd 6, True
            StartExplosion .Enemy(j).Left + EnemyType(.TypeNo).Width * 0.5, .Enemy(j).Top + EnemyType(.TypeNo).Height * 0.5, 4
            StartExplosion .Enemy(j).Left + EnemyType(.TypeNo).Width * 0.5 - 50, .Enemy(j).Top + EnemyType(.TypeNo).Height * 0.5 - 50, 4
            StartExplosion .Enemy(j).Left + EnemyType(.TypeNo).Width * 0.5 + 60, .Enemy(j).Top + EnemyType(.TypeNo).Height * 0.5 + 40, 4
            StartExplosion .Enemy(j).Left + EnemyType(.TypeNo).Width * 0.5 + 40, .Enemy(j).Top + EnemyType(.TypeNo).Height * 0.5 - 60, 4
            BossExplodeCounter = 0
        End If
    End If
End With

End Sub
