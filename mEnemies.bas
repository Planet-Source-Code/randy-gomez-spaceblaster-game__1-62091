Attribute VB_Name = "mEnemies"
Public Type EnemyTypeObject
    Width As Long
    Height As Long
    HitLimit As Integer
    AnimFrames As Integer
    Weapon As Integer
    CollRad As Single
    ResID As String
    ShotYStart As Single
    FirePause As Integer
    IsBoss As Boolean
End Type

Public EnemyType(15) As EnemyTypeObject

Public Type EnemyObject
    OnScreen As Boolean
    PathAngle As Single
    Left As Single
    Top As Single
    HitCount As Integer
    Firing As Boolean
    FireTicker As Long
    FrameNo As Integer
    AnimCounter As Integer
    PathCounter As Long
    ImgRECT As RECT
    CanBeHit As Boolean
    AimAngle As Single
    Ammo(6) As AmmoObject
End Type

Public Type EnemyGroupObject
    Enemy(11) As EnemyObject
    NumEn As Integer
    TypeNo As Integer
    Active As Boolean
    SurfNo As Integer
    PathNo As Integer
    PowerUp As Integer
    PowerUpHolder As Integer
End Type

Public EnGrp(4) As EnemyGroupObject
Dim recDisplay As RECT

Public Sub InitEnemyTypes()

    With EnemyType(0)       'little green shooting guy
        .AnimFrames = 9: .Width = 50: .Height = 50: .CollRad = 22: .HitLimit = 3
        .Weapon = 1: .FirePause = 40: .ShotYStart = 40: .ResID = "ENEMY0"
    End With

    With EnemyType(1)       'forward rolling ship
        .AnimFrames = 20: .Width = 60: .Height = 60: .CollRad = 22: .HitLimit = 3
        .Weapon = 1: .FirePause = 30: .ShotYStart = 50: .ResID = "ENEMY1"
    End With
    
    With EnemyType(2)       'tracker enemy
        .AnimFrames = 15: .Width = 60: .Height = 60: .CollRad = 20: .HitLimit = 2
        .Weapon = 2: .FirePause = 100: .ShotYStart = 25: .ResID = "ENEMY2"
    End With

    With EnemyType(3)       'eggspin enemy
        .AnimFrames = 20: .Width = 60: .Height = 60: .CollRad = 25: .HitLimit = 3
        .Weapon = 0: .ResID = "ENEMY3"
    End With

    With EnemyType(4)       'green enemy
        .AnimFrames = 20: .Width = 60: .Height = 60: .CollRad = 25: .HitLimit = 2
        .Weapon = 0: .ResID = "ENEMY4"
    End With

    With EnemyType(5)       'knobby wings
        .AnimFrames = 12: .Width = 70: .Height = 70: .CollRad = 28: .HitLimit = 3
        .Weapon = 2: .ResID = "ENEMY5"
    End With

    With EnemyType(6)       'v-wing
        .AnimFrames = 9: .Width = 60: .Height = 60: .CollRad = 25: .HitLimit = 2
        .Weapon = 2: .ResID = "ENEMY6"
    End With

    With EnemyType(7)       'arcing flyer
        .AnimFrames = 20: .Width = 70: .Height = 70: .CollRad = 30: .HitLimit = 2
        .Weapon = 2: .ResID = "ENEMY7"
    End With

    With EnemyType(8)       'purple and white guy
        .AnimFrames = 20: .Width = 60: .Height = 60: .CollRad = 25: .HitLimit = 3
        .Weapon = 2: .ResID = "ENEMY8"
    End With

    With EnemyType(9)       'boss 1
        .AnimFrames = 1: .Width = 145: .Height = 170: .CollRad = 70: .HitLimit = 50
        .Weapon = 3: .FirePause = 10: .ResID = "BOSS2": .IsBoss = True
    End With

    With EnemyType(10)       'spinner with arms
        .AnimFrames = 20: .Width = 60: .Height = 60: .CollRad = 25: .HitLimit = 3
        .Weapon = 0: .ResID = "ENEMY9"
    End With

End Sub

Public Sub InitEnemyGroup(GetEnType As Integer, GetPathNo As Integer, GetNumEn As Integer, GetPowerUp As Integer)
Dim j As Integer, k As Integer
Dim SurfWidth As Long

    For j = 0 To UBound(EnGrp)
        With EnGrp(j)
            If .Active = False Then
                .TypeNo = GetEnType
                .NumEn = GetNumEn
                .PathNo = GetPathNo
                .PowerUp = GetPowerUp
                .PowerUpHolder = Int(Rnd * (.NumEn - 1))
                For k = 0 To .NumEn - 1
                    .Enemy(k).PathCounter = 0
                    .Enemy(k).FireTicker = 0
                    .Enemy(k).Firing = False
                    .Enemy(k).FrameNo = 0
                    .Enemy(k).AnimCounter = 0
                    .Enemy(k).HitCount = 0
                    .Enemy(k).ImgRECT.Left = 0
                    .Enemy(k).ImgRECT.Right = EnemyType(GetEnType).Width
                    .Enemy(k).ImgRECT.Bottom = EnemyType(GetEnType).Height
                    .Enemy(k).Top = EnemyType(GetEnType).Height * -1
                    .Enemy(k).Left = GameCtr
                    .Enemy(k).CanBeHit = True
                    .Enemy(k).OnScreen = True
                    .Enemy(k).AimAngle = 0
                    SurfWidth = EnemyType(GetEnType).Width * EnemyType(GetEnType).AnimFrames
                Next k
                .SurfNo = CreateEnemySurf(SurfWidth, EnemyType(GetEnType).Height, EnemyType(GetEnType).ResID)
                .Active = True
                Exit For
            End If
        End With
    Next j
    
End Sub

Public Sub ShowEnemies()
On Error GoTo ShowEnError
Dim i As Integer
Dim j As Integer
Dim k As Integer
Dim blnStillOn As Boolean
Dim EnemyShowTop As Single
    
    For i = 0 To UBound(EnGrp)
        With EnGrp(i)
            If .Active Then
                For j = 0 To .NumEn - 1
                    If .Enemy(j).OnScreen Then
                        
                        If .Enemy(j).Top >= 0 Then
                            TestForShotHitEnemy .TypeNo, .Enemy(j)
                            TestForBombChildHitEnemy .TypeNo, .Enemy(j)
                        End If
                        TestForEnemyHitShip .TypeNo, .Enemy(j)
                        If .Enemy(j).HitCount >= EnemyType(.TypeNo).HitLimit Then
                            If Not EnemyType(.TypeNo).IsBoss Then
                                GameSounds.play_snd 0, True
                                .Enemy(j).OnScreen = False
                                Score = Score + 5 * EnemyType(.TypeNo).HitLimit
                                StartExplosion .Enemy(j).Left + EnemyType(.TypeNo).Width / 2, .Enemy(j).Top + EnemyType(.TypeNo).Height / 2, 1
                                If .PowerUp > 0 And .PowerUpHolder = j Then
                                    StartPowerUp .PowerUp - 1, .Enemy(j).Left + EnemyType(.TypeNo).Width / 2, .Enemy(j).Top + EnemyType(.TypeNo).Height / 2
                                End If
                            Else
                                DoBossExplosion i, j, EnemyShowTop
                            End If
                        End If
        
                        Select Case .PathNo
                            Case 0: DoArcAndDrop i, j
                            Case 1: DoStraightDrop i, j
                            Case 2: DoRollForward i, j
                            Case 3: DoSpiralLoop i, j, 1
                            Case 4: DoSpiralLoop i, j, 2
                            Case 5: DoZigZag i, j, 175, 3, 1
                            Case 6: DoZigZag i, j, 175, 3, -1
                            Case 7: DoLineAndDrop i, j, 180
                            Case 8: DoLineAndDrop i, j, 360
                            Case 9: DoZigZag i, j, 275, 3, 1
                            Case 10: DoZigZag i, j, 275, 3, -1
                            Case 11: DropAndTrack i, j
                            Case 12: DoCrissCross i, j
                            Case 13: DoDiagonalLeft i, j
                            Case 14: DoDiagonalRight i, j
                            Case 15: DoBigArcs i, j
                            Case 16: DoBoss1Path i
                        End Select
                        If .Enemy(j).Top > ScrGame.Bottom Then .Enemy(j).OnScreen = False
                        If .Enemy(j).Left > 874 Then .Enemy(j).OnScreen = False
                        If .Enemy(j).Left < 150 - EnemyType(.TypeNo).Width Then .Enemy(j).OnScreen = False
    
                        If .Enemy(j).Top <= 0 Then
                            EnemyShowTop = 0
                            .Enemy(j).ImgRECT.Top = .Enemy(j).Top * -1
                        ElseIf .Enemy(j).Top + EnemyType(.TypeNo).Height > ScrGame.Bottom Then
                            EnemyShowTop = .Enemy(j).Top
                            .Enemy(j).ImgRECT.Bottom = ScrGame.Bottom - .Enemy(j).Top
                        Else
                            EnemyShowTop = .Enemy(j).Top
                            .Enemy(j).ImgRECT.Top = 0
                            .Enemy(j).ImgRECT.Bottom = EnemyType(.TypeNo).Height
                        End If
                        
                        backbuffer.BltFast .Enemy(j).Left, EnemyShowTop, ddsEnemy(.SurfNo), .Enemy(j).ImgRECT, DDBLTFAST_SRCCOLORKEY Or DDBLTFAST_WAIT
                        
                        If EnemyType(.TypeNo).Weapon > 0 And .Enemy(j).Top > 10 Then
                            If .Enemy(j).Firing = True Then
                                If EnemyType(.TypeNo).Weapon = 1 Then
                                    recDisplay.Left = 0
                                    recDisplay.Right = recDisplay.Left + 10
                                    recDisplay.Top = 0
                                    recDisplay.Bottom = 20
                                ElseIf EnemyType(.TypeNo).Weapon = 2 Then
                                    recDisplay.Left = 10
                                    recDisplay.Right = recDisplay.Left + 10
                                    recDisplay.Top = 0
                                    recDisplay.Bottom = 10
                                ElseIf EnemyType(.TypeNo).Weapon = 3 Then
                                    recDisplay.Left = 10
                                    recDisplay.Right = recDisplay.Left + 10
                                    recDisplay.Top = 10
                                    recDisplay.Bottom = 20
                                End If
                            End If
                            If EnemyType(.TypeNo).IsBoss = False Then
                                If .Enemy(j).Firing = True Then
                                    .Enemy(j).FireTicker = .Enemy(j).FireTicker + 1
                                    If .Enemy(j).FireTicker = EnemyType(.TypeNo).FirePause Then
                                        For k = 0 To 6
                                            If .Enemy(j).Ammo(k).Fired = False Then
                                                .Enemy(j).Ammo(k).Fired = True
                                                .Enemy(j).Ammo(k).xSpan = 5
                                                If .Enemy(j).AimAngle <> 0 Then
                                                    If .Enemy(j).AimAngle > 0 Then
                                                        .Enemy(j).Ammo(k).YMove = Sin(.Enemy(j).AimAngle * PI / 180) * 6
                                                        .Enemy(j).Ammo(k).XMove = Cos(.Enemy(j).AimAngle * PI / 180) * 6
                                                        .Enemy(j).Ammo(k).xCtr = .Enemy(j).Left + (EnemyType(.TypeNo).Width / 2 + Cos(.Enemy(j).AimAngle * PI / 180) * 30 - 5)
                                                        .Enemy(j).Ammo(k).Y = .Enemy(j).Top + EnemyType(.TypeNo).Height / 2 + Sin(.Enemy(j).AimAngle * PI / 180) * 30
                                                    ElseIf .Enemy(j).AimAngle < 0 Then
                                                        .Enemy(j).Ammo(k).YMove = Sin(.Enemy(j).AimAngle * PI / 180) * -6
                                                        .Enemy(j).Ammo(k).XMove = Cos(.Enemy(j).AimAngle * PI / 180) * -6
                                                        .Enemy(j).Ammo(k).xCtr = .Enemy(j).Left + (EnemyType(.TypeNo).Width / 2 - Cos(.Enemy(j).AimAngle * PI / 180) * EnemyType(.TypeNo).ShotYStart - 5)
                                                        .Enemy(j).Ammo(k).Y = .Enemy(j).Top + EnemyType(.TypeNo).Height / 2 - Sin(.Enemy(j).AimAngle * PI / 180) * EnemyType(.TypeNo).ShotYStart
                                                    End If
                                                Else
                                                    .Enemy(j).Ammo(k).YMove = 10
                                                    .Enemy(j).Ammo(k).XMove = 0
                                                    .Enemy(j).Ammo(k).xCtr = .Enemy(j).Left + (EnemyType(.TypeNo).Width / 2 - 5)
                                                    .Enemy(j).Ammo(k).Y = .Enemy(j).Top + EnemyType(.TypeNo).ShotYStart
                                                End If
                                                Exit For
                                            End If
                                        Next k
                                        .Enemy(j).FireTicker = 0
                                    End If
                                End If
                            Else
                                DoBossFiring i, recDisplay.Left, recDisplay.Top
                            End If
                        End If

                    Else
                        .Enemy(j).PathCounter = .Enemy(j).PathCounter + 1
                    End If
                    If .Enemy(j).Firing Then
                        TestforShotHitShip .Enemy(j)
                        ShowEnemyFiring i, .Enemy(j)
                    End If
                Next j
                
                blnStillOn = False
                For j = 0 To .NumEn - 1
                    If .Enemy(j).OnScreen Then blnStillOn = True
                Next j

                If blnStillOn = False And Player.OnScreen And ResetGame = False Then
                    .Active = False
                    DestroyEnemySurf .SurfNo
                End If
            End If
        End With
    Next i

    Exit Sub

ShowEnError:
    MsgBox "Error occurred in ShowEnemies procedure"
    EndIt

End Sub

Public Sub ShowEnemyFiring(GetEGno As Integer, GetEnemy As EnemyObject)
On Error GoTo ShowEnFireError
Dim k As Integer

        With GetEnemy
            For k = 0 To 6
                If .Ammo(k).Fired = True Then
                    .Ammo(k).Y = .Ammo(k).Y + .Ammo(k).YMove
                    .Ammo(k).xCtr = .Ammo(k).xCtr + .Ammo(k).XMove
                    LocalLight.ShowLight 7, .Ammo(k).xCtr - .Ammo(k).XMove * 1.6 + 5, .Ammo(k).Y - .Ammo(k).YMove * 1.6 + 5
                    backbuffer.BltFast .Ammo(k).xCtr, .Ammo(k).Y, ddsEnShot, recDisplay, DDBLTFAST_SRCCOLORKEY Or DDBLTFAST_WAIT
                    If .Ammo(k).Y > ScrGame.Bottom Or .Ammo(k).xCtr > 874 Or .Ammo(k).xCtr < 150 Then
                        .Ammo(k).Fired = False
                    End If
                End If
            Next k
        End With

    Exit Sub
    
ShowEnFireError:
    MsgBox "Error occurred in ShowEnemyFiring procedure"
    EndIt

End Sub

Public Sub ResetEnemies()
Dim j As Integer, k As Integer, m As Integer

    For j = 0 To UBound(EnGrp)
        EnGrp(j).Active = False
        Set ddsEnemy(j) = Nothing
        For k = 0 To EnGrp(j).NumEn - 1
            EnGrp(j).Enemy(k).Firing = False
            For m = 0 To 6
                EnGrp(j).Enemy(k).Ammo(m).Fired = False
            Next m
        Next k
    Next j

End Sub

Private Sub StartPowerUp(PowerUpNo As Integer, EnCtrX As Single, EnCtrY As Single)

    With PowerUp(PowerUpNo)
        .Left = EnCtrX - .Width / 2
        .Top = EnCtrY - .Height / 2
        .OnScreen = True
    End With

End Sub
