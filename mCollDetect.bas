Attribute VB_Name = "mCollDetect"

Public Sub TestForShotHitEnemy(EnTypNo As Integer, GetEnemy As EnemyObject)
'collision detection using the circular method
Dim j As Integer
Dim EnemyCtrX As Single, EnemyCtrY As Single
Dim DistL As Double, DistR As Double

    EnemyCtrX = GetEnemy.Left + EnemyType(EnTypNo).Width / 2
    EnemyCtrY = GetEnemy.Top + EnemyType(EnTypNo).Height / 2

    With Player
        For j = 0 To 15
            If .shipAmmo(j).Fired Then
                If GetEnemy.OnScreen = True And GetEnemy.CanBeHit Then
                    DistL = Sqr((.shipAmmo(j).xLeft - EnemyCtrX) ^ 2 + (.shipAmmo(j).Y - EnemyCtrY) ^ 2)
                    DistR = Sqr((.shipAmmo(j).xRight - EnemyCtrX) ^ 2 + (.shipAmmo(j).Y - EnemyCtrY) ^ 2)
                    If DistL <= EnemyType(EnTypNo).CollRad Or _
                            DistR <= EnemyType(EnTypNo).CollRad Then
                       LocalLight.ShowLight 1, EnemyCtrX, EnemyCtrY
                       GetEnemy.HitCount = GetEnemy.HitCount + Player.LaserPower
                       .shipAmmo(j).Fired = False
                    End If
                End If
            End If
        Next j
    End With

End Sub

Public Sub TestForEnemyHitShip(EnTypNo As Integer, GetEnemy As EnemyObject)
'collision detection using the circular method
Dim Dist As Double
Dim PlayerCtrX As Single, PlayerCtrY As Single
Dim EnemyCtrX As Single, EnemyCtrY As Single

    With Player
        If .OnScreen And .Hit < 6 Then
            PlayerCtrX = .Left + .Width / 2
            PlayerCtrY = .Top + .Height / 2
            EnemyCtrX = GetEnemy.Left + EnemyType(EnTypNo).Width / 2
            EnemyCtrY = GetEnemy.Top + EnemyType(EnTypNo).Height / 2
            If GetEnemy.OnScreen And GetEnemy.CanBeHit Then
                Dist = Sqr((PlayerCtrX - EnemyCtrX) ^ 2 + (PlayerCtrY - EnemyCtrY) ^ 2)
                If Dist <= Player.CollRad + EnemyType(EnTypNo).CollRad Then
                    .Hit = 6
                    GetEnemy.HitCount = EnemyType(EnTypNo).HitLimit
                    StartExplosion PlayerCtrX, PlayerCtrY, 1
                End If
            End If
        End If
    End With

End Sub

Public Sub TestforShotHitShip(GetEnemy As EnemyObject)
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
        If GetEnemy.Ammo(i).Fired Then
            Dist = Sqr((PlayerCtrX - (GetEnemy.Ammo(i).xCtr) + 5) ^ 2 + (PlayerCtrY - (GetEnemy.Ammo(i).Y + 5)) ^ 2)
            If Dist <= Player.CollRad Then
                GetEnemy.Ammo(i).Fired = False
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

Public Sub TestForLaserBeamHit()
Dim PlayerCtrX As Single, PlayerCtrY As Single

    If Player.OnScreen Then
        PlayerCtrX = Player.Left + Player.Width / 2
        PlayerCtrY = Player.Top + Player.Height / 2
        If Player.Top < LaserBeam.Top + 7 And Player.Top + Player.Height > LaserBeam.Top + 10 Then
            Player.Hit = 6
            LocalLight.ShowLight 2, PlayerCtrX, PlayerCtrY
            StartExplosion PlayerCtrX, PlayerCtrY, 1
        End If
    End If

End Sub

Public Sub TestForShipHitPowerUp(PNo As Integer)
'collision detection using the circular method
Dim Dist As Double
Dim PlayerCtrX As Single, PlayerCtrY As Single
Dim PowerCtrX As Single, PowerCtrY As Single

    With Player
        If .OnScreen And .Hit < 6 Then
            PlayerCtrX = .Left + .Width / 2
            PlayerCtrY = .Top + .Height / 2
            PowerCtrX = PowerUp(PNo).Left + PowerUp(PNo).Width / 2
            PowerCtrY = PowerUp(PNo).Top + PowerUp(PNo).Height / 2
            Dist = Sqr((PlayerCtrX - PowerCtrX) ^ 2 + (PlayerCtrY - PowerCtrY) ^ 2)
            If Dist <= Player.CollRad + PowerUp(PNo).CollRad Then
                If PowerUp(PNo).PowerValue > 1 Then
                    Player.LaserPower = PowerUp(PNo).PowerValue
                End If
                If PowerUp(PNo).Life > 0 Then
                    Player.PowerUpLife = PowerUp(PNo).Life
                End If
                PowerUp(PNo).OnScreen = False
                If PNo = 1 Then .Hit = 0
                If PNo = 2 Then .ShieldLife = 10
                If PNo = 3 Then
                    Dim k As Integer
                    .GotBombs = True
                    BombsLeft = 4
                    For k = 0 To UBound(Player.Bomb)
                        Player.Bomb(k).Fired = False
                    Next k
                End If
            End If
        End If
    End With

End Sub

Public Sub TestShotHitAsteroid(AstNo As Integer)
'collision detection using the circular method
Dim j As Integer
Dim AstCtrX As Single, AstCtrY As Single
Dim DistL As Double, DistR As Double
Dim blnHitAst As Boolean

    AstCtrX = Asteroid(AstNo).Left + Asteroid(AstNo).Width / 2
    AstCtrY = Asteroid(AstNo).Top + Asteroid(AstNo).Height / 2

    With Player
        For j = 0 To 15
            If .shipAmmo(j).Fired Then
                If Asteroid(AstNo).OnScreen = True Then
                   DistL = Sqr((.shipAmmo(j).xLeft - AstCtrX) ^ 2 + (.shipAmmo(j).Y - AstCtrY) ^ 2)
                   DistR = Sqr((.shipAmmo(j).xRight - AstCtrX) ^ 2 + (.shipAmmo(j).Y - AstCtrY) ^ 2)
                   If DistL <= 15 Or DistR <= 15 Then blnHitAst = True
                   If blnHitAst Then
                       Asteroid(AstNo).Hit = Asteroid(AstNo).Hit + 1 + Player.LaserPower
                       .shipAmmo(j).Fired = False
                   End If
                End If
            End If
        Next j
    End With

End Sub

Public Sub TestAsteroidHitShip(AstNo As Integer)
'collision detection using the circular method
Dim j As Integer
Dim PlayerCtrX As Single, PlayerCtrY As Single
Dim AstCtrX As Single, AstCtrY As Single
Dim Dist As Double
Dim blnHitAst As Boolean

    With Player
        If .OnScreen And .Hit < 6 Then
            PlayerCtrX = .Left + .Width / 2
            PlayerCtrY = .Top + .Height / 2
            AstCtrX = Asteroid(AstNo).Left + Asteroid(AstNo).Width / 2
            AstCtrY = Asteroid(AstNo).Top + Asteroid(AstNo).Height / 2
            If Asteroid(AstNo).OnScreen Then
                Dist = Sqr((PlayerCtrX - AstCtrX) ^ 2 + (PlayerCtrY - AstCtrY) ^ 2)
                If Dist <= Player.CollRad + 15 Then
                    .Hit = 6
                    Asteroid(AstNo).Hit = Asteroid(AstNo).HitLimit
                    StartExplosion PlayerCtrX, PlayerCtrY, 1
                End If
            End If
        End If
    End With

End Sub

Public Sub TestForShotHitBallShooter(GetEnemy As EnemyObject)
'collision detection using the circular method
Dim j As Integer
Dim EnemyCtrX As Single, EnemyCtrY As Single
Dim DistL As Double, DistR As Double, DistC As Double

    EnemyCtrX = GetEnemy.Left + 12
    EnemyCtrY = GetEnemy.Top + 12

    With Player
        For j = 0 To 15
            If .shipAmmo(j).Fired Then
                If GetEnemy.OnScreen = True And GetEnemy.CanBeHit Then
                    DistL = Sqr((.shipAmmo(j).xLeft - EnemyCtrX) ^ 2 + (.shipAmmo(j).Y - EnemyCtrY) ^ 2)
                    DistR = Sqr((.shipAmmo(j).xRight - EnemyCtrX) ^ 2 + (.shipAmmo(j).Y - EnemyCtrY) ^ 2)
                    If DistL <= 12 Or DistR <= 12 Then
                       LocalLight.ShowLight 1, EnemyCtrX, EnemyCtrY
                       GetEnemy.HitCount = GetEnemy.HitCount + Player.LaserPower
                       .shipAmmo(j).Fired = False
                    End If
                End If
            End If
        Next j
    End With

End Sub


Public Sub TestForBallShooterHitShip(GetEnemy As EnemyObject)
'collision detection using the circular method
Dim Dist As Double
Dim PlayerCtrX As Single, PlayerCtrY As Single
Dim EnemyCtrX As Single, EnemyCtrY As Single

    With Player
        If .OnScreen And .Hit < 6 Then
            PlayerCtrX = .Left + .Width / 2
            PlayerCtrY = .Top + .Height / 2
            EnemyCtrX = GetEnemy.Left + 12
            EnemyCtrY = GetEnemy.Top + 12
            If GetEnemy.OnScreen And GetEnemy.CanBeHit Then
                Dist = Sqr((PlayerCtrX - EnemyCtrX) ^ 2 + (PlayerCtrY - EnemyCtrY) ^ 2)
                If Dist <= Player.CollRad + EnemyType(EnTypNo).CollRad Then
                    .Hit = 6
                    GetEnemy.HitCount = 5
                    StartExplosion PlayerCtrX, PlayerCtrY, 1
                End If
            End If
        End If
    End With

End Sub

Public Sub TestForBombChildHitEnemy(EnTypNo As Integer, GetEnemy As EnemyObject)
'collision detection using the circular method
Dim i As Integer
Dim j As Integer
Dim EnemyCtrX As Single, EnemyCtrY As Single
Dim Dist As Double

    EnemyCtrX = GetEnemy.Left + EnemyType(EnTypNo).Width / 2
    EnemyCtrY = GetEnemy.Top + EnemyType(EnTypNo).Height / 2

    For i = 0 To UBound(Player.Bomb)
        With Player.Bomb(i)
            For j = 0 To UBound(.Child)
                If .Child(j).Fired Then
                    If GetEnemy.OnScreen = True And GetEnemy.CanBeHit Then
                        Dist = Sqr((.Child(j).xCtr - EnemyCtrX) ^ 2 + (.Child(j).Y - EnemyCtrY) ^ 2)
                        If Dist <= EnemyType(EnTypNo).CollRad Then
                           LocalLight.ShowLight 1, EnemyCtrX, EnemyCtrY
                           If EnemyType(EnTypNo).IsBoss Then
                               GetEnemy.HitCount = GetEnemy.HitCount + 3
                           Else
                               GetEnemy.HitCount = EnemyType(EnTypNo).HitLimit
                           End If
                           .Child(j).Fired = False
                        End If
                    End If
                End If
            Next j
        End With
    Next i

End Sub

Public Sub TestForBigGunShotHitShip(GetGunShot As AmmoObject)
'collision detection using the circular method
Dim Dist As Double
Dim PlayerCtrX As Single, PlayerCtrY As Single

    If Player.OnScreen And Player.Hit < 6 Then
        PlayerCtrX = Player.Left + Player.Width / 2
        PlayerCtrY = Player.Top + Player.Height / 2
    Else
        Exit Sub
    End If

    If GetGunShot.Fired Then
        Dist = Sqr((PlayerCtrX - (GetGunShot.xLeft) + 5) ^ 2 + (PlayerCtrY - (GetGunShot.Y + 5)) ^ 2)
        If Dist <= Player.CollRad Then
            GetGunShot.Fired = False
            If Player.ShieldLife = 0 Then
                Player.Hit = Player.Hit + 1
                LocalLight.ShowLight 2, PlayerCtrX, PlayerCtrY
            Else
                Player.ShieldLife = Player.ShieldLife - 1
            End If
        End If
    End If

End Sub
