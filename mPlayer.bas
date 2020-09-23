Attribute VB_Name = "mPlayer"
Public Type AmmoObject
    xCtr As Single
    xLeft As Single
    xRight As Single
    Y As Single
    YMove As Single
    XMove As Single
    xSpan As Single
    Power As Integer
    Fired As Boolean
    SideY As Single
    SideXL As Single
    SideXR As Single
End Type

Public Type TrailObject
    X As Single
    Y As Single
    XMov As Single
    YMov As Single
    LifeTime As Integer
    picno As Integer
End Type

Public Type BombObject
    Left As Single
    Top As Single
    FrameNo As Integer
    Fired As Boolean
    Exploded As Boolean
    CtrX As Single
    CtrY As Single
    ExplodeY As Single
    Child(11) As AmmoObject
End Type

Public BombsLeft As Integer

Public Type ShipObject
    NumLives As Integer
    OnScreen As Boolean
    Left As Single
    Top As Single
    Width As Integer
    Height As Integer
    FireRightX As Integer
    FireLeftX As Integer
    RocketRightX As Integer
    RocketLeftX As Integer
    MaxY As Single
    MinY As Single
    Hit As Integer
    BankCount As Integer
    ImgX As Integer
    Firing As Boolean
    FireTicker As Integer
    LaserPower As Integer
    PowerUpLife As Integer
    GotBombs As Boolean
    Bomb(3) As BombObject
    SpeedLR As Single
    SpeedUD As Single
    ShieldLife As Integer
    shipAmmo(15) As AmmoObject
    CollRad As Single
End Type

Public Player As ShipObject
Private TrailCounter As Integer
Dim recDisplay As RECT
Dim recRocTrails As RECT

Public Sub SetupShip()
Dim j  As Integer
Dim k As Integer

    With Player
        If .NumLives = 0 And Not GameOver Then .NumLives = 5
        .Hit = 0
        .BankCount = 0
        .FireTicker = 0
        .Firing = False
        .Left = GameCtr - 30
        .Top = ScrGame.Bottom - 80
        .ImgX = 0
        .Width = 60
        .Height = 60
        .MinY = 50
        .MaxY = ScrGame.Bottom - 80
        .LaserPower = 1
        .ShieldLife = 0
        .GotBombs = False
        .SpeedLR = 0
        .SpeedUD = 0
        For j = 0 To UBound(.Bomb)
            With .Bomb(j)
                For k = 0 To UBound(.Child)
                    .Child(k).XMove = Cos(k * 0.525) * 8
                    .Child(k).YMove = Sin(k * 0.525) * 8
                Next k
            End With
        Next j
        .CollRad = 27
    End With
    
End Sub


Public Sub ShowShip()
Dim i As Integer, k As Integer
        
    With Player
    
        SetShipPosition
        recDisplay.Right = recDisplay.Left + .Width
        recDisplay.Bottom = 60
        
        If .OnScreen Then
            If .ShieldLife > 0 Then
                recShield.Left = recShield.Left + 80
                If recShield.Left = 800 Then recShield.Left = 0
                recShield.Right = recShield.Left + 80
                backbuffer.BltFast .Left - 10, .Top - 10, ddsShield, recShield, DDBLTFAST_SRCCOLORKEY Or DDBLTFAST_WAIT
            End If
            backbuffer.BltFast .Left, .Top, ddsShip(CurShip), recDisplay, DDBLTFAST_SRCCOLORKEY Or DDBLTFAST_WAIT
       End If
        
         Select Case Player.BankCount
            Case 0
                recDisplay.Left = 0
                .FireLeftX = 18
                .FireRightX = 18
            Case 3
                recDisplay.Left = 60
                .FireLeftX = 15
                .FireRightX = 16
            Case 6
                recDisplay.Left = 120
                .FireLeftX = 15
                .FireRightX = 11
            Case 9
                recDisplay.Left = 180
                .FireLeftX = 13
                .FireRightX = 7
            Case -3
                recDisplay.Left = 240
                .FireLeftX = 17
                .FireRightX = 17
            Case -6
                recDisplay.Left = 300
                .FireLeftX = 14
                .FireRightX = 16
            Case -9
               recDisplay.Left = 360
                .FireLeftX = 8
                .FireRightX = 13
        End Select
    
        recDisplay.Right = recDisplay.Left + .Width
       
        ShowFiring
        If .Firing Then
            .FireTicker = .FireTicker + 1
            If .FireTicker = 12 Then
                For i = 1 To 15
                    If .shipAmmo(i).Fired = False Then
                        .shipAmmo(i).Fired = True
                        .shipAmmo(i).Power = .LaserPower
                        .shipAmmo(i).xCtr = .Left + (.Width / 2)
                        .shipAmmo(i).Y = .Top
                        If .shipAmmo(i).Power = 1 Then
                            .shipAmmo(i).xLeft = .shipAmmo(i).xCtr - .FireLeftX
                            .shipAmmo(i).xRight = .shipAmmo(i).xCtr + .FireRightX
                            LocalLight.ShowLight 3, .shipAmmo(i).xLeft, .Top + 5
                            LocalLight.ShowLight 3, .shipAmmo(i).xRight, .Top + 5
                        ElseIf .shipAmmo(i).Power = 2 Then
                            .shipAmmo(i).xLeft = .shipAmmo(i).xCtr - .FireLeftX
                            .shipAmmo(i).xRight = .shipAmmo(i).xCtr + .FireRightX
                            LocalLight.ShowLight 3, .shipAmmo(i).xCtr + .shipAmmo(i).xSpan, .Top - 2
                            LocalLight.ShowLight 3, .shipAmmo(i).xLeft, .Top + 5
                            LocalLight.ShowLight 3, .shipAmmo(i).xRight, .Top + 5
                        End If
                        .shipAmmo(i).SideY = .Top - 5
                        .shipAmmo(i).YMove = 15
                        .shipAmmo(i).SideXL = .Left + .Width / 2 - 12
                        .shipAmmo(i).SideXR = .Left + .Width / 2
                        GameSounds.play_snd 1, True
                        .PowerUpLife = .PowerUpLife - 1
                        If .PowerUpLife = 0 Then .LaserPower = 1
                        Exit For
                    End If
                Next i
                .FireTicker = 0
            End If
        End If
        
        If .GotBombs Then
            If gblnRMouseButtonUp = True Then
                For i = 0 To UBound(.Bomb)
                    If .Bomb(i).Fired = False And .Bomb(i).Exploded = False Then
                        GameSounds.play_snd 5, True
                        .Bomb(i).Fired = True
                        .Bomb(i).ExplodeY = .Top - 300
                        .Bomb(i).Left = .Left + .Width / 2 - 10
                        .Bomb(i).Top = .Top - 5
                        .Bomb(i).FrameNo = 0
                        For k = 0 To UBound(.Bomb(i).Child)
                            .Bomb(i).Child(k).Fired = True
                        Next k
                        Exit For
                    End If
                Next i
                BombsLeft = BombsLeft - 1
                If BombsLeft = 0 Then .GotBombs = False
            End If
        End If
        ShowBombing
    
    End With

End Sub

Public Sub SetShipPosition()

        'sets ship position with calls to IsKeyDown sub in 'Functions' public module
        'makes use of the Windows API GetKeyState function
        
    With Player
        If .Hit < 6 Then
            If ResetGame = False Then
                If .Left > gintMouseX Then
                    If .Left > ScrGame.Left + 10 Then
                        .Left = .Left - (.Left - gintMouseX) / 10
                    End If
                ElseIf .Left < gintMouseX Then
                    If .Left < ScrGame.Right - (.Width + 10) Then
                        .Left = .Left + (gintMouseX - .Left) / 10
                    End If
                End If
                    
                If .Top > gintMouseY Then
                    If .Top > .MinY Then
                       .Top = .Top - (.Top - gintMouseY) / 20
                       .SpeedUD = -1
                    End If
                ElseIf .Top < gintMouseY Then
                    If .Top < .MaxY Then
                       .Top = .Top + (gintMouseY - .Top) / 20
                       .SpeedUD = 1
                    End If
                End If
                
                
                If MouseMovedLR = 1 Then
                    If .BankCount > -9 Then
                        .BankCount = .BankCount - 1
                    End If
                ElseIf MouseMovedLR = 2 Then
                    If .BankCount < 9 Then
                        .BankCount = .BankCount + 1
                    End If
                ElseIf MouseMovedLR = 0 Then
                    If .BankCount > 0 Then
                        .BankCount = .BankCount - 1
                    ElseIf .BankCount < 0 Then
                        .BankCount = .BankCount + 1
                    End If
                End If
                
                If .SpeedUD < 0 Then
                    ShowRocketTrails
                End If
                
                If gblnLMouseButton Then
                    Player.Firing = True
                Else
                    Player.Firing = False
                End If
            End If
        Else
            .OnScreen = False
            If ResetGame = False Then
                GameSounds.play_snd 4, True
                StartExplosion .Left + .Width / 2, .Top + .Height / 2, 1
                .NumLives = .NumLives - 1
                If .NumLives = 0 Then GameOver = True
                ResetGame = True
            End If
        End If

    End With


End Sub

Private Sub ShowFiring()
Dim i As Integer
Dim recShotDisplay As RECT
    
    recShotDisplay.Left = 0: recShotDisplay.Top = 0
    recShotDisplay.Right = 12: recShotDisplay.Bottom = 24

    With Player
        For i = 1 To 15
            If .shipAmmo(i).Fired = True Then
                .shipAmmo(i).Y = .shipAmmo(i).Y - .shipAmmo(i).YMove
                If .shipAmmo(i).Power = 1 Then
                    recShotDisplay.Left = 0
                    recShotDisplay.Right = 6
                    backbuffer.BltFast .shipAmmo(i).xLeft - 2, .shipAmmo(i).Y, ddsPShot, recShotDisplay, DDBLTFAST_SRCCOLORKEY Or DDBLTFAST_WAIT
                    backbuffer.BltFast .shipAmmo(i).xRight - 3, .shipAmmo(i).Y, ddsPShot, recShotDisplay, DDBLTFAST_SRCCOLORKEY Or DDBLTFAST_WAIT
                ElseIf .shipAmmo(i).Power = 2 Then
                    recShotDisplay.Left = 6
                    recShotDisplay.Right = 12
                    backbuffer.BltFast .shipAmmo(i).xCtr - 3, .shipAmmo(i).Y, ddsPShot, recShotDisplay, DDBLTFAST_SRCCOLORKEY Or DDBLTFAST_WAIT
                    backbuffer.BltFast .shipAmmo(i).xLeft - 3, .shipAmmo(i).Y, ddsPShot, recShotDisplay, DDBLTFAST_SRCCOLORKEY Or DDBLTFAST_WAIT
                    backbuffer.BltFast .shipAmmo(i).xRight - 3, .shipAmmo(i).Y, ddsPShot, recShotDisplay, DDBLTFAST_SRCCOLORKEY Or DDBLTFAST_WAIT
                End If
                
                If .shipAmmo(i).Y < -32 Then
                    .shipAmmo(i).Fired = False
                End If
            End If
        Next i
    End With

End Sub

Private Sub ShowRocketTrails()

    recRocTrails.Bottom = 12
    recRocTrails.Left = recRocTrails.Left + 10
    If recRocTrails.Left = 60 Then recRocTrails.Left = 0
    recRocTrails.Right = recRocTrails.Left + 10
    TrailCounter = 0
    
    backbuffer.BltFast (Player.Left + Player.Width / 2) + (Player.FireRightX * 0.86) + Abs(Player.BankCount / 3) - 13, _
    Player.Top + Player.Height, ddsRocTrails, recRocTrails, DDBLTFAST_SRCCOLORKEY Or DDBLTFAST_WAIT

    backbuffer.BltFast (Player.Left + Player.Width / 2) - (Player.FireLeftX * 0.86) - Abs(Player.BankCount / 3), _
    Player.Top + Player.Height, ddsRocTrails, recRocTrails, DDBLTFAST_SRCCOLORKEY Or DDBLTFAST_WAIT

End Sub

Private Sub ShowBombing()
Dim j As Integer, k As Integer
Dim StillExploding As Boolean
Dim recPBomb As RECT

    recPBomb.Bottom = 20
    
    With Player
        For j = 0 To UBound(.Bomb)
            If .Bomb(j).Fired Then
                .Bomb(j).Top = .Bomb(j).Top - 5
                If .Bomb(j).Top <= .Bomb(j).ExplodeY Then
                    .Bomb(j).Exploded = True
                    .Bomb(j).Fired = False
                    .Bomb(j).CtrX = .Bomb(j).Left + 10
                    .Bomb(j).CtrY = .Bomb(j).Top + 10
                    GammaValue = 50
                    GameSounds.stop_snd 5
                    GameSounds.play_snd 6, True
                    For k = 0 To UBound(.Bomb(j).Child)
                        .Bomb(j).Child(k).xCtr = .Bomb(j).Left + 10
                        .Bomb(j).Child(k).Y = .Bomb(j).Top + 10
                        .Bomb(j).Child(k).Power = 255
                    Next k
                End If
                .Bomb(j).FrameNo = .Bomb(j).FrameNo + 1
                If .Bomb(j).FrameNo = 10 Then .Bomb(j).FrameNo = 0
                recPBomb.Left = .Bomb(j).FrameNo * 20
                recPBomb.Right = recPBomb.Left + 20
                backbuffer.BltFast .Bomb(j).Left, .Bomb(j).Top, ddsPBomb, recPBomb, DDBLTFAST_SRCCOLORKEY Or DDBLTFAST_WAIT
            ElseIf .Bomb(j).Exploded Then
                For k = 0 To UBound(.Bomb(j).Child)
                    .Bomb(j).Child(k).xCtr = .Bomb(j).Child(k).xCtr + .Bomb(j).Child(k).XMove
                    .Bomb(j).Child(k).Y = .Bomb(j).Child(k).Y + .Bomb(j).Child(k).YMove
                    If .Bomb(j).Child(k).Power > 5 Then .Bomb(j).Child(k).Power = .Bomb(j).Child(k).Power - 4
                    backbuffer.SetFillStyle 1
                    backbuffer.SetForeColor RGB(.Bomb(j).Child(k).Power, Int(.Bomb(j).Child(k).Power / 2), .Bomb(j).Child(k).Power)
                    backbuffer.DrawCircle .Bomb(j).CtrX, .Bomb(j).CtrY, .Bomb(j).Child(k).xCtr - .Bomb(j).CtrX - ((k - 1) / 12) * (.Bomb(j).Child(k).xCtr - .Bomb(j).CtrX)
                    LocalLight.ShowLight 3, .Bomb(j).Child(k).xCtr, .Bomb(j).Child(k).Y
                    LocalLight.ShowLight 5, .Bomb(j).Child(k).xCtr - .Bomb(j).Child(k).XMove, .Bomb(j).Child(k).Y - .Bomb(j).Child(k).YMove
                    If .Bomb(j).Child(k).xCtr < 140 Or .Bomb(j).Child(k).xCtr > 884 Then .Bomb(j).Child(k).Fired = False
                    If .Bomb(j).Child(k).Y < -10 Or .Bomb(j).Child(k).Y > 778 Then .Bomb(j).Child(k).Fired = False
                Next k
            End If
            StillExploding = False
            For k = 0 To UBound(.Bomb(j).Child)
                If .Bomb(j).Child(k).Fired Then StillExploding = True
            Next k
            If StillExploding = False Then
                .Bomb(j).Exploded = False
            End If
        Next j
    End With

End Sub


Public Sub ResetPlayerShots()
Dim k As Integer
Dim m As Integer

    Player.Firing = False
    Player.FireTicker = 0
    For k = 0 To UBound(Player.shipAmmo)
        Player.shipAmmo(k).Fired = False
    Next k
    For k = 0 To UBound(Player.Bomb)
        Player.Bomb(k).Fired = False
        For m = 0 To UBound(Player.Bomb(k).Child)
            Player.Bomb(k).Child(m).Fired = False
        Next m
    Next k
    

End Sub

