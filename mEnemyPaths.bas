Attribute VB_Name = "mEnemyPaths"
Public Const PI = 3.141592654


Public Sub DoArcAndDrop(GetEGno As Integer, GetENo As Integer)
Dim CtrX As Single, CtrY As Single
Dim EPAng As Single, ELeft As Single, ETop As Single, EPC As Long
'Path 0
'alternates left and right arc and then straight drop

    If GetENo > 0 Then
        If EnGrp(GetEGno).Enemy(GetENo - 1).PathCounter < 80 Then
            Exit Sub
        End If
    End If
    
    With EnGrp(GetEGno)
        .Enemy(GetENo).PathCounter = .Enemy(GetENo).PathCounter + 1
        EPC = .Enemy(GetENo).PathCounter
        EPAng = .Enemy(GetENo).PathAngle
        ELeft = .Enemy(GetENo).Left
        ETop = .Enemy(GetENo).Top
    End With
    
    If GetENo = 0 Or GetENo / 2 - Int(GetENo / 2) = 0 Then
        If EPC = 1 Then
            EPAng = 3.15
        ElseIf EPC < 220 Then
            CtrX = 600
            CtrY = -54
            EPAng = EPAng + 0.0175 * ((220 - EPC) / 220)
            ELeft = CtrX + 360 * Cos(EPAng)
            ETop = CtrY - 360 * Sin(EPAng)
        ElseIf EPC > 270 And EPC < 470 Then
            EnGrp(GetEGno).Enemy(GetENo).Firing = True
            AnimateEnemy GetEGno, GetENo, 1
            ETop = ETop + (EPC - 220) / 20
        End If
    Else
        If EPC = 1 Then
            EPAng = 0
        ElseIf EPC < 220 Then
            CtrX = 360
            CtrY = -54
            EPAng = EPAng - 0.0175 * ((220 - EPC) / 220)
            ELeft = CtrX + 360 * Cos(EPAng)
            ETop = CtrY - 360 * Sin(EPAng)
        ElseIf EPC > 270 And EPC < 470 Then
            EnGrp(GetEGno).Enemy(GetENo).Firing = True
            AnimateEnemy GetEGno, GetENo, 1
            ETop = ETop + (EPC - 220) / 20
        End If
    End If
    
    With EnGrp(GetEGno)
        .Enemy(GetENo).PathAngle = EPAng
        .Enemy(GetENo).Left = ELeft
        .Enemy(GetENo).Top = ETop
        .Enemy(GetENo).PathCounter = EPC
    End With

End Sub

Public Sub DoStraightDrop(GetEGno As Integer, GetENo As Integer)
'Path 1

    If GetENo > 0 Then
        If EnGrp(GetEGno).Enemy(GetENo - 1).PathCounter < 30 Then
            Exit Sub
        End If
    End If
    
    With EnGrp(GetEGno)
        .Enemy(GetENo).PathCounter = .Enemy(GetENo).PathCounter + 1
        If .Enemy(GetENo).PathCounter < 150 Then
            If GetENo = 0 Or GetENo / 2 - Int(GetENo / 2) = 0 Then
                .Enemy(GetENo).Left = GameCtr - 50 - EnemyType(.TypeNo).Width
            Else
                .Enemy(GetENo).Left = GameCtr + 50
            End If
            .Enemy(GetENo).Top = .Enemy(GetENo).Top + 4
        ElseIf .Enemy(GetENo).PathCounter >= 150 And .Enemy(GetENo).PathCounter < 350 Then
            AnimateEnemy GetEGno, GetENo, 1
            If GetENo = 0 Or GetENo / 2 - Int(GetENo / 2) = 0 Then
                .Enemy(GetENo).Left = .Enemy(GetENo).Left - 2 - GetENo / 2
            Else
                .Enemy(GetENo).Left = .Enemy(GetENo).Left + 2 + GetENo / 2
            End If
        End If
        If EnemyType(.TypeNo).Weapon > 0 Then
            .Enemy(GetENo).Firing = True
        End If
    End With
        
    AnimateEnemy GetEGno, GetENo, 1
    

End Sub

Public Sub DoRollForward(GetEGno As Integer, GetENo As Integer)
'Path 2
'designed for enemy to appear to do vertical outside roll in middle of straight drop
'requires special sprite animation strip

    If GetENo > 0 Then
        If EnGrp(GetEGno).Enemy(GetENo - 1).PathCounter < 90 Then
            Exit Sub
        End If
    End If
    
    With EnGrp(GetEGno)
        .Enemy(GetENo).PathCounter = .Enemy(GetENo).PathCounter + 1
        If .Enemy(GetENo).PathCounter < 90 Then
            .Enemy(GetENo).CanBeHit = True
            .Enemy(GetENo).Left = 220 + GetENo * 520 / (.NumEn - 1)
            .Enemy(GetENo).Top = .Enemy(GetENo).Top + 4
            .Enemy(GetENo).PathAngle = 3.15
        ElseIf .Enemy(GetENo).PathCounter < 175 Then
            .Enemy(GetENo).CanBeHit = False
            .Enemy(GetENo).Top = ScrGame.Bottom / 2 - Sin(.Enemy(GetENo).PathAngle) * 100 - 90
            .Enemy(GetENo).PathAngle = .Enemy(GetENo).PathAngle + 0.063
            If .Enemy(GetENo).PathCounter / 2 - Int(.Enemy(GetENo).PathCounter / 2) = 0 Then
                AnimateEnemy GetEGno, GetENo, 1
            End If
        ElseIf .Enemy(GetENo).PathCounter < 550 Then
            .Enemy(GetENo).CanBeHit = True
            .Enemy(GetENo).Top = .Enemy(GetENo).Top + 4
        End If
    End With

End Sub

Public Sub DoSpiralLoop(GetEGno As Integer, GetENo As Integer, Side As Integer)
Dim EPAng As Single, ELeft As Single, ETop As Single, EPC As Long
'Path 3 and Path 4
'alternates left and right arc and then straight drop

    If GetENo > 0 Then
        If EnGrp(GetEGno).Enemy(GetENo - 1).PathCounter < 35 Then
            Exit Sub
        End If
    End If
    
    With EnGrp(GetEGno)
        .Enemy(GetENo).PathCounter = .Enemy(GetENo).PathCounter + 1
        EPC = .Enemy(GetENo).PathCounter
        EPAng = .Enemy(GetENo).PathAngle
        ELeft = .Enemy(GetENo).Left
        ETop = .Enemy(GetENo).Top
    End With
    
    If Side = 1 Then
        If EPC = 1 Then
            EPAng = 3.15
        ElseIf EPC < 70 Then
            CtrX = 700
            CtrY = EnemyType(EnGrp(GetEGno).TypeNo).Height * -1
            EPAng = EPAng + 0.0175
            ELeft = CtrX + 460 * Cos(EPAng)
            ETop = CtrY - 460 * Sin(EPAng)
            AnimateEnemy GetEGno, GetENo, Side, 10
        ElseIf EPC = 70 Then
            EnGrp(GetEGno).Enemy(GetENo).AnimCounter = 0
        ElseIf EPC > 70 And EPC < 130 Then
            CtrX = 563
            CtrY = 300
            EPAng = EPAng + 0.07
            ELeft = CtrX + 70 * Cos(EPAng)
            ETop = CtrY - 70 * Sin(EPAng)
            AnimateEnemy GetEGno, GetENo, Side, 5
        ElseIf EPC = 130 Then
            EnGrp(GetEGno).Enemy(GetENo).AnimCounter = 0
        ElseIf EPC > 130 And EPC < 400 Then
            CtrX = 790
            CtrY = 626
            EPAng = EPAng + 0.0175
            ELeft = CtrX + 460 * Cos(EPAng)
            ETop = CtrY - 460 * Sin(EPAng)
            AnimateEnemy GetEGno, GetENo, Side, 10
        End If
    ElseIf Side = 2 Then
        If EPC = 1 Then
            EPAng = 0
        ElseIf EPC < 70 Then
            CtrX = 324
            CtrY = EnemyType(EnGrp(GetEGno).TypeNo).Height * -1
            EPAng = EPAng - 0.0175
            ELeft = CtrX + 460 * Cos(EPAng)
            ETop = CtrY - 460 * Sin(EPAng)
            AnimateEnemy GetEGno, GetENo, Side, 10
        ElseIf EPC = 70 Then
            EnGrp(GetEGno).Enemy(GetENo).AnimCounter = 0
        ElseIf EPC > 70 And EPC < 130 Then
            CtrX = 458
            CtrY = 300
            EPAng = EPAng - 0.07
            ELeft = CtrX + 70 * Cos(EPAng)
            ETop = CtrY - 70 * Sin(EPAng)
            AnimateEnemy GetEGno, GetENo, Side, 5
        ElseIf EPC = 130 Then
            EnGrp(GetEGno).Enemy(GetENo).AnimCounter = 0
        ElseIf EPC > 130 And EPC < 400 Then
            CtrX = 234
            CtrY = 626
            EPAng = EPAng - 0.0175
            ELeft = CtrX + 460 * Cos(EPAng)
            ETop = CtrY - 460 * Sin(EPAng)
            AnimateEnemy GetEGno, GetENo, Side, 10
        End If
    End If

    With EnGrp(GetEGno)
        .Enemy(GetENo).PathAngle = EPAng
        .Enemy(GetENo).Left = ELeft
        .Enemy(GetENo).Top = ETop
        .Enemy(GetENo).PathCounter = EPC
    End With

End Sub

Public Sub DoZigZag(GetEGno As Integer, GetENo As Integer, WaveSize As Long, DropSpeed As Single, GetStartDir As Integer)
Dim EPAng As Single, ELeft As Single, ETop As Single
Dim Dist As Single
'Path 5 and Path 6

    If GetENo > 0 Then
        If EnGrp(GetEGno).Enemy(GetENo - 1).PathCounter < 30 Then
            Exit Sub
        End If
    End If
                
    With EnGrp(GetEGno)
        .Enemy(GetENo).PathCounter = .Enemy(GetENo).PathCounter + 1
        EPAng = .Enemy(GetENo).PathAngle
        ELeft = .Enemy(GetENo).Left
        ETop = .Enemy(GetENo).Top
                
        If .Enemy(GetENo).PathCounter = 1 Then
            EPAng = 0
        ElseIf .Enemy(GetENo).PathCounter < 900 Then
            Dist = Cos(EPAng) * WaveSize * GetStartDir
            ELeft = GameCtr + Dist
            ETop = ETop + DropSpeed
            EPAng = EPAng - 0.03
            If Dist >= 0 Then
                If .Enemy(GetENo).FrameNo > 0 Then
                    .Enemy(GetENo).FrameNo = .Enemy(GetENo).FrameNo - 1
                    .Enemy(GetENo).ImgRECT.Left = .Enemy(GetENo).ImgRECT.Left - EnemyType(.TypeNo).Width
                    .Enemy(GetENo).ImgRECT.Right = .Enemy(GetENo).ImgRECT.Left + EnemyType(.TypeNo).Width
                End If
            Else
                If .Enemy(GetENo).FrameNo < EnemyType(.TypeNo).AnimFrames - 1 Then
                    .Enemy(GetENo).FrameNo = .Enemy(GetENo).FrameNo + 1
                    .Enemy(GetENo).ImgRECT.Left = .Enemy(GetENo).ImgRECT.Left + EnemyType(.TypeNo).Width
                    .Enemy(GetENo).ImgRECT.Right = .Enemy(GetENo).ImgRECT.Left + EnemyType(.TypeNo).Width
                End If
            End If
        End If
        
        .Enemy(GetENo).PathAngle = EPAng
        .Enemy(GetENo).Left = ELeft
        .Enemy(GetENo).Top = ETop
        If EnemyType(.TypeNo).Weapon > 0 Then
            .Enemy(GetENo).Firing = True
        End If
    End With

End Sub

Public Sub DoLineAndDrop(GetEGno As Integer, GetENo As Integer, LeftSide As Long)
'Path 7
Dim EnemyStartDrop As Single
    
    With EnGrp(GetEGno)
        .Enemy(GetENo).PathCounter = .Enemy(GetENo).PathCounter + 1
        If .Enemy(GetENo).PathCounter < 80 Then
            AnimateEnemy GetEGno, GetENo, 1
            .Enemy(GetENo).Left = LeftSide + GetENo * 400 / (.NumEn - 1)
            .Enemy(GetENo).Top = .Enemy(GetENo).Top + 4
        ElseIf .Enemy(GetENo).PathCounter > 80 And .Enemy(GetENo).PathCounter < 115 Then
            .Enemy(GetENo).Top = .Enemy(GetENo).Top + (125 - .Enemy(GetENo).PathCounter) / 15
            AnimateEnemy GetEGno, GetENo, 1
        ElseIf .Enemy(GetENo).PathCounter >= 115 And .Enemy(GetENo).PathCounter < 120 Then
            AnimateEnemy GetEGno, GetENo, 1
        ElseIf .Enemy(GetENo).PathCounter >= 120 And .Enemy(GetENo).PathCounter < 500 Then
            AnimateEnemy GetEGno, GetENo, 1
            EnemyStartDrop = Abs((.NumEn - 1) / 2 - GetENo) * 50
            If .Enemy(GetENo).PathCounter - 140 >= EnemyStartDrop Then
                .Enemy(GetENo).Top = .Enemy(GetENo).Top + (.Enemy(GetENo).PathCounter - EnemyStartDrop - 140) / 10
                .Enemy(GetENo).Firing = True
            End If
        End If
    End With

End Sub

Public Sub DropAndTrack(GetEGno As Integer, GetENo As Integer)
Dim PlayerCtrX As Single, PlayerCtrY As Single
Dim EnemyCtrX As Single, EnemyCtrY As Single
    
    If GetENo > 1 Then
        If EnGrp(GetEGno).Enemy(GetENo - 2).PathCounter < 40 Then
            Exit Sub
        End If
    End If
    
    With EnGrp(GetEGno)
        .Enemy(GetENo).PathCounter = .Enemy(GetENo).PathCounter + 1
        If .Enemy(GetENo).PathCounter < 80 Then
            If GetENo = 0 Or GetENo / 2 - Int(GetENo / 2) = 0 Then
                .Enemy(GetENo).Left = GameCtr - 100 - EnemyType(.TypeNo).Width
            Else
                .Enemy(GetENo).Left = GameCtr + 100
            End If
            .Enemy(GetENo).Top = .Enemy(GetENo).Top + 2
        ElseIf .Enemy(GetENo).PathCounter >= 60 And .Enemy(GetENo).PathCounter < 110 Then
            .Enemy(GetENo).Left = .Enemy(GetENo).Left + ((.Enemy(GetENo).Left + EnemyType(.TypeNo).Width / 2) - GameCtr) / 50
            .Enemy(GetENo).Top = .Enemy(GetENo).Top + 2
        ElseIf .Enemy(GetENo).PathCounter >= 110 And .Enemy(GetENo).PathCounter < 280 Then
            .Enemy(GetENo).Top = .Enemy(GetENo).Top + 2
        ElseIf .Enemy(GetENo).PathCounter >= 280 And .Enemy(GetENo).PathCounter < 330 Then
            .Enemy(GetENo).Left = .Enemy(GetENo).Left - ((.Enemy(GetENo).Left + EnemyType(.TypeNo).Width / 2) - GameCtr) / 50
            .Enemy(GetENo).Top = .Enemy(GetENo).Top + 2
        ElseIf .Enemy(GetENo).PathCounter >= 330 And .Enemy(GetENo).PathCounter < 500 Then
            .Enemy(GetENo).Top = .Enemy(GetENo).Top + 2
        End If
        PlayerCtrX = Player.Left + Player.Width / 2
        PlayerCtrY = Player.Top + Player.Height / 2
        EnemyCtrX = .Enemy(GetENo).Left + EnemyType(.TypeNo).Width / 2
        EnemyCtrY = .Enemy(GetENo).Top + EnemyType(.TypeNo).Height / 2

        If EnemyType(.TypeNo).Weapon > 0 Then
            .Enemy(GetENo).Firing = True
        End If

        If (PlayerCtrY - EnemyCtrY) > 0 Then
            If (EnemyCtrX - PlayerCtrX) = 0 Then PlayerCtrX = PlayerCtrX + 1
            .Enemy(GetENo).AimAngle = Atn((PlayerCtrY - EnemyCtrY) / (PlayerCtrX - EnemyCtrX)) * 180 / PI
            If .Enemy(GetENo).AimAngle < 0 Then
             .Enemy(GetENo).FrameNo = 7 - Int((90 + .Enemy(GetENo).AimAngle) / 12.86)
            ElseIf .Enemy(GetENo).AimAngle > 0 Then
             .Enemy(GetENo).FrameNo = 7 + Int((90 - .Enemy(GetENo).AimAngle) / 12.86)
            End If
        Else
            .Enemy(GetENo).AimAngle = 0
            If .Enemy(GetENo).FrameNo > 7 Then
                .Enemy(GetENo).FrameNo = .Enemy(GetENo).FrameNo - 1
            ElseIf .Enemy(GetENo).FrameNo < 7 Then
                .Enemy(GetENo).FrameNo = .Enemy(GetENo).FrameNo + 1
            Else
                .Enemy(GetENo).FrameNo = 7
            End If
        End If
        
        .Enemy(GetENo).ImgRECT.Left = .Enemy(GetENo).FrameNo * EnemyType(.TypeNo).Width
        .Enemy(GetENo).ImgRECT.Right = .Enemy(GetENo).ImgRECT.Left + EnemyType(.TypeNo).Width
    End With
    
    
End Sub

Public Sub DoCrissCross(GetEGno As Integer, GetENo As Integer)

    If GetENo > 0 Then
        If EnGrp(GetEGno).Enemy(GetENo - 1).PathCounter < 40 Then
            Exit Sub
        End If
    End If

    With EnGrp(GetEGno)
        .Enemy(GetENo).PathCounter = .Enemy(GetENo).PathCounter + 1
        If .Enemy(GetENo).PathCounter < 50 Then
            If GetENo / 2 - Int(GetENo / 2) = 0 Then
                .Enemy(GetENo).Left = GameCtr + 300
            Else
                .Enemy(GetENo).Left = GameCtr - 300 - EnemyType(.TypeNo).Width
            End If
            .Enemy(GetENo).Top = .Enemy(GetENo).Top + 3
        ElseIf .Enemy(GetENo).PathCounter >= 50 And .Enemy(GetENo).PathCounter < 152 Then
            If GetENo / 2 - Int(GetENo / 2) = 0 Then
                .Enemy(GetENo).Left = .Enemy(GetENo).Left - Abs(101 - .Enemy(GetENo).PathCounter) / 4
                AnimateEnemy GetEGno, GetENo, 2
            Else
                .Enemy(GetENo).Left = .Enemy(GetENo).Left + Abs(101 - .Enemy(GetENo).PathCounter) / 4
                AnimateEnemy GetEGno, GetENo, 1
            End If
            .Enemy(GetENo).Top = .Enemy(GetENo).Top + 2
            If EnemyType(.TypeNo).Weapon > 0 Then
                .Enemy(GetENo).Firing = True
            End If

        ElseIf .Enemy(GetENo).PathCounter >= 152 And .Enemy(GetENo).PathCounter < 202 Then
            .Enemy(GetENo).Top = .Enemy(GetENo).Top + 3
        ElseIf .Enemy(GetENo).PathCounter >= 202 And .Enemy(GetENo).PathCounter < 304 Then
            If GetENo / 2 - Int(GetENo / 2) = 0 Then
                .Enemy(GetENo).Left = .Enemy(GetENo).Left + Abs(253 - .Enemy(GetENo).PathCounter) / 4
                AnimateEnemy GetEGno, GetENo, 1
            Else
                .Enemy(GetENo).Left = .Enemy(GetENo).Left - Abs(253 - .Enemy(GetENo).PathCounter) / 4
                AnimateEnemy GetEGno, GetENo, 2
            End If
            .Enemy(GetENo).Top = .Enemy(GetENo).Top + 2
            If EnemyType(.TypeNo).Weapon > 0 Then
                .Enemy(GetENo).Firing = True
            End If

        ElseIf .Enemy(GetENo).PathCounter >= 304 And .Enemy(GetENo).PathCounter < 370 Then
            .Enemy(GetENo).Top = .Enemy(GetENo).Top + 3
        End If
    End With

End Sub

Public Sub DoDiagonalLeft(GetEGno As Integer, GetENo As Integer)
Dim CtrX As Single, CtrY As Single

    If GetENo > 0 Then
        If EnGrp(GetEGno).Enemy(GetENo - 1).PathCounter < 13 Then
            Exit Sub
        End If
    End If

    With EnGrp(GetEGno)
        .Enemy(GetENo).PathCounter = .Enemy(GetENo).PathCounter + 1
        If .Enemy(GetENo).PathCounter = 1 Then
            .Enemy(GetENo).FrameNo = 0
        ElseIf .Enemy(GetENo).PathCounter < 30 Then
            .Enemy(GetENo).PathAngle = 3.15
            .Enemy(GetENo).Left = GameCtr - 350
            .Enemy(GetENo).Top = .Enemy(GetENo).Top + 6
        ElseIf .Enemy(GetENo).PathCounter < 41 Then
            CtrX = GameCtr - 250
            CtrY = 105
            .Enemy(GetENo).PathAngle = .Enemy(GetENo).PathAngle + 0.07
            .Enemy(GetENo).Left = CtrX + Cos(.Enemy(GetENo).PathAngle) * 100
            .Enemy(GetENo).Top = CtrY - Sin(.Enemy(GetENo).PathAngle) * 100
            If .Enemy(GetENo).FrameNo < 2 Then
               .Enemy(GetENo).FrameNo = .Enemy(GetENo).FrameNo + 1
               .Enemy(GetENo).ImgRECT.Left = .Enemy(GetENo).FrameNo * EnemyType(.TypeNo).Width
               .Enemy(GetENo).ImgRECT.Right = .Enemy(GetENo).ImgRECT.Left + EnemyType(.TypeNo).Width
            End If
        ElseIf .Enemy(GetENo).PathCounter < 100 Then
            .Enemy(GetENo).Left = .Enemy(GetENo).Left + 6
            .Enemy(GetENo).Top = .Enemy(GetENo).Top + 5
        ElseIf .Enemy(GetENo).PathCounter < 135 Then
            CtrX = 594
            CtrY = 320
            .Enemy(GetENo).PathAngle = .Enemy(GetENo).PathAngle + 0.16
            .Enemy(GetENo).Left = CtrX + Cos(.Enemy(GetENo).PathAngle) * 150
            .Enemy(GetENo).Top = CtrY - Sin(.Enemy(GetENo).PathAngle) * 150
            .Enemy(GetENo).AnimCounter = .Enemy(GetENo).AnimCounter + 1
            If .Enemy(GetENo).AnimCounter = 2 Then
                If .Enemy(GetENo).FrameNo < 19 Then
                   .Enemy(GetENo).FrameNo = .Enemy(GetENo).FrameNo + 1
                   .Enemy(GetENo).ImgRECT.Left = .Enemy(GetENo).FrameNo * EnemyType(.TypeNo).Width
                   .Enemy(GetENo).ImgRECT.Right = .Enemy(GetENo).ImgRECT.Left + EnemyType(.TypeNo).Width
                Else
                   .Enemy(GetENo).FrameNo = 0
                   .Enemy(GetENo).ImgRECT.Left = .Enemy(GetENo).FrameNo * EnemyType(.TypeNo).Width
                   .Enemy(GetENo).ImgRECT.Right = .Enemy(GetENo).ImgRECT.Left + EnemyType(.TypeNo).Width
                End If
                .Enemy(GetENo).AnimCounter = 0
            End If
        ElseIf .Enemy(GetENo).PathCounter < 260 Then
            .Enemy(GetENo).Top = .Enemy(GetENo).Top + 6
            .Enemy(GetENo).FrameNo = 0
            .Enemy(GetENo).ImgRECT.Left = .Enemy(GetENo).FrameNo * EnemyType(.TypeNo).Width
            .Enemy(GetENo).ImgRECT.Right = .Enemy(GetENo).ImgRECT.Left + EnemyType(.TypeNo).Width
        End If
    End With


End Sub

Public Sub DoDiagonalRight(GetEGno As Integer, GetENo As Integer)
Dim CtrX As Single, CtrY As Single

    If GetENo > 0 Then
        If EnGrp(GetEGno).Enemy(GetENo - 1).PathCounter < 13 Then
            Exit Sub
        End If
    End If

    With EnGrp(GetEGno)
        .Enemy(GetENo).PathCounter = .Enemy(GetENo).PathCounter + 1
        If .Enemy(GetENo).PathCounter = 1 Then
            .Enemy(GetENo).FrameNo = 19
        ElseIf .Enemy(GetENo).PathCounter < 30 Then
            .Enemy(GetENo).PathAngle = 0
            .Enemy(GetENo).Left = GameCtr + 350 - EnemyType(.TypeNo).Width
            .Enemy(GetENo).Top = .Enemy(GetENo).Top + 6
        ElseIf .Enemy(GetENo).PathCounter < 41 Then
            CtrX = GameCtr + 250 - EnemyType(.TypeNo).Width
            CtrY = 105
            .Enemy(GetENo).PathAngle = .Enemy(GetENo).PathAngle - 0.07
            .Enemy(GetENo).Left = CtrX + Cos(.Enemy(GetENo).PathAngle) * 100
            .Enemy(GetENo).Top = CtrY - Sin(.Enemy(GetENo).PathAngle) * 100
            If .Enemy(GetENo).FrameNo > 17 Then
               .Enemy(GetENo).FrameNo = .Enemy(GetENo).FrameNo - 1
               .Enemy(GetENo).ImgRECT.Left = .Enemy(GetENo).FrameNo * EnemyType(.TypeNo).Width
               .Enemy(GetENo).ImgRECT.Right = .Enemy(GetENo).ImgRECT.Left + EnemyType(.TypeNo).Width
            End If
        ElseIf .Enemy(GetENo).PathCounter < 100 Then
            .Enemy(GetENo).Left = .Enemy(GetENo).Left - 6
            .Enemy(GetENo).Top = .Enemy(GetENo).Top + 5
        ElseIf .Enemy(GetENo).PathCounter < 135 Then
            CtrX = 370
            CtrY = 320
            .Enemy(GetENo).PathAngle = .Enemy(GetENo).PathAngle - 0.16
            .Enemy(GetENo).Left = CtrX + Cos(.Enemy(GetENo).PathAngle) * 150
            .Enemy(GetENo).Top = CtrY - Sin(.Enemy(GetENo).PathAngle) * 150
            .Enemy(GetENo).AnimCounter = .Enemy(GetENo).AnimCounter + 1
            If .Enemy(GetENo).AnimCounter = 2 Then
                If .Enemy(GetENo).FrameNo > 0 Then
                   .Enemy(GetENo).FrameNo = .Enemy(GetENo).FrameNo - 1
                   .Enemy(GetENo).ImgRECT.Left = .Enemy(GetENo).FrameNo * EnemyType(.TypeNo).Width
                   .Enemy(GetENo).ImgRECT.Right = .Enemy(GetENo).ImgRECT.Left + EnemyType(.TypeNo).Width
                Else
                   .Enemy(GetENo).FrameNo = 19
                   .Enemy(GetENo).ImgRECT.Left = .Enemy(GetENo).FrameNo * EnemyType(.TypeNo).Width
                   .Enemy(GetENo).ImgRECT.Right = .Enemy(GetENo).ImgRECT.Left + EnemyType(.TypeNo).Width
                End If
                .Enemy(GetENo).AnimCounter = 0
            End If
        ElseIf .Enemy(GetENo).PathCounter < 260 Then
            .Enemy(GetENo).Top = .Enemy(GetENo).Top + 6
        End If
    End With


End Sub

Public Sub DoBigArcs(GetEGno As Integer, GetENo As Integer)
Dim CtrX As Single, CtrY As Single

    If GetENo > 0 Then
        If EnGrp(GetEGno).Enemy(GetENo - 1).PathCounter < 20 Then
            Exit Sub
        End If
    End If


    With EnGrp(GetEGno)
        .Enemy(GetENo).PathCounter = .Enemy(GetENo).PathCounter + 1
        If .Enemy(GetENo).PathCounter = 1 Then
            If GetENo / 2 - Int(GetENo / 2) = 0 Then
                .Enemy(GetENo).PathAngle = 2.3625
            Else
                .Enemy(GetENo).ImgRECT.Left = 10 * EnemyType(.TypeNo).Width
                .Enemy(GetENo).ImgRECT.Right = .Enemy(GetENo).ImgRECT.Left + EnemyType(.TypeNo).Width
                .Enemy(GetENo).PathAngle = 0.7875
            End If
        ElseIf .Enemy(GetENo).PathCounter < 200 Then
            If GetENo / 2 - Int(GetENo / 2) = 0 Then
                CtrX = 1512
                CtrY = 700
                .Enemy(GetENo).PathAngle = .Enemy(GetENo).PathAngle + 0.0066
                .Enemy(GetENo).Left = CtrX + Cos(.Enemy(GetENo).PathAngle) * 1000
                .Enemy(GetENo).Top = CtrY - Sin(.Enemy(GetENo).PathAngle) * 1000
            Else
                CtrX = -558
                CtrY = 700
                .Enemy(GetENo).PathAngle = .Enemy(GetENo).PathAngle - 0.0066
                .Enemy(GetENo).Left = CtrX + Cos(.Enemy(GetENo).PathAngle) * 1000
                .Enemy(GetENo).Top = CtrY - Sin(.Enemy(GetENo).PathAngle) * 1000
            End If
            .Enemy(GetENo).AnimCounter = .Enemy(GetENo).AnimCounter + 1
            If .Enemy(GetENo).AnimCounter = 12 Then
                If .Enemy(GetENo).FrameNo < 9 Then
                    .Enemy(GetENo).FrameNo = .Enemy(GetENo).FrameNo + 1
                    .Enemy(GetENo).ImgRECT.Left = .Enemy(GetENo).ImgRECT.Left + EnemyType(.TypeNo).Width
                    .Enemy(GetENo).ImgRECT.Right = .Enemy(GetENo).ImgRECT.Left + EnemyType(.TypeNo).Width
                End If
                .Enemy(GetENo).AnimCounter = 0
            End If
        End If
    End With

End Sub

Public Sub DoBoss1Path(GetEGno As Integer)

    With EnGrp(GetEGno)
        .Enemy(0).PathCounter = .Enemy(0).PathCounter + 1
        If .Enemy(0).PathCounter = 1 Then
            .Enemy(0).AimAngle = 1
        ElseIf .Enemy(0).PathCounter < 200 Then
            .Enemy(0).Top = .Enemy(0).Top + 2
        End If
        If .Enemy(0).Left >= 724 Or .Enemy(0).Left <= 150 Then
            .Enemy(0).AimAngle = .Enemy(0).AimAngle * -1
        End If
        .Enemy(0).Left = .Enemy(0).Left + 1 * .Enemy(0).AimAngle
        If .Enemy(0).PathCounter / 70 - Int(.Enemy(0).PathCounter / 70) = 0 Then
            If .Enemy(0).Firing = True Then
                .Enemy(0).Firing = False
            Else
                .Enemy(0).Firing = True
            End If
        End If
    End With

End Sub

Private Sub AnimateEnemy(EGNo As Integer, ENo As Integer, AnimDir As Integer, Optional FrameSpeed As Integer = 2)
'called by enemy path subs

    With EnGrp(EGNo)
        .Enemy(ENo).AnimCounter = .Enemy(ENo).AnimCounter + 1
        If .Enemy(ENo).AnimCounter = FrameSpeed Then
            If AnimDir = 1 Then
                If .Enemy(ENo).FrameNo < EnemyType(.TypeNo).AnimFrames - 1 Then
                    .Enemy(ENo).FrameNo = .Enemy(ENo).FrameNo + 1
                    .Enemy(ENo).ImgRECT.Left = .Enemy(ENo).ImgRECT.Left + EnemyType(.TypeNo).Width
                    .Enemy(ENo).ImgRECT.Right = .Enemy(ENo).ImgRECT.Left + EnemyType(.TypeNo).Width
                Else
                    .Enemy(ENo).FrameNo = 0
                    .Enemy(ENo).ImgRECT.Left = 0
                    .Enemy(ENo).ImgRECT.Right = EnemyType(.TypeNo).Width
                End If
            ElseIf AnimDir = 2 Then
                If .Enemy(ENo).FrameNo > 1 Then
                    .Enemy(ENo).FrameNo = .Enemy(ENo).FrameNo - 1
                    .Enemy(ENo).ImgRECT.Left = .Enemy(ENo).ImgRECT.Left - EnemyType(.TypeNo).Width
                    .Enemy(ENo).ImgRECT.Right = .Enemy(ENo).ImgRECT.Left + EnemyType(.TypeNo).Width
                Else
                    .Enemy(ENo).FrameNo = EnemyType(.TypeNo).AnimFrames
                    .Enemy(ENo).ImgRECT.Left = (EnemyType(.TypeNo).AnimFrames - 1) * EnemyType(.TypeNo).Width
                    .Enemy(ENo).ImgRECT.Right = .Enemy(ENo).ImgRECT.Left + EnemyType(.TypeNo).Width
                End If
            End If
            .Enemy(ENo).AnimCounter = 0
        End If
    End With

End Sub
