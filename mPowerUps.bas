Attribute VB_Name = "mPowerUps"
Public Type PowerUpObject
    Left As Single
    Top As Single
    Width As Long
    Height As Long
    Speed As Single
    OnScreen As Boolean
    PowerValue As Integer
    Life As Integer
    AnimFrames As Integer
    FrameNo As Integer
    FrameCounter As Integer
    CollRad As Single
    ResID As String
    SurfNo As Integer
End Type

Public PowerUp(3) As PowerUpObject

Public Sub InitPowerUps()

    With PowerUp(0)
        .PowerValue = 2
        .Life = 50
        .AnimFrames = 16
        .Width = 20
        .Height = 20
        .Top = .Height * -1
        .Speed = 1
        .OnScreen = False
        .CollRad = 10
        .ResID = "POWER1"
        .SurfNo = CreatePowerSurf(.Width * .AnimFrames, .Height, .ResID)
    End With

    With PowerUp(1)
        .PowerValue = 1
        .Life = 0
        .AnimFrames = 6
        .Width = 40
        .Height = 30
        .Top = .Height * -1
        .Speed = 1
        .OnScreen = False
        .CollRad = 17
        .ResID = "HEALTH"
        .SurfNo = CreatePowerSurf(.Width * .AnimFrames, .Height, .ResID)
    End With

    With PowerUp(2)
        .PowerValue = 1
        .Life = 0
        .AnimFrames = 20
        .Width = 40
        .Height = 40
        .Left = Rnd * 600 + 250
        .Speed = 1
        .OnScreen = False
        .CollRad = 20
        .ResID = "BATTERY"
        .SurfNo = CreatePowerSurf(.Width * .AnimFrames, .Height, .ResID)
    End With

    With PowerUp(3)
        .PowerValue = 1
        .Life = 0
        .AnimFrames = 16
        .Width = 24
        .Height = 24
        .Top = .Height * -1
        .Speed = 1
        .OnScreen = False
        .CollRad = 12
        .ResID = "BOMBPOWER"
        .SurfNo = CreatePowerSurf(.Width * .AnimFrames, .Height, .ResID)
    End With

End Sub

Public Sub ShowPowerUps()
Dim i As Integer
Dim ChkTime As Single
Dim recDisplay As RECT

    For i = 0 To 3
        With PowerUp(i)
            If .OnScreen Then
                .Top = .Top + .Speed
                TestForShipHitPowerUp i
                If .Top >= ScrGame.Bottom Then
                    .OnScreen = False
                End If
                .FrameCounter = .FrameCounter + 1
                If .FrameCounter = 3 Then
                    .FrameNo = .FrameNo + 1
                    .FrameCounter = 0
                End If
                recDisplay.Left = .FrameNo * .Width
                recDisplay.Right = recDisplay.Left + .Width
                recDisplay.Bottom = .Height
                backbuffer.BltFast .Left, .Top, ddsPower(.SurfNo), recDisplay, DDBLTFAST_SRCCOLORKEY Or DDBLTFAST_WAIT
                LocalLight.ShowLight 4, .Left + .Width / 2, .Top + .Height / 2
                If .FrameNo = .AnimFrames Then .FrameNo = 0
            End If
        End With
    Next i

End Sub

Public Sub ResetPowerUps()
Dim i As Integer

    For i = 0 To UBound(PowerUp)
        PowerUp(i).OnScreen = False
    Next i

End Sub
