Attribute VB_Name = "mStars"
Public Type StarObject
    X As Single
    Y As Single
    Move As Single
End Type

Public SlowStar(50) As StarObject
Public FasterStar(16) As StarObject

Public Sub SetupStars()
Dim i As Integer
    
    'set up position, speed and colour (white) of faster moving stars
    For i = 1 To 16
        FasterStar(i).X = Rnd * (ScrGame.Right - ScrGame.Left) + ScrGame.Left
        FasterStar(i).Y = Rnd * ScrGame.Bottom
        FasterStar(i).Move = Rnd * 1.25 + 0.75
    Next i
        
    'set up position, speed and color(shades of grey) of slow stars
    For i = 0 To 50
        SlowStar(i).X = Rnd * (ScrGame.Right - ScrGame.Left) + ScrGame.Left
        SlowStar(i).Y = Rnd * ScrGame.Bottom
        SlowStar(i).Move = Rnd + 0.5
    Next i

End Sub

Public Sub ShowStars()
Dim i As Integer
Dim ReturnVal As Long
Dim recDisplay As RECT
    
    If Level(CurrentLevel).ShowStars And GameRunning Then
        For i = 0 To 50
            If i < 11 Then
                FasterStar(i).Y = FasterStar(i).Y + FasterStar(i).Move
                recDisplay.Right = 4: recDisplay.Bottom = 4
                ReturnVal = backbuffer.BltFast(FasterStar(i).X, FasterStar(i).Y, ddsFStar, recDisplay, DDBLTFAST_SRCCOLORKEY Or DDBLTFAST_WAIT)
                If FasterStar(i).Y > ScrGame.Bottom Then
                    FasterStar(i).X = Rnd * (ScrGame.Right - ScrGame.Left) + ScrGame.Left
                    FasterStar(i).Y = 0
                    FasterStar(i).Move = Rnd * 1.25 + 0.75
                End If
            End If
            SlowStar(i).Y = SlowStar(i).Y + SlowStar(i).Move
            recDisplay.Right = 1: recDisplay.Bottom = 1
            ReturnVal = backbuffer.BltFast(SlowStar(i).X, SlowStar(i).Y, ddsSStar, recDisplay, DDBLTFAST_SRCCOLORKEY Or DDBLTFAST_WAIT)
            
            If SlowStar(i).Y > ScrGame.Bottom Then
                SlowStar(i).X = Rnd * (ScrGame.Right - ScrGame.Left) + ScrGame.Left
                SlowStar(i).Y = 0
                SlowStar(i).Move = Rnd + 0.5
            End If
        Next i
    End If

End Sub
