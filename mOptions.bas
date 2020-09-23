Attribute VB_Name = "mOptions"
Public OptionsRunning As Boolean

Public Type objOptionItem
    Left As Single
    Top As Single
    Text As String
    TextLeft As Single
    TextTop As Single
    value As Long
    Hilite As Long
End Type

Public OptionsItem(2) As objOptionItem
Private ShipShowCounter As Integer
Private ShipFrame As Integer
Private HiliteBox As RECT

Public SoundVolume As Long
Public MusicVolume As Integer

Dim recDisplay As RECT
Dim DispLeft As Long
Public ActiveOption As Integer

Public Sub DoOptions()
Dim i As Integer, ItemTop As Single

    ShowDoors
    recDisplay.Top = 0: recDisplay.Left = 0
    recDisplay.Right = 400: recDisplay.Bottom = 60
    backbuffer.BltFast GameCtr - 200, 100, ddsTitle, recDisplay, DDBLTFAST_SRCCOLORKEY Or DDBLTFAST_WAIT
    backbuffer.SetFont MenuFnt
    
    recDisplay.Right = 243: recDisplay.Bottom = 40
    For i = 0 To 2
        If gintMouseX > OptionsItem(0).Left And gintMouseX <= OptionsItem(0).Left + 220 And _
            gintMouseY > OptionsItem(i).Top And gintMouseY <= OptionsItem(i).Top + 40 Then
            ActiveOption = i
            If gblnLMouseButton Then
                ProcessOptionsItem
            End If
            If i = 0 And gblnLMouseButtonUp Then
                GameSounds.play_snd 1, True
            End If
        End If
        With OptionsItem(i)
            backbuffer.BltFast .Left, .Top, ddsOptBack, recDisplay, DDBLTFAST_SRCCOLORKEY Or DDBLTFAST_WAIT
            If i = ActiveOption Then
                If .Hilite < 20 Then
                    .Hilite = .Hilite + 1
                End If
            Else
                If .Hilite > 0 Then
                    .Hilite = .Hilite - 1
                End If
            End If
            backbuffer.SetFillStyle 0
            backbuffer.SetForeColor RGB(100 + .Hilite * 7, 100 + .Hilite * 7, 255)
            backbuffer.SetFillColor RGB(0, 0, 100 + .Hilite * 7)
            backbuffer.DrawText .TextLeft, .TextTop, .Text, False
            backbuffer.SetForeColor backbuffer.GetFillColor
            If i <> 2 Then
                backbuffer.DrawBox .Left + 22, .Top + 10, .Left + .value * 2 + 20, .Top + 26
            End If
        End With
    Next i

    backbuffer.SetFillStyle 0
    backbuffer.SetForeColor RGB(150, 150, 255)
    backbuffer.SetFillColor RGB(0, 0, 50)
    If CurShip = 0 Then
        If HiliteBox.Left > 410 Then HiliteBox.Left = HiliteBox.Left - 5
        HiliteBox.Right = HiliteBox.Left + 80
    ElseIf CurShip = 1 Then
        If HiliteBox.Left < 530 Then HiliteBox.Left = HiliteBox.Left + 5
        HiliteBox.Right = HiliteBox.Left + 80
    End If
    backbuffer.DrawRoundedBox HiliteBox.Left, HiliteBox.Top, HiliteBox.Right, HiliteBox.Bottom, 10, 10
    
    recDisplay.Bottom = 60
    ShipShowCounter = ShipShowCounter + 1
    If ShipShowCounter = 25 Then
        ShipFrame = ShipFrame + 1
        Select Case ShipFrame
            Case 0: DispLeft = 60
            Case 1: DispLeft = 120
            Case 2: DispLeft = 180
            Case 3: DispLeft = 120
            Case 4: DispLeft = 60
            Case 5: DispLeft = 240
            Case 6: DispLeft = 300
            Case 7: DispLeft = 360
            Case 8: DispLeft = 300
            Case 9: DispLeft = 240
        End Select
        If ShipFrame = 9 Then ShipFrame = 0
        ShipShowCounter = 0
    End If
    recDisplay.Left = DispLeft
    recDisplay.Right = recDisplay.Left + 60
    backbuffer.BltFast 420, 510, ddsShip(0), recDisplay, DDBLTFAST_SRCCOLORKEY Or DDBLTFAST_WAIT
    backbuffer.BltFast 540, 510, ddsShip(1), recDisplay, DDBLTFAST_SRCCOLORKEY Or DDBLTFAST_WAIT
        
    backbuffer.DrawText 415, 580, "GX-012           CX-010", False
    
    
    DoBackButton
    DisplayCursor

    If (gintMouseX > 420 And gintMouseX < 480) And (gintMouseY > 510 And gintMouseY < 570) Then
        If gblnLMouseButtonUp Then
            CurShip = 0
        End If
    ElseIf (gintMouseX > 540 And gintMouseX < 600) And (gintMouseY > 510 And gintMouseY < 570) Then
        If gblnLMouseButtonUp Then
            CurShip = 1
        End If
    End If
    
End Sub

Public Sub SetupOptions()

    OptionsItem(0).Left = 390
    OptionsItem(0).Top = 200
    OptionsItem(0).TextTop = 175
    OptionsItem(1).Left = 390
    OptionsItem(1).Top = 320
    OptionsItem(1).TextTop = 295
    OptionsItem(2).Left = 390
    OptionsItem(2).Top = 450
    OptionsItem(2).TextTop = 456

    OptionsItem(0).Text = "SOUND VOLUME"
    OptionsItem(0).value = SoundVolume
    OptionsItem(0).TextLeft = 440
    
    OptionsItem(1).Text = "MUSIC VOLUME"
    OptionsItem(1).value = MusicVolume
    OptionsItem(1).TextLeft = 440

    OptionsItem(2).Text = "CHOOSE YOUR SHIP"
    OptionsItem(2).TextLeft = 420

    HiliteBox.Top = 500
    HiliteBox.Bottom = HiliteBox.Top + 80

    If CurShip = 0 Then
        HiliteBox.Left = 410
        HiliteBox.Right = HiliteBox.Left + 80
    ElseIf CurShip = 1 Then
        HiliteBox.Left = 530
        HiliteBox.Right = HiliteBox.Left + 80
    End If

End Sub

Public Sub ProcessOptionsItem()

    If ActiveOption = 0 Then
        If SoundVolume >= 0 And SoundVolume <= 100 Then
            If gintMouseX > OptionsItem(0).Left Then
                OptionsItem(0).value = (gintMouseX - OptionsItem(0).Left) / 2
                If OptionsItem(0).value > 100 Then OptionsItem(0).value = 100
                If OptionsItem(0).value < 0 Then OptionsItem(0).value = 100
                SoundVolume = OptionsItem(0).value
                GameSounds.SetDXVolume SoundVolume
            End If
        End If
     ElseIf ActiveOption = 1 Then
         If MusicVolume >= 0 And MusicVolume <= 100 Then
            If gintMouseX > OptionsItem(1).Left Then
                OptionsItem(1).value = (gintMouseX - OptionsItem(1).Left) / 2
                If OptionsItem(1).value > 100 Then OptionsItem(1).value = 100
                If OptionsItem(1).value < 0 Then OptionsItem(1).value = 100
                MusicVolume = OptionsItem(1).value
                Music.Volume = 0 - (5000 - MusicVolume * 50)
            End If
         End If
     End If

End Sub

Public Sub WriteOptionValuesToFile()
Dim fs, ts, f
Dim j As Integer

    If Dir(App.Path & "\settings.txt") <> "" Then
        Kill App.Path & "\settings.txt"    'if found, delete the .txt file so we can
    End If                              'create a new one

    Set fs = CreateObject("Scripting.FileSystemObject")
    fs.CreateTextFile App.Path & "\settings.txt"            'Create a file
    Set f = fs.getfile(App.Path & "\settings.txt")     'open parameters text file
'    Reading = 1, Writing = 2, Appending = 8
'    TristateUseDefault = -2, TristateTrue = -1, TristateFalse = 0
    Set ts = f.OpenAsTextStream(2, -2)
 
    ts.Writeline "FPS:" & GameSpeed
    ts.Writeline "SOUNDVOL:" & SoundVolume
    ts.Writeline "MUSICVOL:" & MusicVolume
    
    ts.Close
    Set fs = Nothing

End Sub

