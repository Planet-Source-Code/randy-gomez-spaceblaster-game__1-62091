Attribute VB_Name = "mScore"
Public strPlayerName As String
Public Scores(9, 1) As Variant
Public ScoreShowCounter As Long
Public blnShowPrompt As Boolean
Dim recDisplay As RECT

Public Sub GetScores()
On Error GoTo CantLoad
Dim fs, f, ts
Dim i As Integer
Dim strSetScore As String
Dim strName As String
   
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set f = fs.getfile(App.Path & "\ScoreList.txt")
'    Reading = 1, Writing = 2, Appending = 8
'    TristateUseDefault = -2, TristateTrue = -1, TristateFalse = 0
    Set ts = f.OpenAsTextStream(1, -2)

    For i = 0 To 9
        strName = ts.Readline
        strSetScore = Right(strName, Len(strName) - InStr(1, strName, ","))
        strName = Left(strName, InStr(strName, ",") - 1)
        Scores(i, 0) = strName
        Scores(i, 1) = Val(strSetScore)
    Next i

    ts.Close
    Set fs = Nothing

    SortScoresArray
    Exit Sub

CantLoad:
    MsgBox "Error loading scores."
    End

End Sub

Public Sub InsertNewScore()
Dim i As Integer, j As Integer

    For i = 0 To 9
        If Score >= Scores(i, 1) And Score > 0 Then
            If i = 8 Then
                Scores(9, 0) = Scores(8, 0)
                Scores(9, 1) = Scores(8, 1)
            ElseIf i <= 7 Then
                For j = 9 To i + 1 Step -1
                    Scores(j, 0) = Scores(j - 1, 0)
                    Scores(j, 1) = Scores(j - 1, 1)
                Next j
            End If
            Scores(i, 1) = Score
            If strPlayerName > "" Then
                Scores(i, 0) = strPlayerName
            Else
                Scores(i, 0) = "- - -"
            End If
            blnShowPrompt = True
            Exit For
        End If
    Next i

End Sub

Public Sub WriteScoresToFile()
Dim fs, ts, f
Dim j As Integer

    If Dir(App.Path & "\scorelist.txt") <> "" Then
        Kill App.Path & "\scorelist.txt"    'if found, delete the .txt file so we can
    End If                              'create a new one

    Set fs = CreateObject("Scripting.FileSystemObject")
    fs.CreateTextFile App.Path & "\scorelist.txt"            'Create a file
    Set f = fs.getfile(App.Path & "\scorelist.txt")     'open parameters text file
'    Reading = 1, Writing = 2, Appending = 8
'    TristateUseDefault = -2, TristateTrue = -1, TristateFalse = 0
    Set ts = f.OpenAsTextStream(2, -2)
 
    For j = 0 To 9
        ts.Writeline Scores(j, 0) & "," & Scores(j, 1)
    Next j
    
    ts.Close
    Set fs = Nothing

End Sub

Private Sub SortScoresArray()
Dim SaveName As String, SaveScore As Integer
Dim k As Integer, m As Integer

For k = 0 To 8
    For m = 1 To 9
        If Val(Scores(m, 1)) > Val(Scores((m - 1), 1)) Then
            SaveName = Scores((m - 1), 0)
            SaveScore = Scores((m - 1), 1)
            Scores((m - 1), 0) = Scores(m, 0)
            Scores((m - 1), 1) = Scores(m, 1)
            Scores(m, 0) = SaveName
            Scores(m, 1) = SaveScore
        End If
    Next m
Next k

End Sub

Public Sub DoScores()
Dim j As Integer
    

    ShowDoors
    recDisplay.Top = 0: recDisplay.Left = 0
    recDisplay.Right = 400: recDisplay.Bottom = 60
    backbuffer.BltFast GameCtr - 200, 100, ddsTitle, recDisplay, DDBLTFAST_SRCCOLORKEY Or DDBLTFAST_WAIT
    
    If ScoreShowCounter < 76 Then
        If ScoreShowCounter < 35 Then
            ScoreShowCounter = ScoreShowCounter + 1
        End If
        If ScoreShowCounter = 75 Then
            If DoorMove = 2 Then
                GameSounds.stop_snd 2
                GameSounds.play_snd 3
            End If
        End If
    End If

    backbuffer.SetFont MenuFnt
    recDisplay.Right = 200: recDisplay.Bottom = 40
    backbuffer.BltFast 410, 200, ddsMenuBack, recDisplay, DDBLTFAST_SRCCOLORKEY Or DDBLTFAST_WAIT
    
    backbuffer.SetForeColor RGB(255, 255, 0)
    backbuffer.DrawText 470, 206, "TOP TEN", False
    
    For j = 0 To 9
        If Score = Scores(j, 1) And Score > 0 Then
            backbuffer.SetForeColor RGB(200, 200, 0)
            backbuffer.DrawText 580, j * ScoreShowCounter + 260, Scores(j, 1), False
            If blnShowPrompt Then
                backbuffer.DrawText 330, j * ScoreShowCounter + 260, "NAME :", False
            End If
            If strPlayerName > "" Then
                backbuffer.DrawText 410, j * ScoreShowCounter + 260, strPlayerName, False
                Scores(j, 0) = strPlayerName
            End If
        Else
            backbuffer.SetForeColor RGB(140, 140, 0)
            backbuffer.DrawText 410, j * ScoreShowCounter + 260, Scores(j, 0), False
            backbuffer.DrawText 580, j * ScoreShowCounter + 260, Scores(j, 1), False
        End If
    Next j

    DoBackButton
    DisplayCursor

End Sub

