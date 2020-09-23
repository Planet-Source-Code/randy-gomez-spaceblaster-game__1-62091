VERSION 5.00
Begin VB.Form frmGame 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "S p a c e B l a s t e r"
   ClientHeight    =   3090
   ClientLeft      =   0
   ClientTop       =   -105
   ClientWidth     =   4215
   BeginProperty Font 
      Name            =   "Gill Sans MT Condensed"
      Size            =   20.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3090
   ScaleWidth      =   4215
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "frmGame"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyEscape Then
        If IntroRunning Or PauseMenuRunning Then
            If IntroRunning Then GammaValue = -100
            ScreenGamma.UpdateGamma 0, 0, 0
            IntroRunning = False
            ResetEnemies
            InsertNewScore
            SetupMainMenuItems
            MainMenuRunning = True
            PauseMenuRunning = False
            GameRunning = True
            FirstTime = True
            ResetGame = True
            ScrGame.Left = 150
            ScrGame.Right = 874
            CreateBackGrdSurf Level(0).BackGrd
            DoorMove = 2
            Music.StopPlaying
            Music.FileName = App.Path & "\Sounds\menumusic.mid"
            Music.Play
            GameSounds.SetDXVolume SoundVolume
        ElseIf ScoresRunning Then
            ScoresRunning = False
            SetupMainMenuItems
            MainMenuRunning = True
        ElseIf MainMenuRunning Then
            MainMenuRunning = False
            GameRunning = False
            ExitGameCounter = 100
            ExitGame = True
        ElseIf CreditsRunning Then
            CreditsRunning = False
            SetupMainMenuItems
            MainMenuRunning = True
        ElseIf OptionsRunning Then
            OptionsRunning = False
            WriteOptionValuesToFile
            SetupMainMenuItems
            MainMenuRunning = True
        Else
            PauseMenuRunning = True
            SetupPauseMenuItems
        End If
    End If

    If MainMenuRunning Then
        If KeyCode = vbKeyReturn Then
            ProcessMainMenuItem ActiveItem
        End If
    ElseIf PauseMenuRunning Then
        If KeyCode = vbKeyReturn Then
            ProcessPauseMenuItem ActiveItem
        End If
    ElseIf ScoresRunning Then
        If InStr(1, "ABCDEFGHIJKLMNOPQRSTUVWXYZ", UCase(Chr(KeyCode))) <> 0 _
            Or InStr(1, "1234567890", Chr(KeyCode)) <> 0 Then
            If Len(strPlayerName) < 10 Then
                strPlayerName = strPlayerName + Chr(KeyCode)
            End If
        ElseIf KeyCode = vbKeyBack Then
            If strPlayerName > "" Then
                strPlayerName = Left(strPlayerName, Len(strPlayerName) - 1)
            End If
        ElseIf KeyCode = vbKeyReturn Then
            WriteScoresToFile
            blnShowPrompt = False
        End If
    ElseIf OptionsRunning Then
        If KeyCode = vbKeyUp Then
            If ActiveOption > 0 Then ActiveOption = ActiveOption - 1
        ElseIf KeyCode = vbKeyDown Then
            If ActiveOption < 2 Then ActiveOption = ActiveOption + 1
        ElseIf KeyCode = vbKeyRight Then
            If ActiveOption = 0 Then
                If SoundVolume < 100 Then
                    SoundVolume = SoundVolume + 4
                    If SoundVolume > 100 Then SoundVolume = 100
                    OptionsItem(0).value = SoundVolume
                    GameSounds.SetDXVolume SoundVolume
                    GameSounds.play_snd 1, True
                End If
            ElseIf ActiveOption = 1 Then
                If MusicVolume < 100 Then
                    MusicVolume = MusicVolume + 4
                    If MusicVolume > 100 Then MusicVolume = 0
                    OptionsItem(1).value = MusicVolume
                    Music.Volume = 0 - (5000 - MusicVolume * 50)
                End If
            ElseIf ActiveOption = 2 Then
                CurShip = 1
            End If
            
        ElseIf KeyCode = vbKeyLeft Then
            If ActiveOption = 0 Then
                If SoundVolume > 0 Then
                    SoundVolume = SoundVolume - 4
                    If SoundVolume < 0 Then SoundVolume = 0
                    OptionsItem(0).value = SoundVolume
                    GameSounds.SetDXVolume SoundVolume
                    GameSounds.play_snd 1, True
                End If
            ElseIf ActiveOption = 1 Then
                If MusicVolume > 0 Then
                    MusicVolume = MusicVolume - 4
                    If MusicVolume < 0 Then MusicVolume = 0
                    OptionsItem(1).value = MusicVolume
                    Music.Volume = 0 - (5000 - MusicVolume * 50)
                End If
            ElseIf ActiveOption = 2 Then
                CurShip = 0
            End If
        End If
    End If

End Sub

Private Sub Form_Load()

'    ShowCursor 0

End Sub

Private Sub Form_Unload(Cancel As Integer)

    IntroRunning = False
    GameRunning = False
    Set frmGame = Nothing
    End
    
End Sub

