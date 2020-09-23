Attribute VB_Name = "mDDraw"
Public DX As New DirectX7
Public dd As DirectDraw7
Public binit As Boolean

Public primary As DirectDrawSurface7
Public backbuffer As DirectDrawSurface7
Public ScrGame As RECT
Public GameCtr As Single

Public D3D As Direct3D7
Public D3DDevice As Direct3DDevice7
Public DevEnum As Direct3DEnumDevices

Public ddsSidePanelL As DirectDrawSurface7
Public ddsSidePanelR As DirectDrawSurface7
Public PanelPos(1) As Single
Public ddsLDoor As DirectDrawSurface7
Public ddsRDoor As DirectDrawSurface7
Public recDoor As RECT
Public ShowWidth As Single

Public ddsCursor As DirectDrawSurface7

Public ddsFStar As DirectDrawSurface7
Public ddsSStar As DirectDrawSurface7

Public ddsShip(1) As DirectDrawSurface7
Public ddsPShot As DirectDrawSurface7
Public ddsShield As DirectDrawSurface7
Public recShield As RECT
Public ddsPBomb As DirectDrawSurface7
Public ddsSmShip(1) As DirectDrawSurface7
Public ddsRocTrails As DirectDrawSurface7

Public ddsTitle As DirectDrawSurface7
Public ddsRG As DirectDrawSurface7
Public ddsIntroShip As DirectDrawSurface7
Public recIntroShip As RECT
Public ddsStation As DirectDrawSurface7
Public ddsEarth As DirectDrawSurface7

Public ddsEnemy(4) As DirectDrawSurface7
Public recEnemy(4) As RECT
Public ddsEnShot As DirectDrawSurface7

Public ddsMenuBack As DirectDrawSurface7
Public ddsHBarBack As DirectDrawSurface7
Public ddsOptBack As DirectDrawSurface7

Public ddsExplode As DirectDrawSurface7
Public recExplode As RECT
Public ddsTrails As DirectDrawSurface7
Public recTrails As RECT
Public ddsExplode2 As DirectDrawSurface7
Public recExplode2 As RECT
Public ddsSmoke As DirectDrawSurface7
Public recSmoke As RECT
Public ddsSmExplode As DirectDrawSurface7
Public recSmExplode As RECT

Public ddsLaserCannon As DirectDrawSurface7
Public ddsBarrier As DirectDrawSurface7
Public recBarrier As RECT
Public ddsBigGunLeft As DirectDrawSurface7
Public ddsBigGunRight As DirectDrawSurface7
Public recBigGun As RECT
Public ddsGunStation As DirectDrawSurface7
Public recGunStation As RECT
Public ddsBallShooter As DirectDrawSurface7
Public recBallShooter As RECT
Public ddsPower(3) As DirectDrawSurface7
Public ddsAsteroid(1) As DirectDrawSurface7
Public recAsteroid(1) As RECT
Public ddsAstExplode As DirectDrawSurface7
Public recAstExplode As RECT

Public ddsBack(3) As DirectDrawSurface7
Public BackPos(3) As Single

Public ddsSplash As DirectDrawSurface7
Public ddsThanks As DirectDrawSurface7

Public NewBackSurfNo As Integer

Public LoadJPEG As New cLoadResPicture
Private CursorAnimCounter As Integer
Dim CursFramePos As Integer
Dim recDisplay As RECT


Public Function CreateSurfaceFromBMP(DDSurface As DirectDrawSurface7, Width As Long, Height As Long, strResID As String) As Boolean
On Error GoTo errhandle
Dim ddsdF As DDSURFACEDESC2
Dim Surfpic As IPictureDisp
Dim L_dDDCK As DDCOLORKEY           ' Colorkey for making static surfaces transparent

'much of this sub is from examples at Planet Source Code - I modified it to extract bitmaps from a resource file
'many thanks to all responsible for the original code

' Make transparent
    With L_dDDCK
        .Low = 0
        .High = 0
    End With
   
    ddsdF.lFlags = DDSD_CAPS Or DDSD_HEIGHT Or DDSD_WIDTH
    ddsdF.ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN And DDSCAPS_VIDEOMEMORY
        
    ddsdF.lWidth = Width
    ddsdF.lHeight = Height

    Set Surfpic = LoadResPicture(strResID, vbResBitmap)
    SavePicture Surfpic, App.Path & "\TempFile.Bmp"

    Set DDSurface = dd.CreateSurfaceFromFile(App.Path & "\TempFile.bmp", ddsdF)
    Kill App.Path & "\TempFile.Bmp"
    DoEvents
    Set Surfpic = Nothing
    
    DDSurface.SetColorKey DDCKEY_SRCBLT, L_dDDCK
    
    CreateASurfaceFromBMP = True

    Exit Function

errhandle:

    MsgBox Err.Description
    CreateASurfaceFromBMP = False

End Function

Public Function CreateSurfaceFromJPG(DDSurface As DirectDrawSurface7, Width As Long, Height As Long, strResID As Variant) As Boolean
On Error GoTo errhandle
Dim ddsdF As DDSURFACEDESC2
Dim Surfpic As IPictureDisp
Dim L_dDDCK As DDCOLORKEY           ' Colorkey for making static surfaces transparent

' Make transparent
    With L_dDDCK
        .Low = 0
        .High = 0
    End With
   
    ddsdF.lFlags = DDSD_CAPS Or DDSD_HEIGHT Or DDSD_WIDTH
    ddsdF.ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN And DDSCAPS_VIDEOMEMORY
        
    ddsdF.lWidth = Width
    ddsdF.lHeight = Height

    Set Surfpic = LoadJPEG.LoadResPicture(strResID, "JPG")
    SavePicture Surfpic, App.Path & "\TempFile.jpg"

    Set DDSurface = dd.CreateSurfaceFromFile(App.Path & "\TempFile.jpg", ddsdF)
    Kill App.Path & "\TempFile.jpg"
    DoEvents
    Set Surfpic = Nothing
    
    DDSurface.SetColorKey DDCKEY_SRCBLT, L_dDDCK
    
    CreateASurfaceFromJPG = True

    Exit Function

errhandle:

    MsgBox Err.Description
    CreateSurfaceFromJPG = False

End Function

Public Sub SetupDDrawScreen()
Dim i As Integer
Dim RenderState As String
Dim ddsd1 As DDSURFACEDESC2
Dim ddsd3 As DDSURFACEDESC2

    ScrGame.Left = 0
    ScrGame.Right = 1024
    ScrGame.Bottom = 768
    GameCtr = ScrGame.Right / 2
    
    Set dd = DX.DirectDrawCreate("")
    frmGame.Show
    Call dd.SetCooperativeLevel(frmGame.hWnd, DDSCL_FULLSCREEN Or DDSCL_ALLOWMODEX Or DDSCL_EXCLUSIVE)
    Call dd.SetDisplayMode(1024, 768, 16, 0, DDSDM_DEFAULT)
        
    'get the screen surface and create a back buffer too
    ddsd1.lFlags = DDSD_CAPS Or DDSD_BACKBUFFERCOUNT
    ddsd1.ddsCaps.lCaps = DDSCAPS_PRIMARYSURFACE Or DDSCAPS_FLIP Or DDSCAPS_COMPLEX Or DDSCAPS_3DDEVICE
    ddsd1.lBackBufferCount = 1
    Set primary = dd.CreateSurface(ddsd1)
        
    'Get the backbuffer
    Dim caps As DDSCAPS2
    caps.lCaps = DDSCAPS_BACKBUFFER
    Set backbuffer = primary.GetAttachedSurface(caps)
    backbuffer.GetSurfaceDesc ddsd3
    backbuffer.SetFontTransparency True
    
    Set D3D = dd.GetDirect3D
    Set DevEnum = D3D.GetDevicesEnum
 
    For i = 1 To DevEnum.GetCount
        If DevEnum.GetGuid(i) = "IID_IDirect3DRGBDevice" Then
            RenderState = "IID_IDirect3DRGBDevice"
        ElseIf DevEnum.GetGuid(i) = "IID_IDirect3DHALDevice" Then
            RenderState = "IID_IDirect3DHALDevice"
            Exit For
        End If
    Next i

    Set D3DDevice = D3D.CreateDevice(RenderState, backbuffer)

    'init the surfaces
    InitSurfaces
    
    binit = True

End Sub

Public Sub InitSurfaces()
Dim Result As Boolean

    Result = CreateSurfaceFromJPG(ddsSidePanelL, 150, 768, "LPANEL")
    Result = CreateSurfaceFromJPG(ddsSidePanelR, 150, 768, "RPANEL")
    PanelPos(0) = 768

    ShowWidth = 365: DoorMove = 5
    Result = CreateSurfaceFromJPG(ddsLDoor, 365, 768, "LDOOR")
    Result = CreateSurfaceFromJPG(ddsRDoor, 365, 768, "RDOOR")
    
    Result = CreateSurfaceFromBMP(ddsCursor, 400, 40, "CURSOR")
    
    Result = CreateSurfaceFromBMP(ddsShip(0), 420, 60, "SHIP")
    Result = CreateSurfaceFromBMP(ddsShip(1), 420, 60, "SHIP2")
    Result = CreateSurfaceFromBMP(ddsRocTrails, 40, 12, "SHIPROCKETS")
    Result = CreateSurfaceFromBMP(ddsPShot, 12, 24, "SHOT")
    recShield.Right = 800: recShield.Bottom = 80
    Result = CreateSurfaceFromBMP(ddsShield, recShield.Right, recShield.Bottom, "SHIELD")
    Result = CreateSurfaceFromBMP(ddsPBomb, 200, 20, "BOMBOBJ")
    Result = CreateSurfaceFromBMP(ddsSmShip(0), 30, 30, "SMSHIP")
    Result = CreateSurfaceFromBMP(ddsSmShip(1), 30, 30, "SMSHIP2")

    Result = CreateSurfaceFromBMP(ddsFStar, 4, 4, "FASTSTAR")
    Result = CreateSurfaceFromBMP(ddsSStar, 1, 1, "SLOWSTAR")

    Result = CreateSurfaceFromBMP(ddsEnShot, 20, 20, "ENSHOT")

    recTrails.Right = 200: recTrails.Bottom = 10
    Result = CreateSurfaceFromBMP(ddsTrails, recTrails.Right, recTrails.Bottom, "TRAILS")

    Result = CreateSurfaceFromBMP(ddsTitle, 400, 60, "SB")
    Result = CreateSurfaceFromBMP(ddsRG, 200, 400, "RG")

    Result = CreateSurfaceFromBMP(ddsMenuBack, 200, 40, "MENUBACK")

    recIntroShip.Right = 400: recIntroShip.Bottom = 195
    Result = CreateSurfaceFromJPG(ddsIntroShip, recIntroShip.Right, recIntroShip.Bottom, "INTROSHP")

    Result = CreateSurfaceFromBMP(ddsHBarBack, 200, 65, "HBARBACK")
    Result = CreateSurfaceFromBMP(ddsOptBack, 243, 40, "OPTBACK")
    
    Result = CreateSurfaceFromJPG(ddsStation, 300, 300, "STATION")
    Result = CreateSurfaceFromJPG(ddsEarth, 1024, 468, "EARTH")

    recExplode.Right = 1240: recExplode.Bottom = 100
    Result = CreateSurfaceFromBMP(ddsExplode, recExplode.Right, recExplode.Bottom, "EXPLOSION")
    recExplode.Right = 88
    
    recSmoke.Right = 64: recSmoke.Bottom = 8
    Result = CreateSurfaceFromBMP(ddsSmoke, recSmoke.Right, recSmoke.Bottom, "SMOKE")
    recSmoke.Right = 8
    
    recExplode2.Right = 3328: recExplode2.Bottom = 128
    Result = CreateSurfaceFromBMP(ddsExplode2, recExplode2.Right, recExplode2.Bottom, "EXPLOSION2")
    recExplode2.Right = 128

    recSmExplode.Right = 384: recSmExplode.Bottom = 32
    Result = CreateSurfaceFromBMP(ddsSmExplode, recSmExplode.Right, recSmExplode.Bottom, "SMEXPLODE")
    
    Result = CreateSurfaceFromBMP(ddsLaserCannon, 38, 19, "CANNONS")

    recBarrier.Right = 724: recBarrier.Bottom = 85
    Result = CreateSurfaceFromJPG(ddsBarrier, recBarrier.Right, recBarrier.Bottom, "BARRIER")

    recBigGun.Right = 2400: recBigGun.Bottom = 120
    Result = CreateSurfaceFromJPG(ddsBigGunLeft, recBigGun.Right, recBigGun.Bottom, "BIGGUNL")
    Result = CreateSurfaceFromJPG(ddsBigGunRight, recBigGun.Right, recBigGun.Bottom, "BIGGUNR")
    recBigGun.Right = 120

    recGunStation.Right = 724: recGunStation.Bottom = 166
    Result = CreateSurfaceFromJPG(ddsGunStation, recGunStation.Right, recGunStation.Bottom, "GUNSTATION")
    
    recBallShooter.Right = 500: recBallShooter.Bottom = 25
    Result = CreateSurfaceFromBMP(ddsBallShooter, recBallShooter.Right, recBallShooter.Bottom, "BALLSHOOTER")

    Result = CreateSurfaceFromBMP(ddsAsteroid(0), 600, 40, "ASTEROID1")
    Result = CreateSurfaceFromBMP(ddsAsteroid(1), 600, 40, "ASTEROID2")

    recAstExplode.Right = 80: recAstExplode.Bottom = 10
    Result = CreateSurfaceFromBMP(ddsAstExplode, recAstExplode.Right, recAstExplode.Bottom, "ASTEXPLODE")

    Result = CreateSurfaceFromJPG(ddsSplash, 215, 170, "SPLASH")
    Result = CreateSurfaceFromJPG(ddsThanks, 400, 50, "THANKYOU")

End Sub

Public Function CreateEnemySurf(GetWidth As Long, GetHeight As Long, GetResID As String) As Integer
Dim i As Integer
Dim Result As Boolean

    For i = 0 To UBound(ddsEnemy)
        If ddsEnemy(i) Is Nothing Then
            recEnemy(i).Right = GetWidth
            recEnemy(i).Bottom = GetHeight
            Result = CreateSurfaceFromBMP(ddsEnemy(i), recEnemy(i).Right, recEnemy(i).Bottom, GetResID)
            CreateEnemySurf = i
            Exit For
        End If
    Next i
    
End Function

Public Function CreatePowerSurf(GetWidth As Long, GetHeight As Long, GetResID As String) As Integer
Dim i As Integer
Dim Result As Boolean
Dim recSetup As RECT

    For i = 0 To 3
        If ddsPower(i) Is Nothing Then
            recSetup.Right = GetWidth
            recSetup.Bottom = GetHeight
            Result = CreateSurfaceFromBMP(ddsPower(i), recSetup.Right, recSetup.Bottom, GetResID)
            CreatePowerSurf = i
            Exit For
        End If
    Next i
    
End Function

Public Sub CreateBackGrdSurf(SurfName As String)
Dim Result As Boolean
Dim recSetup As RECT
    
    recSetup.Top = 0: recSetup.Right = 730
    recSetup.Bottom = 768
    
    If NewBackSurfNo = 0 Then
        Result = CreateSurfaceFromJPG(ddsBack(0), recSetup.Right, recSetup.Bottom, SurfName)
        Result = CreateSurfaceFromJPG(ddsBack(1), recSetup.Right, recSetup.Bottom, SurfName)
        BackPos(0) = 768
        BackPos(1) = 0
        NewBackSurfNo = 1
    ElseIf NewBackSurfNo = 1 Then
        Result = CreateSurfaceFromJPG(ddsBack(2), recSetup.Right, recSetup.Bottom, SurfName)
        Result = CreateSurfaceFromJPG(ddsBack(3), recSetup.Right, recSetup.Bottom, SurfName)
        BackPos(2) = 768
        BackPos(3) = 0
        NewBackSurfNo = 0
    End If

End Sub

Public Sub DestroyEnemySurf(GetSurfNo As Integer)

    Set ddsEnemy(GetSurfNo) = Nothing

End Sub

Public Sub SetSurfsToNothing()
Dim i As Integer

    For i = 0 To 3
        Set ddsEnemy(i) = Nothing
        Set ddsPower(i) = Nothing
    Next i
    For i = 0 To 1
        Set ddsAsteroid(i) = Nothing
    Next i
    Set ddsBack(0) = Nothing
    Set ddsBack(1) = Nothing
    Set ddsCursor = Nothing
    Set ddsSidePanelL = Nothing
    Set ddsSidePanelR = Nothing
    Set ddsShip(0) = Nothing
    Set ddsShip(1) = Nothing
    Set ddsShield = Nothing
    Set ddsFStar = Nothing
    Set ddsSStar = Nothing
    Set ddsPShot = Nothing
    Set ddsSmShip(0) = Nothing
    Set ddsSmShip(1) = Nothing
    Set ddsTrails = Nothing
    Set ddsTitle = Nothing
    Set ddsRG = Nothing
    Set ddsRocTrails = Nothing
    Set ddsMenuBack = Nothing
    Set ddsHBar = Nothing
    Set ddsHBarBack = Nothing
    Set ddsExplode = Nothing
    Set ddsSplash = Nothing
    Set ddsThanks = Nothing
    Set ddsLDoor = Nothing
    Set ddsRDoor = Nothing
    Set ddsBigGunLeft = Nothing
    Set ddsBigGunRight = Nothing
    Set ddsGunStation = Nothing
    Set ddsPBomb = Nothing

End Sub

Public Function ExModeActive() As Boolean
'This is used to test if we're in the correct resolution.
Dim TestCoopRes As Long
    
    TestCoopRes = dd.TestCooperativeLevel
    
    If (TestCoopRes = DD_OK) Then
        ExModeActive = True
    Else
        ExModeActive = False
    End If

End Function

Public Sub PrepareBack()
Dim ddrval As Long
Dim bRestore As Boolean

    If binit = False Then Exit Sub
    
    ' this will keep us from trying to blt in case we lose the surfaces (alt-tab)
    bRestore = False
    Do Until ExModeActive
        DoEvents
        bRestore = True
    Loop
    
    ' if we lost and got back the surfaces, then restore them
    DoEvents
    If bRestore Then
        bRestore = False
        dd.RestoreAllSurfaces
        InitSurfaces
    End If

'blt black backdrop for screen
'    If IntroRunning Then
        ddrval = backbuffer.BltColorFill(ScrGame, 0)
'    End If
    
End Sub

Public Sub DrawSidePanels()
On Error GoTo DrawSideError
Dim i As Integer

    recDisplay.Left = 0: recDisplay.Right = 150
    
    For i = 0 To 1
        If PanelPos(i) <= 768 Then
            recDisplay.Top = 768 - PanelPos(i)
            recDisplay.Bottom = 768
            ddrval = backbuffer.BltFast(0, 0, ddsSidePanelL, recDisplay, DDBLTFAST_SRCCOLORKEY Or DDBLTFAST_WAIT)
            ddrval = backbuffer.BltFast(874, 0, ddsSidePanelR, recDisplay, DDBLTFAST_SRCCOLORKEY Or DDBLTFAST_WAIT)
        ElseIf PanelPos(i) < 1536 Then
            recDisplay.Top = 0
            recDisplay.Bottom = 768 - (PanelPos(i) - 768)
            ddrval = backbuffer.BltFast(0, PanelPos(i) - 768, ddsSidePanelL, recDisplay, DDBLTFAST_SRCCOLORKEY Or DDBLTFAST_WAIT)
            ddrval = backbuffer.BltFast(874, PanelPos(i) - 768, ddsSidePanelR, recDisplay, DDBLTFAST_SRCCOLORKEY Or DDBLTFAST_WAIT)
        Else
            PanelPos(i) = 0
        End If
        If Not MainMenuRunning And Not ScoresRunning And Not OptionsRunning And Not CreditsRunning Then
            PanelPos(i) = PanelPos(i) + 1
        End If
    Next i
    
    Exit Sub

DrawSideError:
    MsgBox "Error occurred in DrawSidePanels procedure"
    EndIt

End Sub

Public Sub ShowDoors()

        If DoorMove = 1 Then
            If ShowWidth > 0 Then
                ShowWidth = ShowWidth - 5
            End If
        ElseIf DoorMove = 2 Then
            If ShowWidth < 365 Then
                ShowWidth = ShowWidth + 5
            End If
        End If
    
    recDisplay.Top = 0: recDisplay.Bottom = 768
    recDisplay.Left = 365 - ShowWidth
    recDisplay.Right = 365
    backbuffer.BltFast 148, 0, ddsLDoor, recDisplay, DDBLTFAST_SRCCOLORKEY Or DDBLTFAST_WAIT
    
    recDisplay.Left = 0
    recDisplay.Right = ShowWidth
    backbuffer.BltFast GameCtr + (365 - ShowWidth), 0, ddsRDoor, recDisplay, DDBLTFAST_SRCCOLORKEY Or DDBLTFAST_WAIT


End Sub

Public Sub ShowBack()
On Error GoTo ShowBackErr
Dim i As Integer
Dim Low As Integer, High As Integer

    If NewBackSurfNo = 0 Then
        Low = 2
        High = 3
    ElseIf NewBackSurfNo = 1 Then
        Low = 0
        High = 1
    End If

    recDisplay.Left = 0
    recDisplay.Right = 730
    
    For i = Low To High
        If BackPos(i) < 768 Then
            recDisplay.Top = 768 - BackPos(i)
            recDisplay.Bottom = 768
            backbuffer.BltFast 148, 0, ddsBack(i), recDisplay, DDBLTFAST_SRCCOLORKEY Or DDBLTFAST_WAIT
        ElseIf BackPos(i) < 1536 Then
            recDisplay.Top = 0
            recDisplay.Bottom = 768 - (BackPos(i) - 768)
            backbuffer.BltFast 148, BackPos(i) - 768, ddsBack(i), recDisplay, DDBLTFAST_SRCCOLORKEY Or DDBLTFAST_WAIT
        Else
            BackPos(i) = 0
        End If
        BackPos(i) = BackPos(i) + 0.5
    Next i

    Exit Sub
    
ShowBackErr:
    MsgBox "Error occurred in ShowBack procedure"
    EndIt

End Sub

Public Sub DisplayCursor()
On Error GoTo DisCursErr

    recDisplay.Top = 0
    recDisplay.Bottom = 40
    CursorAnimCounter = CursorAnimCounter + 1
    If CursorAnimCounter = 3 Then
        If CursFramePos < 9 Then
            CursFramePos = CursFramePos + 1
        Else
            CursFramePos = 0
        End If
        CursorAnimCounter = 0
    End If
        
    recDisplay.Left = CursFramePos * 40
    recDisplay.Right = recDisplay.Left + 40
    
    backbuffer.BltFast gintMouseX - 20, gintMouseY - 20, ddsCursor, recDisplay, DDBLTFAST_SRCCOLORKEY Or DDBLTFAST_WAIT
    LocalLight.ShowLight 8, gintMouseX, gintMouseY

    Exit Sub
    
DisCursErr:
    MsgBox "Error occurred in DisplayCursor procedure"
    EndIt
    
End Sub
