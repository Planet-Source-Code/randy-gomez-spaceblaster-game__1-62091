Attribute VB_Name = "modGraphics"
Public DX As New DirectX7
Public dd As DirectDraw7
Public primary As DirectDrawSurface7
Public backbuffer As DirectDrawSurface7

Public sship As DirectDrawSurface7
Public sExplo As DirectDrawSurface7
Public sParticle As DirectDrawSurface7
Public sTrailBit As DirectDrawSurface7

Public rShip As RECT
Public rExplo As RECT
Public rParticle As RECT
Public rTrailBit As RECT

Public ddsd1 As DDSURFACEDESC2
Public ddsd3 As DDSURFACEDESC2


Public Function CreateASurface(DirectdrawObject As DirectDraw7, DDSurface As DirectDrawSurface7, Width As Long, Height As Long, SourceFile As String) As Boolean
On Error GoTo errhandle
Dim ddsdF As DDSURFACEDESC2
Dim Surfpic As Picture
Dim L_dDDCK As DDCOLORKEY           ' Colorkey for making static surfaces transparent

'much of this sub is borrowed from examples at Planet Source Code
'many thanks to all responsible

' Make transparent
    With L_dDDCK
        .low = 0
        .high = 0
    End With
   

    ddsdF.lFlags = DDSD_CAPS Or DDSD_HEIGHT Or DDSD_WIDTH
    ddsdF.ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN And DDSCAPS_VIDEOMEMORY
        
    ddsdF.lWidth = Width
    ddsdF.lHeight = Height

    Set Surfpic = LoadPicture(SourceFile)
    SavePicture Surfpic, App.Path & "\TempFile.Bmp"

    Set DDSurface = DirectdrawObject.CreateSurfaceFromFile(App.Path & "\TempFile.bmp", ddsdF)
    Kill App.Path & "\TempFile.Bmp"
    DoEvents
    Set Surfpic = Nothing
    
    DDSurface.SetColorKey DDCKEY_SRCBLT, L_dDDCK
    
    CreateASurface = True

    Exit Function

errhandle:

    CreateASurface = False

End Function

Public Sub InitSurfaces()
Dim Result As Boolean
    
    rShip.Right = 350: rShip.Bottom = 50
    Result = CreateASurface(dd, StoredShip(1).ShSurf, rShip.Right, rShip.Bottom, PathtoBMP & "\ship.bmp")

    rExplo.Right = 576: rExplo.Bottom = 48
    Result = CreateASurface(dd, sExplo, rExplo.Right, rExplo.Bottom, PathtoBMP & "\explo.bmp")

    rParticle.Right = 200: rParticle.Bottom = 10
    Result = CreateASurface(dd, sParticle, rParticle.Right, rParticle.Bottom, PathtoBMP & "\trails.bmp")


End Sub

Private Sub InitGraphics()
On Local Error GoTo errOut
Dim i As Integer
    
    Set dd = DX.DirectDrawCreate("")
    Me.Show
    Call dd.SetCooperativeLevel(Me.hWnd, DDSCL_FULLSCREEN Or DDSCL_ALLOWMODEX Or DDSCL_EXCLUSIVE)
    Call dd.SetDisplayMode(GameWidth, GameHeight, 16, 0, DDSDM_DEFAULT)
        
    'get the screen surface and create a back buffer too
    ddsd1.lFlags = DDSD_CAPS Or DDSD_BACKBUFFERCOUNT
    ddsd1.ddsCaps.lCaps = DDSCAPS_PRIMARYSURFACE Or DDSCAPS_FLIP Or DDSCAPS_COMPLEX
    ddsd1.lBackBufferCount = 1
    Set primary = dd.CreateSurface(ddsd1)
    
    'Get the backbuffer
    Dim caps As DDSCAPS2
    caps.lCaps = DDSCAPS_BACKBUFFER
    Set backbuffer = primary.GetAttachedSurface(caps)
    backbuffer.GetSurfaceDesc ddsd3
    
    Dim fnt As New StdFont
    fnt.Size = 16
    fnt.Name = "Arial Bold"
    backbuffer.SetFont fnt
    backbuffer.SetForeColor vbWhite
    
    'init the surfaces
    InitSurfaces
    
    binit = True
    bRunning = True
    DoGameLoop

errOut:
    
    MsgBox "Error!"
    EndIt

End Sub

Private Sub DoGameLoop()
On Error GoTo errOut
    'This is the main loop. It only runs while brunning=true
    
    Do
        Blt
        RegGameSpeed
        DoEvents
    
    Loop Until bRunning = False

errOut:
    
    MsgBox "Error!"
    EndIt

End Sub

