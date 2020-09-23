Attribute VB_Name = "mDInput"
'Note from R.G.:
'I modified this module for this project to include mouse handling only
'I found DirectX keyboard handling too difficult to work with (e.g. timing issues)
'It was easier to use the game form KeyDown events instead
'I am still grateful to Ryan Clark for the mouse stuff though!
'**************************************************************
'
' THIS WORK, INCLUDING THE SOURCE CODE, DOCUMENTATION
' AND RELATED MEDIA AND DATA, IS PLACED INTO THE PUBLIC DOMAIN.
'
' THE ORIGINAL AUTHOR IS RYAN CLARK.
'
' THIS SOFTWARE IS PROVIDED AS-IS WITHOUT WARRANTY
' OF ANY KIND, NOT EVEN THE IMPLIED WARRANTY OF
' MERCHANTABILITY. THE AUTHOR OF THIS SOFTWARE,
' ASSUMES _NO_ RESPONSIBILITY FOR ANY CONSEQUENCE
' RESULTING FROM THE USE, MODIFICATION, OR
' REDISTRIBUTION OF THIS SOFTWARE.
'
'**************************************************************
'
' This file was downloaded from The Game Programming Wiki.
' Come and visit us at http://gpwiki.org
'
'**************************************************************

Option Explicit

'dX Variables
Dim mobjDI As DirectInput
Dim mobjDIMouse As DirectInputDevice
Dim mobjDIMState As DIMOUSESTATE

Const MOUSE_SPEED = 1.5               'Speed of mouse cursor movement
Const CURSOR_RADIUS = 3             'Radius of mouse cursor circle

Global gintMouseX As Integer           'X Coordinate of the mouse cursor
Global gintMouseY As Integer           'Y Coordinate of the mouse cursor
Global MouseMovedLR As Integer
Global gblnLMouseButton As Boolean     'Is the left mouse button being pressed?
Global gblnRMouseButton As Boolean     'Is the right mouse button being pressed?
Global gblnLMouseButtonUp As Boolean   'Was the left mouse button just released?
Global gblnRMouseButtonUp As Boolean   'Was the right mouse button just released?

Public Sub Initialize(frmInit As Form)

    'Create the direct input object
    On Local Error GoTo DIERROR
    Set mobjDI = DX.DirectInputCreate()
        
   'Aquire the mouse as the diMouse device
    On Local Error GoTo DIMOUSEERROR
    Set mobjDIMouse = mobjDI.CreateDevice("GUID_SysMouse")
    
    'Get mouse input exclusively, but only when in foreground mode
    mobjDIMouse.SetCommonDataFormat DIFORMAT_MOUSE
    mobjDIMouse.SetCooperativeLevel frmInit.hWnd, DISCL_FOREGROUND Or DISCL_EXCLUSIVE
    mobjDIMouse.Acquire
    
    'Initialize the mouse variables
    gintMouseX = 512
    gintMouseY = 700
    gblnLMouseButton = False
    gblnRMouseButton = False
    
    'Exit before error code
    On Error GoTo 0
    Exit Sub
    
'Error handlers
DIERROR:
    Terminate
    MsgBox "Error initializing DirectInput.  Please report this to support@rookscape.com along with your system specifications, the log.txt file, and any information you think may be helpful."
    End
DIMOUSEERROR:
    Terminate
    MsgBox "Can't acquire mouse.  Please report this to support@rookscape.com along with your system specifications, the log.txt file, and any information you think may be helpful."
    End
    
End Sub

Public Sub Terminate()
    
    'Unaquire and destroy
    On Local Error GoTo DITERMERROR
    mobjDIMouse.Unacquire
    Set mobjDIMouse = Nothing
    Set mobjDI = Nothing
    
    'Exit before error code
    On Error GoTo 0
    Exit Sub
    
'Error handlers
DITERMERROR:
    Terminate
    MsgBox "Error terminating DirectInput.  Please report this to support@rookscape.com along with your system specifications, the log.txt file, and any information you think may be helpful."
    End
    
End Sub

Public Sub RefreshMouseState()

    'Get the current state of the mouse
    On Local Error Resume Next
    mobjDIMouse.GetDeviceStateMouse mobjDIMState
    
    'R.G. note - I added this code to detect how or if the mouse is moving horizontally
    If mobjDIMState.X = 0 Then
        MouseMovedLR = 0
    ElseIf mobjDIMState.X > 0 Then
        MouseMovedLR = 1
    ElseIf mobjDIMState.X < 0 Then
        MouseMovedLR = 2
    End If
    
    'If we've been forced to unaquire, try to reaquire
    If Err.Number <> 0 Then mobjDIMouse.Acquire
    'If this fails, exit sub
    If Err.Number <> 0 Then Exit Sub
    On Error GoTo 0
    
    'Adjust the mouse cursor x coordinate
    On Local Error GoTo DIMOUSESCANERROR
    gintMouseX = gintMouseX + mobjDIMState.X * MOUSE_SPEED
    If gintMouseX < 150 Then gintMouseX = 150
    If gintMouseX > 874 Then gintMouseX = 874
    
    'Adjust the mouse cursor y coordinate
    gintMouseY = gintMouseY + mobjDIMState.Y * MOUSE_SPEED
    If gintMouseY < 50 Then gintMouseY = 50
    If gintMouseY > 727 Then gintMouseY = 727
    
    'Check the left mouse button state
    gblnLMouseButtonUp = False
    If mobjDIMState.buttons(0) <> 0 Then gblnLMouseButton = True
    If mobjDIMState.buttons(0) = 0 Then
        'If it WAS down, but not anymore, set released
        If gblnLMouseButton = True Then gblnLMouseButtonUp = True
        gblnLMouseButton = False
    End If
    
    'Check the right mouse button state
    gblnRMouseButtonUp = False
    If mobjDIMState.buttons(1) <> 0 Then gblnRMouseButton = True
    If mobjDIMState.buttons(1) = 0 Then
        'If it WAS down, but not anymore, set released
        If gblnRMouseButton = True Then gblnRMouseButtonUp = True
        gblnRMouseButton = False
    End If
        
    'Exit before error code
    On Error GoTo 0
    Exit Sub
    
'Error handlers
DIMOUSESCANERROR:
    Terminate
    MsgBox "Error polling mouse for input.  Please report this to support@rookscape.com along with your system specifications, the log.txt file, and any information you think may be helpful."
    End
                
End Sub


Public Sub ResetMousePosition()

    

End Sub
