Attribute VB_Name = "mFPS"

'variables for Timing Loop
Public Declare Function QueryPerformanceFrequency Lib "kernel32" (lpFrequency As Currency) As Long
Public Declare Function QueryPerformanceCounter Lib "kernel32" (lpPerformanceCount As Currency) As Long
Public curFreq As Currency
Public curStart As Currency
Public curEnd As Currency
Public curFPS As Currency
Public dblResult As Double
Public TempFPS As Single
Public FPS As Single
Public TargetFPS As Single

Public Sub SetFrameRate(Value As Long)
    
    TargetFPS = Value
    QueryPerformanceFrequency curFreq

End Sub

Public Function GetFPS() As Single
    
    GetFPS = FPS

End Function

Public Sub UpdateFPS()

'do delay first
Dim bRunning As Boolean
    bRunning = False

Do While bRunning = False
    QueryPerformanceCounter curStart

    If (curStart - curEnd) / curFreq >= 1 / TargetFPS Then
        bRunning = True
    End If

Loop
    TempFPS = TempFPS + 1
    QueryPerformanceCounter curEnd

'check fps
    QueryPerformanceCounter curStart
If (curStart - curFPS) / curFreq >= 1 Then
    FPS = TempFPS
    TempFPS = 0
    QueryPerformanceCounter curFPS
End If

End Sub
