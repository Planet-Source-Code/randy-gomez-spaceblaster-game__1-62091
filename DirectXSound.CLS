VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDirectSound"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Private ds As DirectSound
Private dsBuffer() As DirectSoundBuffer
Private dsbuffercount As Single
Private sound As Boolean

Public Sub SetDXVolume(Volume As Long)
Dim i As Integer
'This formula allows you to specify a volume between 0-100; similiar to a percentage

    For i = 0 To 6
        dsBuffer(i).SetVolume (Volume * 60 - 6000)
    Next i

End Sub

Public Sub Load_Sound(FileName As String)
Dim bufferDesc As DSBUFFERDESC
Dim waveFormat As WAVEFORMATEX
     
     bufferDesc.lFlags = DSBCAPS_CTRLVOLUME Or DSBCAPS_STATIC
     waveFormat.nFormatTag = WAVE_FORMAT_PCM
     waveFormat.nChannels = 2
     waveFormat.lSamplesPerSec = 22050
     waveFormat.nBitsPerSample = 16
     waveFormat.nBlockAlign = waveFormat.nBitsPerSample / 8 * waveFormat.nChannels
     waveFormat.lAvgBytesPerSec = waveFormat.lSamplesPerSec * waveFormat.nBlockAlign

    dsbuffercount = dsbuffercount + 1
    ReDim Preserve dsBuffer(dsbuffercount - 1) As DirectSoundBuffer
    Set dsBuffer(dsbuffercount - 1) = ds.CreateSoundBufferFromFile(FileName, bufferDesc, waveFormat)

End Sub

Public Function Init_sound(hWnd As Long) As Boolean
On Error GoTo SoundOut:

    Set ds = DX.DirectSoundCreate("")
    
    ds.SetCooperativeLevel hWnd, DSSCL_PRIORITY
    
    sound = True 'success
    Init_sound = True
    Exit Function

'no sound card, sound file deleted or something just turn sound off

SoundOut:
    sound = False
    Init_sound = False

End Function

Sub play_snd(snd, Optional force As Boolean = False)

    If sound = False Then Exit Sub
    
    If force Then
        dsBuffer(snd).SetCurrentPosition 1
        dsBuffer(snd).Play DSBPLAY_DEFAULT
    Else
        If dsBuffer(snd).GetStatus = 0 Then
            dsBuffer(snd).SetCurrentPosition 1
            dsBuffer(snd).Play DSBPLAY_DEFAULT
        End If
    End If

End Sub

Sub pan_snd(snd, Optional center As Long = 0)

    dsBuffer(snd).SetPan center

End Sub

Sub stop_snd(snd)

    If dsBuffer(snd).GetStatus = DSBSTATUS_PLAYING Then
        dsBuffer(snd).Stop
    End If

End Sub

Function is_playing(snd) As Boolean

    If dsBuffer(snd).GetStatus = DSBSTATUS_PLAYING Then
        is_playing = True
    Else
        is_playing = False
    End If

End Function
