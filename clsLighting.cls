VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLighting"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private Type LIGHTSETTINGS
   LightSurface As DirectDrawSurface7
   Width As Long
   Height As Long
End Type

Private Lights() As LIGHTSETTINGS
Private mvarLightCount As Long
'-----------------------------------------------------------------------------

Public Property Get LightCount() As Long
   LightCount = mvarLightCount
End Property
'-----------------------------------------------------------------------------


Public Function CreateLight(ByVal Radius As Byte, ByVal LightSize As Byte, RColorMod As Single, _
GColorMod As Single, BColorMod As Single, Optional ByVal MinBrightness As Long = 0, _
Optional ByVal Maxbrightness As Long = 255) As Long

On Local Error GoTo Error

Dim dRect As RECT
Dim Height As Long
Dim Counter As Long
Dim SurfaceDesc As DDSURFACEDESC2
Dim Brightness As Long

mvarLightCount = mvarLightCount + 1

ReDim Preserve Lights(1 To mvarLightCount) As LIGHTSETTINGS
 
Set Lights(mvarLightCount).LightSurface = Nothing

Height = (Radius * LightSize) * 2
  
If Sqr(Height ^ 2) Mod 2 <> 0 Then
   CreateLight = 0
   Exit Function
End If

dRect.Bottom = Height
dRect.Right = Height
Lights(mvarLightCount).Width = Height
Lights(mvarLightCount).Height = Height

SurfaceDesc.lFlags = DDSD_CAPS Or DDSD_HEIGHT Or DDSD_WIDTH
SurfaceDesc.ddsCaps.lCaps = DDSCAPS_TEXTURE
SurfaceDesc.ddsCaps.lCaps2 = DDSCAPS2_TEXTUREMANAGE
SurfaceDesc.lTextureStage = 0
SurfaceDesc.lWidth = Height
SurfaceDesc.lHeight = Height
 
Set Lights(mvarLightCount).LightSurface = dd.CreateSurface(SurfaceDesc)

Brightness = Maxbrightness - MinBrightness
 
For Counter = 1 To Radius
   Lights(mvarLightCount).LightSurface.SetForeColor RGB2DX(Int(Counter * (Brightness * RColorMod / Radius)) + MinBrightness, Counter * (Brightness * GColorMod / Radius) + MinBrightness, Counter * (Brightness * BColorMod / Radius) + MinBrightness)
   Lights(mvarLightCount).LightSurface.SetFillColor RGB2DX(Int(Counter * (Brightness * RColorMod / Radius)) + MinBrightness, Counter * (Brightness * GColorMod / Radius) + MinBrightness, Counter * (Brightness * BColorMod / Radius) + MinBrightness)
   Lights(mvarLightCount).LightSurface.DrawCircle Height / 2, Height / 2, ((Radius - (Counter - 1)) * LightSize)
Next Counter

CreateLight = mvarLightCount

Exit Function
Error:
   CreateLight = 0

End Function
'-----------------------------------------------------------------------------


Public Function ShowLight(ByVal LightIndex As Integer, ByVal x As Long, ByVal y As Long) As Boolean

On Error GoTo Error
 
Dim mtlLight(3) As D3DTLVERTEX
Dim lColor As Long

D3DDevice.SetRenderState D3DRENDERSTATE_COLORKEYENABLE, True
lColor = DX.CreateColorRGBA(1, 1, 1, 0.5)
 
DX.CreateD3DTLVertex x - Lights(LightIndex).Width / 2, y - Lights(LightIndex).Height / 2, 0, 1, lColor, 0, 0, 0, mtlLight(0)
DX.CreateD3DTLVertex x + Lights(LightIndex).Width / 2, y - Lights(LightIndex).Height / 2, 0, 1, lColor, 0, 1, 0, mtlLight(1)
DX.CreateD3DTLVertex x - Lights(LightIndex).Width / 2, y + Lights(LightIndex).Height / 2, 0, 1, lColor, 0, 0, 1, mtlLight(2)
DX.CreateD3DTLVertex x + Lights(LightIndex).Width / 2, y + Lights(LightIndex).Height / 2, 0, 1, lColor, 0, 1, 1, mtlLight(3)

D3DDevice.BeginScene
   D3DDevice.SetRenderState D3DRENDERSTATE_ALPHABLENDENABLE, True

   D3DDevice.SetRenderState D3DRENDERSTATE_DESTBLEND, D3DBLEND_ONE
   D3DDevice.SetRenderState D3DRENDERSTATE_SRCBLEND, D3DBLEND_ONE
            
   D3DDevice.SetTexture 0, Lights(LightIndex).LightSurface
   D3DDevice.DrawPrimitive D3DPT_TRIANGLESTRIP, D3DFVF_TLVERTEX, mtlLight(0), 4, D3DDP_DEFAULT
            
   D3DDevice.SetRenderState D3DRENDERSTATE_ALPHABLENDENABLE, False
D3DDevice.EndScene

ShowLight = True

Exit Function
Error:
   ShowLight = False

End Function
'-----------------------------------------------------------------------------


Public Function DrawLight(ByVal LightIndex As Integer, x As Long, y As Long)

On Error GoTo Error

Dim CKey As DDCOLORKEY
Dim rRect As RECT

rRect.Left = 0
rRect.Top = 0
rRect.Bottom = Lights(LightIndex).Height
rRect.Right = Lights(LightIndex).Width

CKey.High = 0
CKey.Low = 0
Lights(LightIndex).LightSurface.SetColorKey DDCKEY_SRCBLT, CKey
backbuffer.BltFast x, y, Lights(LightIndex).LightSurface, rRect, DDBLTFAST_DONOTWAIT Or DDBLTFAST_SRCCOLORKEY

DrawLight = True

Exit Function
Error:
   DrawLight = False
   
End Function
'-----------------------------------------------------------------------------


Public Function CreateLightFromFileFX(ByVal FileName As String, Optional Width As Long = 256, Optional Height As Long = 256) As Long
 
On Error GoTo Error

Dim dRect As RECT
Dim Counter As Long
Dim SurfaceDesc As DDSURFACEDESC2

mvarLightCount = mvarLightCount + 1
 
ReDim Preserve Lights(1 To mvarLightCount) As LIGHTSETTINGS
 
Set Lights(mvarLightCount).LightSurface = Nothing

If Sqr(Height * Width) Mod 2 <> 0 Then
   CreateLightFromFileFX = 0
   Exit Function
End If
 
dRect.Bottom = Height
dRect.Right = Width
Lights(mvarLightCount).Height = Height
Lights(mvarLightCount).Width = Width
 
SurfaceDesc.lFlags = DDSD_CAPS Or DDSD_HEIGHT Or DDSD_WIDTH
SurfaceDesc.ddsCaps.lCaps = DDSCAPS_TEXTURE
SurfaceDesc.ddsCaps.lCaps2 = DDSCAPS2_TEXTUREMANAGE
SurfaceDesc.lTextureStage = 0
SurfaceDesc.lWidth = Width
SurfaceDesc.lHeight = Height
 
Set Lights(mvarLightCount).LightSurface = dd.CreateSurfaceFromFile(FileName, SurfaceDesc)

CreateLightFromFileFX = mvarLightCount

Exit Function
Error:
   CreateLightFromFileFX = 0
   
End Function
'-----------------------------------------------------------------------------

Public Sub DeleteLight()

Set Lights(mvarLightCount).LightSurface = Nothing
mvarLightCount = mvarLightCount - 1
ReDim Preserve Lights(1 To mvarLightCount) As LIGHTSETTINGS

End Sub
'-----------------------------------------------------------------------------

Public Function RGB2DX(r As Long, G As Long, B As Long) As Long
    RGB2DX = DX.CreateColorRGBA(CSng((1 / 255) * r), CSng((1 / 255) * G), CSng((1 / 255) * B), 0)
End Function

