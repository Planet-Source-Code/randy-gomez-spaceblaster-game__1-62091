VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsGamma"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private GammaSupport As Boolean
Private GammaControler As DirectDrawGammaControl
Private GammaRamp As DDGAMMARAMP
Private OriginalRamp As DDGAMMARAMP

Public Event ErrorOccur(ByVal Procedure As String, ByVal ErrorCode As Long)

Public Sub UpdateGamma(ByVal Red As Integer, ByVal Green As Integer, ByVal Blue As Integer)

On Local Error GoTo Error

Dim Counter As Integer

For Counter = 0 To 255
    If Red < 0 Then GammaRamp.Red(Counter) = ConvToSignedValue(ConvToUnSignedValue(OriginalRamp.Red(Counter)) * (100 - Abs(Red)) / 100)
    If Red = 0 Then GammaRamp.Red(Counter) = OriginalRamp.Red(Counter)
    If Red > 0 Then GammaRamp.Red(Counter) = ConvToSignedValue(65535 - ((65535 - ConvToUnSignedValue(OriginalRamp.Red(Counter))) * (100 - Red) / 100))
    If Green < 0 Then GammaRamp.Green(Counter) = ConvToSignedValue(ConvToUnSignedValue(OriginalRamp.Green(Counter)) * (100 - Abs(Green)) / 100)
    If Green = 0 Then GammaRamp.Green(Counter) = OriginalRamp.Green(Counter)
    If Green > 0 Then GammaRamp.Green(Counter) = ConvToSignedValue(65535 - ((65535 - ConvToUnSignedValue(OriginalRamp.Green(Counter))) * (100 - Green) / 100))
    If Blue < 0 Then GammaRamp.Blue(Counter) = ConvToSignedValue(ConvToUnSignedValue(OriginalRamp.Blue(Counter)) * (100 - Abs(Blue)) / 100)
    If Blue = 0 Then GammaRamp.Blue(Counter) = OriginalRamp.Blue(Counter)
    If Blue > 0 Then GammaRamp.Blue(Counter) = ConvToSignedValue(65535 - ((65535 - ConvToUnSignedValue(OriginalRamp.Blue(Counter))) * (100 - Blue) / 100))
Next

GammaControler.SetGammaRamp DDSGR_DEFAULT, GammaRamp

Exit Sub
Error:
   RaiseEvent ErrorOccur("MEgammaRamp: UpdateGamma", Err.Number)

End Sub

Public Function CreateGammaRamp() As Boolean

On Local Error GoTo Error

Dim Hard As DDCAPS, Soft As DDCAPS

dd.GetCaps Hard, Soft

If (Hard.lCaps2 And DDCAPS2_PRIMARYGAMMA) = 0 Then
   CreateGammaRamp = False
Else
   Set GammaControler = primary.GetDirectDrawGammaControl
   GammaControler.GetGammaRamp DDSGR_DEFAULT, OriginalRamp
      
   CreateGammaRamp = True
End If

Exit Function
Error:
   RaiseEvent ErrorOccur("MEgammaRamp: CreateGammaRamp", Err.Number)

End Function

Private Function ConvToSignedValue(ByVal value As Long) As Integer

On Local Error GoTo Error

If value <= 32767 Then
   ConvToSignedValue = CInt(value)
   Exit Function
End If
ConvToSignedValue = CInt(value - 65535)

Exit Function
Error:
   RaiseEvent ErrorOccur("MEgammaRamp: ConvToSignedValue", Err.Number)

End Function

Private Function ConvToUnSignedValue(ByVal value As Integer) As Long

On Local Error GoTo Error

If value >= 0 Then
   ConvToUnSignedValue = value
   Exit Function
End If
ConvToUnSignedValue = value + 65535

Exit Function
Error:
   RaiseEvent ErrorOccur("MEgammaRamp: ConvToUnSignedValue", Err.Number)

End Function
