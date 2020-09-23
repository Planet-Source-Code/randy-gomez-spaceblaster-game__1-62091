Attribute VB_Name = "mShots"

Public Sub ShowshipFiring()
Dim i As Integer

        
    With Player
        
        For i = 1 To 15
            If .shipAmmo(i).Fired = True Then
                .shipAmmo(i).Y = .shipAmmo(i).Y - .shipAmmo(i).Move
                If .shipAmmo(i).Power = 0 Then
                    backbuffer.BltFast .shipAmmo(i).x, .shipAmmo(i).Y, ddsPShot, recPShot, DDBLTFAST_SRCCOLORKEY Or DDBLTFAST_WAIT
                ElseIf .shipAmmo(i).Power = 1 Then
                    .shipAmmo(i).SideXR = .shipAmmo(i).SideXR + .shipAmmo(i).Move * Cos(1.05)
                    .shipAmmo(i).SideXL = .shipAmmo(i).SideXL - .shipAmmo(i).Move * Cos(1.05)
                    .shipAmmo(i).SideY = .shipAmmo(i).SideY + .shipAmmo(i).Move * Sin(1.05) * direction
                    BitBlt frmGame.hdc, .shipAmmo(i).x, .shipAmmo(i).Y, 6, 24, frmSprites.picShot(.shipAmmo(i).picno).hdc, 0, 0, vbSrcPaint
                    BitBlt frmGame.hdc, .shipAmmo(i).SideXL, .shipAmmo(i).SideY, 17, 24, frmSprites.picShot(.shipAmmo(i).picno + 1).hdc, 0, 0, vbSrcPaint
                    BitBlt frmGame.hdc, .shipAmmo(i).SideXR, .shipAmmo(i).SideY, 17, 24, frmSprites.picShot(.shipAmmo(i).picno + 2).hdc, 0, 0, vbSrcPaint
                End If
                
                If .shipAmmo(i).Y < -32 Then
                    .shipAmmo(i).Fired = False
                End If
            End If
        Next i
        
    End With

End Sub

