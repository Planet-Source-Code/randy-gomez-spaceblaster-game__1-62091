Attribute VB_Name = "mPaths"

Public Sub DoRightArcAndDrop(GetEnemy As EnemyObject)
Dim CtrX As Single, CtrY As Single

    With GetEnemy
        .PathCounter = .PathCounter + 1
        If .PathCounter = 1 Then
            .PathAngle = 3.15
        ElseIf .PathCounter < 250 Then
            CtrX = 600
            CtrY = -60
            .PathAngle = .PathAngle + 0.0175 * ((250 - .PathCounter) / 250)
            .Left = CtrX + 280 * Cos(.PathAngle)
            .Top = CtrY - 280 * Sin(.PathAngle)
        ElseIf .PathCounter > 270 And .PathCounter < 370 Then
            .Top = .Top + (.PathCounter - 220) / 10
        End If
    End With

End Sub

Public Sub DoLeftArcAndDrop(GetEnemy As EnemyObject)
Dim CtrX As Single, CtrY As Single

    With GetEnemy
        .PathCounter = .PathCounter + 1
        If .PathCounter = 1 Then
            .PathAngle = 0
        ElseIf .PathCounter < 250 Then
            CtrX = 360
            CtrY = -60
            .PathAngle = .PathAngle - 0.0175 * ((250 - .PathCounter) / 250)
            .Left = CtrX + 280 * Cos(.PathAngle)
            .Top = CtrY - 280 * Sin(.PathAngle)
        ElseIf .PathCounter > 270 And .PathCounter < 370 Then
            .Top = .Top + (.PathCounter - 220) / 10
        End If
    End With

End Sub


