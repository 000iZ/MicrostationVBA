Public Sub edit_layer_color()
    Dim oLvl As Level
    Const targetLvl As String: targetLvl = "Level 52 - Annotation"
    Const newColor As Integer: newColor = 6

    For Each oLvl In ActiveDesignFile.Levels
        If oLvl.name = targetLvl Then
            oLvl.ElementColor = newColor
        End If
    Next

    ActiveDesignFile.Levels.Rewrite

End Sub
