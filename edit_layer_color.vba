Public Sub edit_layer_color()
    Dim oLvl As Level
    Const targetLvl As String: targetLvl = "Defpoints"
    Const newColor As Integer: newColor = 6

    MsgBox "Changing the color of Level '" & targetLvl & "' to " & newColor

    For Each oLvl In ActiveDesignFile.Levels
        If oLvl.name = targetLvl Then
            oLvl.ElementColor = newColor
        
        End If
    
    Next oLvl

    ActiveDesignFile.Levels.Rewrite

End Sub
