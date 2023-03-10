Public Sub edit_layer_color()
    Dim oLvl As Level
    Dim oScan As ElementScanCriteria
    Dim oEE As ElementEnumerator
    
    Const targetLvl As String = "Defpoints"
    Const newColor As Integer = 6

    Set oLvl = ActiveModelReference.Levels(targetLvl)
    'If (oLvl Is Nothing) Then
    '    MsBox "Level '" & targetLvl & "' does not exist"
    '    Exit Sub
    '
    'End If

    Set oScan = New ElementScanCriteria
    oScan.ExcludeAllLevels
    oScan.IncludeLevel oLvl
    
    Set oEE = ActiveModelReference.Scan(oScan)

    MsgBox "Changing the color of Level '" & targetLvl & "' to " & newColor

    ' This only changes the by-level color,  elements not using by-level will not be alterated
    oLvl.ElementColor = newColor
    ActiveDesignFile.Levels.Rewrite


    ' Find all elements on the target level that have not had their color alterated
    Do While oEE.MoveNext
        If oEE.Current.Color <> newColor Then
            oEE.Current.Color = newColor
            
            oEE.Current.Rewrite

        End If
    
    Loop

    ActiveDesignFile.Levels.Rewrite

End Sub
