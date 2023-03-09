Public Sub edit_layer_color()
    Dim oLvl As Level
    Dim oScan As ElementScanCriteria
    Dim oEE As ElementEnumerator
    
    Const targetLvl As String: targetLvl = "Defpoints"
    Const newColor As Integer: newColor = 6

    Set oScan = New ElementScanCriteria
    oScan.ExcludeAllLevels
    oScan.IncludeLevel targetLvl
    
    Set oEE = ActiveModelReference.Scan(oScan)

    MsgBox "Changing the color of Level '" & targetLvl & "' to " & newColor

    For Each oLvl In ActiveDesignFile.Levels
        If oLvl.name = targetLvl Then
            ' This only changes the by-level color 
            ' Elements not using by-level will not be alterated
            oLvl.ElementColor = newColor 
        
            
            Do While oEE.MoveNext
                If oEE.Current.Color <> newColor Then
                    oEE.Current.Color = newColor
                    
                    oEE.Current.Rewrite

                End If
            
            Loop

        End If
    
    Next oLvl

    ActiveDesignFile.Levels.Rewrite

End Sub
