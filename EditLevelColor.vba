Public Sub EditLevelColor()
    Dim oLvl As Level
    Dim oScan As ElementScanCriteria
    Dim oEE As ElementEnumerator
    

    Const targetLvl As String = "Defpoints"  ' Level name
    Const newColor As Integer = 6  ' Microstation's color codes
    'MsgBox "Changing the color of Level '" & targetLvl & "' to " & newColor


    ' Create a level object for a level in the current drawing
    If targetIsValidLevelName(targetLvl) Then
        Set oLvl = ActiveModelReference.Levels(targetLvl)
    Else
        MsgBox "Invalid level name"
        
        Exit Sub
    End If


    ' Create a search criteria object consisting of only the target level
    Set oScan = New ElementScanCriteria
    oScan.ExcludeAllLevels
    oScan.IncludeLevel oLvl
    

    ' Create an element enumerator object that uses the search criteria
    Set oEE = ActiveModelReference.Scan(oScan)


    ' Changes the by-level color; elements not using by-level will be unaffected
    oLvl.ElementColor = newColor
    ActiveDesignFile.Levels.Rewrite


    ' Changes the color of elements not using by-level color
    Do While oEE.MoveNext
        If oEE.Current.Color <> newColor Then
            oEE.Current.Color = newColor
            
            oEE.Current.Rewrite

        End If
    
    Loop

    ActiveDesignFile.Levels.Rewrite  ' Update drawing


End Sub


Public Function targetIsValidLevelName(ByVal targetLvl As String) As Boolean
    targetIsValidLevelName = False

    On Error GoTo err_IsValidLevelName
    Dim oLvl As Level
    Set oLvl = ActiveDesignFile.Levels(targetLvl)

    If oLvl Is Nothing Then
        targetIsValidLevelName = False

    Else
        targetIsValidLevelName = True

    End If
    
    Exit Function

err_IsValidLevelName:
    Select Case Err.Number

    Case 5:
        '   Level not found
        Resume Next
        
    Case Else
        MsgBox "targetIsValidLevelName failed"

    End Select

End Function