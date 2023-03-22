Public Sub EditLevelColor()
    Dim oLvl As Level
    Dim oScan As ElementScanCriteria
    Dim oEE As ElementEnumerator
    

    Const targetLvl As String = "c_Dimensions"  ' Level name
    Const newColor As Integer = 2  ' Microstation's color codes
    'MsgBox "Changing the color of Level '" & targetLvl & "' to " & newColor


    'Attempt to create a level object for target level in the current drawing
    On Error GoTo Err_InvalidLevelName  'Label statement if an error occurs
    Set oLvl = ActiveModelReference.Levels(targetLvl)
    
    
Err_InvalidLevelName:
    'MsgBox "Error code: " & Err.Number

    Select Case Err.Number  'Handle errors based on the error code
    Case 5:  'Invalid procedure call or argument
        MsgBox "Cannot find Level '" & targetLvl & "'. No changes will be made."
        Exit Sub
        
    Case Else  'Other errors, the target level may be a library reference level
        MsgBox "An error has occurred. '" & targetLvl & "' may be unused in this file"
        Exit Sub

    End Select


    ' Create a search criteria object consisting of only the target level
    Set oScan = New ElementScanCriteria
    oScan.ExcludeAllLevels
    oScan.IncludeLevel oLvl
    

    ' Create an element enumerator object that uses the search criteria
    Set oEE = ActiveModelReference.Scan(oScan)


    ' Change the by-level color; elements not using by-level will be unaffected
    oLvl.ElementColor = newColor
    ActiveDesignFile.Levels.Rewrite


    ' Change the color of elements not using by-level color
    Do While oEE.MoveNext
        If oEE.Current.Color <> newColor Then
            oEE.Current.Color = newColor
            
            oEE.Current.Rewrite

        End If
    
    Loop

    ActiveDesignFile.Levels.Rewrite  ' Update drawing


End Sub

