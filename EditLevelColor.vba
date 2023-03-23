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
    Dim boxTitle As String: boxTitle = "Error: "  'Title for error message box

    Select Case Err.Number  'Handle errors based on the error code
    'If no error occurrs, code of 0 is given
    Case 0:

    'Run-time error '-2147221504 (80040000)': Cannot modify library level attribute'
    Case -2147221504:
        MsgBox "Level '" & targetLvl & "' is a library level and cannot be modified.", vbExclamation, boxTitle & Err.Number
        Exit Sub

    'Run-time error '5': Invalid procedure call or argument
    'Run-time error '-2147024809 (80070057): Class not registered'
    Case 5 Or -2147024809:
        MsgBox "Level '" & targetLvl & "' cannot be found, or is unused.", vbExclamation, boxTitle & Err.Number
        Exit Sub
        
    'Other errors
    Case Else
        MsgBox "An unknown error has occurred.", vbExclamation, boxTitle & Err.Number
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

