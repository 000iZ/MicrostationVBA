' README
' This script will change the color of all elements on a level you specify
' If used without Batch Process, changes will only affect the current file
' Refer to the sections below for instructions on how to use this script
' All text beginning with ' are commented out and ignored by the script
'
'
' LEVEL
' To specify the level, edit the word in quotation marks for targetLvl below
'   eg.  Const targetLvl As String = "os_base"  will choose the level os_base
'
' The level's name must be in quotation marks, but is not case-sensitive
' An error may occur if the specified level doesn't exist or isn't used in the 
'   file, and no changes will be made if that is the case
'
'
' COLOR
' To specify the color, edit the number code for newColor below
'   eg.  Const newColor As Integer = 2  will change the ByLevel color to 2, blue
'
' The ByLevel color of the level you specified will be updated, and all elements
'   on this level not using ByLevel color will be made to use ByLevel
'

Public Sub EditLevelColor()
    Dim oLvl As Level
    Dim oScan As ElementScanCriteria
    Dim oEE As ElementEnumerator
    

    'Edit the lines below to specify a level and the color you want to change to
    Const targetLvl As String = "c_Dimensions"  'Level name
    Const newColor As Integer = 2  'Microstation's color codes
    MessageCenter.AddMessage "Changing the color of Level '" & _
                                targetLvl & "' to " & newColor


    'Attempt to create a level object for target level in the current drawing
    On Error GoTo Err_InvalidLevelName  'Label statement if an error occurs
    Set oLvl = ActiveModelReference.Levels(targetLvl)
    
    
Err_InvalidLevelName:
    'Constructing an appropriate error message
    Dim errorMsg As String: errorMsg = "Error Code: " & Err.Number & _
                                        " (" & Err.HelpContext & ")" & _
                                        vbNewLine & Err.Description
    
    Select Case Err.Number  'Handle errors based on the error code
    'If no error occurrs, code of 0 is given
    Case 0:

    'Run-time error '-2147221504 (80040000)': Cannot modify library level attribute'
    Case -2147221504:
        MessageCenter.AddMessage "Level '" & targetLvl & _
                                    "' is a library level and " & _
                                    "cannot be modified.", _
                                    errorMsg, msdMessageCenterPriorityWarning
        Exit Sub

    'Run-time error '5': Invalid procedure call or argument
    'Run-time error '-2147024809 (80070057): Class not registered'
    Case 5 Or -2147024809:
        MessageCenter.AddMessage "Level '" & targetLvl & _
                                    "' cannot be found, or is unused.", _
                                    errorMsg, msdMessageCenterPriorityWarning
        Exit Sub
        
    'Other errors
    Case Else
        MessageCenter.AddMessage "An unknown error has occurred.", _
                                    errorMsg, msdMessageCenterPriorityWarning
        Exit Sub

    End Select


    'Create a search criteria object consisting of only the target level
    Set oScan = New ElementScanCriteria
    oScan.ExcludeAllLevels
    oScan.IncludeLevel oLvl
    

    'Create an element enumerator object that uses the search criteria
    Set oEE = ActiveModelReference.Scan(oScan)


    'Change the by-level color; elements not using by-level will be unaffected
    oLvl.ElementColor = newColor
    ActiveDesignFile.Levels.Rewrite


    'Change the color of elements not using by-level color
    Do While oEE.MoveNext
        If oEE.Current.Color <> newColor Then
            oEE.Current.Color = -1  '-1 here makes the element use by-level
            
            oEE.Current.Rewrite

        End If
    
    Loop

    ActiveDesignFile.Levels.Rewrite  'Update drawing


End Sub
