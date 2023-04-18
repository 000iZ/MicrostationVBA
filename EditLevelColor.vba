'README-----------------------------------------------------------------------
'
' This script will change the color of all elements on a level you specify
' If used without Batch Process, changes will only affect the current file
' Refer to the sections below for instructions on how to use this script
' All text beginning with ' are commented out and ignored by the script
' Messages regarding errors, if any, are shown at the bottom of Microstation,
'   near the View Toggles and the coordinates for X Axis, Y Axis
'
'
' LEVEL
' To specify the level, type a word in quotation marks for targetLvl below
'   eg.  Const targetLvl As String = "os_base"  will choose the level os_base
'
' The level's name must be in quotation marks, but is not case-sensitive
' An error may occur if the specified level cannot be found, and no changes
'   will be made if that is the case
'
'
' COLOR
' To specify the color, type a number code for newColor below
'   eg.  Const newColor As Integer = 2  will change the ByLevel color to 2, blue
'
' The ByLevel color of the level you specified will be updated, and all elements
'   on this level not using ByLevel color will be made to use ByLevel
'
'
' LINE WEIGHT
' You may also edit the LINE WEIGHT of the specified level through this script
'   eg.  Const editLineWeight As Boolean = True  allows you to edit the weight
'   eg.  Const newLineWeight As Integer = 5  will change the ByLevel weight to 5
'
' If you do not wish to edit the line weight, type False for the boolean


Public Sub EditLevelColor()
'SETUP------------------------------------------------------------------------
    Dim oLvl As Level
    Dim oScan As ElementScanCriteria
    Dim oEE As ElementEnumerator

    'Edit the lines below to specify a LEVEL and the COLOR you want to change to
    Const targetLvl As String = "os_Base"  'Level name, not case-sensitive
    Const newColor As Integer = 3  'Specify a Microstation color code

    'Edit the lines below if you want to change the LINE WEIGHT
    Const editLineWeight As Boolean = False  'Type True or False
    Const newLineWeight As Integer = 5  'Specify a Microstation line weight code
    
    
    'Allows library levels to be edited by modifying a Configuration Variable
    'If edited, library levels are copied into the master-file
    ActiveWorkspace.AddConfigurationVariable _
                                        "MS_LEVEL_ALLOW_LIBRARY_LEVEL_EDIT", "1"

    
    'Attempt to create a level object for target level in the current drawing
    On Error GoTo Err_InvalidLevelName  'Label statement if an error occurs
    Set oLvl = ActiveModelReference.Levels(targetLvl)


'ERROR HANDLING---------------------------------------------------------------
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


'SEARCH FOR ELEMENTS----------------------------------------------------------
    'Create a search criteria object consisting of only the target level
    Set oScan = New ElementScanCriteria
    oScan.ExcludeAllLevels
    oScan.IncludeLevel oLvl
    

    'Create an element enumerator object that uses the search criteria
    Set oEE = ActiveModelReference.Scan(oScan)

'MAKE CHANGES-----------------------------------------------------------------
    'Change the ByLevel color
    oLvl.ElementColor = newColor
    
    
    'Change the ByLevel line weight, if the choice was made to edit weights
    If editLineWeight = True Then
        oLvl.ElementLineWeight = newLineWeight
        
    End If
    
    ActiveDesignFile.Levels.Rewrite  'Update drawing


    'Change the color of elements not using ByLevel color
    Do While oEE.MoveNext
        If oEE.Current.Color <> newColor Then
            oEE.Current.Color = -1  '-1 makes the element use ByLevel
            
            oEE.Current.Rewrite  'Update elements

        End If
        
        If editLineWeight = True Then
            oEE.Current.LineWeight = -1  '-1 makes the element use ByLevel
            
            oEE.Current.Rewrite  'Update elements
            
        End If
    
    Loop

    ActiveDesignFile.Levels.Rewrite  'Update drawing


End Sub
