Sub Bmrsingle_text_edit()
    Dim startPoint As Point3d
    Dim point As Point3d, point2 As Point3d
    Dim lngTemp As Long
    Dim oMessage As CadInputMessage

'   Send a keyin that can be a command string
    CadInputQueue.SendKeyin "TEXTEDITOR MODIFY "

'   Coordinates are in master units
    startPoint.X = -186.570019664368
    startPoint.Y = -106.127431877135
    startPoint.Z = 0#

'   Send a data point to the current command
    point.X = startPoint.X
    point.Y = startPoint.Y
    point.Z = startPoint.Z
    CadInputQueue.SendDataPoint point, 1

'   XX-MacroComment_Header-XX

'   Start a command
    CadInputQueue.SendCommand "TEXTEDITOR PLAYCOMMAND CLEAR_ANCHOR_CARET"

    CadInputQueue.SendCommand "TEXTEDITOR PLAYCOMMAND SET_INSERT_CARET LINE 0 CHARACTER 7"

'   XX-MacroComment_KeyDownPerCharacter-XX

    CadInputQueue.SendCommand "TEXTEDITOR PLAYCOMMAND KEY_DOWN KEY_CODE 0x02 CONTROL_KEY_STATE UP SHIFT_KEY_STATE UP ALT_KEY_STATE UP"

'   XX-MacroComment_KeyDownPerCharacter-XX

    CadInputQueue.SendCommand "TEXTEDITOR PLAYCOMMAND KEY_DOWN KEY_CODE 0x02 CONTROL_KEY_STATE UP SHIFT_KEY_STATE UP ALT_KEY_STATE UP"

'   XX-MacroComment_KeyDownPerCharacter-XX

    CadInputQueue.SendCommand "TEXTEDITOR PLAYCOMMAND KEY_DOWN KEY_CODE 0x02 CONTROL_KEY_STATE UP SHIFT_KEY_STATE UP ALT_KEY_STATE UP"

'   XX-MacroComment_InsertText-XX

    CadInputQueue.SendKeyin "TEXTEDITOR PLAYCOMMAND INSERT_TEXT ""aaa"""

    point.X = startPoint.X - 4.43033321140325
    point.Y = startPoint.Y + 4.43479063392277
    point.Z = startPoint.Z
    CadInputQueue.SendDataPoint point, 1

    CommandState.StartDefaultCommand
End Sub