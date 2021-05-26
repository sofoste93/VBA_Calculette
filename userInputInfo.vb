Sub setAge(input)
'
' description.
'
' @since 1.0.0
' @param {type} [name] description.
' @return {type} [name] description.
' @see dependencies
'
Dim age As String
Dim third_box As Variant  ' since we don't know which data type is this \/ Yes or No

' Input. First Box
age = InputBox("Hello! How old are you? ", "Age Information", "33")

' Output with messageBox -
MsgBox age, vbInformation, "Your age"

' second box
Call MsgBox "Now with 'Call'", vbInformation, "Am the second Box!"

' third box
third_box = MsgBox "Third Box", vbYesNo, "Continue?"

Stop  ' here vbYesNo Yes=6, No=7 - - - use the debugger to find this out    

End Sub