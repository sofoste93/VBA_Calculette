
Public Sub miniCalculette()

Dim firstNum As Double
Dim secondNum As Double
Dim addi, subt, mult, divi As Double

' enabling user input
firstNum = InputBox("Enter a number", "Calculator", "0123456789")
secondNum = InputBox("Enter a second number", "Calculator", "0123456789")

addi = addNum(firstNum, secondNum)
subt = subNum(firstNum, secondNum)
mult = mulNum(firstNum, secondNum)
divi = divNum(firstNum, secondNum)

' output the result in a Message Box
MsgBox ("Addition result: " & addi & vbCrLf & "Subtraction result: " & subt _
& vbCrLf & "Multiplication result: " & mult & vbCrLf & "Division result: " & divi), _
vbInformation, "All In One Results"


End Sub

