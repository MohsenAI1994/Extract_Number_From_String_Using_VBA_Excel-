Function Extract_Number(ByVal inputString As String):

Dim I As Integer
Dim INPUT_LEN As Integer
Dim RESULT As String

' Initialize result as an empty string
RESULT = ""
    
INPUT_LEN = Len(inputString)

' Loop through each character in the input string
For I = 1 To INPUT_LEN
    ' Check if the character is a digit or a decimal point
    If IsNumeric(Mid(inputString, I, 1)) Or currentChar = "." Then
        RESULT = RESULT & Mid(inputString, I, 1)
    End If
Next I

' If result is not empty, convert it to a number
    If RESULT <> "" Then
        Extract_Number = CDbl(RESULT)
    Else
        ' If no number is found, return 0
        Extract_Number = 0
    End If

End Function
