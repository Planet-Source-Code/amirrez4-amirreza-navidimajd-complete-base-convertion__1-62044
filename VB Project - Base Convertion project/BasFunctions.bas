Attribute VB_Name = "BasFunctions"
Public Function IsValidChar(InputBase As Integer, Char As String) As Boolean
    Dim CharSet As String
    Select Case InputBase
            Case 2: CharSet = "01"
            Case 8: CharSet = "01234567"
            Case 10: CharSet = "0123456789"
            Case 16: CharSet = "0123456789ABCDEF"
    End Select
    IsValidChar = InStr(1, CharSet, Char, vbTextCompare)
End Function


Public Function GetBase(I As Integer) As Integer
Select Case I
        Case 0: GetBase = 2
        Case 1: GetBase = 8
        Case 2: GetBase = 10
        Case 3: GetBase = 16
End Select
End Function
