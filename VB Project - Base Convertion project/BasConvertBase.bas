Attribute VB_Name = "BasConvertBase"
Option Explicit

Public Enum Operations
    Multiply = 0
    Divide = 1
    Add = 2
    Subtract = 3
End Enum

Private Function DecimalToAll(InputDecimal As Long, OutPutBase As Integer) As String
    Dim IntRemainder As Long, StrRemainder As String, StrTempNumber As String

    Do While InputDecimal >= 1
        IntRemainder = InputDecimal Mod OutPutBase
        If OutPutBase = 16 Then
                If IntRemainder >= 10 And IntRemainder <= 15 Then
                    StrRemainder = Chr(IntRemainder + 55)
                Else
                    StrRemainder = CStr(IntRemainder)
                End If
        Else
                    StrRemainder = CStr(IntRemainder)
        End If

        InputDecimal = InputDecimal \ OutPutBase
        StrTempNumber = StrRemainder & StrTempNumber
    Loop
    DecimalToAll = StrTempNumber
End Function

Private Function AllToDecimal(StrInput As String, InputBase As Integer) As Long
        Dim LastNum As String, InputLen As Integer, Pow As Long, DecNum As Long, LastBit As Long
        
        InputLen = Len(StrInput)
    Do
        LastNum = Mid(StrInput, InputLen, 1)
        InputLen = InputLen - 1
        
        If InputBase = 16 Then
            If Asc(LastNum) >= 65 And Asc(LastNum) <= 70 Then LastNum = CStr(Asc(LastNum) - 55)
        End If
        
        LastBit = CLng(LastNum) * InputBase ^ Pow
        
        DecNum = LastBit + DecNum
        Pow = Pow + 1
    Loop Until InputLen = 0
    AllToDecimal = DecNum
End Function


Public Function ConvertAll(StrInput As String, InputBase As Integer, OutPutBase As Integer)
'OutPut is Decimal
    'binarytodecimal==>2,10
    'OctalToDecimal==>8,10
    'HexadecimalToDecimal==>16,10

'Input Is Decimal
    'DecimalToBinary==>10,2
    'DecimalToOctal==>10,8
    'DecimalToHexaDecimal==>10,16
    
'Others
    'BinaryToOctal==>2,8
    'BinaryToHexadecimal==>2,16
    
    'OctalToBinary==>8,2
    'OctalToHexadecimal==>8,16
    
    'HexadecimalToBinary==>16,2
    'HexadecimalToOctal==>16,8
    
    If InputBase = 10 And OutPutBase = 10 Then ConvertAll = StrInput: Exit Function
    
    If InputBase = 10 Then
        ConvertAll = DecimalToAll(CLng(StrInput), OutPutBase)
        Exit Function
    End If
    
    If OutPutBase = 10 Then
        ConvertAll = AllToDecimal(StrInput, InputBase)
        Exit Function
    End If
    
    ConvertAll = DecimalToAll(AllToDecimal(StrInput, InputBase), OutPutBase)
End Function

Public Function FloatDecimalToBinary(FloatDecimal As Double, Optional WordSize As Integer = 100) As String
    Dim IntegerPart As Integer, FloatingPart As Double
    Dim Tstr As String
    Dim TwordSize As Integer
    
    IntegerPart = CInt(FloatDecimal)
    FloatingPart = FloatDecimal - IntegerPart
    
    Tstr = DecimalToAll(CLng(IntegerPart), 2) & "."
    
    Do Until FloatingPart = Int(FloatingPart)
        FloatingPart = FloatingPart * 2
        Tstr = Tstr & Int(FloatingPart)
        TwordSize = TwordSize + 1
        If TwordSize >= WordSize Then Exit Do
    Loop
    FloatDecimalToBinary = Tstr
End Function

Public Function Operations(Number1 As String, Number2 As String, InputBase As Integer, Operator As Operations) As String
    Dim FirstNumber As Long, SecondNumber As Long
    FirstNumber = ConvertAll(Number1, InputBase, 10)
    SecondNumber = ConvertAll(Number2, InputBase, 10)

    Dim Temp As Long
    Select Case Operator
        Case 0:
            Temp = FirstNumber * SecondNumber
        Case 1:
            Temp = FirstNumber / SecondNumber
        Case 2:
            Temp = FirstNumber + SecondNumber
        Case 3:
            Temp = FirstNumber - SecondNumber
    End Select
    
    Operations = ConvertAll(CStr(Temp), 10, InputBase)
End Function
