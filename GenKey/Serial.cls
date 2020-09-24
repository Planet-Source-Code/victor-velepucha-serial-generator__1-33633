VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Serial"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private Const CHAR_ARRAY = "0123456789ABCDEF"
Private Const CHAR_MASK = 69

Public Function VerificarSerial(Code As String, User As String, Email As String, strSerial As String) As Collection
    Dim Resultado(1 To 6)
    Dim strKey As String
    Dim tmpKey(1 To 8) As Byte
    Dim crcTotal As Long
    Dim i As Integer
    Dim tmpStr As String
    Dim c As New Collection
    Dim crcCode As Long, crcUser As Long, crcEmail As Long, crcKey As Long
    
    On Error GoTo EH:
    
    'taking away - signs
    strKey = Left(strSerial, 4) & Mid(strSerial, 6, 4) & Mid(strSerial, 11, 4) & Right(strSerial, 4)
    
    'Get the byte array from the serial
    For i = 1 To 8
        tmpKey(i) = UnBin(BinPad(Bin(GetPos(Mid(strKey, 2 * i - 1, 1))), 4) & BinPad(Bin(GetPos(Mid(strKey, 2 * i, 1))), 4))
        tmpKey(i) = tmpKey(i) Xor CHAR_MASK
    Next
    
    'Let's see if user CRC is right...
    crcCode = UnBin(BinPad(Bin(tmpKey(1)), 8) & BinPad(Bin(tmpKey(2)), 8))
    crcUser = UnBin(BinPad(Bin(tmpKey(3)), 8) & BinPad(Bin(tmpKey(4)), 8))
    crcEmail = UnBin(BinPad(Bin(tmpKey(5)), 8) & BinPad(Bin(tmpKey(6)), 8))
    
    'If not, bye bye
    If (CRC16(Code) <> crcCode) Or (CRC16(User) <> crcUser) Or (CRC16(Email) <> crcEmail) Then
        c.Add False, "ValidKey"
        Set VerificarSerial = c
        Set c = Nothing
        Exit Function
    End If
    
    'Now Check the whole CRC
    crcTotal = UnBin(BinPad(Bin(tmpKey(7)), 8) & BinPad(Bin(tmpKey(8)), 8))
        
    For i = 1 To 6
        tmpStr = tmpStr & Chr(tmpKey(i))
    Next
                
    'If not, bye bye
    If CRC16(tmpStr) <> crcTotal Then
        c.Add False, "ValidKey"
        Set VerificarSerial = c
        Set c = Nothing
        Exit Function
    End If
    
    'Everything is OK, so return the result...
    
    
    c.Add True, "ValidKey"
    c.Add tmpKey(3), "ProgramCode"
    c.Add UnBin(Left(BinPad(Bin(tmpKey(4)), 8), 4)), "AppMajor"
    c.Add UnBin(Right(BinPad(Bin(tmpKey(4)), 8), 4)), "AppMinor"
    c.Add UnBin(Left(BinPad(Bin(tmpKey(5)), 8), 4)) & "/" & UnBin(Right(BinPad(Bin(tmpKey(5)), 8), 4)), "KeyDate"
    c.Add tmpKey(6), "Reserved"
    
    Resultado(1) = True
    Resultado(2) = tmpKey(3)
    Resultado(3) = UnBin(Left(BinPad(Bin(tmpKey(4)), 8), 4))
    Resultado(4) = UnBin(Right(BinPad(Bin(tmpKey(4)), 8), 4))
    Resultado(5) = UnBin(Left(BinPad(Bin(tmpKey(5)), 8), 4)) & "/" & UnBin(Right(BinPad(Bin(tmpKey(5)), 8), 4))
    Resultado(6) = tmpKey(6)
    
    Set VerificarSerial = c
    
    Set c = Nothing
    Exit Function
EH:
    c.Add False, "ValidKey"
    Set VerificarSerial = c
    Set c = Nothing
End Function

Public Function VerificarLicencia(Code As String, strSerial As String) As Collection
    Dim strKey As String
    Dim tmpKey(1 To 8) As Byte
    Dim crcTotal As Long
    Dim i As Integer
    Dim tmpStr As String
    Dim c As New Collection
    Dim crcCode As Long, crcKey As Long
    
    On Error GoTo EH:
    
    'taking away - signs
    strKey = Left(strSerial, 4) & Mid(strSerial, 6, 4) & Mid(strSerial, 11, 4) & Right(strSerial, 4)
    
    'Get the byte array from the serial
    For i = 1 To 8
        tmpKey(i) = UnBin(BinPad(Bin(GetPos(Mid(strKey, 2 * i - 1, 1))), 4) & BinPad(Bin(GetPos(Mid(strKey, 2 * i, 1))), 4))
        tmpKey(i) = tmpKey(i) Xor CHAR_MASK
    Next
    
    'Let's see if user CRC is right...
    crcCode = UnBin(BinPad(Bin(tmpKey(1)), 8) & BinPad(Bin(tmpKey(2)), 8))
    
    'If not, bye bye
    If (CRC16(Code) <> crcCode) Then
        c.Add False, "ValidKey"
        Set VerificarLicencia = c
        Set c = Nothing
        Exit Function
    End If
    
    'Now Check the whole CRC
    crcTotal = UnBin(BinPad(Bin(tmpKey(7)), 8) & BinPad(Bin(tmpKey(8)), 8))
        
    For i = 1 To 6
        tmpStr = tmpStr & Chr(tmpKey(i))
    Next
                
    'If not, bye bye
    If CRC16(tmpStr) <> crcTotal Then
        c.Add False, "ValidKey"
        Set VerificarLicencia = c
        Set c = Nothing
        Exit Function
    End If
    
    'Everything is OK, so return the result...
    
    
    c.Add True, "ValidKey"
    c.Add tmpKey(1), "OK"
    
    Set VerificarLicencia = c
    
    Set c = Nothing
    Exit Function
EH:
    c.Add False, "ValidKey"
    Set VerificarLicencia = c
    Set c = Nothing
End Function











Private Function CRC16(B As String) As Long
    Dim Power(0 To 7) As Byte
    Dim i As Integer, j As Integer
    Dim ByteVal As Byte
    Dim TestBit As Boolean
    Dim CRC As Long

    For i = 0 To 7
        Power(i) = 2 ^ i
    Next i
    CRC = 0
    For i = 1 To Len(B)
        ByteVal = Asc(Mid$(B, i, 1))
        For j = 7 To 0 Step -1
            TestBit = ((CRC And 32768) = 32768) Xor ((ByteVal And Power(j)) = Power(j))
            CRC = ((CRC And 32767&) * 2&)
            If TestBit Then CRC = CRC Xor &H1021
        Next j
    Next i
    CRC16 = CRC
End Function




Private Function GetChar(Char As Byte) As String
    GetChar = Mid(CHAR_ARRAY, Char Mod Len(CHAR_ARRAY) + 1, 1)
End Function

Private Function GetPos(Char As String) As Byte
    GetPos = InStr(CHAR_ARRAY, Char) - 1
End Function









'Functions for handling BinaryStr
'Taken brom VBExt.dll by NetGopher, posted on PSC
'Some modifications made

Private Function BinPad(strBinary As String, FieldLen As Integer) As String
    BinPad = String(FieldLen - Len(strBinary), "0") & strBinary
End Function

Private Function Reradix(Expression, FromBase, ToBase) As Variant


    Dim Value As Long
    Dim InDigit(), OutDigit()
    Dim InExpr() As Byte, OutExpr() As Byte
    Dim Power As Integer, Multiple As Long, RetVal As String
    
    FromBase = CInt(FromBase)
    ToBase = CInt(ToBase)
    If FromBase < 2 Or FromBase > 32 Or ToBase < 2 Or ToBase > 32 Then Err.Raise 5
    
    If Val(Expression) = 0 Then
        If ToBase = 10 Then
            Reradix = 0
        Else
            Reradix = "0"
        End If
        Exit Function
    End If
    
    
    InDigit() = Array(0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, _
                             0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, _
                             0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, _
                             0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 0, 0, 0, 0, 0, 0, _
                             0, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20, 21, 22, 23, 24, _
                             25, 26, 27, 28, 29, 30, 31, 0, 0, 0, 0, 0, 0, 0, 0, 0)
    
    OutDigit() = Array(48, 49, 50, 51, 52, 53, 54, 55, 56, 57, 65, 66, 67, 68, 69, 70, _
                               71, 72, 73, 74, 75, 76, 77, 78, 79, 80, 81, 82, 83, 84, 85, 86)
    
    InExpr() = StrConv(StrReverse(UCase(CStr(Expression))), vbFromUnicode)

    Value = 0
    For Power = LBound(InExpr) To UBound(InExpr)
        Value = Value + FromBase ^ Power * InDigit(InExpr(Power))
    Next

    Power = 0
    While Value >= ToBase ^ Power
        Power = Power + 1
    Wend
    Power = Power - 1
    
    ReDim OutExpr(0 To Power) As Byte
    While Power >= 0
        Multiple = ToBase ^ Power
        While Value >= Multiple
            OutExpr(Power) = OutExpr(Power) + 1
            Value = Value - Multiple
        Wend
        OutExpr(Power) = OutDigit(OutExpr(Power))
        Power = Power - 1
    Wend
    
    RetVal = StrReverse(StrConv(OutExpr(), vbUnicode))
    If ToBase = 10 Then
        Reradix = CLng(RetVal)
    Else
        Reradix = RetVal
    End If

End Function

' shorthand for converting number to binary
' (uses Radix function)
'
Private Function Bin(Number) As String

    Bin = Reradix(Number, 10, 2)
End Function

' converts binary string into decimal number
'
Private Function UnBin(ByVal BinaryString As String) As Long

    UnBin = Reradix(BinaryString, 2, 10)
End Function

