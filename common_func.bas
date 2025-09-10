Attribute VB_Name = "common_func"
Option Explicit

Public Const MAXINT As Integer = (2 ^ 15) - 1

'' シートの存在確認
Public Function CheckSheetExist(ByVal SheetName As String)
    Dim ws As Variant
    
    For Each ws In Worksheets
        If LCase(ws.Name) = LCase(SheetName) Then
            CheckSheetExist = True
            Exit Function
        End If
    Next

    CheckSheetExist = False
End Function

'' 32 ビット実数（単精度）変換
Public Function HexToSingle(sHex As String) As Single
  Dim sTemp As String
  Dim iSign, iExponent As Integer
  Dim fTemp, fFraction As Single
  
  ' 符号 1ビット
  sTemp = Mid(sHex, 1, 1)
  fTemp = Val("&H" & sTemp) And &H8
  iSign = IIf(fTemp = 8, -1, 1)
  
  ' 指数部 8ビット
  sTemp = Mid(sHex, 1, 3)
  fTemp = Val("&H" & sTemp) And &H7F8
  iExponent = fTemp / 2 ^ 3 - 127  '32ビット形式のバイアス=127
  
  ' 仮数部 23ビット
  sTemp = Mid(sHex, 3, 6)
  fTemp = Val("&H" & sTemp) And &H7FFFFF
  fFraction = 1# + (fTemp / 2 ^ 23)
  
  HexToSingle = iSign * fFraction * 2 ^ iExponent
  
End Function
 
'' 64 ビット実数（倍精度）変換
Public Function HexToDouble(sHex As String) As Double
  Dim sTemp As String
  Dim iSign, iExponent As Integer
  Dim fTemp, fFraction As Double
  
  ' 符号 1ビット
  sTemp = Mid(sHex, 1, 1)
  fTemp = Val("&H" & sTemp) And &H8
  iSign = IIf(fTemp = 8, -1, 1)
    
  ' 指数部 11ビット
  sTemp = Mid(sHex, 1, 3)
  fTemp = Val("&H" & sTemp) And &H7FF
  iExponent = fTemp - 1023  '64ビット形式のバイアス=1023
  
  ' 仮数部 52ビット
  sTemp = Mid(sHex, 4, 13)
  fTemp = CDbl("&H" & sTemp)
  fFraction = 1 + (fTemp / 2 ^ 52)
    
  HexToDouble = iSign * fFraction * 2 ^ iExponent
  
End Function






