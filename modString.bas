Attribute VB_Name = "modString"
'String to Hex String Conversion.
Function StringToHex(str1 As String) As String
Dim lengtestring
Dim HexString As String
Dim Hvalue As String
Dim Cnt As Integer
lengtestring = Len(str1)
For Cnt = 1 To lengtestring
Hvalue = Asc(Mid(str1, Cnt, 1))
'Kleiner dan 16 => plaats 0 voor :=> 0F ...
If Hvalue < 16 Then
Hvalue = "0" & Hex(Hvalue)
Else
Hvalue = Hex(Hvalue)
End If
HexString = HexString & Hvalue & "."
Next Cnt
StringToHex = HexString
End Function
