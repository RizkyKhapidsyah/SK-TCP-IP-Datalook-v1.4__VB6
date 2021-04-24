Attribute VB_Name = "DC"
Public Versie As Byte

Function StringZoek(str1 As String) As Boolean
Dim result

If Versie = 0 Then
'Origineel = "$MyINFO"... 0.9977
result = InStr(1, str1, "$MyINFO")
ElseIf Versie = 1 Then
result = InStr(1, str1, "$Version 1,0005|$GetNickList|$MyINFO")
End If
If result <> 0 Then
StringZoek = True
'MsgBox "ok"
Else
StringZoek = False
End If
End Function

Function Verander_HD(str As String, MB As String) As String
'Replace string.
'strbegin = Mid(str, 1, X - 1)
MB = (Val(MB) * 1.08)
strVerander = MB & 999 & 999
'Zoeken naar 5é $ teken
If Versie = 0 Then
ZoekD = 3
Else
ZoekD = 33
End If
For teller = 1 To 5
dollar = InStr(ZoekD, str, "$", vbTextCompare)
'MsgBox "$ gevonden op:" & dollar & "teller: " & teller
ZoekD = dollar + 1
Next teller
'Laastste maal zoeken.. 6é $ teken
dollar = InStr(ZoekD, str, "$", vbTextCompare)
'Wissen data in text1 voor buffer..MBYTE
StrBegin = Mid(str, 1, ZoekD - 1)
StrEinde = Mid(str, dollar, Len(str))
str = StrBegin + strVerander + StrEinde
Verander_HD = str
End Function
