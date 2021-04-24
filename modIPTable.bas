Attribute VB_Name = "modIPTable"

Const MAX_IP = 5   'To make a buffer... i dont think you have more than 5 ip on your pc..

Public IP_count As Byte
Public MyIP(MAX_IP) As String

Type IPINFO
     dwAddr As Long   ' IP address
    dwIndex As Long '  interface index
    dwMask As Long ' subnet mask
    dwBCastAddr As Long ' broadcast address
    dwReasmSize  As Long ' assembly size
    unused1 As Integer ' not currently used
    unused2 As Integer '; not currently used
End Type

Type MIB_IPADDRTABLE
    dEntrys As Long   'number of entries in the table
    mIPInfo(MAX_IP) As IPINFO  'array of IP address entries
End Type

Type IP_Array
    mBuffer As MIB_IPADDRTABLE
    BufferLen As Long
End Type

Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Public Declare Function GetIpAddrTable Lib "IPHlpApi" (pIPAdrTable As Byte, pdwSize As Long, ByVal Sort As Long) As Long


'converts a Long  to a string
Public Function ConvertAddressToString(longAddr As Long) As String
    Dim myByte(3) As Byte
    Dim cnt As Long
    CopyMemory myByte(0), longAddr, 4
    For cnt = 0 To 3
        ConvertAddressToString = ConvertAddressToString + CStr(myByte(cnt)) + "."
    Next cnt
    ConvertAddressToString = Left$(ConvertAddressToString, Len(ConvertAddressToString) - 1)
End Function

Public Sub Get_My_IP_Adresses()
Dim Ret As Long
Dim bBytes() As Byte
Dim Listing As MIB_IPADDRTABLE


 GetIpAddrTable ByVal 0&, Ret, True

    If Ret <= 0 Then Exit Sub
    ReDim bBytes(0 To Ret - 1) As Byte
    'retrieve the data
    GetIpAddrTable bBytes(0), Ret, False
      
    'Get the first 4 bytes to get the entry's.. ip installed
    CopyMemory Listing.dEntrys, bBytes(0), 4
    'MsgBox "IP's found : " & Listing.dEntrys    => Founded ip installed on your PC..
   IP_count = Listing.dEntrys
    For tel = 0 To Listing.dEntrys - 1
        'Copy whole structure to Listing..
       ' MsgBox bBytes(tel) & "."
        CopyMemory Listing.mIPInfo(tel), bBytes(4 + (tel * Len(Listing.mIPInfo(0)))), Len(Listing.mIPInfo(tel))
        MyIP(tel) = ConvertAddressToString(Listing.mIPInfo(tel).dwAddr)
    Next

'MsgBox ConvertAddressToString(Listing.mIPInfo(1).dwAddr)
End Sub
