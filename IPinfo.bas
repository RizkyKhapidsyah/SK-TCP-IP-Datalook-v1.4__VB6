Attribute VB_Name = "IP"
'Colors Server - Client
Public RTColorRemote As Double
Public RTColorLocaal As Double

Public Type IPX
    Name As String
    ConnectToIP As String
    ListenOnIP As String
    ConnectToPORT As Integer
    ListenOnPORT As Integer
End Type

Public IPinfo As IPX

Public DialogInfo As Byte
