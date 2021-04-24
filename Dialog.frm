VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form Dialog 
   BackColor       =   &H80000004&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Connecting setup"
   ClientHeight    =   1230
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   4620
   ForeColor       =   &H00C00000&
   Icon            =   "Dialog.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1230
   ScaleWidth      =   4620
   ShowInTaskbar   =   0   'False
   Begin VB.ComboBox cmbIPs 
      BackColor       =   &H80000018&
      Height          =   315
      Left            =   120
      TabIndex        =   9
      Text            =   "Choose your IP"
      Top             =   840
      Width           =   2415
   End
   Begin MSWinsockLib.Winsock MyIP 
      Left            =   4140
      Top             =   840
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.TextBox txtListenOnPort 
      BackColor       =   &H80000018&
      Height          =   285
      Left            =   2640
      TabIndex        =   6
      Text            =   "0"
      Top             =   840
      Width           =   615
   End
   Begin VB.TextBox txtConToPort 
      Height          =   285
      Left            =   2640
      TabIndex        =   5
      Text            =   "0"
      Top             =   240
      Width           =   615
   End
   Begin VB.TextBox txtConnectTO_IP 
      Height          =   285
      Left            =   120
      TabIndex        =   4
      Text            =   "Connecting to IP"
      Top             =   240
      Width           =   2415
   End
   Begin VB.CommandButton CancelButton 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   3360
      TabIndex        =   1
      Top             =   600
      Width           =   1215
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "OK"
      Height          =   375
      Left            =   3360
      TabIndex        =   0
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label Label4 
      Caption         =   "Port"
      Height          =   255
      Left            =   2640
      TabIndex        =   8
      Top             =   600
      Width           =   495
   End
   Begin VB.Label Label3 
      Caption         =   "Port"
      Height          =   255
      Left            =   2640
      TabIndex        =   7
      Top             =   0
      Width           =   615
   End
   Begin VB.Label Label2 
      Caption         =   "Connecting to IP:"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   0
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "Listening on IP: your IP (read)"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   600
      Width           =   2295
   End
End
Attribute VB_Name = "Dialog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Dim ChangeTXT As Boolean

Private Sub CancelButton_Click()

IP.DialogInfo = 2
Unload Me
End Sub

Private Sub Form_Load()
Dim cnt As Byte
ChangeTXT = False
IP.DialogInfo = 0
modIPTable.Get_My_IP_Adresses

For cnt = 0 To modIPTable.IP_count - 1
    cmbIPs.AddItem modIPTable.MyIP(cnt)
Next
End Sub



Private Sub Form_Terminate()
If IP.DialogInfo = 0 Then IP.DialogInfo = 2
Unload Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
If IP.DialogInfo = 0 Then IP.DialogInfo = 2
End Sub

Private Sub OKButton_Click()
'Check if the values are made..
'Check for IP con TO ip & Name caption
On Error GoTo SocketError
If txtConnectTO_IP = "Connecting to IP" _
Or txtListenOnPort = 0 Then
MsgBox "please enter valid IP Port number " & vbCrLf & _
"                         Or" _
& vbCrLf & "enter a new connect IP address !"
Else
If cmbIPs.ListIndex < 0 Then GoTo ErrorIP
'Set the input data's
IP.IPINFO.ConnectToIP = txtConnectTO_IP
IP.IPINFO.ConnectToPORT = txtConToPort
IP.IPINFO.ListenOnIP = cmbIPs.List(cmbIPs.ListIndex)
IP.IPINFO.Name = "Connect to " & txtConnectTO_IP
IP.IPINFO.ListenOnPORT = txtListenOnPort
IP.DialogInfo = 1
Unload Me
End If
Exit Sub
SocketError:
    MsgBox "Error , port number to high ! ", vbCritical
Exit Sub
ErrorIP:
    MsgBox "Error, Choose Local IP address !!", vbCritical
End Sub



Private Sub txtListeningOn_Change()

End Sub

