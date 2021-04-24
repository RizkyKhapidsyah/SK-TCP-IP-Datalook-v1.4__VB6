VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MsComCtl.ocx"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "Mswinsck.ocx"
Begin VB.Form frmDocument 
   BackColor       =   &H80000016&
   Caption         =   "frmDocument"
   ClientHeight    =   6405
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7710
   Icon            =   "frmDocument.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6405
   ScaleWidth      =   7710
   Begin VB.Frame fraClientDC 
      Height          =   2295
      Left            =   2880
      TabIndex        =   31
      Top             =   3000
      Width           =   2655
      Begin VB.CommandButton cmdClientApply 
         Caption         =   "ok"
         Height          =   255
         Left            =   120
         TabIndex        =   34
         Top             =   1920
         Width           =   2415
      End
      Begin VB.Frame fraCDC 
         Caption         =   "Client Data Change "
         Height          =   1335
         Left            =   120
         TabIndex        =   33
         Top             =   480
         Width           =   2415
         Begin VB.TextBox txtCDC_Change 
            Height          =   285
            Left            =   120
            TabIndex        =   38
            Top             =   990
            Width           =   2175
         End
         Begin VB.TextBox txtCDC_Source 
            Height          =   285
            Left            =   120
            TabIndex        =   36
            Top             =   480
            Width           =   2175
         End
         Begin VB.Label Label8 
            Caption         =   "Change to:"
            Height          =   255
            Left            =   120
            TabIndex        =   37
            Top             =   800
            Width           =   1335
         End
         Begin VB.Label Label7 
            Caption         =   "Source data:"
            Height          =   255
            Left            =   120
            TabIndex        =   35
            Top             =   240
            Width           =   2055
         End
      End
      Begin VB.CheckBox chkCDC 
         Caption         =   "Enable Change"
         Height          =   255
         Left            =   120
         TabIndex        =   32
         Top             =   240
         Width           =   2295
      End
   End
   Begin VB.Frame fraServerDC 
      Height          =   2295
      Left            =   0
      TabIndex        =   23
      Top             =   3000
      Width           =   2775
      Begin VB.Frame fraSDC 
         Caption         =   "Server Data Change "
         Height          =   1335
         Left            =   120
         TabIndex        =   26
         Top             =   480
         Width           =   2535
         Begin VB.TextBox txtSDC_Change 
            Height          =   285
            Left            =   120
            TabIndex        =   28
            Top             =   960
            Width           =   2295
         End
         Begin VB.TextBox txtSDC_Source 
            Height          =   285
            Left            =   120
            TabIndex        =   27
            Top             =   450
            Width           =   2295
         End
         Begin VB.Label Label6 
            Caption         =   "Source data:"
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   30
            Top             =   195
            Width           =   1335
         End
         Begin VB.Label Label5 
            Caption         =   "Change to:"
            Height          =   255
            Left            =   120
            TabIndex        =   29
            Top             =   750
            Width           =   1095
         End
      End
      Begin VB.CommandButton cmdServerAppl 
         Caption         =   "ok"
         Height          =   255
         Left            =   120
         TabIndex        =   25
         Top             =   1920
         Width           =   2535
      End
      Begin VB.CheckBox chkSDC 
         Caption         =   "Enable Change"
         Height          =   255
         Left            =   120
         TabIndex        =   24
         Top             =   240
         Width           =   2295
      End
   End
   Begin VB.Frame fraExtra 
      Appearance      =   0  'Flat
      BackColor       =   &H80000000&
      ForeColor       =   &H80000008&
      Height          =   1455
      Left            =   5520
      TabIndex        =   12
      Top             =   1200
      Width           =   2175
      Begin VB.PictureBox Picture5 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   150
         Picture         =   "frmDocument.frx":0CCA
         ScaleHeight     =   195
         ScaleWidth      =   195
         TabIndex        =   22
         Top             =   130
         Width           =   195
      End
      Begin VB.PictureBox Picture4 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   150
         Picture         =   "frmDocument.frx":11DC
         ScaleHeight     =   195
         ScaleWidth      =   195
         TabIndex        =   21
         Top             =   370
         Width           =   195
      End
      Begin VB.PictureBox Picture3 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   150
         Picture         =   "frmDocument.frx":16EE
         ScaleHeight     =   195
         ScaleWidth      =   195
         TabIndex        =   20
         Top             =   610
         Width           =   195
      End
      Begin VB.PictureBox Picture2 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   150
         Picture         =   "frmDocument.frx":1C00
         ScaleHeight     =   195
         ScaleWidth      =   195
         TabIndex        =   16
         Top             =   850
         Width           =   195
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   0
         Left            =   150
         Picture         =   "frmDocument.frx":2112
         ScaleHeight     =   195
         ScaleWidth      =   195
         TabIndex        =   15
         Top             =   1090
         Width           =   195
      End
      Begin VB.Label lblSDC 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Disconnect"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   19
         Top             =   600
         Width           =   1935
      End
      Begin VB.Label lblCDC 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Client data Change"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   18
         Top             =   360
         Width           =   1935
      End
      Begin VB.Label lblCls 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Clearscreen"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   17
         Top             =   120
         Width           =   1935
      End
      Begin VB.Label lblHelp 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Help"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Top             =   840
         Width           =   1935
      End
      Begin VB.Label lblExit 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Exit Submenu"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   1080
         Width           =   1935
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Extra"
      Height          =   495
      Left            =   0
      TabIndex        =   11
      Top             =   0
      Width           =   615
   End
   Begin VB.Timer tmrSpeed 
      Interval        =   500
      Left            =   4560
      Top             =   600
   End
   Begin VB.CheckBox Check1 
      Caption         =   "View Hex"
      Height          =   255
      Left            =   1920
      TabIndex        =   9
      Top             =   720
      Width           =   1335
   End
   Begin RichTextLib.RichTextBox RTBox 
      Height          =   1815
      Left            =   0
      TabIndex        =   6
      Top             =   1080
      Width           =   5415
      _ExtentX        =   9551
      _ExtentY        =   3201
      _Version        =   393217
      BackColor       =   14737632
      ReadOnly        =   -1  'True
      ScrollBars      =   2
      TextRTF         =   $"frmDocument.frx":2624
   End
   Begin MSWinsockLib.Winsock LOCAAL 
      Left            =   5040
      Top             =   600
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock REMOTE 
      Left            =   5040
      Top             =   120
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSComctlLib.StatusBar Status 
      Align           =   2  'Align Bottom
      Height          =   315
      Left            =   0
      TabIndex        =   5
      Top             =   6090
      Width           =   7710
      _ExtentX        =   13600
      _ExtentY        =   556
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Text            =   "Status"
            TextSave        =   "Status"
            Object.ToolTipText     =   "Status Connection"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Text            =   "UP:"
            TextSave        =   "UP:"
            Object.ToolTipText     =   "Upload speed at this connection"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Text            =   "Down:"
            TextSave        =   "Down:"
            Object.ToolTipText     =   "Download Speed at this connection"
         EndProperty
      EndProperty
   End
   Begin VB.TextBox txtLIS_ONPORT 
      BackColor       =   &H00808080&
      Enabled         =   0   'False
      Height          =   285
      Left            =   3600
      TabIndex        =   4
      Top             =   240
      Width           =   735
   End
   Begin VB.TextBox txtCON_PORT 
      BackColor       =   &H00808080&
      Enabled         =   0   'False
      Height          =   285
      Left            =   3600
      TabIndex        =   3
      Top             =   0
      Width           =   735
   End
   Begin VB.TextBox txtLIST_ON 
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   1680
      TabIndex        =   2
      Text            =   "Listen on port "
      Top             =   240
      Width           =   1815
   End
   Begin VB.TextBox txtCON_TO 
      BackColor       =   &H00C0C0C0&
      Enabled         =   0   'False
      ForeColor       =   &H00800000&
      Height          =   285
      Left            =   1680
      TabIndex        =   1
      Text            =   "connecting to"
      Top             =   0
      Width           =   1815
   End
   Begin VB.Label Label1 
      Caption         =   "Connect to:"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   720
      TabIndex        =   10
      Top             =   0
      Width           =   855
   End
   Begin VB.Label Label4 
      Caption         =   "Server data"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   960
      TabIndex        =   8
      Top             =   720
      Width           =   855
   End
   Begin VB.Label Label3 
      Caption         =   "Client data"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   7
      Top             =   720
      Width           =   735
   End
   Begin VB.Line Line1 
      BorderColor     =   &H8000000D&
      BorderWidth     =   3
      Index           =   0
      X1              =   0
      X2              =   5280
      Y1              =   600
      Y2              =   600
   End
   Begin VB.Label Label2 
      Caption         =   "Listening on:"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   720
      TabIndex        =   0
      Top             =   240
      Width           =   1095
   End
End
Attribute VB_Name = "frmDocument"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public blnDialogShow As Boolean
Public blnConnect As Boolean
Dim strDataLOC As String
Dim strDataREM As String
Dim blnREMOTE As Boolean
Dim blnLOC As Boolean
Dim blnDataREM As Boolean
Dim blnDataLOC As Boolean
Dim speedUP, speedDOWN As Long
Public RTColorRemote As Double
Public RTColorLocaal As Double
Public RemCon As Boolean
Public blnDataExist As Boolean

'Retrieve the IP & data

Private Type LocalIP
    L_addr As String
    L_port As Integer
End Type

Private Type RemoteIP
    R_addr As String
    R_port As Integer
End Type

Dim MyLocalIP As LocalIP
Dim MyRemIP As RemoteIP



Public Sub DisconnectPort()
If blnConnect = True Then
    
    Status.Panels(1) = "Disconnected"
    LOCAAL.Close
    REMOTE.Close
    blnREMOTE = False  'Kan dus nog niets verzenden!
    blnDataREM = False
    tmrSpeed.Enabled = False
    blnConnect = False
    End If

End Sub
Public Sub connectPort()
If blnConnect = False Then

    Status.Panels(1) = "Listen @ Client"
    Me.txtLIST_ON = IP.IPINFO.ListenOnIP
    LOCAAL.Bind MyLocalIP.L_port, MyLocalIP.L_addr
    LOCAAL.Listen
    tmrSpeed.Enabled = True
    blnConnect = True
End If
End Sub


Private Sub cmdClientApply_Click()
'Controle if checkbox is checked..
If chkCDC.Value = 1 Then

    If txtCDC_Source = "" Or txtCDC_Change = "" Then
        MsgBox "enter data in the textbox's or Disable the enable checkbutton !", vbCritical
        Exit Sub
    End If
    fraClientDC.Visible = False
Else
    
    fraClientDC.Visible = False
End If
End Sub

Private Sub cmdServerAppl_Click()
If chkSDC.Value = 1 Then
    
    If txtSDC_Source = "" Or txtSDC_Change = "" Then
        MsgBox "enter data in the textbox's or Disable the enable checkbutton !", vbCritical
        Exit Sub
    End If
    fraServerDC.Visible = False
Else
    
    fraServerDC.Visible = False
End If
End Sub

Private Sub Command1_Click()

If fraExtra.Visible = False Then
fraExtra.Top = Command1.Top + 500
fraExtra.Left = Command1.Left + 10
fraExtra.Visible = True
Else
fraExtra.Visible = False
End If

End Sub





Private Sub Form_Load()
fraExtra.Visible = False
fraServerDC.Visible = False
fraClientDC.Visible = False
On Error GoTo Errors:
'
  'Check if you want to open a file or make a new connection ..
 If Settings.Fileopen = True Then
    Form_Resize
    Settings.Fileopen = False
       Exit Sub
 End If
 
 blnConnect = True
 
 'Get the IP's local & remote incl ports..
 MyLocalIP.L_addr = IP.IPINFO.ListenOnIP
 MyLocalIP.L_port = Val(IP.IPINFO.ListenOnPORT)
 MyRemIP.R_addr = IP.IPINFO.ConnectToIP
 MyRemIP.R_port = Val(IP.IPINFO.ConnectToPORT)
 
 
 Me.Caption = IP.IPINFO.Name
 Me.txtCON_PORT = MyRemIP.R_port
 Me.txtCON_TO = MyRemIP.R_addr
 'Me.txtLIST_ON = IP.IPinfo.ListenOnIP
 Me.txtLIS_ONPORT = MyLocalIP.L_port
 'blnDialogShow = False
 'Statusbar change..
 
 Status.Panels(1) = "Waiting..."
 Status.Panels(2) = " 0 Bit/s"
 Status.Panels(3) = " 0 Bit/s"
 'Set WSLIST to listen..
 LOCAAL.Close
 Me.txtLIST_ON = MyLocalIP.L_addr
 'LOCAAL.LocalPort = IP.IPINFO.ListenOnPORT
 LOCAAL.Bind MyLocalIP.L_port, MyLocalIP.L_addr
 LOCAAL.Listen
 Status.Panels(1) = "Listen @ client"
 blnDataREM = False
 blnREMOTE = False
 
  Form_Resize
    speedUP = 0
    speedDOWN = 0

'Display text in RTBox if checkview Data isn't enabled

'Color Change
Label3(0).ForeColor = IP.RTColorLocaal
Label4.ForeColor = IP.RTColorRemote
Label3(0) = Label3(0)
Label4 = Label4
Exit Sub
Errors:
    MsgBox "An Error Occured , maybe that the local port you want to use , is already in use.., try other port !", vbCritical
    Unload Me
End Sub



Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblExit.ForeColor = 0

End Sub

Private Sub Form_Resize()
    On Error Resume Next
    Line1(0).X2 = Me.Width - 80
    RTBox.Top = 950
    RTBox.Left = 50
    RTBox.Width = Me.Width - 200
    RTBox.Height = Me.Height - 1700
End Sub







Private Sub fraExtra_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblExit.ForeColor = 0
lblHelp.ForeColor = 0
lblSDC.ForeColor = 0
lblCls.ForeColor = 0
lblCDC.ForeColor = 0
End Sub



Private Sub lblCDC_Click()
fraClientDC.Top = Command1.Top + 500
fraClientDC.Left = Command1.Left + 10
fraClientDC.Visible = True
fraExtra.Visible = False
End Sub

Private Sub lblCDC_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblExit.ForeColor = 0
lblHelp.ForeColor = 0
lblSDC.ForeColor = 0
lblCls.ForeColor = 0
lblCDC.ForeColor = vbWhite
End Sub

Private Sub lblCls_Click()
RTBox.Text = ""
fraExtra.Visible = False
End Sub

Private Sub lblCls_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblExit.ForeColor = 0
lblHelp.ForeColor = 0
lblSDC.ForeColor = 0
lblCDC.ForeColor = 0
lblCls.ForeColor = vbWhite
End Sub

Private Sub lblExit_Click()
fraExtra.Visible = False
End Sub

Private Sub lblExit_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblExit.ForeColor = vbWhite
lblHelp.ForeColor = 0
lblCls.ForeColor = 0
lblCDC.ForeColor = 0
lblSDC.ForeColor = 0
End Sub

Private Sub lblHelp_Click()
frmHelp.Show
fraExtra.Visible = False
End Sub

Private Sub lblHelp_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblHelp.ForeColor = vbWhite
lblExit.ForeColor = 0
lblCls.ForeColor = 0
lblSDC.ForeColor = 0
lblCDC.ForeColor = 0
End Sub

Private Sub lblSDC_Click()
fraServerDC.Top = Command1.Top + 500
fraServerDC.Left = Command1.Left + 10
fraExtra.Visible = False

If blnConnect = True Then
        'Disconnect
        lblSDC.Caption = "Connect"
        DisconnectPort
Else
        'Connect
        lblSDC.Caption = "Disconnect"
        connectPort
End If
'fraServerDC.Visible = True
End Sub

Private Sub lblSDC_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblExit.ForeColor = 0
lblHelp.ForeColor = 0
lblCls.ForeColor = 0
lblCDC.ForeColor = 0
lblSDC.ForeColor = vbWhite
End Sub

Private Sub LOCAAL_ConnectionRequest(ByVal requestID As Long)
    ' Check if the control's State is closed. If not,
    ' close the connection before accepting the new
    ' connection.
    If LOCAAL.State <> sckClosed Then _
    LOCAAL.Close
    ' Accept the request with the requestID
    ' parameter.
    LOCAAL.Accept requestID
    Status.Panels(1) = "in progress"
    Connecteer
End Sub
Sub Connecteer()
REMOTE.Connect MyRemIP.R_addr, MyRemIP.R_port
'Wachten van data binnenhalen locaal.. indien host nog niet gevonden..
RemCon = False
End Sub
Private Sub LOCAAL_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
' Idien adres - port in gebruik is..
If Number = 10048 Then
MsgBox "Port is in use.. choose other port"
Dialog.Show
Exit Sub
End If
Status.Panels(1) = " Error!! "
RTBox.SelColor = vbWhite
RTBox.SelText = "Client Error !"
RTBox.SelText = "ERROR :> " & Number & "  " & Description & "  " & Scode & "  " & Source
If Number = sckConnectAborted Then
LOCAAL.Close
End If

End Sub

Private Sub LOCAAL_Close()
LOCAAL.Close
REMOTE.Close
Status.Panels(1) = "Disconnected!"
'cmd_CON_DISCON.Caption = "Connect"
blnREMOTE = False  'Kan dus nog niets verzenden!
blnDataREM = False
blnConnect = False
tmrSpeed.Enabled = False
End Sub


Private Sub Form_Unload(Cancel As Integer)
LOCAAL.Close
REMOTE.Close
tmrSpeed.Enabled = False
End Sub
Private Sub LOCAAL_DataArrival(ByVal bytesTotal As Long)
'This will accept the data from the Client application.
Dim str As String
Status.Panels(1) = "in progress.."

LOCAAL.GetData strDataLOC, vbByte, bytesTotal
str = strDataLOC
'------------------------------------------------
'Changing the string str
'------------------------------------------------
'Put the data in the RT Box Color RED
If Settings.EnableTXT = True Then
RTBox.SelStart = Len(RTBox.TextRTF)
RTBox.SelColor = IP.RTColorLocaal
'Check of it must be send to RTBOX in Hex Format
If Check1.Value = 1 Then
str = StringToHex(strDataLOC)
End If

RTBox.SelText = str & vbCrLf
End If
'record the speed connection
speedUP = speedUP + Len(strDataLOC)
'------------------------------------------------
' NEW *******************************************
' Change data from client appl before it will send to the server
If chkCDC.Value = 1 Then
    
    Debug.Print strDataLOC
    If InStr(1, strDataLOC, txtCDC_Source, vbTextCompare) <> 0 Then
        strDataLOC = Replace(strDataLOC, txtCDC_Source, txtCDC_Change)
        RTBox.SelColor = &H800080
        RTBox.SelStart = Len(RTBox.TextRTF)
        RTBox.SelText = "************ Source Data found from Client Appl & changed ! *********" & _
        vbCrLf & strDataLOC & "************** end changed part *************" & vbCrLf
        
    End If
End If
' -----------------------------------------------
'Send the data to Server socket
SendToServer (strDataLOC)
End Sub



Sub SendToServer(Data As String)
While (blnDataREM = False)
'wanneer vorige data nog niet volledig verzonden is..
DoEvents
Wend
'kijken of er geconecteerd is naar de server..
While (RemCon = False)
DoEvents
Wend

REMOTE.SendData (Data)
End Sub



Private Sub LOCAAL_SendComplete()
blnDataExist = False
End Sub

Private Sub REMOTE_DataArrival(ByVal bytesTotal As Long)
Dim str As String

blnDataExist = True

REMOTE.GetData strDataREM, vbString, bytesTotal
str = strDataREM
'Check to change to hex value..
If Check1.Value = 1 Then
str = StringToHex(strDataREM)
End If
'-----------------------------------------
RTBox.SelColor = vbBlue
RTBox.SelStart = Len(RTBox.TextRTF)
RTBox.SelText = str
'*********** NEW PART  Server data change ****************
If chkSDC.Value = 1 Then
    
     If InStr(1, strDataREM, txtSDC_Source, vbBinaryCompare) <> 0 Then
        strDataREM = Replace(strDataREM, txtSDC_Source, txtSDC_Change)
        LOCAAL.SendData strDataREM
        RTBox.SelColor = &H808080
        RTBox.SelStart = Len(RTBox.TextRTF)
        RTBox.SelText = "************* String found in Server Data & changed ! *********" & _
        vbCrLf & strDataREM & "*************** end changed part *************" & vbCrLf
        
    End If
End If
'Record the speed DOWNLOAD
speedDOWN = speedDOWN + Len(strDataREM)
LOCAAL.SendData strDataREM
End Sub
Private Sub REMOTE_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
Status.Panels(1) = "Error "
RTBox.SelColor = vbWhite
RTBox.SelText = "Server Error !" & vbCrLf
RTBox.SelText = "ERROR :> " & Number & "  " & Description & "  " & Scode & "  " & Source
If Number = 11001 Then
    MsgBox "Hostadress not found !"
End If
lblSDC_Click
End Sub

Private Sub REMOTE_SendComplete()
blnDataREM = True
End Sub

Private Sub REMOTE_SendProgress(ByVal bytesSent As Long, ByVal bytesRemaining As Long)
blnDataREM = False
End Sub
Private Sub REMOTE_Connect()
blnREMOTE = True
blnDataREM = True
RemCon = True
End Sub

Private Sub REMOTE_Close()
REMOTE.Close


LOCAAL.Close
Status.Panels(1) = "Disconnected"
'cmd_CON_DISCON.Caption = "Connect"
blnConnect = True
tmrSpeed.Enabled = False
End Sub

Private Sub RTBox_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblExit.ForeColor = 0
lblHelp.ForeColor = 0
lblSDC.ForeColor = 0
lblCls.ForeColor = 0
lblCDC.ForeColor = 0
End Sub

Private Sub tmrSpeed_Timer()
Dim UP, DOWN As Long

'Show the upload speed in the box
If (speedUP * 2) < 1000 Then
Status.Panels(2) = "UP: " & (speedUP * 2) & " Bytes/s"
Else
UP = (speedUP * 2) / 1000
UP = Round(UP, 2)
Status.Panels(2) = "UP: " & UP & " kB/s"
End If
If (speedDOWN * 2) < 1000 Then
Status.Panels(3) = "DN: " & (speedDOWN * 2) & " Bytes/s"
Else
DOWN = (speedDOWN * 2) / 1000
DOWN = Round(DOWN, 2)
Status.Panels(3) = "DN: " & DOWN & " kB/s"

End If
speedUP = 0
speedDOWN = 0
End Sub

Private Sub txtLIST_ON_Change()
txtLIST_ON = MyLocalIP.L_addr

End Sub
