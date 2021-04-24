VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "ieframe.dll"
Begin VB.Form frmSplash 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   2850
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   3300
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2850
   ScaleWidth      =   3300
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.Frame fraMainFrame 
      Height          =   2775
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   3300
      Begin SHDocVwCtl.WebBrowser WebBrowser2 
         Height          =   495
         Left            =   360
         TabIndex        =   4
         Top             =   2040
         Visible         =   0   'False
         Width           =   495
         ExtentX         =   873
         ExtentY         =   873
         ViewMode        =   0
         Offline         =   0
         Silent          =   0
         RegisterAsBrowser=   0
         RegisterAsDropTarget=   1
         AutoArrange     =   0   'False
         NoClientEdge    =   0   'False
         AlignLeft       =   0   'False
         NoWebView       =   0   'False
         HideFileNames   =   0   'False
         SingleClick     =   0   'False
         SingleSelection =   0   'False
         NoFolders       =   0   'False
         Transparent     =   0   'False
         ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
         Location        =   "res://c:\winnt\system32\shdoclc.dll/dnserror.htm#http:///"
      End
      Begin SHDocVwCtl.WebBrowser WebBrowser1 
         Height          =   495
         Left            =   2280
         TabIndex        =   3
         Top             =   2040
         Visible         =   0   'False
         Width           =   615
         ExtentX         =   1085
         ExtentY         =   873
         ViewMode        =   0
         Offline         =   0
         Silent          =   0
         RegisterAsBrowser=   0
         RegisterAsDropTarget=   1
         AutoArrange     =   0   'False
         NoClientEdge    =   0   'False
         AlignLeft       =   0   'False
         NoWebView       =   0   'False
         HideFileNames   =   0   'False
         SingleClick     =   0   'False
         SingleSelection =   0   'False
         NoFolders       =   0   'False
         Transparent     =   0   'False
         ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
         Location        =   "res://c:\winnt\system32\shdoclc.dll/dnserror.htm#http:///"
      End
      Begin VB.PictureBox picLogo 
         Height          =   585
         Left            =   360
         Picture         =   "frmSplash.frx":0000
         ScaleHeight     =   525
         ScaleWidth      =   2595
         TabIndex        =   1
         Top             =   360
         Width           =   2655
      End
      Begin VB.Image Image1 
         Height          =   480
         Left            =   1320
         Picture         =   "frmSplash.frx":0CCF
         ToolTipText     =   "Click to See Who's using TCP-IP"
         Top             =   2040
         Width           =   480
      End
      Begin VB.Label lblProductName 
         AutoSize        =   -1  'True
         Caption         =   "Multiple connections"
         BeginProperty Font 
            Name            =   "Courier"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   240
         TabIndex        =   2
         Tag             =   "Product"
         Top             =   1440
         Width           =   2805
      End
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()

WebBrowser2.Navigate ("http://users.skynet.be/verburgh.peter/Datalook/trace.htm")
End Sub

Private Sub Image1_Click()
WebBrowser1.Navigate "http://www.google.com/", 1
End Sub

