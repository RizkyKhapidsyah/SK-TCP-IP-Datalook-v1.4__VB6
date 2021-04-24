VERSION 5.00
Begin VB.Form frmHelp 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Help "
   ClientHeight    =   5205
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   5370
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5205
   ScaleWidth      =   5370
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame3 
      Caption         =   "Server data Change "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000001&
      Height          =   1455
      Left            =   120
      TabIndex        =   7
      Top             =   3240
      Width           =   5175
      Begin VB.Label Label7 
         Caption         =   "Same as Client data change .."
         ForeColor       =   &H000000C0&
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   1080
         Width           =   3255
      End
      Begin VB.Label Label6 
         Caption         =   "Example: "
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   840
         Width           =   1335
      End
      Begin VB.Label Label1 
         Caption         =   "When data comes from the server , you have the possibility to change it before it will send to your program / Client application"
         Height          =   495
         Left            =   120
         TabIndex        =   8
         Top             =   360
         Width           =   4815
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Client data Change"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000001&
      Height          =   2055
      Left            =   120
      TabIndex        =   3
      Top             =   960
      Width           =   5175
      Begin VB.Label Label5 
         Caption         =   "Example:"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   6
         Top             =   840
         Width           =   1815
      End
      Begin VB.Label Label4 
         Caption         =   $"frmHelp.frx":0000
         ForeColor       =   &H000000C0&
         Height          =   855
         Left            =   120
         TabIndex        =   5
         Top             =   1080
         Width           =   3495
      End
      Begin VB.Label Label3 
         Caption         =   $"frmHelp.frx":00B5
         Height          =   495
         Left            =   120
         TabIndex        =   4
         Top             =   360
         Width           =   4815
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Clearscreen"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000001&
      Height          =   735
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   5175
      Begin VB.Label Label2 
         Caption         =   "Clears the data in the textbox"
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   360
         Width           =   3615
      End
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "OK"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   4800
      Width           =   5175
   End
   Begin VB.Label Label5 
      Caption         =   "Example:"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   240
      TabIndex        =   9
      Top             =   4080
      Width           =   1815
   End
End
Attribute VB_Name = "frmHelp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit



Private Sub OKButton_Click()
Unload Me
End Sub
