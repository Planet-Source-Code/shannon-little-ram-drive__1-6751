VERSION 5.00
Begin VB.Form frmAbout 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "About NeoTrix Program"
   ClientHeight    =   5925
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6390
   Icon            =   "frmAbout.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MousePointer    =   1  'Arrow
   ScaleHeight     =   5925
   ScaleWidth      =   6390
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdOK 
      BackColor       =   &H00C0C0C0&
      Caption         =   " &Ok"
      Default         =   -1  'True
      Height          =   375
      Left            =   5040
      TabIndex        =   0
      Top             =   120
      Width           =   1215
   End
   Begin VB.TextBox txtDisclaimer 
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   2775
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   3000
      Width           =   6135
   End
   Begin VB.Label lblDisclaimerLicense 
      BackStyle       =   0  'Transparent
      Caption         =   "Disclaimer/License:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   240
      TabIndex        =   4
      Top             =   2640
      Width           =   2655
   End
   Begin VB.Label lblntProgramName 
      BackStyle       =   0  'Transparent
      Caption         =   "ntRAM Drive"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   240
      TabIndex        =   3
      Top             =   1200
      Width           =   5775
   End
   Begin VB.Label lblCopyright 
      BackStyle       =   0  'Transparent
      Caption         =   "Copyright 2000  Shannon Little"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   360
      TabIndex        =   2
      Top             =   1800
      Width           =   2535
   End
   Begin VB.Image imgNTLogo 
      Height          =   1125
      Left            =   120
      Picture         =   "frmAbout.frx":1CFA
      Top             =   120
      Width           =   4830
   End
   Begin VB.Label lblDescription 
      BackStyle       =   0  'Transparent
      Caption         =   $"frmAbout.frx":138D4
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   360
      TabIndex        =   1
      Top             =   2040
      Width           =   5895
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdOK_Click()
    Unload frmAbout
End Sub

Private Sub Form_Load()
    txtDisclaimer.Text = strDisclaimer
End Sub
