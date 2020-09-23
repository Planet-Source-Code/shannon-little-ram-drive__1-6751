VERSION 5.00
Begin VB.Form frmDisclaimer 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Disclaimer - Please read carefully"
   ClientHeight    =   4770
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5145
   Icon            =   "frmDisclaimer.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4770
   ScaleWidth      =   5145
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdAgree 
      Caption         =   "&Agree"
      Enabled         =   0   'False
      Height          =   375
      Left            =   2520
      TabIndex        =   2
      Top             =   4320
      Width           =   1215
   End
   Begin VB.CommandButton cmdDisagree 
      Caption         =   "&Disagree"
      Default         =   -1  'True
      Height          =   375
      Left            =   3840
      TabIndex        =   3
      Top             =   4320
      Width           =   1215
   End
   Begin VB.CheckBox chkAgree 
      Caption         =   "I have read and agree to the disclaimer/liscense above."
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   3840
      Width           =   4335
   End
   Begin VB.TextBox txtDisclaimerText 
      BackColor       =   &H00C0C0C0&
      Height          =   3255
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   480
      Width           =   4935
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00FFFFFF&
      X1              =   120
      X2              =   5040
      Y1              =   4200
      Y2              =   4200
   End
   Begin VB.Label lblDisclaimerLiscenseTitle 
      Caption         =   "Disclaimer/License for ntRam Drive"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   45
      Width           =   4935
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      BorderWidth     =   2
      X1              =   5040
      X2              =   120
      Y1              =   4200
      Y2              =   4200
   End
End
Attribute VB_Name = "frmDisclaimer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub chkAgree_Click()
    'Disables the agree button if the checkbox is not checked
    If chkAgree = vbChecked Then
        cmdAgree.Enabled = True
    Else
        cmdAgree.Enabled = False
    End If
End Sub

Private Sub cmdAgree_Click()
    Dim lngTemp As Long
    'Save setting that the disclaimer has been agreed to
    SetKey "ATD", "TRUE"
    For lngTemp = 5145 To 0 Step -50
        Height = lngTemp
    Next lngTemp
    Unload frmDisclaimer
    frmMain.Show
End Sub

Private Sub cmdDisagree_Click()
    Unload frmDisclaimer
    End
End Sub

Private Sub Form_Load()
    txtDisclaimerText.Text = strDisclaimer
End Sub

