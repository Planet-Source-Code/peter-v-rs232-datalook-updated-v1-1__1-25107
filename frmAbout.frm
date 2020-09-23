VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Serial datareceiver  sends data to the wanted application !"
   ClientHeight    =   2535
   ClientLeft      =   2340
   ClientTop       =   1935
   ClientWidth     =   5730
   ClipControls    =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1749.702
   ScaleMode       =   0  'User
   ScaleWidth      =   5380.766
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   345
      Left            =   2040
      TabIndex        =   0
      Top             =   2040
      Width           =   1260
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      BorderStyle     =   6  'Inside Solid
      Index           =   1
      X1              =   112.686
      X2              =   5337.57
      Y1              =   1325.218
      Y2              =   1325.218
   End
   Begin VB.Label lblDescription 
      Caption         =   "Created by Verburgh Peter  peter.verburgh2@yucom.be                       "
      ForeColor       =   &H00000000&
      Height          =   570
      Left            =   1320
      TabIndex        =   1
      Top             =   1200
      Width           =   2685
   End
   Begin VB.Label lblTitle 
      Caption         =   "RS232 ASCII receiver"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   480
      Left            =   600
      TabIndex        =   2
      Top             =   240
      Width           =   4245
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      Index           =   0
      X1              =   98.6
      X2              =   4507.448
      Y1              =   1325.218
      Y2              =   1325.218
   End
   Begin VB.Label lblVersion 
      Caption         =   "Version 1.1"
      Height          =   225
      Left            =   1920
      TabIndex        =   3
      Top             =   840
      Width           =   1245
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub cmdOK_Click()
  Unload Me
End Sub

