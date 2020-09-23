VERSION 5.00
Begin VB.Form Appl 
   Caption         =   "Application To use with."
   ClientHeight    =   1365
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   1365
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   2280
      TabIndex        =   3
      Top             =   960
      Width           =   2295
   End
   Begin VB.CommandButton cmdApply 
      Caption         =   "Apply"
      Height          =   375
      Left            =   0
      TabIndex        =   2
      Top             =   960
      Width           =   2175
   End
   Begin VB.TextBox txtFile 
      Height          =   375
      Left            =   0
      TabIndex        =   1
      Top             =   480
      Width           =   4575
   End
   Begin VB.Label Label1 
      Caption         =   "Give te Path  and executable file to use with."
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4335
   End
End
Attribute VB_Name = "Appl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdApply_Click()
'--
modSettings.strApplication = txtFile.Text
modINI.WriteINI modSettings.ThisDir & "Settings.ini", "ApplicationUsed", "APPLICATION", modSettings.strApplication
Unload Me
End Sub

Private Sub cmdCancel_Click()
Unload Me
End Sub

Private Sub Form_Load()
'---
txtFile = modSettings.strApplication

End Sub
