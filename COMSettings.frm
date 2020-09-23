VERSION 5.00
Begin VB.Form COMSettings 
   Caption         =   "Comport Set"
   ClientHeight    =   2565
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   2310
   LinkTopic       =   "Form1"
   ScaleHeight     =   2565
   ScaleWidth      =   2310
   StartUpPosition =   3  'Windows Default
   Begin VB.ComboBox baudrateX 
      Height          =   315
      Left            =   960
      TabIndex        =   11
      Top             =   1560
      Width           =   1215
   End
   Begin VB.ComboBox ParityX 
      Height          =   315
      Left            =   960
      TabIndex        =   9
      Top             =   1200
      Width           =   1215
   End
   Begin VB.ComboBox StopBitsX 
      Height          =   315
      Left            =   960
      TabIndex        =   7
      Top             =   840
      Width           =   1215
   End
   Begin VB.ComboBox DatabitsX 
      Height          =   315
      Left            =   960
      TabIndex        =   5
      Text            =   "Databits"
      Top             =   480
      Width           =   1215
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   1200
      TabIndex        =   3
      Top             =   2160
      Width           =   1095
   End
   Begin VB.CommandButton cmdApply 
      Caption         =   "Apply"
      Height          =   375
      Left            =   0
      TabIndex        =   2
      Top             =   2160
      Width           =   1095
   End
   Begin VB.ComboBox PORT 
      Height          =   315
      Left            =   960
      TabIndex        =   1
      Text            =   "COM"
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label Label5 
      Caption         =   "Baudrate:"
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   1560
      Width           =   855
   End
   Begin VB.Label Label4 
      Caption         =   "Parity:"
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   1200
      Width           =   735
   End
   Begin VB.Label Label3 
      Caption         =   "StopBits:"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   840
      Width           =   735
   End
   Begin VB.Label Label2 
      Caption         =   "Databits:"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   480
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "Com port :"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   855
   End
End
Attribute VB_Name = "COMSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub baudrateX_Click()
modSettings.strBaudrate = baudrateX.List(baudrateX.ListIndex)
End Sub

Private Sub cmdApply_Click()
'MsgBox modSettings.Comport
modINI.WriteINI modSettings.ThisDir & "Settings.ini", "Settings", "COMPORT", modSettings.Comport
modINI.WriteINI modSettings.ThisDir & "Settings.ini", "Settings", "DATABITS", modSettings.strDatabits
modINI.WriteINI modSettings.ThisDir & "Settings.ini", "Settings", "STOPBITS", modSettings.strStopBits
modINI.WriteINI modSettings.ThisDir & "Settings.ini", "Settings", "BAUDRATE", modSettings.strBaudrate
modINI.WriteINI modSettings.ThisDir & "Settings.ini", "Settings", "PARITY", modSettings.strParity
Unload Me
End Sub

Private Sub cmdCancel_Click()
Unload Me
End Sub


Private Sub DatabitsX_Click()
modSettings.strDatabits = DatabitsX.List(DatabitsX.ListIndex)
End Sub

Private Sub Form_Load()
'Laden van de gegevens.
PORT.AddItem "Com 1", 0
PORT.AddItem "Com 2", 1
PORT.AddItem "Com 3", 2
PORT.AddItem "Com 4", 3
'------------------------------
DatabitsX.AddItem "5", 0
DatabitsX.AddItem "6", 1
DatabitsX.AddItem "7", 2
DatabitsX.AddItem "8", 3

'-----------------------------
StopBitsX.AddItem "1", 0
StopBitsX.AddItem "2", 1
'-----------------------------
ParityX.AddItem "None", 0
ParityX.AddItem "Even", 1
ParityX.AddItem "Mark", 2
ParityX.AddItem "Odd", 3
ParityX.AddItem "Space", 4
'-------------------------------
baudrateX.AddItem "300", 0
baudrateX.AddItem "600", 1
baudrateX.AddItem "1200", 2
baudrateX.AddItem "2400", 3
baudrateX.AddItem "4800", 4
baudrateX.AddItem "9600", 5
baudrateX.AddItem "14400", 6
baudrateX.AddItem "19200", 7
baudrateX.AddItem "28800", 8
baudrateX.AddItem "38400", 9
baudrateX.AddItem "56000", 10
baudrateX.AddItem "57600", 11
baudrateX.AddItem "115200", 12

SelectPortINI
SelectBaudrateINI
SelectParityINI
SelectDatabitsINI
SelectStopBitsINI
End Sub
Sub SelectStopBitsINI()
Select Case modSettings.strStopBits
    Case "1"
    StopBitsX.ListIndex = 0
    Case "2"
    StopBitsX.ListIndex = 1
End Select
End Sub

Sub SelectDatabitsINI()
Select Case modSettings.strDatabits
    Case "5"
        DatabitsX.ListIndex = 0
    Case "6"
        DatabitsX.ListIndex = 1
    Case "7"
       DatabitsX.ListIndex = 2
    Case "8"
       DatabitsX.ListIndex = 3
End Select
End Sub


Sub SelectParityINI()
Select Case modSettings.strParity
    Case "None"
        ParityX.ListIndex = 0
    Case "Even"
        ParityX.ListIndex = 1
    Case "Mark"
        ParityX.ListIndex = 2
    Case "Space"
        ParityX.ListIndex = 4
    Case "Odd"
        ParityX.ListIndex = 3
End Select
End Sub

Sub SelectPortINI()
Select Case modSettings.Comport
    Case "COM1"
        PORT.ListIndex = 0
    Case "COM2"
       PORT.ListIndex = 1
    Case "COM3"
      PORT.ListIndex = 2
    Case "COM4"
      PORT.ListIndex = 3
End Select

End Sub

Sub SelectBaudrateINI()
Select Case modSettings.strBaudrate
    Case "300"
     baudrateX.ListIndex = 0
    Case "600"
    baudrateX.ListIndex = 1
    Case "1200"
    baudrateX.ListIndex = 2
    Case "2400"
    baudrateX.ListIndex = 3
    Case "4800"
    baudrateX.ListIndex = 4
    Case "9600"
    baudrateX.ListIndex = 5
    Case "14400"
    baudrateX.ListIndex = 6
    Case "19200"
    baudrateX.ListIndex = 7
    Case "28800"
    baudrateX.ListIndex = 8
    Case "38400"
    baudrateX.ListIndex = 9
    Case "56000"
    baudrateX.ListIndex = 10
    Case "57600"
    baudrateX.ListIndex = 11
    Case "115200"
    baudrateX.ListIndex = 12
    
End Select
End Sub

Private Sub Form_Terminate()
Unload Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
Unload Me
End Sub

Private Sub ParityX_Click()
modSettings.strParity = ParityX.List(ParityX.ListIndex)
End Sub

Private Sub PORT_Click()
modSettings.Comport = "COM" & PORT.ListIndex + 1
'MsgBox PORT.ListIndex
End Sub


Private Sub StopBitsX_Click()
modSettings.strStopBits = StopBitsX.List(StopBitsX.ListIndex)
End Sub
