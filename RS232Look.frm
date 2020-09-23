VERSION 5.00
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Begin VB.Form RS232Look 
   Caption         =   "Form1"
   ClientHeight    =   780
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   780
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer tmrApplCheck 
      Interval        =   1000
      Left            =   3480
      Top             =   240
   End
   Begin MSCommLib.MSComm MSComm1 
      Left            =   1080
      Top             =   120
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
      Handshaking     =   1
   End
   Begin VB.PictureBox Picture1 
      Height          =   495
      Left            =   120
      Picture         =   "RS232Look.frx":0000
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   0
      Top             =   120
      Width           =   495
   End
   Begin VB.Menu mnuRS232 
      Caption         =   "mnuSetting"
      Begin VB.Menu About 
         Caption         =   "About"
      End
      Begin VB.Menu none 
         Caption         =   "-"
      End
      Begin VB.Menu COM_Settings 
         Caption         =   "COM Settings"
      End
      Begin VB.Menu none1 
         Caption         =   "-"
      End
      Begin VB.Menu UseAPP 
         Caption         =   "Use with Appl."
      End
      Begin VB.Menu none2 
         Caption         =   "-"
      End
      Begin VB.Menu StartService 
         Caption         =   "Start Service"
      End
      Begin VB.Menu StopService 
         Caption         =   "Stop Service"
      End
      Begin VB.Menu none3 
         Caption         =   "-"
      End
      Begin VB.Menu ExitProg 
         Caption         =   "Exit"
      End
   End
End
Attribute VB_Name = "RS232Look"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'This program is made by Peter Verburgh.
'-------------------------------------------------
'This program shows an icon in taskbar...
'and my program listens on a serial port that you can choose,
'for incoming data & send it to a specified program..
'example , with my program , you starts Excell, my program got
'now that handle , startadres op that application, and now if data
'entered in the serial port , ex. COM1 , the data would be send to
'the Excell application..
'Remark : It sends only the ASCII data..because i use the Sendkeys function..
'but with some extra code you could send other data to that handle..
'Questions : mail me at Peter.verburgh2@yucom.be
'Now with errorhandling
'And checking (API-processes)of the started application has exited..
'so this appl would stop recieving data.
'-----------------------------------------------------------------------------
Const TH32CS_SNAPHEAPLIST = &H1
Const TH32CS_SNAPPROCESS = &H2
Const TH32CS_SNAPTHREAD = &H4
Const TH32CS_SNAPMODULE = &H8
Const TH32CS_SNAPALL = (TH32CS_SNAPHEAPLIST Or TH32CS_SNAPPROCESS Or TH32CS_SNAPTHREAD Or TH32CS_SNAPMODULE)
Const TH32CS_INHERIT = &H80000000
Const MAX_PATH As Integer = 260
Private Type PROCESSENTRY32
    dwSize As Long
    cntUsage As Long
    th32ProcessID As Long
    th32DefaultHeapID As Long
    th32ModuleID As Long
    cntThreads As Long
    th32ParentProcessID As Long
    pcPriClassBase As Long
    dwFlags As Long
    szExeFile As String * MAX_PATH
End Type
Private Declare Function CreateToolhelp32Snapshot Lib "kernel32" (ByVal lFlags As Long, ByVal lProcessID As Long) As Long
Private Declare Function Process32First Lib "kernel32" (ByVal hSnapShot As Long, uProcess As PROCESSENTRY32) As Long
Private Declare Function Process32Next Lib "kernel32" (ByVal hSnapShot As Long, uProcess As PROCESSENTRY32) As Long
Private Declare Sub CloseHandle Lib "kernel32" (ByVal hPass As Long)

Dim hSnapShot As Long, uProcess As PROCESSENTRY32

Dim Port1 As Integer
Dim ReturnValue
Dim Parity2 As String
Dim Stopbits1 As Integer
Dim blnStat As Boolean
Dim blnApplRun As Boolean


Public Sub CreateIcon()
    Dim Tic As NOTIFYICONDATA
    Tic.cbSize = Len(Tic)
    Tic.hwnd = Picture1.hwnd
    Tic.uID = 1&
    Tic.uFlags = NIF_DOALL
    Tic.uCallbackMessage = WM_MOUSEMOVE
    Tic.hIcon = Picture1.Picture
    Tic.szTip = "RS232 DataReceiver " & Chr$(0)
    erg = Shell_NotifyIcon(NIM_ADD, Tic)
End Sub

Public Sub DeleteIcon()
    Dim Tic As NOTIFYICONDATA
    Tic.cbSize = Len(Tic)
    Tic.hwnd = Picture1.hwnd
    Tic.uID = 1&
    erg = Shell_NotifyIcon(NIM_DELETE, Tic)
End Sub

Private Sub About_Click()
 frmAbout.Show
End Sub

Private Sub COM_Settings_Click()
'MsgBox modSettings.Comport & " " & modSettings.strDatabits
COMSettings.Show
End Sub

Private Sub ExitProg_Click()
Unload Me
End Sub

Private Sub Form_Load()
CreateIcon

modSettings.ThisDir = CurDir  'IF COMPILED !!!
blnApplRun = True   'No Attached program is running..
tmrApplCheck.Enabled = False
modSettings.ThisDir = modSettings.ThisDir & "\"
modSettings.Comport = modINI.sGetINI(modSettings.ThisDir & "Settings.ini", "Settings", "COMPORT", "?")
modSettings.strBaudrate = modINI.sGetINI(modSettings.ThisDir & "Settings.ini", "Settings", "BAUDRATE", "?")
modSettings.strDatabits = modINI.sGetINI(modSettings.ThisDir & "Settings.ini", "Settings", "DATABITS", "?")
modSettings.strParity = modINI.sGetINI(modSettings.ThisDir & "Settings.ini", "Settings", "PARITY", "?")
modSettings.strStopBits = modINI.sGetINI(modSettings.ThisDir & "Settings.ini", "Settings", "STOPBITS", "?")
modSettings.strApplication = modINI.sGetINI(modSettings.ThisDir & "Settings.ini", "ApplicationUsed", "APPLICATION", "?")
'-----------------------------------------------------------------
Me.Hide
StopService.Enabled = False
blnStat = False
End Sub

Private Sub Form_Terminate()
DeleteIcon
End Sub

Private Sub Form_Unload(Cancel As Integer)
DeleteIcon
End Sub
Private Sub Picture1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    X = X / Screen.TwipsPerPixelX


    Select Case X
        Case WM_LBUTTONDOWN
        
        Case WM_RBUTTONDOWN
        
        PopupMenu mnuRS232
        Case WM_MOUSEMOVE
        
        Case WM_LBUTTONDBLCLK
        
    End Select
End Sub

Private Sub StartService_Click()
On Error GoTo Error1
blnStat = True
StopService.Enabled = True
StartService.Enabled = False
'Starten  seriele Communication...
Settings
MSComm1.CommPort = Port1

MSComm1.Handshaking = comXOnXoff
MSComm1.InBufferCount = 0
portset = modSettings.strBaudrate & "," & Parity2 & "," & modSettings.strDatabits & "," & Stopbits1
MSComm1.Settings = portset

MSComm1.PortOpen = True
'---- Reading data & sending to the specified program
ReturnValue = Shell(modSettings.strApplication, 1)   ' Run Application.
AppActivate ReturnValue
'Debug.Print RetunValue
tmrApplCheck.Enabled = True
While blnStat = True
        If MSComm1.InBufferCount Then
        On Error GoTo AppClosed
        AppActivate ReturnValue
        SendKeys MSComm1.Input
         'Write to Client
    End If
     DoEvents
     If blnApplRun = False Then
           GoTo ErrorProgClosed
      End If
     
     Wend
Exit Sub
AppClosed:
MsgBox "Error , Application is Closed!", vbCritical
blnStat = False
StopService.Enabled = False
StartService.Enabled = True
Exit Sub
Error1:
MsgBox "Error , Comport settings are not correct !", vbCritical
ErrorProgClosed:
MsgBox "Error, Application Where Data must send to is Closed ! ", vbCritical
blnStat = False
StopService.Enabled = False
StartService.Enabled = True
Call StopService_Click
blnApplRun = True
End Sub

Sub Settings()
Select Case modSettings.Comport
    Case "COM1"
        Port1 = 1
    Case "COM2"
       Port1 = 2
    Case "COM3"
      Port1 = 3
    Case "COM4"
      Port1 = 4
End Select
'-------------------- Stopbits ------------
Stopbits1 = Val(modSettings.strStopBits)
'--------------------- PARITY -------------
Parity2 = UCase(Mid(modSettings.strParity, 1, 1))
End Sub



Private Sub StopService_Click()
StopService.Enabled = False
StartService.Enabled = True
MSComm1.PortOpen = False
blnStat = False
tmrApplCheck.Enabled = False
End Sub

Private Sub tmrApplCheck_Timer()
blnApplRun = CheckApplication(ReturnValue)
End Sub

Private Sub UseAPP_Click()
'Check APP............ ToDo.. yet
Appl.Show
End Sub

Public Function CheckApplication(ByVal handle As Long) As Boolean
'This api calls look if the ProcessID exist ..
'If the user Close the program - window where the data must be
'send then it must send an error to the user & stop accepting data
'from the serial port !!!

Dim blnCheck As Boolean
blnCheck = False
hSnapShot = CreateToolhelp32Snapshot(TH32CS_SNAPALL, 0&)
    'set the length of our ProcessEntry-type
    uProcess.dwSize = Len(uProcess)
    'Retrieve information about the first process encountered in our system snapshot
    r = Process32First(hSnapShot, uProcess)
    Do While r
        'Debug.Print "Handle " & uProcess.th32ProcessID
        'Retrieve information about the next process recorded in our system snapshot
        If uProcess.th32ProcessID = handle Then
            blnCheck = True
        End If
        r = Process32Next(hSnapShot, uProcess)
    Loop
    'close our snapshot handle
    CloseHandle hSnapShot
    CheckApplication = blnCheck
End Function
