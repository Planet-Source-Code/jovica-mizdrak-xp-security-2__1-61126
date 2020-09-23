VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Xp Security 2 (BETA)"
   ClientHeight    =   5715
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6195
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5715
   ScaleWidth      =   6195
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin VB.CommandButton Command5 
      Caption         =   "Change Icon"
      Height          =   495
      Left            =   3000
      TabIndex        =   0
      Top             =   240
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.CheckBox chkDisablePA 
      Caption         =   "Windows Task Manager"
      Height          =   255
      Index           =   15
      Left            =   3960
      TabIndex        =   35
      Top             =   3840
      Value           =   1  'Checked
      Width           =   2175
   End
   Begin VB.CheckBox chkDisablePA 
      Caption         =   "Windows Firewall"
      Height          =   255
      Index           =   14
      Left            =   3960
      TabIndex        =   34
      Top             =   3480
      Value           =   1  'Checked
      Width           =   2175
   End
   Begin VB.CheckBox chkDisablePA 
      Caption         =   "Windows Security Center"
      Height          =   255
      Index           =   13
      Left            =   3960
      TabIndex        =   33
      Top             =   3120
      Value           =   1  'Checked
      Width           =   2175
   End
   Begin VB.CheckBox chkDisablePA 
      Caption         =   "Automatic Updates"
      Height          =   255
      Index           =   12
      Left            =   3960
      TabIndex        =   32
      Top             =   2760
      Value           =   1  'Checked
      Width           =   2175
   End
   Begin VB.CheckBox chkDisablePA 
      Caption         =   "Add or Remove Programs"
      Height          =   255
      Index           =   11
      Left            =   3960
      TabIndex        =   31
      Top             =   2400
      Value           =   1  'Checked
      Width           =   2175
   End
   Begin VB.Timer tmrDisablePA 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   0
      Top             =   0
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Lock Windows"
      Height          =   375
      Left            =   3960
      TabIndex        =   30
      Top             =   4560
      Width           =   2175
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Disable Task Manager"
      Height          =   255
      Left            =   4200
      TabIndex        =   29
      Top             =   3840
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Frame Frame2 
      Caption         =   "More Security"
      Height          =   255
      Left            =   3960
      TabIndex        =   28
      Top             =   4200
      Width           =   2175
   End
   Begin VB.CheckBox chkDisablePA 
      Caption         =   "User Accounts 2"
      Height          =   255
      Index           =   10
      Left            =   3960
      TabIndex        =   27
      Top             =   2040
      Value           =   1  'Checked
      Width           =   2175
   End
   Begin VB.CheckBox chkDisablePA 
      Caption         =   "User Accounts"
      Height          =   255
      Index           =   9
      Left            =   3960
      TabIndex        =   26
      Top             =   1680
      Value           =   1  'Checked
      Width           =   2175
   End
   Begin VB.CheckBox chkDisablePA 
      Caption         =   "Taskbar Properties"
      Height          =   255
      Index           =   8
      Left            =   1920
      TabIndex        =   25
      Top             =   4560
      Value           =   1  'Checked
      Width           =   1935
   End
   Begin VB.CheckBox chkDisablePA 
      Caption         =   "Internet Properties"
      Height          =   255
      Index           =   7
      Left            =   1920
      TabIndex        =   24
      Top             =   4200
      Value           =   1  'Checked
      Width           =   1935
   End
   Begin VB.CheckBox chkDisablePA 
      Caption         =   "Folder Options"
      Height          =   255
      Index           =   6
      Left            =   1920
      TabIndex        =   23
      Top             =   3840
      Value           =   1  'Checked
      Width           =   1935
   End
   Begin VB.CheckBox chkDisablePA 
      Caption         =   "Display Properties"
      Height          =   255
      Index           =   5
      Left            =   1920
      TabIndex        =   22
      Top             =   3480
      Value           =   1  'Checked
      Width           =   1935
   End
   Begin VB.CheckBox chkDisablePA 
      Caption         =   "System Properties"
      Height          =   255
      Index           =   4
      Left            =   1920
      TabIndex        =   21
      Top             =   3120
      Value           =   1  'Checked
      Width           =   1935
   End
   Begin VB.Frame Frame1 
      Caption         =   "Disable Keys"
      Height          =   255
      Left            =   120
      TabIndex        =   20
      Top             =   1200
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Deactivate"
      Enabled         =   0   'False
      Height          =   495
      Index           =   1
      Left            =   2160
      TabIndex        =   19
      Top             =   5160
      Width           =   1935
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Activate"
      Height          =   495
      Index           =   0
      Left            =   120
      TabIndex        =   18
      Top             =   5160
      Width           =   1935
   End
   Begin VB.Frame frameDPA 
      Caption         =   "Disable Programm Access"
      Height          =   255
      Left            =   1800
      TabIndex        =   17
      Top             =   1200
      Width           =   4335
   End
   Begin VB.CheckBox chkDisablePA 
      Caption         =   "Administrative Tools"
      Height          =   255
      Index           =   3
      Left            =   1920
      TabIndex        =   16
      Top             =   2760
      Value           =   1  'Checked
      Width           =   1935
   End
   Begin VB.CheckBox chkDisablePA 
      Caption         =   "CMD Prompt"
      Height          =   255
      Index           =   2
      Left            =   1920
      TabIndex        =   15
      Top             =   2400
      Value           =   1  'Checked
      Width           =   1935
   End
   Begin VB.CheckBox chkDisablePA 
      Caption         =   "Control Panel"
      Height          =   255
      Index           =   1
      Left            =   1920
      TabIndex        =   14
      Top             =   2040
      Value           =   1  'Checked
      Width           =   1935
   End
   Begin VB.CheckBox chkDisablePA 
      Caption         =   "Registry Editor"
      Height          =   255
      Index           =   0
      Left            =   1920
      TabIndex        =   13
      Top             =   1680
      Value           =   1  'Checked
      Width           =   1935
   End
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   3255
      Left            =   120
      ScaleHeight     =   3255
      ScaleWidth      =   1575
      TabIndex        =   3
      Top             =   1680
      Width           =   1575
      Begin VB.CheckBox chkDisable 
         Caption         =   "WIN POPUP"
         Height          =   255
         Index           =   8
         Left            =   0
         TabIndex        =   12
         Top             =   2880
         Value           =   1  'Checked
         Width           =   1935
      End
      Begin VB.CheckBox chkDisable 
         Caption         =   "APP POPUP"
         Height          =   255
         Index           =   7
         Left            =   0
         TabIndex        =   11
         Top             =   2520
         Value           =   1  'Checked
         Width           =   1935
      End
      Begin VB.CheckBox chkDisable 
         Caption         =   "WIN + L"
         Height          =   255
         Index           =   6
         Left            =   0
         TabIndex        =   10
         Top             =   2160
         Value           =   1  'Checked
         Width           =   1935
      End
      Begin VB.CheckBox chkDisable 
         Caption         =   "CTRL + ESCAPE"
         Height          =   255
         Index           =   5
         Left            =   0
         TabIndex        =   9
         Top             =   1800
         Value           =   1  'Checked
         Width           =   1935
      End
      Begin VB.CheckBox chkDisable 
         Caption         =   "ALT + F4"
         Height          =   255
         Index           =   4
         Left            =   0
         TabIndex        =   8
         Top             =   1440
         Value           =   1  'Checked
         Width           =   1935
      End
      Begin VB.CheckBox chkDisable 
         Caption         =   "ALT + ENTER"
         Height          =   255
         Index           =   3
         Left            =   0
         TabIndex        =   7
         Top             =   1080
         Value           =   1  'Checked
         Width           =   1935
      End
      Begin VB.CheckBox chkDisable 
         Caption         =   "ALT + SPACE"
         Height          =   255
         Index           =   2
         Left            =   0
         TabIndex        =   6
         Top             =   720
         Value           =   1  'Checked
         Width           =   1935
      End
      Begin VB.CheckBox chkDisable 
         Caption         =   "ALT + TAB"
         Height          =   255
         Index           =   1
         Left            =   0
         TabIndex        =   5
         Top             =   360
         Value           =   1  'Checked
         Width           =   1935
      End
      Begin VB.CheckBox chkDisable 
         Caption         =   "ALT + ESCAPE"
         Height          =   255
         Index           =   0
         Left            =   0
         TabIndex        =   4
         Top             =   0
         Value           =   1  'Checked
         Width           =   1935
      End
   End
   Begin VB.PictureBox Picture2 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   1080
      Left            =   0
      Picture         =   "Form1.frx":0000
      ScaleHeight     =   1080
      ScaleWidth      =   6195
      TabIndex        =   2
      Top             =   0
      Width           =   6195
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Done"
      Height          =   495
      Left            =   4200
      TabIndex        =   1
      Top             =   5160
      Width           =   1935
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00FFFFFF&
      Index           =   1
      X1              =   0
      X2              =   6240
      Y1              =   5050
      Y2              =   5050
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000003&
      Index           =   0
      X1              =   0
      X2              =   6240
      Y1              =   5040
      Y2              =   5040
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      Index           =   0
      X1              =   0
      X2              =   6240
      Y1              =   1080
      Y2              =   1080
   End
   Begin VB.Menu mSysPopup 
      Caption         =   "SysPopup"
      Visible         =   0   'False
      Begin VB.Menu mShow 
         Caption         =   "Show"
      End
      Begin VB.Menu mSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mExit 
         Caption         =   "Exit"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


'this API is necessary to make sure that menu will disappear if user clicks outside of it
Private Declare Function SetForegroundWindow Lib "user32" (ByVal hwnd As Long) As Long
Dim hhkLowLevelKybd As Long, iExit As Boolean

Private Sub Command2_Click()
Form_Unload CInt(iExit)
End Sub

Private Sub Command3_Click()
Dim WindowToFind As Long    'Window Handle
WindowToFind& = FindWindow("RegEdit_RegEdit", "Registry Editor")
' Look for the Window
Call ShowWindow(WindowToFind&, SW_HIDE)
Call SendMessageLong(WindowToFind&, WM_CLOSE, 0&, 0&)
End Sub

Private Sub Command4_Click()
On Error GoTo err:
   Dim sPath As String
   Dim cNewDesktop As New cDesktop
   cNewDesktop.Create DESKTOP_NAME
   sPath = App.Path & "\Tools\ldesk.exe"
   cNewDesktop.StartProcess sPath
   Form_Unload CInt(iExit)
Exit Sub
err:
MsgBox "Error number: " & err.Number & vbNewLine & "Description: " & err.Description, vbCritical
End Sub

Private Sub Command5_Click()
    TrayChangeIcon Form1, App.Path & "\Res\2.ico", "XP Security"
End Sub

Private Sub Command6_Click()
On Error GoTo err:
Static bolTaskMgr As Boolean
If bolTaskMgr = False Then
  Dim x As Long
    x = FindWindow("#32770", "Windows Task Manager")
        DoEvents
            x = FindWindow("#32770", "Windows Task Manager")
            Call ShowWindow(x, SW_HIDE)
            Call SendMessageLong(x, WM_CLOSE, 0&, 0&)
bolTaskMgr = True
Command6.Caption = "Enable Task Manager"
SetAttr "C:\windows\system32\taskmgr.exe", vbHidden + vbSystem
Open "C:\windows\system32\taskmgr.exe" For Binary As #1
Else
Close #1
bolTaskMgr = False
Command6.Caption = "Disable Task Manager"
End If
Exit Sub
err:
MsgBox "Error number: " & err.Number & vbNewLine & "Description: " & err.Description, vbCritical
End Sub

Private Sub Form_Initialize()
    InitControlsXP
End Sub

Private Sub Command1_Click(Index As Integer)
Select Case Index
    Case 0
        hhkLowLevelKybd = SetWindowsHookEx(WH_KEYBOARD_LL, AddressOf LowLevelKeyboardProc, App.hInstance, 0)
        Command1(0).Enabled = False
        Command1(1).Enabled = True
        Command6_Click
        tmrDisablePA.Enabled = True
    Case 1
        UnhookWindowsHookEx hhkLowLevelKybd
        hhkLowLevelKybd = 0
        Command1(0).Enabled = True
        Command1(1).Enabled = False
        Command6_Click
        tmrDisablePA.Enabled = False
End Select
End Sub

Private Sub Form_Load()
On Error GoTo err:
    SetIcon Me.hwnd, "AA0", True
    iExit = False
    'Load the system tray feature
    TrayAddIcon Form1, App.Path & "\Res\1.ico", "XP Security"
Exit Sub
err:
MsgBox "Error number: " & err.Number & vbNewLine & "Description: " & err.Description, vbCritical
End Sub

Private Sub Form_MouseMove(button As Integer, Shift As Integer, x As Single, Y As Single)
    'All mouse events including balloon click
    Dim Result As Long
    Dim cEventx As Single
    Dim cEventy As Single
    cEventx = x / Screen.TwipsPerPixelX
    cEventy = Y / Screen.TwipsPerPixelY

Select Case cEventx Xor cEventy
    Case MouseMove
        'Debug.Print "MouseMove"
    Case LeftUp
        'Debug.Print "Left Up"
    Case LeftDown
        'Debug.Print "LeftDown"
    Case LeftDbClick
        'Debug.Print "LeftDbClick"
        TrayRemoveIcon
        Me.WindowState = 0
        Me.Show
    Case MiddleUp
        'Debug.Print "MiddleUp"
    Case MiddleDown
        'Debug.Print "MiddleDown"
    Case MiddleDbClick
        'Debug.Print "MiddleDbClick"
    Case RightUp
        'Debug.Print "RightUp"
        'now show it
        PopupMenu mSysPopup, , , , mShow
    Case RightDown
        'Debug.Print "RightDown"
        'make sure that menu will disappear if user clicks outside of it
        Result = SetForegroundWindow(Me.hwnd)
    Case RightDbClick
        'Debug.Print "RightDbClick"
    Case BalloonClick
        'Debug.Print "Balloon Click"

    End Select
End Sub

Private Sub Form_Resize()
If Me.WindowState = vbMinimized Then
        Me.Hide
        TrayAddIcon Form1, App.Path & "\Res\1.ico", "XP Security"
End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
Cancel = iExit
If Cancel = True Then
    TrayRemoveIcon
    If hhkLowLevelKybd <> 0 Then UnhookWindowsHookEx hhkLowLevelKybd
Else
    Cancel = True
    Me.Hide
    TrayAddIcon Form1, App.Path & "\Res\1.ico", "XP Security"
End If
End Sub

Private Sub mShow_Click()
'If bolPassword Then
'    Me.WindowState = 0
'    Me.Show
'Else
'    frmPassword.Show
'End If
End Sub

Private Sub mExit_Click()
    iExit = True
    Unload Me
    End
End Sub

Private Sub tmrDisablePA_Timer()
On Error Resume Next
Dim WindowToFind As Long    'Window Handle
Dim ChildWin As Long
Dim ParentWin As Long
Dim RunDll As String
RunDll = "C:\WINDOWS\system32\rundll32.exe"
If Form1.chkDisablePA(0).Value = 1 Then
    WindowToFind& = FindWindow("RegEdit_RegEdit", "Registry Editor") ' Look for "Registry Editor"
    Call ShowWindow(WindowToFind&, SW_HIDE)
    Call SendMessageLong(WindowToFind&, WM_CLOSE, 0&, 0&)
End If

If Form1.chkDisablePA(1).Value = 1 Then
WindowToFind& = FindWindow("CabinetWClass", "Control Panel") ' Look for "Control Panel"
    Call ShowWindow(WindowToFind&, SW_HIDE)
    Call SendMessageLong(WindowToFind&, WM_CLOSE, 0&, 0&)
End If

If Form1.chkDisablePA(2).Value = 1 Then
    WindowToFind& = FindWindow("ConsoleWindowClass", "C:\WINDOWS\system32\cmd.exe") ' Look for "CMD"
    Call ShowWindow(WindowToFind&, SW_HIDE)
    Call SendMessageLong(WindowToFind&, WM_CLOSE, 0&, 0&)
End If

If Form1.chkDisablePA(3).Value = 1 Then ' Administrative Tools
    WindowToFind& = FindWindow("CabinetWClass", "Administrative Tools")
    Call ShowWindow(WindowToFind&, SW_HIDE)
    Call SendMessageLong(WindowToFind&, WM_CLOSE, 0&, 0&)
End If

If Form1.chkDisablePA(4).Value = 1 Then
ParentWin& = FindWindow("RunDLL", RunDll) 'System Properties
ChildWin& = FindWindowEx(ParentWin&, 0&, "#32770", "System Properties")
    Call ShowWindow(ChildWin&, SW_HIDE)
    Call SendMessageLong(ChildWin&, WM_CLOSE, 0&, 0&)
End If

If Form1.chkDisablePA(5).Value = 1 Then
ParentWin& = FindWindow("RunDLL", RunDll) ' Display Properties
ChildWin& = FindWindowEx(ParentWin&, 0&, "#32770", "Display Properties")
    Call ShowWindow(ChildWin&, SW_HIDE)
    Call SendMessageLong(ChildWin&, WM_CLOSE, 0&, 0&)
End If

If Form1.chkDisablePA(6).Value = 1 Then 'Folder Options
ParentWin& = FindWindow("MSGlobalFolderOptionsStub", "Folder Options")
ChildWin& = FindWindowEx(ParentWin&, 0&, "#32770", "Folder Options")
    Call ShowWindow(ChildWin&, SW_HIDE)
    Call SendMessageLong(ChildWin&, WM_CLOSE, 0&, 0&)
End If

If Form1.chkDisablePA(7).Value = 1 Then
ParentWin& = FindWindow("RunDLL", RunDll) 'Internet Properties
ChildWin& = FindWindowEx(ParentWin&, 0&, "#32770", "Internet Properties")
    Call ShowWindow(ChildWin&, SW_HIDE)
    Call SendMessageLong(ChildWin&, WM_CLOSE, 0&, 0&)
End If

If Form1.chkDisablePA(8).Value = 1 Then 'Taskbar and Start Menu Properties
ParentWin& = FindWindow("Static", "Taskbar and Start Menu Properties")
ChildWin& = FindWindowEx(ParentWin&, 0&, "#32770", "Taskbar and Start Menu Properties")
    Call ShowWindow(ChildWin&, SW_HIDE)
    Call SendMessageLong(ChildWin&, WM_CLOSE, 0&, 0&)
End If

If Form1.chkDisablePA(9).Value = 1 Then 'User Accounts
    WindowToFind& = FindWindow("HTML Application Host Window Class", "User Accounts")
    Call ShowWindow(WindowToFind&, SW_HIDE)
    Call SendMessageLong(WindowToFind&, WM_CLOSE, 0&, 0&)
End If

If Form1.chkDisablePA(10).Value = 1 Then 'User Accounts2
ParentWin& = FindWindow("RunDLL", RunDll)
ChildWin& = FindWindowEx(ParentWin&, 0&, "#32770", "User Accounts")
    Call ShowWindow(ChildWin&, SW_HIDE)
    Call SendMessageLong(ChildWin&, WM_CLOSE, 0&, 0&)
End If

If Form1.chkDisablePA(11).Value = 1 Then '"Add or Remove Programs"
WindowToFind& = FindWindow("NativeHWNDHost", "Add or Remove Programs")
    Call ShowWindow(WindowToFind&, SW_HIDE)
    Call SendMessageLong(WindowToFind&, WM_CLOSE, 0&, 0&)
End If

If Form1.chkDisablePA(12).Value = 1 Then
ParentWin& = FindWindow("RunDLL", RunDll)
ChildWin& = FindWindowEx(ParentWin&, 0&, "#32770", "Automatic Updates")
    Call ShowWindow(ChildWin&, SW_HIDE)
    Call SendMessageLong(ChildWin&, WM_CLOSE, 0&, 0&)
End If

If Form1.chkDisablePA(13).Value = 1 Then
WindowToFind& = FindWindow("wscui_class", "Windows Security Center")
    Call ShowWindow(WindowToFind&, SW_HIDE)
    Call SendMessageLong(WindowToFind&, WM_CLOSE, 0&, 0&)
End If

If Form1.chkDisablePA(14).Value = 1 Then
ParentWin& = FindWindow("RunDLL", RunDll)
ChildWin& = FindWindowEx(ParentWin&, 0&, "#32770", "Windows Firewall")
    Call ShowWindow(ChildWin&, SW_HIDE)
    Call SendMessageLong(ChildWin&, WM_CLOSE, 0&, 0&)
End If


If Form1.chkDisablePA(15).Value = 1 Then
DoEvents
WindowToFind& = FindWindow("#32770", "Windows Task Manager")
Call ShowWindow(WindowToFind&, SW_HIDE)
Call SendMessageLong(WindowToFind&, WM_CLOSE, 0&, 0&)
End If
End Sub
