VERSION 5.00
Begin VB.Form frmTrayIcon 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Tray Icon"
   ClientHeight    =   1095
   ClientLeft      =   45
   ClientTop       =   615
   ClientWidth     =   4755
   Icon            =   "frmTrayIcon.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1095
   ScaleWidth      =   4755
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.PictureBox picHook 
      Height          =   495
      Left            =   405
      ScaleHeight     =   435
      ScaleWidth      =   675
      TabIndex        =   0
      Top             =   210
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileHide 
         Caption         =   "&Hide"
      End
      Begin VB.Menu mnuFileShow 
         Caption         =   "&Show"
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "E&xit"
      End
   End
End
Attribute VB_Name = "frmTrayIcon"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False






'Do you have a beautiful wall paper but has got lot of desktop icons on it, only
'to destroy its beauty. Then this program will help you. On clicking the tray icon
'of this program all the desktop icons are transfered to a window and on
'clicking the tray icon again will restore all the icons. Put a short cut in the
'startup to load it everytime windows is loaded.



'Thanks for downloading my code,
'please visit my site for downloading free Quran MP3 player
'at http://www.qurantrans.f2s.com ,
'Mail me your comments at shafeekalfa@hotmail.com
'Thank you
'Shafeek Mohammed



' Constants for methods associated with tray icons
Private Const NIM_ADD = &H0
Private Const NIM_MODIFY = &H1
Private Const WM_MOUSEMOVE = &H200
Private Const NIF_MESSAGE = &H1
Private Const NIF_ICON = &H2
Private Const NIF_TIP = &H4
' Constants for determining event associated with tray icon
Private Const WM_LBUTTONDOWN = &H201
Private Const WM_RBUTTONDOWN = &H204

'Function used to add,delete or modify tray icon
Private Declare Function Shell_NotifyIcon Lib "shell32" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, pnid As NOTIFYICONDATA) As Boolean

'Before calling the above function this stucture has to be filled
Private Type NOTIFYICONDATA
    cbSize As Long
    hwnd As Long
    uId As Long
    uFlags As Long
    ucallbackMessage As Long
    hIcon As Long
    szTip As String * 64
End Type

'Declare a variable for the above stucture
Dim nidTray As NOTIFYICONDATA

'Change the parent of a control or object
Private Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
'Find a handle of particular window by name
Private Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long

Dim mblnToggle As Boolean 'use to remember the last state of desktop icons

Dim hwndDesktopIcons As Long ' to store handle to window of desktop icons
Dim hwndDesktop  As Long 'to store handle to window of desktop
Private Sub Form_Load()
    'Get the handle of Desktop
    hwndDesktop = GetDesktopWindow
    'Get the handle of Desktop Icons
    hwndDesktopIcons = FindWindowEx(0&, 0&, "Progman", vbNullString)
    'Fill the tray stucture
    With nidTray
        .cbSize = Len(nidTray)
        .hwnd = picHook.hwnd
        .uId = 1&
        .uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
        .ucallbackMessage = WM_MOUSEMOVE
        .hIcon = Me.Icon 'Icon of the tray icon
        .szTip = "Hide Desktop" & Chr$(0)  'Tooltip text
    End With
    Shell_NotifyIcon NIM_ADD, nidTray   'Add the item to tray
    
    'if you want to hide the desktop whenever the program starts uncomment this line
    HideDesktop
End Sub

Private Sub mnuFileExit_Click()
    ShowDesktop ' on exit restore the desktop
    End
End Sub

Private Sub mnuFileHide_Click()
    HideDesktop  'Transfer the desktop icons to Form1 and change tran icon
End Sub

Private Sub mnuFileShow_Click()
    ShowDesktop 'Restore the desktop icons and change tran icon
End Sub

Private Sub picHook_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
    Dim msg As Long
    ' Get the message ID
    msg = x / Screen.TwipsPerPixelX
    
    ' Determines if any of the following events have occured
    Select Case msg
        Case WM_LBUTTONDOWN: ' Left Button MouseDown
            If mblnToggle = False Then
                HideDesktop
            Else
                ShowDesktop
            End If
            mblnToggle = Not mblnToggle
        Case WM_RBUTTONDOWN: ' Right Button MouseDown
               PopupMenu mnuFile 'Show menu
    End Select

End Sub
Sub HideDesktop()
    SetParent hwndDesktopIcons, Form1.hwnd 'desktop icons transfered to Form1
    
    'modify the tray icon
    nidTray.cbSize = Len(nidTray)
    nidTray.hIcon = Form1.Icon
    nidTray.hwnd = picHook.hwnd
    nidTray.szTip = "Show Desktop" & Chr$(0)
    nidTray.uId = 1&
    Shell_NotifyIcon NIM_MODIFY, nidTray
    Form1.Visible = True
    Form1.WindowState = vbMinimized
End Sub
'
Public Sub ShowDesktop()
    'Restore the desktop icons to desktop
    SetParent hwndDesktopIcons, hwndDesktop
    
    'modify the tray icon
    nidTray.cbSize = Len(nidTray)
    nidTray.hIcon = Me.Icon
    nidTray.hwnd = picHook.hwnd
    nidTray.szTip = "Hide Desktop" & Chr$(0)
    nidTray.uId = 1&
    Shell_NotifyIcon NIM_MODIFY, nidTray
    
    Form1.Visible = False
End Sub
