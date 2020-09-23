Attribute VB_Name = "MTrayBar"
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''        By: Peptido
''      Date: Dec 21 1999
''
''   Purpose: Showing an icon in the system traybar
''
''   Functions:
''
''    ChangeTraybarIcon: Changes the icon currently showing in the traybar.
''        Parameters: NewIcon: Here you must pass a picture object
''
''    ChangeTraybarTip: Changes the tooltip of the Traybar Icon
''        Parameters: Tip: Pass the string you want to show as a tooltip
''
''    FormMessages: This procedure receives windows messages. Here you will
''        put the code you want to run when the user interacts with the traybar
''        icon.
''
''    PutInTraybar: This procedure is the one that actually shows the icon.
''        Parameters: Tip: String you want to show as a tooltip
''                    hWnd: Handle to the form that owns the icon
''                    Icon: Icon you want to show in the traybar
''
''    RemoveFromTraybar: Removes the icon from the traybar. You must call this
''        before exiting, or you will get a GPF
''
''
''    Known Bugs: None
''
''    Reamrks: Using this module, you can't show more than one icon from within
''             the same application. I'll publish one soon that lets you do it.
''
''             In Runtime, never stop your application using the Stop button, or
''             VB will crash. If you have to stop, call RemoveFromTraybar from the
''             Immediate Window
''
''
''    Please send any comments, suggestions or bug reports to:
''        peptido@insideo.com.ar
''


'Functions to Subclass a window
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

'Function to use the System Traybar
Private Declare Function Shell_NotifyIcon Lib "shell32.dll" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, lpData As NOTIFYICONDATA) As Long


'Traybar Icon Structures
Private Type NOTIFYICONDATA
  cbSize As Long
  hWnd As Long
  uID As Long
  uFlags As Long
  uCallbackMessage As Long
  hIcon As Long
  szTip As String * 64
End Type


'Window Subclassing Constants
Private Const GWL_WNDPROC = (-4)

'Traybar Constants
Private Const NIF_MESSAGE = &H1
Private Const NIF_ICON = &H2
Private Const NIF_TIP = &H4
Private Const NIM_ADD = &H0
Private Const NIM_DELETE = &H2
Private Const NIM_MODIFY = &H1
Private Const WM_TRAYBAR = 5130

'Mouse Messages Constants
Private Const WM_MBUTTONUP = &H208
Private Const WM_MBUTTONDOWN = &H207
Private Const WM_LBUTTONDOWN = &H201
Private Const WM_LBUTTONUP = &H202
Private Const WM_LBUTTONDBLCLK = &H203
Private Const WM_MOUSEMOVE = &H200
Private Const WM_RBUTTONDOWN = &H204
Private Const WM_RBUTTONUP = &H205

Private MyhWnd As Long
Private OldSubClassProc As Long
Private iData As NOTIFYICONDATA

Public Sub PutInTraybar(Tip As String, hWnd As Long, Icon As Long)

Call SubClassWindow(hWnd)

MyhWnd = hWnd

With iData
  .cbSize = Len(iData)
  .hWnd = MyhWnd
  .uID = 9999
  .uFlags = NIF_ICON + NIF_TIP + NIF_MESSAGE
  .uCallbackMessage = WM_TRAYBAR
  .hIcon = Icon
  .szTip = Tip & Chr$(0)
End With

Call Shell_NotifyIcon(NIM_ADD, iData)

End Sub

Public Sub RemoveFromTraybar()
  Call Shell_NotifyIcon(NIM_DELETE, iData)
  UnSubClassWindow (MyhWnd)
End Sub

Public Sub ChangeTrayBarTip(Tip As String)
  iData.szTip = Tip & Chr$(0)
  Call Shell_NotifyIcon(NIM_MODIFY, iData)
End Sub

Public Function FormMessages(ByVal hWnd As Long, ByVal iMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

If iMsg = WM_TRAYBAR Then
  Select Case lParam
    Case WM_MOUSEMOVE
      'Write here the code you want to run when the mouse is over your icon
    Case WM_LBUTTONDOWN
      'Write here the code you want to run when the Left Mouse Button is pressed
    Case WM_LBUTTONUP
      'Write here the code you want to run when the Left Mouse Button is released (click)
    Case WM_LBUTTONDBLCLK
      'Write here the code you want to run when there has been a Double Click
    Case WM_RBUTTONDOWN
      'Write here the code you want to run when the Right Mouse Button is pressed
    Case WM_RBUTTONUP
      'Write here the code you want to run when the Right Mouse Button is released
    Case WM_MBUTTONDOWN
      'Write here the code you want to run when the Middle Mouse Button is pressed
    Case WM_MBUTTONUP
      'Write here the code you want to run when the Middle Mouse Button is released
    Case Else
      Exit Function
  End Select
End If

FormMessages = CallWindowProc(OldSubClassProc, hWnd, iMsg, wParam, ByVal lParam)

End Function

Private Function SubClassWindow(hWnd As Long) As Long
Dim ProcOld As Long

ProcOld = SetWindowLong(hWnd, GWL_WNDPROC, AddressOf FormMessages)

OldSubClassProc = ProcOld
SubClassWindow = ProcOld

End Function

Private Sub UnSubClassWindow(hWnd As Long)
  Call SetWindowLong(hWnd, GWL_WNDPROC, OldSubClassProc)
End Sub

Public Sub ChangeTrayBarIcon(NewIcon As Long)
  iData.hIcon = NewIcon
  Call Shell_NotifyIcon(NIM_MODIFY, iData)
End Sub

