Attribute VB_Name = "TrayIcon"
Option Explicit
Private Const APP_SYSTRAY_ID               As Integer = 999 'unique identifier
Private Const NOTIFYICON_VERSION           As Long = &H3
Private Const NIF_MESSAGE                  As Long = &H1
Private Const NIF_ICON                     As Long = &H2
Private Const NIF_TIP                      As Long = &H4
Private Const NIF_STATE                    As Long = &H8
Private Const NIF_INFO                     As Long = &H10
Private Const NIF_DOALL                    As Double = NIF_MESSAGE Or NIF_ICON Or NIF_TIP
Private Const NIM_ADD                      As Long = &H0
Private Const NIM_MODIFY                   As Long = &H1
Private Const NIM_DELETE                   As Long = &H2
Private Const NIM_SETFOCUS                 As Long = &H3
Private Const NIM_SETVERSION               As Long = &H4
Private Const NIM_VERSION                  As Long = &H5
Private Const NIS_HIDDEN                   As Long = &H1
Private Const NIS_SHAREDICON               As Long = &H2
'icon flags
Private Const NIIF_NONE                    As Long = &H0
Private Const NIIF_INFO                    As Long = &H1
Private Const NIIF_WARNING                 As Long = &H2
Private Const NIIF_ERROR                   As Long = &H3
Private Const NIIF_GUID                    As Long = &H5
Private Const NIIF_ICON_MASK               As Long = &HF
Private Const NIIF_NOSOUND                 As Long = &H10
Private Const WM_USER                      As Long = &H400
Private Const NIN_BALLOONSHOW              As Double = (WM_USER + 2)
Private Const NIN_BALLOONHIDE              As Double = (WM_USER + 3)
Private Const NIN_BALLOONTIMEOUT           As Double = (WM_USER + 4)
Private Const NIN_BALLOONUSERCLICK         As Double = (WM_USER + 5)
'Icon
Private Const BALOON_ICO_NONE              As Integer = 0
Public Const BALOON_ICO_INFO               As Integer = 1
Private Const BALOON_ICO_WARNING           As Integer = 2
Private Const BALOON_ICO_ERROR             As Integer = 3
'shell version / NOTIFIYICONDATA struct size constants
Private Const NOTIFYICONDATA_V1_SIZE       As Long = 88      'pre-5.0 structure size
Private Const NOTIFYICONDATA_V2_SIZE       As Long = 488     'pre-6.0 structure size
Private Const NOTIFYICONDATA_V3_SIZE       As Long = 504     '6.0+ structure size
Private NOTIFYICONDATA_SIZE                As Long
Private Type GUID
  Data1                                    As Long
  Data2                                    As Integer
  Data3                                    As Integer
  Data4(7)                                 As Byte
End Type
Private Type NOTIFYICONDATA
  cbSize                                   As Long
  hWnd                                     As Long
  uID                                      As Long
  uFlags                                   As Long
  uCallbackMessage                         As Long
  hIcon                                    As Long
  szTip                                    As String * 128
  dwState                                  As Long
  dwStateMask                              As Long
  szInfo                                   As String * 256
  uTimeoutAndVersion                       As Long
  szInfoTitle                              As String * 64
  dwInfoFlags                              As Long
  guidItem                                 As GUID
End Type
Public Const WM_LBUTTONUP                  As Long = &H202   '514
Public Const WM_RBUTTONUP                  As Long = &H205   '517
Private Const WM_MOUSEMOVE                 As Long = &H200   '512
Private Declare Function Shell_NotifyIcon Lib "shell32.dll" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, _
                                                                                       lpData As NOTIFYICONDATA) As Long
Private Declare Function GetFileVersionInfoSize Lib "version.dll" Alias "GetFileVersionInfoSizeA" (ByVal lptstrFilename As String, _
                                                                                                   lpdwHandle As Long) As Long
Private Declare Function GetFileVersionInfo Lib "version.dll" Alias "GetFileVersionInfoA" (ByVal lptstrFilename As String, _
                                                                                           ByVal dwHandle As Long, _
                                                                                           ByVal dwLen As Long, _
                                                                                           lpData As Any) As Long
Private Declare Function VerQueryValue Lib "version.dll" Alias "VerQueryValueA" (pBlock As Any, _
                                                                                 ByVal lpSubBlock As String, _
                                                                                 lpBuffer As Any, _
                                                                                 nVerSize As Long) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, _
                                                                     Source As Any, _
                                                                     ByVal Length As Long)

Private Function IsShellVersion(ByVal version As Long) As Boolean

  'returns True if the Shell version
  '(shell32.dll) is equal or later than
  'the value passed as 'version'
  
  Dim nBufferSize As Long
  Dim nUnused     As Long
  Dim lpBuffer    As Long
  Dim nVerMajor   As Integer
  Dim bBuffer()   As Byte
  Const sDLLFile  As String = "shell32.dll"

  nBufferSize = GetFileVersionInfoSize(sDLLFile, nUnused)
  If nBufferSize > 0 Then
    ReDim bBuffer(nBufferSize - 1) As Byte
    Call GetFileVersionInfo(sDLLFile, 0&, nBufferSize, bBuffer(0))
    If VerQueryValue(bBuffer(0), "\", lpBuffer, nUnused) = 1 Then
      CopyMemory nVerMajor, ByVal lpBuffer + 10, 2
      IsShellVersion = nVerMajor >= version
    End If  'VerQueryValue
  End If  'nBufferSize

End Function

Private Sub SetShellVersion()

  Select Case True
   Case IsShellVersion(6)
    NOTIFYICONDATA_SIZE = NOTIFYICONDATA_V3_SIZE '6.0+ structure size
   Case IsShellVersion(5)
    NOTIFYICONDATA_SIZE = NOTIFYICONDATA_V2_SIZE 'pre-6.0 structure size
   Case Else
    NOTIFYICONDATA_SIZE = NOTIFYICONDATA_V1_SIZE 'pre-5.0 structure size
  End Select

End Sub

Public Sub ShellTrayAdd(ByVal strTip As String)
  
  Dim nid As NOTIFYICONDATA

  If NOTIFYICONDATA_SIZE = 0 Then
    SetShellVersion
  End If
  'set up the type members
  With nid
    .cbSize = NOTIFYICONDATA_SIZE
    .hWnd = Form1.picTrayIco.hWnd
    .uID = APP_SYSTRAY_ID
    .uFlags = NIF_MESSAGE Or NIF_ICON Or NIF_TIP
    .dwState = NIS_SHAREDICON
    .hIcon = Form1.picTrayIco.Picture
    .uCallbackMessage = WM_MOUSEMOVE '###############
    'szTip is the tooltip shown when the
    'mouse hovers over the systray icon.
    'Terminate it since the strings are
    'fixed-length in NOTIFYICONDATA.
    .szTip = strTip & vbNullChar
    .uTimeoutAndVersion = NOTIFYICON_VERSION
  End With
  'add the icon ...
  Call Shell_NotifyIcon(NIM_ADD, nid)
  '... and inform the system of the
  'NOTIFYICON version in use
  Call Shell_NotifyIcon(NIM_SETVERSION, nid)

End Sub

Public Sub ShellTrayChangeIcon(ByVal strTip As String)

  Dim nid As NOTIFYICONDATA

  With nid
    .cbSize = NOTIFYICONDATA_SIZE
    .hWnd = Form1.picTrayIco.hWnd
    .uID = APP_SYSTRAY_ID
    .uFlags = NIF_DOALL
    .dwState = NIS_SHAREDICON
    .hIcon = Form1.picTrayIco.Picture
    .uCallbackMessage = WM_MOUSEMOVE
    .szTip = strTip & vbNullChar
    .uTimeoutAndVersion = NOTIFYICON_VERSION
  End With
  Call Shell_NotifyIcon(NIM_MODIFY, nid)

End Sub

Public Sub ShellTrayModifyTip(ByVal nIconIndex As Long, _
                              ByVal strTitle As String, _
                              ByVal strMess As String)
                              
  Dim nid As NOTIFYICONDATA

  If NOTIFYICONDATA_SIZE = 0 Then
    SetShellVersion
  End If
  With nid
    .cbSize = NOTIFYICONDATA_SIZE
    .hWnd = Form1.picTrayIco.hWnd
    .uID = APP_SYSTRAY_ID
    .uFlags = NIF_INFO
    .dwInfoFlags = nIconIndex
    'InfoTitle is the balloon tip title,
    'and szInfo is the message displayed.
    'Terminating both with vbNullChar prevents
    'the display of the unused padding in the
    'strings defined as fixed-length in NOTIFYICONDATA.
    .szInfoTitle = strTitle & vbNullChar
    .szInfo = strMess & vbNullChar
  End With
  Call Shell_NotifyIcon(NIM_MODIFY, nid)

End Sub

Public Sub ShellTrayRemove()

  Dim nid As NOTIFYICONDATA

  If NOTIFYICONDATA_SIZE = 0 Then
    SetShellVersion
  End If
  With nid
    .cbSize = NOTIFYICONDATA_SIZE
    .hWnd = Form1.picTrayIco.hWnd
    .uID = APP_SYSTRAY_ID
  End With
  Call Shell_NotifyIcon(NIM_DELETE, nid)

End Sub
