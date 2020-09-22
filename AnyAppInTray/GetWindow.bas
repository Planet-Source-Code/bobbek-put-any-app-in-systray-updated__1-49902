Attribute VB_Name = "Show_Hide"
Option Explicit
Public Type POINTAPI
  X                                As Long
  Y                                As Long
End Type
Private Const GWL_STYLE            As Long = (-16)
Private Const GWW_HINSTANCE        As Long = (-6)
Private Const GWW_ID               As Long = (-12)
'---------- FOR CLOSING A WINDOW ------------
Private Const SW_HIDE              As Integer = 0
Private Const SW_MAXIMIZE          As Integer = 3
Public Const SW_SHOW               As Integer = 5
Private Const SW_MINIMIZE          As Integer = 6
Private Const SW_RESTORE           As Integer = 9
Public RestrictedhWnd(1 To 5)      As Long
Private hwC                        As Long
Public HiddenWindows(0 To 11)      As Long
Private WndoText                   As String
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hWnd As Long, _
                                                                            ByVal lpString As String, _
                                                                            ByVal cch As Long) As Long
Private Declare Function WindowFromPointXY Lib "user32" Alias "WindowFromPoint" (ByVal xPoint As Long, _
                                                                                 ByVal yPoint As Long) As Long
Private Declare Function GetParent Lib "user32" (ByVal hWnd As Long) As Long
Public Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, _
                                                                     ByVal lpWindowName As String) As Long
Public Declare Function SetForegroundWindow Lib "user32" (ByVal hWnd As Long) As Long
Public Declare Function ShowWindow Lib "user32.dll" (ByVal hWnd As Long, _
                                                     ByVal nCmdShow As Long) As Long
Public Declare Function GetInputState Lib "user32" () As Long

Public Sub ShowApps()

  Dim i      As Integer
  Dim Result As Long
  
  With Form1
    For i = 0 To 10
      If LenB(.mnuHiddenWindow(i).Caption) Then
        Result = ShowWindow(HiddenWindows(i), SW_SHOW)
        .mnuHiddenWindow(i).Caption = vbNullString
        .mnuHiddenWindow(i).Visible = False
      End If
    Next i
  End With

End Sub

Public Sub WindowSPY()
  Dim pt32        As POINTAPI
  Dim ptx         As Long
  Dim pty         As Long
  Dim sWindowText As String * 100
  Dim hWndOver    As Long
  Dim r           As Long
  Dim OlhWnd      As Long
  Dim i           As Integer
  
  Call GetCursorPos(pt32)
  ptx = pt32.X
  pty = pt32.Y
  hWndOver = WindowFromPointXY(ptx, pty)
  Do
    OlhWnd = hWndOver
    hWndOver = GetParent(hWndOver)
    If GetInputState <> 0 Then DoEvents
  Loop Until hWndOver = 0
  hWndOver = OlhWnd
  r = GetWindowText(hWndOver, sWindowText, 100) 'TextLenght
  WndoText = Left$(sWindowText, r)
  'End If
  For i = 1 To 5
    If hWndOver = RestrictedhWnd(i) Then
      Select Case i
       Case 1
        MsgBox "I can't hide Myself!" & vbNewLine & "How you going to find me? lol.", vbInformation + vbOKOnly, "AnyApp - Warning - Can't Do That!"
       Case Else
        MsgBox "This is a RESTRICTED Window!" & vbNewLine & "Windows Wouldn't Be Too Happy If You Hid This Window.", vbInformation + vbOKOnly, "AnyApp - Warning - Can't Do That!"
      End Select
      Exit Sub
    End If
  Next i
  hwC = 0
  Do
    hwC = hwC + 1
  Loop Until hwC = 11 Or Form1.mnuHiddenWindow(hwC).Caption = vbNullString
  If hwC = 11 Then
    MsgBox "There Are Already 10 Hidden Windows!" & vbNewLine & "Can Not Hide Any More" & vbNewLine & "Right Click Tray Icon To Unhide Other Windows First." & vbNewLine & "Thankyou.", vbExclamation + vbOKOnly, "AnyApp - Warning - Too Many Hidden Windows!"
    Exit Sub
  End If
  ShowWindow hWndOver, SW_HIDE
  Form1.mnuHiddenWindow(hwC).Caption = WndoText
  Form1.mnuHiddenWindow(hwC).Visible = True
  HiddenWindows(hwC) = hWndOver
  SetMenuIcon Form1.hWnd, hWndOver, 1, hwC + 3, 0

End Sub
