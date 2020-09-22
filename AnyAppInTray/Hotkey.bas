Attribute VB_Name = "HotKey"
Option Explicit
Private Const WM_HOTKEY        As Long = &H312
Private Const GWL_WNDPROC      As Integer = -4
Private glWinRet               As Long
Private retVal(6 To 12)        As Boolean
Public nn                      As Integer
Private Declare Function RegisterHotkey Lib "user32" Alias "RegisterHotKey" (ByVal hWnd As Long, _
                                                                             ByVal ID As Long, _
                                                                             ByVal fsModifiers As Long, _
                                                                             ByVal vk As Long) As Long
Public Declare Function UnregisterHotKey Lib "user32" (ByVal hWnd As Long, _
                                                       ByVal ID As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, _
                                                                            ByVal nIndex As Long, _
                                                                            ByVal dwNewLong As Long) As Long
Private Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, _
                                                                              ByVal hWnd As Long, _
                                                                              ByVal msg As Long, _
                                                                              ByVal wParam As Long, _
                                                                              ByVal lParam As Long) As Long

Public Function CallbackMsgs(ByVal wHwnd As Long, _
                             ByVal wMsg As Long, _
                             ByVal wp_id As Long, _
                             ByVal lp_id As Long) As Long

  If wMsg = WM_HOTKEY Then
    Call DoFunctions(wp_id)
    CallbackMsgs = 1
    Exit Function
  End If
  CallbackMsgs = CallWindowProc(glWinRet, wHwnd, wMsg, wp_id, lp_id)

End Function

Public Sub DoFunctions(ByVal vKeyID As Byte)

  ' Sub : DoFunction
  ' Activated by the Function "CallbackMsgs()" whenever
  ' a hotkey is pressed.
  ' Important Notes :
  ' Do not include any msgboxes or Modal forms in
  ' this procedure, else if you include then by
  ' pressing the Hotkey twice/thrice the application
  ' will be terminated abnormally.
  '
  ' But if it is a requirement for you to include the
  ' Modal forms or msgbox in this procedure then put
  ' the RegisterHotKey() API before hiding the Form
  ' and put the UnRegisterHotKey() API before Showing
  ' the form.

  DoEvents
  ' When the Hotkey is pressed once
  ' check if the Dofunctions() has completed
  ' before the CallbackMsgs().
  ' This check is not required if the form is
  ' minimized in the SysTray ...
  Debug.Print "vkey id" & vKeyID
  '<:-):SUGGESTION: Active Debug should be removed from final code.
  If vKeyID = 0 Or vKeyID = 1 Or vKeyID = 2 Or vKeyID = 3 Then
    Call WindowSPY
  End If

End Sub

Public Sub RegHotkey()
  Dim arr() As String

  arr() = Split("&H75, &H76, &H77, &H78, &H79, &H7A, &H7B", ",")
  For nn = 12 To 6 Step -1
    retVal(nn) = RegisterHotkey(Form1.hWnd, 2, 0, arr(nn - 6))
    If retVal(nn) Then
      Exit For
    End If
  Next nn
  Form1.Caption = "Press Hotkey F" & nn & " To Hide A Window!"
  'Subclassing the form to get the Windows callback msgs.
  glWinRet = SetWindowLong(Form1.hWnd, GWL_WNDPROC, AddressOf CallbackMsgs)

End Sub
