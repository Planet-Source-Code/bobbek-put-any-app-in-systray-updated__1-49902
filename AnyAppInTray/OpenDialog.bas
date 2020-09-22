Attribute VB_Name = "OpenDialog"
Option Explicit

Private Type OPENFILENAME
  lStructSize         As Long
  hwndOwner           As Long
  hInstance           As Long
  lpstrFilter         As String
  lpstrCustomFilter   As String
  nMaxCustFilter      As Long
  nFilterIndex        As Long
  lpstrFile           As String
  nMaxFile            As Long
  lpstrFileTitle      As String
  nMaxFileTitle       As Long
  lpstrInitialDir     As String
  lpstrTitle          As String
  flags               As Long
  nFileOffset         As Integer
  nFileExtension      As Integer
  lpstrDefExt         As String
  lCustData           As Long
  lpfnHook            As Long
  lpTemplateName      As String
End Type
Private Declare Function GetOpenFileName Lib "comdlg32.dll" Alias "GetOpenFileNameA" (pOpenfilename As OPENFILENAME) As Long

Public Sub ChooseIcon()

  Dim LoadName As String

  With Form1
    LoadName = ShowOpen(.hWnd, "Icon Files (*.ico)" & vbNullChar & "*.ico", "Select New Icon", False)
    If LoadName = vbNullString Then
      Exit Sub
    End If
    .picNewTrayIco.Picture = LoadPicture(LoadName)
    .Icon = .picNewTrayIco.Picture
    .picTrayIco.Picture = .picNewTrayIco.Picture
    ShellTrayChangeIcon (.txtTip.Text)
    SaveSetting "AnyApp", "ICON", "Path", LoadName
  End With

End Sub

Public Function ShowOpen(ByVal AppHwnd As Long, _
                         ByVal Filter As String, _
                         ByVal Title As String, _
                         Optional ByVal Multiple As Boolean) As String

  Dim OpenF As OPENFILENAME

  On Error GoTo ErrorLoc
  OpenF.flags = &H4 ' no open as readonly box
  If Multiple Then
    OpenF.hwndOwner = AppHwnd 'set the window handle
  End If
  With OpenF
    .lpstrFile = String$(500, vbNullChar)
    .lpstrFileTitle = String$(500, vbNullChar)
    .lpstrFilter = Filter
    .lpstrTitle = Title
    .lStructSize = Len(OpenF)
    .nMaxFile = 501
    .nMaxFileTitle = 501
  End With 'OpenF
  If GetOpenFileName(OpenF) Then
    ShowOpen = Replace$(OpenF.lpstrFile, vbNullChar, vbNullString)
   Else
ErrorLoc:
    ShowOpen = vbNullString 'No file error
  End If

End Function
