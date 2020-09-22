Attribute VB_Name = "SetMenuIconModule"
Option Explicit
Private Const WM_GETICON      As Long = &H7F
Private hAppIcon              As Long
Private Declare Function GetMenu Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function GetSubMenu Lib "user32" (ByVal hMenu As Long, _
                                                  ByVal nPos As Long) As Long
Private Declare Function GetMenuItemID Lib "user32" (ByVal hMenu As Long, _
                                                     ByVal nPos As Long) As Long
Private Declare Function SetMenuItemBitmaps Lib "user32" (ByVal hMenu As Long, _
                                                          ByVal nPosition As Long, _
                                                          ByVal wFlags As Long, _
                                                          ByVal hBitmapUnchecked As Long, _
                                                          ByVal hBitmapChecked As Long) As Long

Public Sub SetMenuIcon(ByVal FrmHwnd As Long, _
                       ByVal FromApphWnd As Long, _
                       ByVal MainMenuNumber As Long, _
                       ByVal MenuItemNumber As Long, _
                       ByVal flags As Long)
  
  Dim lngMenu       As Long
  Dim lngSubMenu    As Long
  Dim lngMenuItemID As Long

  On Error Resume Next
  'hAppIcon = SendMessage(FromApphWnd, WM_GETICON, 0, 0)
  hAppIcon = Form1.Icon.Handle
  lngMenu = GetMenu(FrmHwnd)
  lngSubMenu = GetSubMenu(lngMenu, MainMenuNumber)
  lngMenuItemID = GetMenuItemID(lngSubMenu, MenuItemNumber)
  SetMenuItemBitmaps lngMenu, lngMenuItemID, flags, hAppIcon, hAppIcon
  On Error GoTo 0
End Sub
