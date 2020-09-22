VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "System Tray Manager"
   ClientHeight    =   2010
   ClientLeft      =   150
   ClientTop       =   720
   ClientWidth     =   3885
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2010
   ScaleWidth      =   3885
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtTip 
      Height          =   285
      Left            =   120
      TabIndex        =   5
      Text            =   "AnyApp"
      Top             =   1200
      Width           =   3675
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Change Tool Tip"
      Height          =   375
      Left            =   2280
      TabIndex        =   4
      Top             =   1560
      Width           =   1515
   End
   Begin VB.Frame Frame1 
      Caption         =   "Current Icon"
      Height          =   855
      Left            =   120
      TabIndex        =   2
      Top             =   0
      Width           =   1215
      Begin VB.PictureBox picNewTrayIco 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   240
         Left            =   360
         Picture         =   "Form1.frx":030A
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   3
         Top             =   240
         Width           =   240
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Change Tray Icon"
      Height          =   375
      Left            =   2280
      TabIndex        =   1
      Top             =   360
      Width           =   1515
   End
   Begin VB.PictureBox picTrayIco 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   4320
      Picture         =   "Form1.frx":0894
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   0
      Top             =   720
      Width           =   495
   End
   Begin VB.Label Label1 
      Caption         =   "Current Tool Tip:"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   960
      Width           =   1455
   End
   Begin VB.Menu mnuPopup 
      Caption         =   "mnuPopup"
      Visible         =   0   'False
      Begin VB.Menu mnuAbout 
         Caption         =   "About"
      End
      Begin VB.Menu mnuOpt 
         Caption         =   "Options"
      End
      Begin VB.Menu mnuSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHiddenWindow 
         Caption         =   ""
         Index           =   0
         Visible         =   0   'False
      End
      Begin VB.Menu mnuSep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      Begin VB.Menu mnuSndtry 
         Caption         =   "Send To Tray"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private bClose     As Boolean

'Thanks to "Peter H" for updating this program:
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

'Put a default icon path in 'GetSetting',
    'so it shows the icon before any icon is
    'selected by the user...

'I (Peter H) have also inserted
    'Option Explicit'
    'a hotkey and
    'removed the timer. (Was used for getting KeyAssyncState)
    'Email me and I will send it to you. (He did, and here it is)

'This update done by "Peter H"
'Original proggie by "Bobbek"

Private Sub Command1_Click()
  ChooseIcon
End Sub

Private Sub Command2_Click()
  SaveSetting "AnyApp", "TIP", "STRING", txtTip.Text
  ShellTrayChangeIcon (txtTip.Text)
End Sub

Private Sub Form_Load()
  Dim i As Integer

  On Error Resume Next
  RegHotkey
  RestrictedhWnd(1) = Me.hWnd 'Can't Hide Itself
  RestrictedhWnd(2) = FindWindow("Shell_TrayWnd", vbNullString) 'Don't want to Hide taskbar, do you?
  
  With picNewTrayIco
    .Picture = LoadPicture(GetSetting("AnyApp", "ICON", "Path", App.Path & "\flash.ico"))
    picTrayIco.Picture = .Picture
    Me.Icon = .Picture
  End With
  
  txtTip.Text = GetSetting("AnyApp", "TIP", "STRING")
  If txtTip.Text = vbNullString Then
    txtTip.Text = "AnyApp"
  End If
  
  ShellTrayAdd (txtTip.Text)
  
  For i = 1 To 11
    Load mnuHiddenWindow(i)
    mnuHiddenWindow(i).Visible = False
  Next i
  
  ShellTrayModifyTip BALOON_ICO_INFO, "AnyApp Instructions", "Press [F9] with mouse over window you want to hide" & vbNewLine & _
                                      "Right Click to bring up menu."
  On Error GoTo 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
  If bClose Then
    ShowApps
    ShellTrayRemove
    UnregisterHotKey Me.hWnd, (nn)
   Else
    Me.Hide
    Cancel = 1
  End If
End Sub

Private Sub mnuAbout_Click()
  MsgBox "AnyApp: Allows You to hide program windows quickly and painlessly" & vbNewLine & "Just Press F" & nn & " with the mouse over a window you want to hide", vbExclamation + vbOKOnly, "AnyApp - Information - General"
End Sub

Private Sub mnuExit_Click()
  bClose = True
  Unload Me
End Sub

Private Sub mnuHiddenWindow_Click(Index As Integer)
  ShowWindow HiddenWindows(Index), SW_SHOW
  mnuHiddenWindow(Index).Caption = vbNullString
  mnuHiddenWindow(Index).Visible = False
End Sub

Private Sub mnuOpt_Click()
  Form1.Visible = True
End Sub

Private Sub mnuSndtry_Click()
  Me.Hide
End Sub

Private Sub picTrayIco_MouseMove(Button As Integer, _
                                 Shift As Integer, _
                                 X As Single, _
                                 Y As Single)
  Dim butt   As Long
 
  butt = X / Screen.TwipsPerPixelX
  Select Case butt
   Case WM_LBUTTONUP
    Me.Show ' ShellTrayModifyTip BALOON_ICO_INFO, "Hello There!", "Try Clicking RIGHT Button..."
   Case WM_RBUTTONUP
    SetForegroundWindow Me.hWnd  ' To make popup go away when clicked somewherelse
    PopupMenu mnuPopup
    'Show Baloon
   Case Else
    Exit Sub
  End Select
End Sub
