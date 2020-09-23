VERSION 5.00
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   Caption         =   "Change Read-only Properties at Runtime"
   ClientHeight    =   3150
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5160
   LinkTopic       =   "Form1"
   ScaleHeight     =   3150
   ScaleWidth      =   5160
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      Caption         =   "Taskbar"
      Height          =   1215
      Left            =   2640
      TabIndex        =   11
      Top             =   120
      Width           =   2415
      Begin VB.OptionButton OpTaskbar 
         Caption         =   "Dont Show in Taskbar"
         Height          =   255
         Index           =   1
         Left            =   360
         TabIndex        =   13
         Top             =   720
         Width           =   1935
      End
      Begin VB.OptionButton OpTaskbar 
         Caption         =   "Show in Taskbar"
         Height          =   255
         Index           =   0
         Left            =   360
         TabIndex        =   12
         Top             =   360
         Value           =   -1  'True
         Width           =   1935
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Borderstyle"
      Height          =   2895
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   2415
      Begin VB.OptionButton OptBorderstyle 
         Caption         =   "Sizable ToolWindow"
         Height          =   255
         Index           =   5
         Left            =   360
         TabIndex        =   10
         Top             =   2160
         Width           =   1935
      End
      Begin VB.OptionButton OptBorderstyle 
         Caption         =   "Fixed ToolWindow"
         Height          =   255
         Index           =   4
         Left            =   360
         TabIndex        =   9
         Top             =   1800
         Width           =   1935
      End
      Begin VB.OptionButton OptBorderstyle 
         Caption         =   "Fixed Dialog"
         Height          =   255
         Index           =   3
         Left            =   360
         TabIndex        =   8
         Top             =   1440
         Width           =   1935
      End
      Begin VB.OptionButton OptBorderstyle 
         Caption         =   " Sizable"
         Height          =   255
         Index           =   2
         Left            =   360
         TabIndex        =   7
         Top             =   1080
         Value           =   -1  'True
         Width           =   1935
      End
      Begin VB.OptionButton OptBorderstyle 
         Caption         =   " Fixed Single"
         Height          =   255
         Index           =   1
         Left            =   360
         TabIndex        =   6
         Top             =   720
         Width           =   1935
      End
      Begin VB.OptionButton OptBorderstyle 
         Caption         =   "None"
         Height          =   255
         Index           =   0
         Left            =   360
         TabIndex        =   5
         Top             =   360
         Width           =   1935
      End
   End
   Begin VB.CheckBox Check3 
      Caption         =   "Control Box"
      Height          =   255
      Left            =   3000
      TabIndex        =   3
      Top             =   2160
      Value           =   1  'Checked
      Width           =   1575
   End
   Begin VB.CheckBox Check2 
      Caption         =   "Maximise button"
      Height          =   255
      Left            =   3000
      TabIndex        =   2
      Top             =   1800
      Value           =   1  'Checked
      Width           =   1575
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Minimise button"
      Height          =   255
      Left            =   3000
      TabIndex        =   1
      Top             =   1440
      Value           =   1  'Checked
      Width           =   1575
   End
   Begin VB.CommandButton cmdReset 
      Caption         =   "Reset"
      Height          =   375
      Left            =   2640
      TabIndex        =   0
      Top             =   2640
      Width           =   2415
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function LockWindowUpdate Lib "user32" (ByVal hwndLock As Long) As Long
Private Declare Function GetDesktopWindow Lib "user32" () As Long
'API for handling Window styles
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
'Window style constants
Private Const WS_CAPTION = &HC00000
Private Const WS_SYSMENU = &H80000
Private Const WS_MAXIMIZEBOX = &H10000
Private Const WS_MINIMIZEBOX = &H20000
Private Const WS_EX_APPWINDOW = &H40000
Private Const WS_EX_DLGMODALFRAME = &H1&
Private Const WS_EX_TOOLWINDOW = &H80&
Private Const WS_EX_WINDOWEDGE = &H100&
Private Const GWL_EXSTYLE = (-20)
Private Const GWL_STYLE = (-16)
'variables to hold original Window style
Dim StartStyle As Long
Dim StartExStyle As Long


Private Sub Check1_Click()
    'minimise box
    If Check1.Value = 1 Then
        SetWindowLong Me.hwnd, GWL_STYLE, GetWindowLong(Me.hwnd, GWL_STYLE) Or (WS_MINIMIZEBOX)
    Else
        SetWindowLong Me.hwnd, GWL_STYLE, GetWindowLong(Me.hwnd, GWL_STYLE) And (Not WS_MINIMIZEBOX)
    End If
    'force repaint via resizing up, then down one pixel
    Me.Height = Me.Height - 15
    Me.Height = Me.Height + 15
End Sub
Private Sub Check2_Click()
    'maximise box
    If Check2.Value = 1 Then
        SetWindowLong Me.hwnd, GWL_STYLE, GetWindowLong(Me.hwnd, GWL_STYLE) Or (WS_MAXIMIZEBOX)
    Else
        SetWindowLong Me.hwnd, GWL_STYLE, GetWindowLong(Me.hwnd, GWL_STYLE) And (Not WS_MAXIMIZEBOX)
    End If
    'force repaint via resizing up, then down one pixel
    Me.Height = Me.Height - 15
    Me.Height = Me.Height + 15
End Sub

Private Sub Check3_Click()
    'controlbox
    If Check3.Value = 1 Then
        SetWindowLong Me.hwnd, GWL_STYLE, GetWindowLong(Me.hwnd, GWL_STYLE) Or (WS_SYSMENU)
    Else
        SetWindowLong Me.hwnd, GWL_STYLE, GetWindowLong(Me.hwnd, GWL_STYLE) And (Not WS_SYSMENU)
    End If
    'force repaint via resizing up, then down one pixel
    Me.Height = Me.Height - 15
    Me.Height = Me.Height + 15
End Sub

Private Sub cmdReset_Click()
    Check1.Value = 1
    Check2.Value = 1
    Check3.Value = 1
    OpTaskbar(0).Value = True
    OptBorderstyle(2).Value = True
End Sub

Private Sub Form_Load()
    'save original style so we can reset
    StartStyle = GetWindowLong(Me.hwnd, GWL_STYLE)
    StartExStyle = GetWindowLong(Me.hwnd, GWL_EXSTYLE)
    
End Sub

Private Sub OpTaskbar_Click(Index As Integer)
    'show/hide taskbar - causes slight screen flash
    Me.Visible = False 'removes existing item from taskbar
    If Index = 0 Then
        SetWindowLong Me.hwnd, GWL_EXSTYLE, GetWindowLong(Me.hwnd, GWL_EXSTYLE) Or WS_EX_APPWINDOW
    Else
        SetWindowLong Me.hwnd, GWL_EXSTYLE, GetWindowLong(Me.hwnd, GWL_EXSTYLE) And (Not WS_EX_APPWINDOW)
    End If
    Me.Visible = True
End Sub

Private Sub OptBorderstyle_Click(Index As Integer)
    SetWindowLong Me.hwnd, GWL_EXSTYLE, StartExStyle
    SetWindowLong Me.hwnd, GWL_STYLE, StartStyle
    OpTaskbar(0).Value = True
    Check3.Value = 1
    Select Case Index
        Case 0 'none
            SetWindowLong Me.hwnd, GWL_STYLE, WS_SYSMENU
        Case 1 'fixed single
            Check1.Value = 0
            Check2.Value = 0
            SetWindowLong Me.hwnd, GWL_STYLE, (WS_CAPTION Or WS_SYSMENU)
        Case 2 'sizable
            Check1.Value = 1
            Check2.Value = 1
        Case 3 'dialog
            SetWindowLong Me.hwnd, GWL_STYLE, (WS_CAPTION Or WS_SYSMENU)
            SetWindowLong Me.hwnd, GWL_EXSTYLE, WS_EX_DLGMODALFRAME 'And (Not WS_EX_APPWINDOW)
            OpTaskbar(1).Value = True
        Case 4 'fixed toolwindow
            SetWindowLong Me.hwnd, GWL_STYLE, (WS_CAPTION Or WS_SYSMENU)
            SetWindowLong Me.hwnd, GWL_EXSTYLE, (WS_EX_TOOLWINDOW Or WS_EX_WINDOWEDGE) And (Not WS_EX_APPWINDOW)
        Case 5 'sizable toolwindow
            SetWindowLong Me.hwnd, GWL_EXSTYLE, (StartExStyle Or WS_EX_TOOLWINDOW) And (Not WS_EX_APPWINDOW)
    End Select
    'force repaint via resizing up, then down one pixel, then setting visible
    Me.Height = Me.Height - 15
    Me.Height = Me.Height + 15
    Me.Visible = True 'show
End Sub
