VERSION 5.00
Begin VB.Form frmMain 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00808080&
   BorderStyle     =   0  'None
   Caption         =   "MDI Skin Example"
   ClientHeight    =   7725
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9660
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   515
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   644
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Display Child Form 2"
      Height          =   255
      Left            =   120
      TabIndex        =   17
      Top             =   1080
      Width           =   3855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Display Child Form"
      Height          =   255
      Left            =   120
      TabIndex        =   16
      Top             =   720
      Width           =   3855
   End
   Begin VB.PictureBox Picture1 
      Align           =   3  'Align Left
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      Height          =   7395
      Left            =   0
      ScaleHeight     =   7395
      ScaleWidth      =   30
      TabIndex        =   13
      Top             =   300
      Width           =   30
   End
   Begin VB.PictureBox picBottom 
      Align           =   2  'Align Bottom
      AutoRedraw      =   -1  'True
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      Height          =   30
      Left            =   0
      MousePointer    =   7  'Size N S
      ScaleHeight     =   30
      ScaleWidth      =   9660
      TabIndex        =   12
      Top             =   7695
      Width           =   9660
   End
   Begin VB.PictureBox picRight 
      Align           =   4  'Align Right
      AutoRedraw      =   -1  'True
      BackColor       =   &H00404040&
      Height          =   7395
      Left            =   9630
      MousePointer    =   9  'Size W E
      ScaleHeight     =   7395
      ScaleWidth      =   30
      TabIndex        =   11
      Top             =   300
      Width           =   30
   End
   Begin VB.PictureBox picTitle 
      Align           =   1  'Align Top
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   300
      Left            =   0
      ScaleHeight     =   20
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   644
      TabIndex        =   0
      Top             =   0
      Width           =   9660
      Begin VB.PictureBox titleButton 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   300
         Left            =   960
         Picture         =   "Form1.frx":0000
         ScaleHeight     =   20
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   20
         TabIndex        =   6
         Top             =   0
         Visible         =   0   'False
         Width           =   300
      End
      Begin VB.PictureBox titleRight 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   300
         Left            =   1440
         Picture         =   "Form1.frx":04F2
         ScaleHeight     =   20
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   3
         TabIndex        =   5
         Top             =   0
         Visible         =   0   'False
         Width           =   45
      End
      Begin VB.PictureBox titleBack 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   300
         Left            =   960
         Picture         =   "Form1.frx":0624
         ScaleHeight     =   20
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   1
         TabIndex        =   4
         Top             =   0
         Visible         =   0   'False
         Width           =   15
      End
      Begin VB.PictureBox titleButton2 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   300
         Left            =   1080
         Picture         =   "Form1.frx":06B6
         ScaleHeight     =   20
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   20
         TabIndex        =   3
         Top             =   360
         Visible         =   0   'False
         Width           =   300
      End
      Begin VB.PictureBox picMin 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   300
         Left            =   360
         Picture         =   "Form1.frx":0BA8
         ScaleHeight     =   20
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   20
         TabIndex        =   2
         ToolTipText     =   "Minimize"
         Top             =   0
         Width           =   300
      End
      Begin VB.PictureBox picExit 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   300
         Left            =   0
         Picture         =   "Form1.frx":109A
         ScaleHeight     =   20
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   20
         TabIndex        =   1
         ToolTipText     =   "Exit"
         Top             =   0
         Width           =   300
      End
      Begin VB.PictureBox titleLeft 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   300
         Left            =   0
         Picture         =   "Form1.frx":158C
         ScaleHeight     =   20
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   59
         TabIndex        =   7
         Top             =   0
         Visible         =   0   'False
         Width           =   885
      End
      Begin VB.Label lblTitle 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "My Title"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   195
         Index           =   0
         Left            =   3600
         TabIndex        =   9
         Top             =   0
         Width           =   690
      End
      Begin VB.Label lblTitle 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "My Title"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   1
         Left            =   3600
         TabIndex        =   10
         Top             =   0
         Width           =   930
      End
      Begin VB.Image Image1 
         Height          =   300
         Left            =   2280
         Picture         =   "Form1.frx":23DE
         Top             =   0
         Width           =   45
      End
   End
   Begin VB.PictureBox picResize 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   300
      Left            =   4320
      MousePointer    =   8  'Size NW SE
      Picture         =   "Form1.frx":2510
      ScaleHeight     =   300
      ScaleWidth      =   300
      TabIndex        =   8
      Top             =   2760
      Width           =   300
   End
   Begin VB.TextBox txtCaption 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   120
      TabIndex        =   14
      Text            =   "MDI Skin Example"
      Top             =   360
      Width           =   2295
   End
   Begin VB.Timer TimerCaption 
      Interval        =   1
      Left            =   4080
      Top             =   480
   End
   Begin VB.CommandButton cmdTemp 
      Caption         =   "Change Title"
      Default         =   -1  'True
      Height          =   315
      Left            =   2520
      TabIndex        =   15
      Top             =   360
      Width           =   1455
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function ReleaseCapture Lib "user32" () As Long
Private Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Private Const WM_NCLBUTTONDOWN = &HA1
Private Const HTCAPTION = 2
Dim MouseIsDown(1) As Boolean
Dim p As Integer, m As Integer, n As Integer

Public Function MoveForm(TheForm As Form)
    Dim ret
    ReleaseCapture
    SendMessage TheForm.hwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&
End Function

Private Sub cmdTemp_Click()
    Me.Caption = txtCaption
End Sub

Private Sub Command1_Click()
    frmChildForm.Show
    SetParent frmChildForm.hwnd, hwnd
End Sub

Private Sub Command2_Click()
    frmChildForm2.Show
    SetParent frmChildForm2.hwnd, hwnd
End Sub

Private Sub Form_Load()
lblTitle(0).Top = (picTitle.ScaleHeight / 2) - (lblTitle(0).Height / 2)
lblTitle(0).Left = 60
lblTitle(1).Top = (picTitle.ScaleHeight / 2) - (lblTitle(1).Height / 2) + 2
lblTitle(1).Left = 62
picTitle.PaintPicture titleLeft, 0, 0
Dim tWidth
tWidth = Me.Width
Me.Visible = False
Me.Width = Screen.Width
For i = titleLeft.Width To Screen.Width
    picTitle.PaintPicture titleBack, i, 0
Next i
Me.Width = tWidth
Me.Visible = True
End Sub

Private Sub Form_Resize()
On Error Resume Next
Image1.Left = picTitle.Width - titleRight.Width
picExit.Left = picTitle.ScaleWidth - 3 - picExit.Width
picMin.Left = picTitle.ScaleWidth - 6 - (picExit.Width * 2)
picResize.Top = Me.ScaleHeight - picResize.Height - 3
picResize.Left = Me.ScaleWidth - picResize.Width - 2
lblTitle(0).Width = picTitle.ScaleWidth - (picTitle.ScaleWidth - picMin.Left) - lblTitle(0).Left
lblTitle(1).Width = picTitle.ScaleWidth - (picTitle.ScaleWidth - picMin.Left) - lblTitle(1).Left
End Sub

Private Sub Image1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    MoveForm Me
End Sub

Private Sub lblTitle_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    MoveForm Me
End Sub

Private Sub picBottom_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    n = 1
End Sub

Private Sub picBottom_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error GoTo 1
    If n = 1 Then
        Me.Height = Y + Me.Height
    End If
    Exit Sub
1
    Height = 630
End Sub

Private Sub picBottom_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    n = 0
    If Height < 630 Then Height = 630
    Form_Resize
End Sub

Private Sub picExit_Click()
    End
End Sub

Private Sub picExit_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    MouseIsDown(0) = True
    picExit = titleButton2
End Sub

Private Sub picExit_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If MouseIsDown(0) And X >= 0 And X <= picExit.ScaleWidth And Y >= 0 And Y <= picExit.ScaleHeight Then
        picExit = titleButton2
    Else
        picExit = titleButton
    End If
End Sub

Private Sub picExit_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    MouseIsDown(0) = False
    picExit = titleButton
End Sub

Private Sub picMin_Click()
    Me.WindowState = 1
End Sub

Private Sub picMin_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    MouseIsDown(1) = True
    picMin = titleButton2
End Sub

Private Sub picMin_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If MouseIsDown(1) And X >= 0 And X <= picMin.ScaleWidth And Y >= 0 And Y <= picMin.ScaleHeight Then
        picMin = titleButton2
    Else
        picMin = titleButton
    End If
End Sub

Private Sub picMin_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    MouseIsDown(1) = False
    picMin = titleButton
End Sub

Private Sub picResize_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    p = 1
End Sub

Private Sub picResize_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error GoTo 1
    If p = 1 Then
        Width = X + Width
        Height = Y + Height
    End If
    Exit Sub
1
    Width = 2640
    Height = 630
End Sub

Private Sub picResize_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    p = 0
    If Width < 2640 Then Width = 2640
    If Height < 630 Then Height = 630
    Form_Resize
End Sub

Private Sub picRight_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    m = 1
End Sub

Private Sub picRight_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error GoTo 1
    If m = 1 Then
        Me.Width = X + Me.Width
    End If
    Exit Sub
1
    Width = 2640
End Sub

Private Sub picRight_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    m = 0
    If Width < 2640 Then Width = 2640
    Form_Resize
End Sub

Private Sub picTitle_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    MoveForm Me
End Sub

Private Sub TimerCaption_Timer()
    If lblTitle(0) <> Me.Caption Then lblTitle(0) = Me.Caption
    If lblTitle(1) <> Me.Caption Then lblTitle(1) = Me.Caption
End Sub

Private Sub txtCaption_GotFocus()
    cmdTemp.Default = True
End Sub

Private Sub txtCaption_LostFocus()
    cmdTemp.Default = False
End Sub
