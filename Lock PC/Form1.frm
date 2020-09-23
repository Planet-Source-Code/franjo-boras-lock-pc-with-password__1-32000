VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   4650
   ClientLeft      =   150
   ClientTop       =   150
   ClientWidth     =   4845
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   238
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4650
   ScaleWidth      =   4845
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdAbout 
      Caption         =   "&About"
      Height          =   320
      Left            =   2760
      TabIndex        =   11
      Top             =   4200
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "&Exit"
      Height          =   320
      Left            =   240
      TabIndex        =   10
      Top             =   4200
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.Frame Frame2 
      Height          =   1455
      Left            =   120
      TabIndex        =   6
      Top             =   2160
      Width           =   4575
      Begin VB.CommandButton cmdUnlock 
         Caption         =   "&Unlock PC"
         Default         =   -1  'True
         Height          =   320
         Left            =   1320
         TabIndex        =   9
         Top             =   960
         Width           =   1815
      End
      Begin VB.TextBox txtUnlock 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   238
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   288
         Left            =   1320
         TabIndex        =   7
         Top             =   600
         Width           =   1812
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Enter password:"
         Height          =   195
         Left            =   1560
         TabIndex        =   8
         Top             =   240
         Width           =   1185
      End
   End
   Begin VB.Frame Frame1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1935
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   4575
      Begin VB.Label Label4 
         Caption         =   "Unlock your PC with password wich you entered in Login window."
         Height          =   495
         Left            =   600
         TabIndex        =   5
         Top             =   1320
         Width           =   3495
      End
      Begin VB.Label Label3 
         Caption         =   "Your PC now is locked. Keyboard is disabled, mouse is disabled, everthing is disabled outside this window."
         Height          =   615
         Left            =   600
         TabIndex        =   4
         Top             =   600
         Width           =   3615
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "ATTENTION!"
         Height          =   195
         Left            =   1920
         TabIndex        =   3
         Top             =   240
         Width           =   915
      End
   End
   Begin VB.Label lblPassword 
      AutoSize        =   -1  'True
      Caption         =   "lblPassword"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   1920
      TabIndex        =   1
      Top             =   4920
      Visible         =   0   'False
      Width           =   840
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Height          =   255
      Left            =   1320
      TabIndex        =   0
      Top             =   3720
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.Menu mnufile 
      Caption         =   "File"
      Visible         =   0   'False
      Begin VB.Menu mnucrta 
         Caption         =   "-"
      End
      Begin VB.Menu mnucrta9 
         Caption         =   "-"
      End
      Begin VB.Menu mnuabout 
         Caption         =   "About..."
      End
      Begin VB.Menu mnucrta1 
         Caption         =   "-"
      End
      Begin VB.Menu mnucrt2 
         Caption         =   "-"
      End
      Begin VB.Menu mnushowme 
         Caption         =   "Show program"
      End
      Begin VB.Menu mnucrta3 
         Caption         =   "-"
      End
      Begin VB.Menu mnucrta4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuexit 
         Caption         =   "Exit"
      End
      Begin VB.Menu mnucrta5 
         Caption         =   "-"
      End
      Begin VB.Menu mnucrta6 
         Caption         =   "-"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" (ByVal uAction As Long, ByVal uParam As Long, ByRef lpvParam As Any, ByVal fuWinIni As Long) As Long
Private Const SPI_SCREENSAVERRUNNING = 97

' Clip mouse Declarations
Private Type RECT
     Left As Long
     Top As Long
     Right As Long
     Bottom As Long
End Type
Private Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long
Private Declare Function ClipCursor Lib "user32" (lpRect As Any) As Long
Private Declare Function ClipCursorByNum Lib "user32" Alias "ClipCursor" (ByVal num As Long) As Long


Private Sub cmdAbout_Click()
Form3.Show
End Sub

Private Sub cmdExit_Click()
End
End Sub

Private Sub cmdUnlock_Click()
If txtUnlock.Text = lblPassword.Caption Then
Label1.Visible = True
Label1.Caption = "Password is good!!!"
    
    cmdUnlock.Enabled = False
 cmdExit.Visible = True
   cmdAbout.Visible = True

    
    
    

    ' Free the mouse
    ClipCursorByNum 0&

    ' Tell the system no screen saver is running.
    SystemParametersInfo SPI_SCREENSAVERRUNNING, False, 0, 0
Else
Label1.Visible = True
Label1.Caption = "Invalid password, try again!"
txtUnlock.SetFocus
SendKeys "{Home}+{End}"
End If

End Sub

Private Sub Command1_Click()

End Sub

Private Sub Form_Load()
MoveMouse 100, 100


    cmdUnlock.Enabled = False
  
 LockPC
End Sub

Private Sub mnuabout_Click()
Form3.Show
End Sub








Sub LockPC()

Dim window As RECT

 
    cmdUnlock.Enabled = True

 

    ' Restrict the mouse to this window.
    GetWindowRect hWnd, window
    ClipCursor window

    ' Tell the system a screen saver is running.
    SystemParametersInfo SPI_SCREENSAVERRUNNING, True, 0, 0
End Sub




