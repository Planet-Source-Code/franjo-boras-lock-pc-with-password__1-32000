VERSION 5.00
Begin VB.Form Form3 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "About..."
   ClientHeight    =   2610
   ClientLeft      =   30
   ClientTop       =   270
   ClientWidth     =   3960
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   7.5
      Charset         =   238
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Form3.frx":0000
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2610
   ScaleWidth      =   3960
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "&Ok"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   1560
      TabIndex        =   4
      Top             =   2160
      Width           =   852
   End
   Begin VB.Frame Frame1 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1572
      Left            =   600
      TabIndex        =   0
      Top             =   120
      Width           =   2772
      Begin VB.Label Label3 
         Caption         =   "Croatia"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   238
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Left            =   1200
         TabIndex        =   3
         Top             =   600
         Width           =   732
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "E-mail me"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   238
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   1080
         MouseIcon       =   "Form3.frx":0442
         MousePointer    =   99  'Custom
         TabIndex        =   2
         ToolTipText     =   "Send mail to Franjo Boras"
         Top             =   1080
         Width           =   675
      End
      Begin VB.Label Label2 
         Caption         =   "Program created: Franjo Boras"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   238
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Left            =   240
         TabIndex        =   1
         Top             =   360
         Width           =   2412
      End
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   120
      Picture         =   "Form3.frx":074C
      Top             =   120
      Width           =   480
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Private Sub Command1_Click()
Unload Me
End Sub

Private Sub Label5_Click()
ShellExecute 0, "open", "mailto:franjo.boras@sb.tel.hr ; boras@vip.hr", "", "", 0
End Sub
