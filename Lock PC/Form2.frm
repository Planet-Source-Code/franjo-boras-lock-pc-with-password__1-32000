VERSION 5.00
Begin VB.Form Form2 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Lock your PC"
   ClientHeight    =   735
   ClientLeft      =   30
   ClientTop       =   270
   ClientWidth     =   3315
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   7.5
      Charset         =   238
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Form2.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   735
   ScaleWidth      =   3315
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   650
      Left            =   2280
      Top             =   480
   End
   Begin VB.Label Label1 
      Caption         =   "Thank you for using Lock Your PC"
      Height          =   372
      Left            =   480
      TabIndex        =   0
      Top             =   240
      Width           =   3012
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Timer1_Timer()
End
End Sub
