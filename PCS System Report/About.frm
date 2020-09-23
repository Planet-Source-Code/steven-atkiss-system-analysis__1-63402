VERSION 5.00
Begin VB.Form About 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "  About PCS"
   ClientHeight    =   5295
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   6855
   BeginProperty Font 
      Name            =   "Arial Narrow"
      Size            =   14.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   353
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   457
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   WhatsThisHelp   =   -1  'True
   Begin VB.Timer Timer1 
      Interval        =   10
      Left            =   900
      Top             =   3780
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Enabled         =   0   'False
      Height          =   315
      Index           =   2
      Left            =   3180
      TabIndex        =   2
      Top             =   2940
      Width           =   3495
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Enabled         =   0   'False
      Height          =   315
      Index           =   1
      Left            =   3180
      TabIndex        =   1
      Top             =   1860
      Width           =   3495
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Enabled         =   0   'False
      Height          =   315
      Index           =   0
      Left            =   3180
      TabIndex        =   0
      Top             =   780
      Width           =   3495
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      Height          =   1695
      Left            =   360
      Picture         =   "About.frx":0000
      Top             =   0
      Width           =   2490
   End
End
Attribute VB_Name = "About"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private MoveToPosition As Single
Private MoveToPos(2)

Private Sub Form_Load()

    MoveToPos(0) = 0
    MoveToPos(1) = 76
    MoveToPos(2) = 148
    
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

Dim LP As Single

    For LP = 0 To 2
        If X > Label1(LP).Left And X < Label1(LP).Left + Label1(LP).Width Then
            If Y > Label1(LP).Top And Y < Label1(LP).Top + Label1(LP).Height Then
                MoveToPosition = LP
            End If
        End If
    Next LP
    
End Sub

Private Sub Timer1_Timer()

    If Image1.Top > MoveToPos(MoveToPosition) Then 'move up
        Image1.Top = Image1.Top - 5
    End If
    
    If Image1.Top < MoveToPos(MoveToPosition) Then 'move down
        Image1.Top = Image1.Top + 5
    End If
    
End Sub
