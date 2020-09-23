VERSION 5.00
Begin VB.Form frmDialog 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "    PCS"
   ClientHeight    =   2250
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   9465
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2250
   ScaleWidth      =   9465
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton CmdExit 
      Appearance      =   0  'Flat
      Height          =   855
      Left            =   8040
      Picture         =   "frmDialog.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Close Information Window"
      Top             =   1320
      Width           =   1335
   End
   Begin VB.Label LblDialog 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   2640
      TabIndex        =   0
      Top             =   240
      Width           =   5295
   End
   Begin VB.Image Image1 
      Height          =   1695
      Left            =   60
      Picture         =   "frmDialog.frx":3D02
      Top             =   240
      Width           =   2490
   End
End
Attribute VB_Name = "frmDialog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CmdExit_Click()
Intro.Enabled = True
Me.Hide
End Sub

Private Sub Form_Load()
SetWindowPos Me.hWnd, -1, 0, 0, 0, 0, 3
End Sub
