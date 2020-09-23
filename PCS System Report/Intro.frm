VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Intro 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "  PCS System Resorces Analisys"
   ClientHeight    =   6345
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   10815
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   12
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6345
   ScaleWidth      =   10815
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox ChkTips 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Show Page Tips"
      Height          =   375
      Left            =   6060
      TabIndex        =   49
      Top             =   5640
      Value           =   1  'Checked
      Width           =   2415
   End
   Begin VB.Frame FrmPage 
      BackColor       =   &H00FFFFFF&
      Height          =   3435
      Index           =   5
      Left            =   120
      TabIndex        =   42
      Top             =   1860
      Width           =   10575
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Sound Card"
         Height          =   315
         Index           =   12
         Left            =   360
         TabIndex        =   46
         Top             =   1260
         Width           =   2415
      End
      Begin VB.Label LblOut 
         BackStyle       =   0  'Transparent
         Caption         =   "Retrieving."
         Height          =   315
         Index           =   12
         Left            =   2820
         TabIndex        =   45
         Top             =   1920
         Width           =   6795
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Screen Resolution"
         Height          =   315
         Index           =   11
         Left            =   360
         TabIndex        =   44
         Top             =   1920
         Width           =   2415
      End
      Begin VB.Label LblOut 
         BackStyle       =   0  'Transparent
         Caption         =   "Retrieving."
         Height          =   315
         Index           =   11
         Left            =   2820
         TabIndex        =   43
         Top             =   1260
         Width           =   6795
      End
   End
   Begin VB.Frame FrmPage 
      BackColor       =   &H00FFFFFF&
      Height          =   3435
      Index           =   4
      Left            =   120
      TabIndex        =   35
      Top             =   1860
      Width           =   10575
      Begin VB.CommandButton Command2 
         BackColor       =   &H00FFFFFF&
         Height          =   915
         Left            =   8100
         Picture         =   "Intro.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   40
         Top             =   2340
         Width           =   2295
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Run Test Again"
         Height          =   315
         Index           =   9
         Left            =   8040
         TabIndex        =   41
         Top             =   2040
         Width           =   2415
      End
      Begin VB.Label LblOut 
         BackStyle       =   0  'Transparent
         Caption         =   "Retrieving."
         Height          =   315
         Index           =   10
         Left            =   2820
         TabIndex        =   39
         Top             =   1740
         Width           =   6795
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Connection Method"
         Height          =   315
         Index           =   8
         Left            =   360
         TabIndex        =   38
         Top             =   1740
         Width           =   2415
      End
      Begin VB.Label LblOut 
         BackStyle       =   0  'Transparent
         Caption         =   "Retrieving."
         Height          =   315
         Index           =   9
         Left            =   2820
         TabIndex        =   37
         Top             =   1380
         Width           =   6795
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Net Connection Status"
         Height          =   315
         Index           =   7
         Left            =   360
         TabIndex        =   36
         Top             =   1380
         Width           =   2415
      End
   End
   Begin VB.Frame FrmPage 
      BackColor       =   &H00FFFFFF&
      Height          =   3435
      Index           =   3
      Left            =   120
      TabIndex        =   20
      Top             =   1860
      Width           =   10575
      Begin VB.Label LblOut 
         BackStyle       =   0  'Transparent
         Caption         =   "Retrieving."
         Height          =   315
         Index           =   8
         Left            =   3600
         TabIndex        =   34
         Top             =   360
         Width           =   6795
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "CPU Description"
         Height          =   315
         Index           =   6
         Left            =   480
         TabIndex        =   33
         Top             =   360
         Width           =   3075
      End
      Begin VB.Label LblOut 
         BackStyle       =   0  'Transparent
         Caption         =   "Retrieving."
         Height          =   315
         Index           =   7
         Left            =   3600
         TabIndex        =   32
         Top             =   2880
         Width           =   3435
      End
      Begin VB.Label LblOut 
         BackStyle       =   0  'Transparent
         Caption         =   "Retrieving."
         Height          =   315
         Index           =   6
         Left            =   3600
         TabIndex        =   31
         Top             =   2520
         Width           =   3435
      End
      Begin VB.Label LblOut 
         BackStyle       =   0  'Transparent
         Caption         =   "Retrieving."
         Height          =   315
         Index           =   5
         Left            =   3600
         TabIndex        =   30
         Top             =   2040
         Width           =   3435
      End
      Begin VB.Label LblOut 
         BackStyle       =   0  'Transparent
         Caption         =   "Retrieving."
         Height          =   315
         Index           =   4
         Left            =   3600
         TabIndex        =   29
         Top             =   1680
         Width           =   3435
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Page Memory Available"
         Height          =   315
         Index           =   5
         Left            =   480
         TabIndex        =   28
         Top             =   2880
         Width           =   3075
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Total Page Memory"
         Height          =   315
         Index           =   4
         Left            =   480
         TabIndex        =   27
         Top             =   2520
         Width           =   3075
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Virtual memory Available"
         Height          =   315
         Index           =   3
         Left            =   480
         TabIndex        =   26
         Top             =   2040
         Width           =   3075
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Total Virtual Memory"
         Height          =   315
         Index           =   2
         Left            =   480
         TabIndex        =   25
         Top             =   1680
         Width           =   3075
      End
      Begin VB.Label LblOut 
         BackStyle       =   0  'Transparent
         Caption         =   "Retrieving."
         Height          =   315
         Index           =   3
         Left            =   3600
         TabIndex        =   24
         Top             =   1200
         Width           =   3435
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Physical Memory Available"
         Height          =   315
         Index           =   1
         Left            =   480
         TabIndex        =   23
         Top             =   1200
         Width           =   3075
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Total Physical Memory"
         Height          =   315
         Index           =   0
         Left            =   480
         TabIndex        =   22
         Top             =   840
         Width           =   3075
      End
      Begin VB.Label LblOut 
         BackStyle       =   0  'Transparent
         Caption         =   "Retrieving."
         Height          =   315
         Index           =   2
         Left            =   3600
         TabIndex        =   21
         Top             =   840
         Width           =   3435
      End
   End
   Begin VB.DriveListBox Drive1 
      Height          =   390
      Left            =   3120
      TabIndex        =   19
      Top             =   5700
      Visible         =   0   'False
      Width           =   1635
   End
   Begin VB.Frame FrmPage 
      BackColor       =   &H00FFFFFF&
      Height          =   3435
      Index           =   2
      Left            =   120
      TabIndex        =   17
      Top             =   1860
      Width           =   10575
      Begin MSComctlLib.ListView LstDriveInfo 
         Height          =   2955
         Left            =   180
         TabIndex        =   18
         Top             =   300
         Width           =   10215
         _ExtentX        =   18018
         _ExtentY        =   5212
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   0
         NumItems        =   6
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Drive"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Type"
            Object.Width           =   3069
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "File System"
            Object.Width           =   7832
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Total Capacity"
            Object.Width           =   6068
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Space Used"
            Object.Width           =   6068
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "Space Free"
            Object.Width           =   6068
         EndProperty
      End
   End
   Begin VB.Timer Timer1 
      Left            =   2220
      Top             =   5640
   End
   Begin VB.Frame FrmPage 
      BackColor       =   &H00FFFFFF&
      Height          =   3435
      Index           =   1
      Left            =   120
      TabIndex        =   9
      Top             =   1860
      Width           =   10575
      Begin VB.CommandButton Command1 
         BackColor       =   &H00FFFFFF&
         Height          =   915
         Left            =   6780
         Picture         =   "Intro.frx":307E
         Style           =   1  'Graphical
         TabIndex        =   14
         ToolTipText     =   "Run CPU Test"
         Top             =   2400
         Width           =   3255
      End
      Begin VB.ListBox LstProcess 
         Appearance      =   0  'Flat
         Height          =   1650
         ItemData        =   "Intro.frx":60FC
         Left            =   300
         List            =   "Intro.frx":60FE
         TabIndex        =   13
         Top             =   1680
         Width           =   5835
      End
      Begin VB.Label LblOut 
         BackStyle       =   0  'Transparent
         Caption         =   "Retrieving."
         Height          =   315
         Index           =   13
         Left            =   1920
         TabIndex        =   48
         Top             =   960
         Width           =   5055
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Current Mode:"
         Height          =   315
         Index           =   1
         Left            =   360
         TabIndex        =   47
         Top             =   960
         Width           =   1935
      End
      Begin VB.Label LblOut 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Waiting For Test"
         Height          =   315
         Index           =   1
         Left            =   6300
         TabIndex        =   16
         Top             =   1860
         Width           =   4095
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Run CPU Utilization Test"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   6720
         TabIndex        =   15
         Top             =   1440
         Width           =   3255
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Running Process's"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   360
         TabIndex        =   12
         Top             =   1320
         Width           =   2115
      End
      Begin VB.Label LblOut 
         BackStyle       =   0  'Transparent
         Caption         =   "Retrieving."
         Height          =   315
         Index           =   0
         Left            =   360
         TabIndex        =   11
         Top             =   600
         Width           =   10035
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Operating System:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   0
         Left            =   360
         TabIndex        =   10
         Top             =   300
         Width           =   1935
      End
   End
   Begin VB.CommandButton CmdPrev 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   855
      Left            =   8700
      Picture         =   "Intro.frx":6100
      Style           =   1  'Graphical
      TabIndex        =   8
      ToolTipText     =   "Previous"
      Top             =   5400
      Width           =   915
   End
   Begin VB.CommandButton CmdColour 
      Appearance      =   0  'Flat
      Height          =   855
      Left            =   4950
      Picture         =   "Intro.frx":9306
      Style           =   1  'Graphical
      TabIndex        =   7
      ToolTipText     =   "Change Text colour"
      Top             =   5400
      Width           =   915
   End
   Begin VB.CommandButton CmdExit 
      Appearance      =   0  'Flat
      Height          =   855
      Left            =   180
      Picture         =   "Intro.frx":C408
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   "Exit"
      Top             =   5400
      Width           =   915
   End
   Begin VB.CommandButton CmdNext 
      Appearance      =   0  'Flat
      Height          =   855
      Left            =   9720
      Picture         =   "Intro.frx":F60E
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "Next"
      Top             =   5400
      Width           =   915
   End
   Begin VB.Frame FrmPage 
      BackColor       =   &H00FFFFFF&
      Height          =   3435
      Index           =   0
      Left            =   120
      TabIndex        =   1
      Top             =   1860
      Width           =   10575
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   $"Intro.frx":12814
         Height          =   1095
         Left            =   240
         TabIndex        =   4
         Top             =   1380
         Width           =   10095
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   $"Intro.frx":12917
         Height          =   555
         Left            =   240
         TabIndex        =   2
         Top             =   360
         Width           =   9615
      End
   End
   Begin VB.Label LblPagedescription 
      BackStyle       =   0  'Transparent
      Caption         =   "General Information, And Procedure Description"
      Height          =   615
      Left            =   5580
      TabIndex        =   3
      Top             =   1020
      Width           =   4875
   End
   Begin VB.Label LblPageNo 
      BackStyle       =   0  'Transparent
      Caption         =   "Page 1 of 8"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   5580
      TabIndex        =   0
      Top             =   300
      Width           =   2775
   End
   Begin VB.Image ImgPcsLogo 
      Height          =   1695
      Left            =   60
      Picture         =   "Intro.frx":129BD
      Top             =   60
      Width           =   5310
   End
End
Attribute VB_Name = "Intro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Wmi As Object, Locator As Object
Private PrevCpuTime As Long, SampleRate As Long
Private Sub Command4_Click()

    
    
End Sub

Private Sub CmdColour_Click()
    If VisAid = False Then
        VisAid = True
    Else
        VisAid = False
    End If
    
    Set PassForm = Me
    SetColour
End Sub

Private Sub CmdExit_Click()
End
End Sub

Private Sub CmdNext_Click()
    CurrentPage = CurrentPage + 1
    UpdatePage (CurrentPage)
End Sub

Private Sub CmdPrev_Click()
    CurrentPage = CurrentPage - 1
    If CurrentPage < 0 Then CurrentPage = 0
    UpdatePage (CurrentPage)
End Sub

Private Sub Command1_Click()

    SampleRate = 1 'in seconds
    Timer1.Interval = SampleRate * 1000
    Set Locator = CreateObject("WbemScripting.SWbemLocator")
    Set Wmi = Locator.ConnectServer
    Intro.Enabled = False
    frmDialog.CmdExit.Visible = False
    Timer1.Enabled = True
    Timer1_Timer
End Sub

Private Sub Command2_Click()
TestNetConnection
End Sub

Private Sub Command3_Click()
About.Show
End Sub

Private Sub Form_Load()

    PageTitle(0) = "General Information And Procedure Information."
    PageTitle(1) = "Windows Version, Active Process's And CPU Utilization."
    PageTitle(2) = "Drive Information."
    PageTitle(3) = "CPU And Memory Information."
    PageTitle(4) = "Net Connection Status"
    PageTitle(5) = "Sound And Graphics Information."
    UpdatePage (0)
    
    VisAid = False
    Set PassForm = Me
    SetColour
End Sub

Private Sub DoInstructions(PageNo)

If ChkTips.Value = 0 Then Exit Sub

Select Case PageNo
Case 0
frmDialog.LblDialog.Caption = "Click OK to Close These Information Windows, Then Click Next (The Arrow Pointing Right) To Contiue To The next Page. You Can Also Stop These Tip Boxes By Removing The TickFrom Show Page Tips."
Me.Enabled = False
frmDialog.Show

Case 1
frmDialog.LblDialog.Caption = "This Page Is Displaying Your Windows Version And Running Programs. Click On The Large Tick to Run A Processor Test. Once Complete click Next."
Me.Enabled = False
frmDialog.Show

Case 2
frmDialog.LblDialog.Caption = "You Will Now Be Shown The Information Regarding Your Hardrive(s) And Disk Drives, Used Space And Free Space. Have A Quick Look Through Then Click Next."
Me.Enabled = False
frmDialog.Show

Case 3
frmDialog.LblDialog.Caption = "This Page Will Show You What Your Processor Is And How Much Memory You Have. The Most Refered To Memory Is Your Physical Memory. Scan Through Then Click Next"
Me.Enabled = False
frmDialog.Show

Case 4
frmDialog.LblDialog.Caption = "You will Now See You Internet Connection Status And How You Are Connected, eg LAN Or Modem etc. Click Next."
Me.Enabled = False
frmDialog.Show
End Select

End Sub

Private Sub UpdatePage(PageNo As Single)
    
    LblPagedescription.Caption = PageTitle(PageNo)
    LblPageNo.Caption = "Page " & PageNo + 1 & " Of 8"
    FrmPage(PageNo).ZOrder (0)
    
    If CurrentPage = 0 Then
        
    Else
        CmdPrev.Enabled = True
    End If
    
    Select Case PageNo
    Case 0
        CmdPrev.Enabled = False
        CmdNext.Enabled = True
        DoInstructions (PageNo)
    Case 1
        LblOut(0).Caption = sOperatingSystemString
        CmdNext.Enabled = True
        If CPUTestComplete = False Then CmdNext.Enabled = False
        Dim SMode As String
        
        Select Case StartupMode
        Case 0
        SMode = "Normal Mode"
        Case 1
        SMode = "Safe Mode"
        Case 2
        SMode = "Safe Mode With Network Support"
        End Select
        
        LblOut(13).Caption = SMode
        
        LstProcess.Clear
        CmdPrev.Enabled = True
        GetProccess
        
        DoInstructions (PageNo)
    Case 2
        CmdNext.Enabled = True
        LstDriveInfo.ListItems.Clear
        RunDriveTest
        DoInstructions (PageNo)
    Case 3
        CmdNext.Enabled = True
        GetMemory
        GetCPUDescription
        DoInstructions (PageNo)
    Case 4
        CmdNext.Enabled = True
        DoInstructions (PageNo)
        TestNetConnection
    Case 5
        CmdNext.Enabled = False
        GetSoundInfo
        LblOut(12).Caption = Screen.Width / 15 & " x " & Screen.Height / 15
    End Select
End Sub

Private Sub TestNetConnection()
    Dim MyConnectionStatus As tConnectionStatus
        Dim ConectStatus As String
        GetConnectionInfo MyConnectionStatus 'See this Sub routine For more info...
    If MyConnectionStatus.Connected = True Then
        ConectStatus = "Connected"
        CmdNext.Enabled = True
    Else
        ConectStatus = "Not Connected"
        CmdNext.Enabled = True
    End If
    
    LblOut(9).Caption = ConectStatus
    LblOut(10).Caption = MyConnectionStatus.ConnectionType
End Sub

Private Sub OptConf1_Click(Index As Integer)
End Sub

Private Sub TxtManual_KeyPress(Index As Integer, KeyAscii As Integer)
End Sub


Private Sub ListView1_BeforeLabelEdit(Cancel As Integer)

End Sub

Private Sub Timer1_Timer()
    
    Dim Procs As Object, Proc As Object
    Dim CpuTime, Utilization As Single
    Set Procs = Wmi.InstancesOf("Win32_Process")
    
    frmDialog.Show
    
    For Each Proc In Procs


        If Proc.processid = 0 Then 'System Idle Process
            CpuTime = Proc.KernelModeTime / 10000000

            If PrevCpuTime <> 0 Then
                 Utilization = 1 - (CpuTime - PrevCpuTime) / SampleRate
                If Utilization < 0 Then Utilization = 0
                
            End If
            PrevCpuTime = CpuTime
        End If
    Next
    
    If AveCnt < 10 Then
        AveStat(AveCnt) = Format(Utilization, "0.00")
        Debug.Print AveStat(AveCnt)
        AveCnt = AveCnt + 1
        frmDialog.LblDialog.Caption = "CPU Test " & AveCnt * 10 & "% Complete"
        
        
    Else
        
        If AveCnt = 10 Then LblOut(1).Caption = "CPU Average Utilization " & GetAverage
        Timer1.Enabled = False
        AveCnt = 0
        CPUTestComplete = True
        CmdNext.Enabled = True
        frmDialog.CmdExit.Visible = True
        frmDialog.LblDialog.Caption = "CPU Test Complete. Close This Window"
    End If
    
End Sub

Public Function GetAverage() As String
On Error Resume Next
Dim LP As Single
Dim AveOut As Double

For LP = 0 To 9
    AveOut = AveOut + AveStat(LP)
    'Debug.Print AveStat(LP)
Next LP

AveOut = AveOut / 9

GetAverage = Format(AveOut, "0.00%")

End Function

