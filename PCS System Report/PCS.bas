Attribute VB_Name = "Module1"
Option Explicit
Public VisAid As Boolean
Public PassForm As Form
Public Control As Control
Public PageTitle(8) As String
Public CurrentPage As Single
Public CPUTestComplete As Boolean
Public AveStat(10)
Public AveCnt As Single

'Form OnTop
Public Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long

'Registery Stuff
Public Declare Function RegCloseKey Lib "advapi32" (ByVal HKey As Long) As Long
Public Declare Function RegOpenKeyEx Lib "advapi32" Alias "RegOpenKeyExA" (ByVal HKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As Long) As Long
Public Declare Function RegQueryValueEx Lib "advapi32" Alias "RegQueryValueExA" (ByVal HKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, ByVal lpData As String, lpcbData As Long) As Long
Public Declare Function RegSetValueEx Lib "advapi32" Alias "RegSetValueExA" (ByVal HKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, ByVal lpData As String, ByVal cbData As Long) As Long
Public Declare Function RegFlushKey Lib "advapi32" (ByVal HKey As Long) As Long
Public Declare Function RegConnectRegistry Lib "advapi32.dll" Alias "RegConnectRegistryA" (ByVal lpMachineName As String, ByVal HKey As Long, phkResult As Long) As Long
Public Declare Function RegCreateKey Lib "advapi32.dll" Alias "RegCreateKeyA" (ByVal HKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Public Declare Function RegSetValue Lib "advapi32.dll" Alias "RegSetValueA" (ByVal HKey As Long, ByVal lpSubKey As String, ByVal dwType As Long, ByVal lpData As String, ByVal cbData As Long) As Long
Public Declare Function RegSetValueExString Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal HKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, ByVal lpValue As String, ByVal cbData As Long) As Long
Public Declare Function RegSetValueExLong Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal HKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, lpValue As Long, ByVal cbData As Long) As Long
    
    Public Const READ_CONTROL = &H20000
    Public Const SYNCHRONIZE = &H100000
    Public Const STANDARD_RIGHTS_ALL = &H1F0000
    Public Const STANDARD_RIGHTS_READ = READ_CONTROL
    Public Const STANDARD_RIGHTS_WRITE = READ_CONTROL
    Public Const KEY_QUERY_VALUE = &H1
    Public Const KEY_SET_VALUE = &H2
    Public Const KEY_CREATE_SUB_KEY = &H4
    Public Const KEY_ENUMERATE_SUB_KEYS = &H8
    Public Const KEY_NOTIFY = &H10
    Public Const KEY_CREATE_LINK = &H20
    Public Const KEY_ALL_ACCESS = ((STANDARD_RIGHTS_ALL Or KEY_QUERY_VALUE Or KEY_SET_VALUE Or KEY_CREATE_SUB_KEY Or KEY_ENUMERATE_SUB_KEYS Or KEY_NOTIFY Or KEY_CREATE_LINK) And (Not SYNCHRONIZE))
    Public Const KEY_READ = ((STANDARD_RIGHTS_READ Or KEY_QUERY_VALUE Or KEY_ENUMERATE_SUB_KEYS Or KEY_NOTIFY) And (Not SYNCHRONIZE))
    Public Const KEY_EXECUTE = ((KEY_READ) And (Not SYNCHRONIZE))
    Public Const KEY_WRITE = ((STANDARD_RIGHTS_WRITE Or KEY_SET_VALUE Or KEY_CREATE_SUB_KEY) And (Not SYNCHRONIZE))
    
    Public Const HKEY_CLASSES_ROOT = &H80000000
    Public Const HKEY_CURRENT_USER = &H80000001
    Public Const HKEY_LOCAL_MACHINE = &H80000002
    Public Const HKEY_USERS = &H80000003
    Public Const HKEY_PERFORMANCE_DATA = &H80000004
    
    Public Const ERROR_SUCCESS = 0&
    Public Const ERROR_BADDB = 1009&
    Public Const ERROR_BADKEY = 1010&
    Public Const ERROR_CANTOPEN = 1011&
    Public Const ERROR_CANTREAD = 1012&
    Public Const ERROR_CANTWRITE = 1013&
    Public Const ERROR_OUTOFMEMORY = 14&
    Public Const ERROR_INVALID_PARAMETER = 87&
    Public Const ERROR_ACCESS_DENIED = 5&
    
    Public Const REG_NONE = 0
    Public Const REG_SZ = 1
    Public Const REG_EXPAND_SZ = 2
    Public Const REG_BINARY = 3
    Public Const REG_DWORD = 4
    Public Const REG_DWORD_LITTLE_ENDIAN = 4
    Public Const REG_DWORD_BIG_ENDIAN = 5
    


'Memory Information
Private Declare Sub GlobalMemoryStatus Lib "kernel32" (lpBuffer As MEMORYSTATUS)

Private Type MEMORYSTATUS
    dwLength As Long
    dwMemoryLoad As Long
    dwTotalPhys As Long
    dwAvailPhys As Long
    dwTotalPageFile As Long
    dwAvailPageFile As Long
    dwTotalVirtual As Long
    dwAvailVirtual As Long
    End Type


'Drive Information
Public m_cDiskSize As Currency
Public m_cDiskUsed As Currency
Public m_cDiskFree As Currency
Public m_fFreePercent As Single
Public m_lSerial As Long
Public m_sVolume As String
Public m_sFileSystem As String
Public m_sAllDrives As String
Public m_sDriveType As String
Public Const FS_CASE_IS_PRESERVED = &H2
Public Const FS_CASE_SENSITIVE = &H1
Public Const FS_UNICODE_STORED_ON_DISK = &H4
Public Const FS_PERSISTENT_ACLS = &H8
Public Const FS_FILE_COMPRESSION = &H10
Public Const FS_VOL_IS_COMPRESSED = &H8000


Public Declare Function GetDiskFreeSpace Lib "kernel32" Alias "GetDiskFreeSpaceA" (ByVal lpRootPathName As String, lpSectorsPerCluster As Long, lpBytesPerSector As Long, lpNumberOfFreeClusters As Long, lpTotalNumberOfClusters As Long) As Long
Public Declare Function GetVolumeInformation Lib "kernel32" Alias "GetVolumeInformationA" (ByVal lpRootPathName As String, ByVal lpVolumeNameBuffer As String, ByVal nVolumeNameSize As Long, lpVolumeSerialNumber As Long, lpMaximumComponentLength As Long, lpFileSystemFlags As Long, ByVal lpFileSystemNameBuffer As String, ByVal nFileSystemNameSize As Long) As Long
Public Declare Function GetLogicalDriveStrings Lib "kernel32" Alias "GetLogicalDriveStringsA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long
Public Declare Function GetDriveType Lib "kernel32" Alias "GetDriveTypeA" (ByVal nDrive As String) As Long
  
'Windows Version
Public Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" (lpVersionInformation As OSVERSIONINFO) As Long
Public Declare Function GetVersionExA Lib "kernel32" (lpVersionInformation As OSVERSIONINFOEX) As Long


Public Type OSVERSIONINFO
    dwOSVersionInfoSize As Long
    dwMajorVersion As Long
    dwMinorVersion As Long
    dwBuildNumber As Long
    dwPlatformId As Long
    szCSDVersion As String * 128
    End Type


Public Type OSVERSIONINFOEX
    dwOSVersionInfoSize As Long
    dwMajorVersion As Long
    dwMinorVersion As Long
    dwBuildNumber As Long
    dwPlatformId As Long
    szCSDVersion As String * 128
    wServicePackMajor As Integer
    wServicePackMinor As Integer
    wSuiteMask As Integer
    wProductType As Byte
    wReserved As Byte
    End Type
    Public Const VER_PLATFORM_WIN32s = 0
    Public Const VER_PLATFORM_WIN32_WINDOWS = 1
    Public Const VER_PLATFORM_WIN32_NT = 2

'Net Connection Info
Public Declare Function InternetGetConnectedState Lib "wininet.dll" (ByRef lpdwFlags As Long, ByVal dwReserved As Long) As Integer
Public Const INTERNET_CONNECTION_PROXY = &H4
Public Const INTERNET_CONNECTION_CONFIGURED = &H40
Public Const INTERNET_CONNECTION_LAN = &H2
Public Const INTERNET_CONNECTION_MODEM = &H1
Public Const INTERNET_RAS_INSTALLED = &H10
Public Const INTERNET_CONNECTION_OFFLINE = &H20


Public Type tConnectionStatus
    Connected As Boolean
    ConnectionType As String
    RASInstalled As Boolean
End Type
    
'Sound Card
Public Declare Function waveOutGetDevCaps Lib "winmm.dll" Alias "waveOutGetDevCapsA" (ByVal uDeviceID As Long, lpCaps As WAVEOUTCAPS, ByVal uSize As Long) As Long
Public Const MAXPNAMELEN = 32


Public Type WAVEOUTCAPS
    wMid As Integer
    wPid As Integer
    vDriverVersion As Long
    szPname As String * MAXPNAMELEN
    dwFormats As Long
    wChannels As Integer
    dwSupport As Long
    End Type

'Start Mode
Public Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long



Public Enum StartupModeConstants
    NormalMode = 0
    SafeMode = 1
    SafeModeWithNetworking = 2
End Enum



Public Function StartupMode() As StartupModeConstants
    StartupMode = GetSystemMetrics(&H43)
End Function
Public Function GetConnectionInfo(ConnectionStatus As tConnectionStatus) As Boolean
    Dim pdFlags& 'Dimensionalize pdFlags as Long data Type


    If InternetGetConnectedState(pdFlags&, 0) Then
        'Call InternetGetConnectedState to initi
        '     alize pdFlags with the current connectio
        '     n information flags
        GetConnectionInfo = True
        'InterNetGetConnectedState function was
        '     successful, return true


        If (pdFlags& And INTERNET_CONNECTION_PROXY) = INTERNET_CONNECTION_PROXY Then
            'Perform a Bitwise And operation to dete
            '     rmine if the variable pdFlags specifies
            '     the Internet_Connection_Proxy constant
            ConnectionStatus.ConnectionType = "Local system uses a proxy server To connect to the Internet."
            'Initialize ConnectionStatus's Connectio
            '     nType member with the appropriate connec
            '     tion description
            ConnectionStatus.Connected = True
            'Initialize this structures Connected me
            '     mber
        End If


        If (pdFlags& And INTERNET_CONNECTION_CONFIGURED) = INTERNET_CONNECTION_CONFIGURED Then
            ConnectionStatus.Connected = True
            ConnectionStatus.ConnectionType = "Connection Found, On\OffLine Udetermind."
        End If


        If (pdFlags& And INTERNET_CONNECTION_LAN) = INTERNET_CONNECTION_LAN Then
            ConnectionStatus.ConnectionType = "Local Area Network Connection."
            ConnectionStatus.Connected = True
        End If


        If (pdFlags& And INTERNET_CONNECTION_MODEM) = INTERNET_CONNECTION_MODEM Then
            ConnectionStatus.ConnectionType = "Modem Connection."
            ConnectionStatus.Connected = True
        End If


        If (pdFlags& And INTERNET_CONNECTION_OFFLINE) = INTERNET_CONNECTION_OFFLINE Then
            ConnectionStatus.ConnectionType = "Local system is offline."
            ConnectionStatus.Connected = False
        End If


        If (pdFlags& And INTERNET_RAS_INSTALLED) = INTERNET_RAS_INSTALLED Then
            ConnectionStatus.RASInstalled = True
        End If
    End If
End Function


Public Sub GetDiskSpace(ByVal sDrive As String)
    ' this will calculate the drive specs fo
    '     r the drive and report total size,
    ' size used and size available as well a
    '     s used %
    Dim lResult As Long
    Dim lSectorPerCluster As Long
    Dim lBytesPerSector As Long
    Dim lFreeClusters As Long
    Dim lTotalClusters As Long
    
    ' call the API and get the information
    lResult = GetDiskFreeSpace(sDrive, lSectorPerCluster, lBytesPerSector, lFreeClusters, _
    lTotalClusters)
    
    ' perform the various calculations requi
    '     red
    m_cDiskSize = CCur(lTotalClusters) * CCur(lSectorPerCluster) * CCur(lBytesPerSector)
    m_cDiskFree = CCur(lFreeClusters) * CCur(lSectorPerCluster) * CCur(lBytesPerSector)
    m_cDiskUsed = m_cDiskSize - m_cDiskFree
    


    If m_cDiskSize <> 0 Then
        m_fFreePercent = m_cDiskFree / m_cDiskSize * 100
    Else
        m_fFreePercent = 0
    End If
End Sub

Public Sub RunDriveTest()
Dim LP As Single
Dim DriveLetter As String

With Intro

    For LP = 0 To .Drive1.ListCount - 1
        DriveLetter = Left(.Drive1.List(LP), 1) & ":\"
        .LstDriveInfo.ListItems.Add , , DriveLetter
        
        GetTypeOfDrive (DriveLetter)
        .LstDriveInfo.ListItems(LP + 1).ListSubItems.Add 1, , m_sDriveType
        
        GetVolumeInfo (DriveLetter)
        .LstDriveInfo.ListItems(LP + 1).ListSubItems.Add 2, , m_sFileSystem
        
        GetDiskSpace (DriveLetter)
        .LstDriveInfo.ListItems(LP + 1).ListSubItems.Add 3, , Format(m_cDiskSize / 1048576, "#,#.##")
        .LstDriveInfo.ListItems(LP + 1).ListSubItems.Add 4, , Format(m_cDiskUsed / 1048576, "#,#.##")
        If m_fFreePercent > 0 Then
            .LstDriveInfo.ListItems(LP + 1).ListSubItems.Add 5, , Format(m_cDiskFree / 1048576, "#,#.##") & " (" & Format(m_fFreePercent, "#.##") & "%)"
        Else
            .LstDriveInfo.ListItems(LP + 1).ListSubItems.Add 5, , "0.00 (0.00%)"
        End If
    Next LP

End With
End Sub

Public Sub GetTypeOfDrive(ByVal sDrive As String)


    Select Case GetDriveType(sDrive)
        Case Is = 2
        m_sDriveType = "Removable"
        Case Is = 3
        m_sDriveType = "Fixed"
        Case Is = 4
        m_sDriveType = "Remote"
        Case Is = 5
        m_sDriveType = "CD-Rom"
        Case Is = 6
        m_sDriveType = "RAM Disk"
        Case Else
        m_sDriveType = "Unknown"
    End Select
End Sub


Public Sub GetVolumeInfo(ByVal sDrive As String)
    Dim sBuffer As String
    Dim sSysName As String
    Dim lResult As Long
    Dim lSysFlags As Long
    Dim lComponentLength As Long
    
    sBuffer = String$(256, 0)
    sSysName = String$(256, 0)
    lResult = GetVolumeInformation(sDrive, sBuffer, 255, m_lSerial, lComponentLength, lSysFlags, sSysName, 255)
    


    If lResult = 0 Then
        ' unable to get information
        m_sVolume = "Unable To retrieve information"
        m_sFileSystem = "Unable To retrieve information"
        m_lSerial = 0
    Else
        ' retrieve the information
        m_sVolume = Left$(sBuffer, InStr(sBuffer, Chr$(0)) - 1)
        m_sFileSystem = Left$(sSysName, InStr(sSysName, Chr$(0)) - 1)
    End If
End Sub



Public Sub SetColour()

Dim TxtColour As Long

If VisAid = True Then
    TxtColour = vbBlack
Else
    TxtColour = RGB(85, 140, 250)
End If

For Each Control In PassForm
    If TypeOf Control Is Label Then
        Control.ForeColor = TxtColour
    End If
    If TypeOf Control Is ListBox Then
        Control.ForeColor = TxtColour
    End If
    If TypeOf Control Is CheckBox Then
        Control.ForeColor = TxtColour
    End If
    If TypeOf Control Is TextBox Then
        Control.ForeColor = TxtColour
    End If
    If TypeOf Control Is OptionButton Then
        Control.ForeColor = TxtColour
    End If
    If TypeOf Control Is ListView Then
        Control.ForeColor = TxtColour
    End If
Next

End Sub


Public Sub GetProccess()
    
    Dim objWMIService As Object
    Dim colItems As Object
    Dim objItem As Object
    Dim strCheck As String
    
    Const strComputer = ""
    Const WMIFLAGFORWARDONLY = 32
    Const WMIFLAGRETURNIMMEDIATELY = 16
        
    strCheck = vbNullString
    Set objWMIService = GetObject("winmgmts:" & strComputer & "\root\cimv2")
    
    Set colItems = objWMIService.ExecQuery("SELECT * FROM Win32_Process WHERE NAME LIKE '" & "%" & "'", , WMIFLAGRETURNIMMEDIATELY + WMIFLAGFORWARDONLY)
    
    For Each objItem In colItems
        Intro.LstProcess.AddItem objItem.Name
        DoEvents
    Next
    
    Set objWMIService = Nothing
    Set colItems = Nothing

End Sub

Public Function sOperatingSystemString() As String

    Dim sOSString As String
    Dim lMaj As Long
    Dim lMin As Long
    Dim lPID As Long
    
    Dim osvVersionInfo As OSVERSIONINFO
    Dim osvexVersionInfo As OSVERSIONINFOEX
    osvVersionInfo.dwOSVersionInfoSize = Len(osvVersionInfo)
    osvexVersionInfo.dwOSVersionInfoSize = Len(osvexVersionInfo)


    If GetVersionEx(osvVersionInfo) <> 0 Then
        lMaj = osvVersionInfo.dwMajorVersion
        lMin = osvVersionInfo.dwMinorVersion
        lPID = osvVersionInfo.dwPlatformId


        Select Case lPID
            Case VER_PLATFORM_WIN32_WINDOWS '9x/ME
            If lMaj = 4 And lMin = 0 Then sOSString = "Windows 95,"
            If lMaj = 4 And lMin = 10 Then sOSString = "Windows 98,"
            If lMaj = 4 And lMin = 90 Then sOSString = "Windows ME,"


            If InStr(osvVersionInfo.szCSDVersion, "A") > 0 Then
                sOSString = sOSString & " Second Edition,"
            End If


            If InStr(osvVersionInfo.szCSDVersion, "C") > 0 Then
                sOSString = sOSString & " OSR2,"
            End If
            Case VER_PLATFORM_WIN32_NT 'NT Based


            If GetVersionExA(osvexVersionInfo) <> 0 Then
                lMaj = osvexVersionInfo.dwMajorVersion
                lMin = osvexVersionInfo.dwMinorVersion
                lPID = osvexVersionInfo.dwPlatformId
                If lMaj = 4 And lMin = 0 Then sOSString = "Windows NT 4.0,"
                If lMaj = 5 And lMin = 0 Then sOSString = "Windows 2000,"
                If lMaj = 5 And lMin = 1 Then sOSString = "Windows XP,"
                If lMaj = 5 And lMin = 2 Then sOSString = "Windows Server 2003 Family,"
                


                If osvexVersionInfo.wSuiteMask = &H300 Then
                    sOSString = sOSString & " Home Edition,"
                Else


                    If lMin = 1 Then
                        sOSString = sOSString & " Professional,"
                    End If
                End If
                
                sOSString = sOSString & " " & osvexVersionInfo.szCSDVersion
                
                sOSString = Left(sOSString, InStr(sOSString, Chr(0)) - 1)
                
                sOSString = sOSString & ", Version "
                
                sOSString = sOSString & osvexVersionInfo.dwMajorVersion
                sOSString = sOSString & "."
                sOSString = sOSString & osvexVersionInfo.dwMinorVersion
                sOSString = sOSString & "."
                sOSString = sOSString & osvexVersionInfo.dwBuildNumber
                
            End If
        End Select
    sOperatingSystemString = sOSString
    Exit Function
Else
    sOperatingSystemString = ""
    Exit Function
End If

End Function

Public Sub GetMemory()
    On Error Resume Next
    
    Dim MEL As MEMORYSTATUS
    GlobalMemoryStatus MEL
    
    With Intro
        .LblOut(2).Caption = Format(MEL.dwTotalPhys / 1048576, "#,#,#.##") & " Mb"
        .LblOut(3).Caption = Format(MEL.dwAvailPhys / 1048576, "#,#,#.##") & " Mb"
        .LblOut(4).Caption = Format(MEL.dwTotalVirtual / 1048576, "#,#,#.##") & " Mb"
        .LblOut(5).Caption = Format(MEL.dwAvailVirtual / 1048576, "#,#,#.##") & " Mb"
        .LblOut(6).Caption = Format(MEL.dwTotalPageFile / 1048576, "#,#,#.##") & " Mb"
        .LblOut(7).Caption = Format(MEL.dwAvailPageFile / 1048576, "#,#,#.##") & " Mb"
    End With
    
End Sub

Public Function GetCPUDescription()

Dim HKey As Long
Dim C As Long
Dim R As Long
Dim S As String
Dim T As Long

R = RegOpenKeyEx(HKEY_LOCAL_MACHINE, "HARDWARE\DESCRIPTION\System\CentralProcessor\0", 0, KEY_READ, HKey)
C = 255
S = String(C, Chr(0))
R = RegQueryValueEx(HKey, "ProcessorNameString", 0, T, S, C)


Intro.LblOut(8).Caption = Trim(Left(S, C - 1))

End Function

Public Sub GetSoundInfo()
    Dim X As WAVEOUTCAPS
    waveOutGetDevCaps 0, X, Len(X)
    
    Intro.LblOut(11).Caption = X.szPname

End Sub

