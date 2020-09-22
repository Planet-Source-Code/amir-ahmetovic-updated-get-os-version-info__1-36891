Attribute VB_Name = "mGetOSName"
Option Explicit

Private Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" ( _
     lpVersionInformation As OSVERSIONINFOEX) As Long
Private Const VS_FF_DEBUG = &H1&
Private Const VS_FF_INFOINFERRED = &H10&
Private Const VS_FF_PATCHED = &H4&
Private Const VS_FF_PRERELEASE = &H2&
Private Const VER_NT_SERVER = &H3
Private Const VER_NT_WORKSTATION = &H1
Private Const VER_PLATFORM_WIN32_NT = 2
Private Const VER_PLATFORM_WIN32_WINDOWS = 1
Private Const VER_PLATFORM_WIN32s = 0
Private Const VER_PLATFORMID = &H8
Private Const VER_SERVER_NT = &H80000000
Private Const VER_SUITE_DATACENTER = &H80
Private Const VER_SUITE_ENTERPRISE = &H2
Private Const VER_SUITE_PERSONAL = &H200
Private Const VER_WORKSTATION_NT = &H40000000
Private Type OSVERSIONINFOEX
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
Private b_Win3x As Boolean
Private b_Win95 As Boolean
Private b_Win95OSR2 As Boolean
Private b_Win98 As Boolean
Private b_Win98SE As Boolean
Private b_WinMe As Boolean
Private b_WinNT As Boolean
Private b_WinNT4 As Boolean
Private b_Win2000 As Boolean
Private b_WinXP As Boolean
Dim OSV As OSVERSIONINFOEX
Public Function WordLo(LongIn As Long) As Integer
   ' Low word retrieved by masking off high word.
   ' If low word is too large, twiddle sign bit.
   If (LongIn And &HFFFF&) > &H7FFF Then
      WordLo = (LongIn And &HFFFF&) - &H10000
   Else
      WordLo = LongIn And &HFFFF&
   End If
End Function

Public Function GetOSName() As String
Dim nNull
Dim OS As String
OSV.dwOSVersionInfoSize = Len(OSV)
Call GetVersionEx(OSV)

   nNull = InStr(OSV.szCSDVersion, vbNullChar)
   If nNull > 1 Then
      OSV.szCSDVersion = Left(OSV.szCSDVersion, nNull - 1)
   ElseIf nNull = 1 Then
      OSV.szCSDVersion = "None"
   End If

Select Case OSV.dwPlatformId
    Case VER_PLATFORM_WIN32_WINDOWS
        If OSV.dwMajorVersion = 4 And OSV.dwMinorVersion = 0 And WordLo(OSV.dwBuildNumber) = "950" Then
            OS = "Microsoft Windows 95 "
            b_Win95 = True
          ElseIf OSV.dwMajorVersion = 4 And OSV.dwMinorVersion = 0 And WordLo(OSV.dwBuildNumber) = "1111" And Trim(OSV.szCSDVersion = "B") Then
            OS = "Microsoft Windows 95 OSR2 "
            b_Win95OSR2 = True
          ElseIf OSV.dwMajorVersion = 4 And OSV.dwMinorVersion = 10 And WordLo(OSV.dwBuildNumber) = "1998" Then
            OS = "Microsoft Windows 98 "
            b_Win98 = True
          ElseIf OSV.dwMajorVersion = 4 And OSV.dwMinorVersion = 10 And WordLo(OSV.dwBuildNumber) = "2222" And Trim(OSV.szCSDVersion) = "A" Then
            OS = "Microsoft Windows 98 Second Edition "
            b_Win98SE = True
          ElseIf OSV.dwMajorVersion = 4 And OSV.dwMinorVersion = 90 And WordLo(OSV.dwBuildNumber) = "3000" Then
            OS = "Microsoft Windows Me "
            b_WinMe = True
        End If
    
    Case VER_PLATFORM_WIN32_NT
        If OSV.dwMajorVersion = 4 And OSV.dwMinorVersion = 0 And WordLo(OSV.dwBuildNumber) = "1381" Then
            OS = "Microsoft Windows NT "
            b_WinNT = True
          ElseIf OSV.dwMajorVersion = 4 And OSV.dwMinorVersion = 0 And WordLo(OSV.dwBuildNumber) = "1381" And Trim(OSV.szCSDVersion = "Service Pack 5") Then
            OS = "Microsoft Windows NT 4.0 "
            b_WinNT4 = True
          ElseIf OSV.dwMajorVersion = 5 And OSV.dwMinorVersion = 0 And WordLo(OSV.dwBuildNumber) = "2195" Then
            OS = "Microsoft Windows 2000 "
            b_Win2000 = True
          ElseIf OSV.dwMajorVersion = 5 And OSV.dwMinorVersion = 1 And WordLo(OSV.dwBuildNumber) = "2600" Then
            OS = "Microsoft Windows XP "
            b_WinXP = True
        End If
        
        If OSV.wProductType = VER_NT_WORKSTATION Then
            If OSV.dwMajorVersion = 4 And OSV.dwMinorVersion = 0 Then OS = OS & "Workstation "
            If OSV.wSuiteMask And VER_SUITE_PERSONAL Then
                OS = OS & "Home Edition "
              Else
                OS = OS & "Professional"
            End If
          ElseIf OSV.wProductType = VER_NT_SERVER Then
            If OSV.wSuiteMask And VER_SUITE_DATACENTER Then
                OS = OS & "Data Center "
              ElseIf OSV.wSuiteMask And VER_SUITE_ENTERPRISE Then
                OS = OS & "Enterprise "
              Else
                OS = OS & "Server "
            End If
        End If
        
    Case VER_PLATFORM_WIN32s
        OS = "Microsoft Windows 3.1 "
        b_Win3x = True
End Select
GetOSName = OS
End Function

Public Function GetOSBuildVer() As Long
OSV.dwOSVersionInfoSize = Len(OSV)
Call GetVersionEx(OSV)

GetOSBuildVer = WordLo(OSV.dwBuildNumber)
End Function

Public Function GetOSMajorVer() As Long
OSV.dwOSVersionInfoSize = Len(OSV)
Call GetVersionEx(OSV)

GetOSMajorVer = OSV.dwMajorVersion
End Function

Public Function GetOSMinorVer() As Long
OSV.dwOSVersionInfoSize = Len(OSV)
Call GetVersionEx(OSV)

GetOSMinorVer = OSV.dwMinorVersion
End Function

Public Function GetOSPlatformID() As Long
OSV.dwOSVersionInfoSize = Len(OSV)
Call GetVersionEx(OSV)

GetOSPlatformID = OSV.dwPlatformId
End Function

Public Function GetOSCSDVer() As String
Dim nNull
OSV.dwOSVersionInfoSize = Len(OSV)
Call GetVersionEx(OSV)

   nNull = InStr(OSV.szCSDVersion, vbNullChar)
   If nNull > 1 Then
      OSV.szCSDVersion = Left(OSV.szCSDVersion, nNull - 1)
   ElseIf nNull = 1 Then
      OSV.szCSDVersion = ""
   End If

GetOSCSDVer = Trim(OSV.szCSDVersion)
End Function

Public Function GetOSPlatformType() As String
OSV.dwOSVersionInfoSize = Len(OSV)
Call GetVersionEx(OSV)

Select Case OSV.dwPlatformId
    Case VER_PLATFORM_WIN32_WINDOWS
        GetOSPlatformType = "Windows 9x/Me"
    Case VER_PLATFORM_WIN32_NT
        GetOSPlatformType = "Windows NT/2000/XP"
    Case VER_PLATFORM_WIN32s
        GetOSPlatformType = "Windows 3.x"
End Select
End Function

Public Function GetOSServicePack() As String
OSV.dwOSVersionInfoSize = Len(OSV)
Call GetVersionEx(OSV)

If OSV.wServicePackMajor > 0 Then
    GetOSServicePack = "SP" & OSV.wServicePackMajor
  Else
    GetOSServicePack = "None"
End If
End Function

Public Function GetOSSPMajorVer() As Long
OSV.dwOSVersionInfoSize = Len(OSV)
Call GetVersionEx(OSV)

GetOSSPMajorVer = OSV.wServicePackMajor
End Function

Public Function GetOSSPMinorVer() As Long
OSV.dwOSVersionInfoSize = Len(OSV)
Call GetVersionEx(OSV)

GetOSSPMinorVer = OSV.wServicePackMinor
End Function

Public Property Get IsWindows3x() As Boolean
    IsWindows3x = b_Win3x
End Property

Public Property Get IsWindows95() As Boolean
    IsWindows95 = b_Win95
End Property

Public Property Get IsWindows95OSR2() As Boolean
    IsWindows95OSR2 = b_Win95OSR2
End Property

Public Property Get IsWindows98() As Boolean
    IsWindows98 = b_Win98
End Property

Public Property Get IsWindows98SE() As Boolean
    IsWindows98SE = b_Win98SE
End Property

Public Property Get IsWindowsMe() As Boolean
    IsWindowsMe = b_WinMe
End Property

Public Property Get IsWindowsNT() As Boolean
    IsWindowsNT = b_WinNT
End Property

Public Property Get IsWindowsNT4() As Boolean
    IsWindowsNT4 = b_WinNT4
End Property

Public Property Get IsWindows2000() As Boolean
    IsWindows2000 = b_Win2000
End Property

Public Property Get IsWindowsXP() As Boolean
    IsWindowsXP = b_WinXP
End Property
