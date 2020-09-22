VERSION 5.00
Begin VB.Form frmGetOSInfo 
   AutoRedraw      =   -1  'True
   BackColor       =   &H8000000A&
   Caption         =   "Get OS Version Info"
   ClientHeight    =   6345
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5955
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6345
   ScaleWidth      =   5955
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton frmExit 
      Caption         =   "Exit"
      Height          =   345
      Left            =   3975
      TabIndex        =   2
      Top             =   6000
      Width           =   1980
   End
   Begin VB.CommandButton cmdCopy 
      Caption         =   "Copy"
      Height          =   345
      Left            =   15
      TabIndex        =   1
      Top             =   6000
      Width           =   1980
   End
   Begin VB.TextBox txtOS 
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   5970
      Left            =   15
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   0
      Width           =   5955
   End
End
Attribute VB_Name = "frmGetOSInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCopy_Click()
txtOS.SelStart = 192
txtOS.SelLength = Len(txtOS.Text)

Clipboard.Clear
Clipboard.SetText txtOS.SelText

MsgBox "OS Version Info copied to clipboard", vbOKOnly, "Info copied"
End Sub

Private Sub Form_Load()
txtOS = "************************************" & vbCrLf & _
        "* Made by: Amir Ahmetovic          *" & vbCrLf & _
        "* Year: 2002                       *" & vbCrLf & _
        "* Published on: Planet Source Code *" & vbCrLf & _
        "************************************" & vbCrLf & vbCrLf

txtOS = txtOS & "OS Name : " & GetOSName & vbCrLf
txtOS = txtOS & "Platform: " & GetOSPlatformType & vbCrLf
txtOS = txtOS & "----------------------------------------------" & vbCrLf
txtOS = txtOS & "OS Major Version: " & GetOSMajorVer & vbCrLf
txtOS = txtOS & "OS Minor Version: " & GetOSMinorVer & vbCrLf
txtOS = txtOS & "OS Build Version: " & GetOSBuildVer & vbCrLf
txtOS = txtOS & "----------------------------------------------" & vbCrLf
txtOS = txtOS & "OS Platform ID: " & GetOSPlatformID & vbCrLf
txtOS = txtOS & "OS CSD String : " & GetOSCSDVer & vbCrLf
txtOS = txtOS & "----------------------------------------------" & vbCrLf
txtOS = txtOS & "OS Service Pack    : " & GetOSServicePack & vbCrLf
txtOS = txtOS & "OS SP Major Version: " & GetOSSPMajorVer & vbCrLf
txtOS = txtOS & "OS SP Minor Version: " & GetOSSPMinorVer & vbCrLf
txtOS = txtOS & "----------------------------------------------" & vbCrLf
txtOS = txtOS & "Is Windows 3.x    : " & IIf(IsWindows3x, "Yes", "No") & vbCrLf
txtOS = txtOS & "Is Windows 95     : " & IIf(IsWindows95, "Yes", "No") & vbCrLf
txtOS = txtOS & "Is Windows 95 OSR2: " & IIf(IsWindows95OSR2, "Yes", "No") & vbCrLf
txtOS = txtOS & "Is Windows 98     : " & IIf(IsWindows98, "Yes", "No") & vbCrLf
txtOS = txtOS & "Is Windows 98 SE  : " & IIf(IsWindows98SE, "Yes", "No") & vbCrLf
txtOS = txtOS & "Is Windows Me     : " & IIf(IsWindowsMe, "Yes", "No") & vbCrLf
txtOS = txtOS & "Is Windows NT     : " & IIf(IsWindowsNT, "Yes", "No") & vbCrLf
txtOS = txtOS & "Is Windows NT 4.0 : " & IIf(IsWindowsNT4, "Yes", "No") & vbCrLf
txtOS = txtOS & "Is Windows 2000   : " & IIf(IsWindows2000, "Yes", "No") & vbCrLf
txtOS = txtOS & "Is Windows XP     : " & IIf(IsWindowsXP, "Yes", "No")
End Sub

Private Sub frmExit_Click()
End
End Sub
