VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MEInstallDX"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Type VERSION
   lMajor As Long
   lMinor As Long
End Type
'-----------------------------------------------------------------------------


Public Function GetInstalledVersion() As VERSION

AddLog "Getting installed version of DirectX"

DirectXSetupGetVersion GetInstalledVersion.lMajor, GetInstalledVersion.lMinor

GetInstalledVersion.lMajor = GetInstalledVersion.lMajor - (GetInstalledVersion.lMajor And &H40000)
GetInstalledVersion.lMinor = (GetInstalledVersion.lMinor - (100 * GetInstalledVersion.lMajor))

AddLog "Version is " & GetInstalledVersion.lMajor & "." & GetInstalledVersion.lMinor

End Function
'-----------------------------------------------------------------------------


Public Function InstallNeeded() As Boolean

InstallNeeded = False

AddLog "Testing for installation"

If GetInstalledVersion.lMajor < 7 Then
   InstallNeeded = True
   AddLog "Installation needed"
Else
   AddLog "No installation required"
End If

End Function
'-----------------------------------------------------------------------------


Public Sub InstallDX(ByVal RedistFolder As String)
   AddLog "Installing directX"
   DirectXSetup mvarWindowHandle, RedistFolder, 65560
End Sub
'-----------------------------------------------------------------------------


Public Sub RestartWindows()
   AddLog "Restarting windows"
   ExitWindowsEx 2, 0
End Sub
'-----------------------------------------------------------------------------
