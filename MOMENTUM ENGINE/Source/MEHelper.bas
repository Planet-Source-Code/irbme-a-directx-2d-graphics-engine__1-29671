Attribute VB_Name = "MEGlobals"
Option Explicit

'Main objects
Public DX As New DirectX7
Public DDraw As DirectDraw7
Public D3D As Direct3D7
Public D3DDevice As Direct3DDevice7

'Backbuffer
Public mvarBackColor As Long
Public mvarForeColor As Long
Public mvarBackBufferCount As Long
Public BackBufferSurf As DirectDrawSurface7
Public BackbufferDesc As DDSURFACEDESC2

'Screen
Public mvarResolutionX As Long
Public mvarResolutionY As Long
Public mvarColorDepth As Long
Public mvarWindowHandle As Long
Public ScreenSurf As DirectDrawSurface7
Public ScreenDesc As DDSURFACEDESC2

'Surfaces
Public surface() As DirectDrawSurface7
Public SurfaceDesc() As DDSURFACEDESC2
Public mvarNumberSurfaces As Long

'Font
Public mvarFontName As String
Public mvarFontSize As Long
Public mvarFontBackColor As Long
Public mvarFontTransparent As Boolean
Public FontX As New StdFont

'Gamma Correction
Public GammaSupport As Boolean
Public GammaControler As DirectDrawGammaControl
Public GammaRamp As DDGAMMARAMP
Public OriginalRamp As DDGAMMARAMP
Public CurrRed As Integer
Public CurrGreen As Integer
Public CurrBlue As Integer

'Frame rate
Public FramesDone As Integer
Public LastTimeCount As Long
Public LastFrameRate As Integer

'Log file
Public mvarFileName As String
Public mvarFileNumber As Integer
Public mvarAutoLog As Boolean
'-----------------------------------------------------------------------------

Public mTimer As METimer
Public mSubclass As MESubclass


Type POINTAPI
   x As Long
   Y As Long
End Type

Public Type LARGE_INTEGER
   T As Long
   R As Long
End Type
'-----------------------------------------------------------------------------


Public Declare Function DeleteFile Lib "kernel32" Alias "DeleteFileA" (ByVal lpFileName As String) As Long
Public Declare Function ClientToScreen Lib "user32" (ByVal hWnd As Long, lpPoint As POINTAPI) As Long
Public Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Public Declare Function ScreenToClient Lib "user32" (ByVal hWnd As Long, lpPoint As POINTAPI) As Long
Public Declare Function SetCursorPos Lib "user32" (ByVal x As Long, ByVal Y As Long) As Long
Public Declare Function ShowCursor Lib "user32" (ByVal bShow As Long) As Long

Public Declare Function DirectXSetup Lib "dsetup.dll" Alias "DirectXSetupA" (ByVal hWnd As Long, ByVal lpszRootPath As String, ByVal dwFlags As Long) As Long
Public Declare Function DirectXSetupGetVersion Lib "dsetup.dll" (dwVersion As Long, dwRevision As Long) As Long

Public Declare Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Public Declare Function ExitWindowsEx Lib "user32" (ByVal uFlags As Long, ByVal dwReserved As Long) As Long
'-----------------------------------------------------------------------------


Public Function LargeIntToCurrency(ByRef liInput As LARGE_INTEGER) As Currency
    CopyMemory LargeIntToCurrency, liInput, LenB(liInput)
    LargeIntToCurrency = LargeIntToCurrency * 10000
End Function
'-----------------------------------------------------------------------------


Public Function RGB2DX(R As Long, G As Long, B As Long) As Long
    RGB2DX = DX.CreateColorRGBA(CSng((1 / 255) * R), CSng((1 / 255) * G), CSng((1 / 255) * B), 0)
End Function
'-----------------------------------------------------------------------------


Public Sub AddLog(ByVal LogLine As String)
   If mvarAutoLog And Len(mvarFileName) > 3 Then
      Print #mvarFileNumber, LogLine
   End If
End Sub
'-----------------------------------------------------------------------------
