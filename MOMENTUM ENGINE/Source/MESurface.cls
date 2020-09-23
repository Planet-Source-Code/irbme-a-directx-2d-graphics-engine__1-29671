VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MESurface"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit


Public Event ErrorOccur(ByVal Procedure As String, ByVal ErrorCode As Long)
'-----------------------------------------------------------------------------


Private mvarColorKeyHi As Long
Private mvarColorKeyLo As Long
'-----------------------------------------------------------------------------


Public Property Let ColorKeyLo(ByVal lData As Long)
    mvarColorKeyLo = lData
    AddLog "Colorkeylo = " & lData
End Property

Public Property Get ColorKeyLo() As Long
    ColorKeyLo = mvarColorKeyLo
End Property
'-----------------------------------------------------------------------------


Public Property Let ColorKeyHi(ByVal lData As Long)
    mvarColorKeyHi = lData
   AddLog "Colorkeyhi = " & lData
End Property

Public Property Get ColorKeyHi() As Long
    ColorKeyHi = mvarColorKeyHi
End Property
'-----------------------------------------------------------------------------


Public Property Get NumberSurfaces() As Long
    NumberSurfaces = mvarNumberSurfaces
End Property
'-----------------------------------------------------------------------------


Public Sub LoadSurface(ByVal Filename As String, Optional ByVal Width As Long = 0, Optional ByVal Height As Long = 0)

On Local Error GoTo Error

Dim CKey As DDCOLORKEY

AddLog "Load surface sub called"

mvarNumberSurfaces = mvarNumberSurfaces + 1
AddLog "Number of surfaces = " & mvarNumberSurfaces

AddLog "Redimensioning dynamic arrays with reservation of data"
ReDim Preserve surface(1 To mvarNumberSurfaces)
ReDim Preserve SurfaceDesc(1 To mvarNumberSurfaces)

AddLog "Clearing surface"
Set surface(mvarNumberSurfaces) = Nothing

AddLog "Setting flags"
SurfaceDesc(mvarNumberSurfaces).lFlags = DDSD_CAPS Or DDSD_HEIGHT Or DDSD_WIDTH
SurfaceDesc(mvarNumberSurfaces).ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN
SurfaceDesc(mvarNumberSurfaces).ddsCaps.lCaps = DDSCAPS_TEXTURE
SurfaceDesc(mvarNumberSurfaces).ddsCaps.lCaps2 = DDSCAPS2_TEXTUREMANAGE
SurfaceDesc(mvarNumberSurfaces).lTextureStage = 1

If Width > 0 And Height > 0 Then
   AddLog "Using width " & Width & " and height " & Height
   SurfaceDesc(mvarNumberSurfaces).lWidth = Width
   SurfaceDesc(mvarNumberSurfaces).lHeight = Height
End If

AddLog "Setting surface"
Set surface(mvarNumberSurfaces) = DDraw.CreateSurfaceFromFile(Filename, SurfaceDesc(mvarNumberSurfaces))

AddLog "Setting color keys"
CKey.high = mvarColorKeyHi
CKey.low = mvarColorKeyLo

surface(mvarNumberSurfaces).SetColorKey DDCKEY_SRCBLT, CKey

AddLog "Load surface was successful"

Exit Sub
Error:
   AddLog "Load surface failed"
   RaiseEvent ErrorOccur("MESurface: Loasdsurface", Err.Number)

End Sub
'-----------------------------------------------------------------------------


Public Sub DrawSurface(ByVal DestX As Long, ByVal DestY As Long, ByVal SurfaceIndex As Integer, ByVal Transparent As Boolean)

On Local Error GoTo Error

Dim SrcRect As RECT
Dim RetVal As Long

With SrcRect
   .Left = 0
   .Top = 0
   .Bottom = SurfaceDesc(SurfaceIndex).lHeight
   .Right = SurfaceDesc(SurfaceIndex).lWidth
End With

If Transparent Then
   RetVal = BackBufferSurf.BltFast(DestX, DestY, surface(SurfaceIndex), SrcRect, DDBLTFAST_SRCCOLORKEY Or DDBLTFAST_WAIT)
Else
   RetVal = BackBufferSurf.BltFast(DestX, DestY, surface(SurfaceIndex), SrcRect, DDBLTFAST_WAIT)
End If

Exit Sub
Error:
   AddLog "Drawing of surface failed"
   RaiseEvent ErrorOccur("MESurface: DrawSurface", Err.Number)

End Sub
'-----------------------------------------------------------------------------
