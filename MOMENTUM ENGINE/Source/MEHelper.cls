VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MEHelper"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit


Public Function ExModeActive() As Boolean

AddLog "Testing surfaces"

If DDraw.TestCooperativeLevel = DD_OK Then
   ExModeActive = True
   AddLog "Surfaces OK"
Else
   ExModeActive = False
   AddLog "Surfaces lost"
End If
     
End Function
'-----------------------------------------------------------------------------


Public Sub RestoreSurfaces()
   DDraw.RestoreAllSurfaces
End Sub
'-----------------------------------------------------------------------------


Public Function UpdateFrameRate() As Integer

FramesDone = FramesDone + 1

If DX.TickCount >= LastTimeCount + 1000 Then
   UpdateFrameRate = CInt(FramesDone)
   LastFrameRate = CInt(FramesDone)
   LastTimeCount = DX.TickCount
   FramesDone = 0
Else
   UpdateFrameRate = LastFrameRate
End If

End Function
'-----------------------------------------------------------------------------


Sub DrawBox(X1 As Long, Y1 As Long, X2 As Long, Y2 As Long)
   BackBufferSurf.DrawBox X1, Y1, X2, Y2
End Sub

Sub DrawCircle(X1 As Long, Y1 As Long, Radius As Long)
   BackBufferSurf.DrawCircle X1, Y1, Radius
End Sub

Sub DrawRoundedBox(X1 As Long, Y1 As Long, X2 As Long, Y2 As Long, RoundW As Long, RoundH As Long)
   BackBufferSurf.DrawRoundedBox X1, Y1, X2, Y2, RoundW, RoundH
End Sub

Sub DrawLine(X1 As Long, Y1 As Long, X2 As Long, Y2 As Long)
   BackBufferSurf.DrawLine X1, Y1, X2, Y2
End Sub
'-----------------------------------------------------------------------------


Function Random(Min As Long, Max As Long) As Long

Randomize Timer
Random = Int(Rnd * (Max - Min)) + Min

End Function
'-----------------------------------------------------------------------------


Sub GetCursor(ByRef x As Long, ByRef Y As Long)

Dim Point As POINTAPI

GetCursorPos Point
ScreenToClient mvarWindowHandle, Point

x = Point.x
Y = Point.Y

End Sub
'-----------------------------------------------------------------------------


Sub SetCursor(ByVal x As Long, ByVal Y As Long)
   
Dim Point As POINTAPI

ClientToScreen mvarWindowHandle, Point
   
SetCursorPos Point.x, Point.Y

End Sub
'-----------------------------------------------------------------------------


Sub CursorHide()
   Call ShowCursor(False)
   AddLog "Cursor hidden"
End Sub
'-----------------------------------------------------------------------------


Sub CursorShow()
   Call ShowCursor(True)
   AddLog "Cursor shown"
End Sub
'-----------------------------------------------------------------------------


Sub KillFile(ByVal Filename As String)
   AddLog Filename & " was deleted"
   DeleteFile Filename
End Sub
'-----------------------------------------------------------------------------


Function DDrawErrorDesc(ByVal ErrorNumber As Long, ByRef ExactErrorReturn As String) As String

Dim ReturnS As String

Select Case ErrorNumber
    Case DD_OK
         ReturnS = "The request completed successfully."
         ExactErrorReturn = "DD_OK"
    Case DDERR_ALREADYINITIALIZED
         ReturnS = "The object has already been initialized."
         ExactErrorReturn = "DDERR_ALREADYINITIALIZED"
    Case DDERR_BLTFASTCANTCLIP
         ReturnS = "A DirectDrawClipper object is attached to a source surface that has passed into a call to the DirectDrawSurface7.BltFast method."
         ExactErrorReturn = "DDERR_BLTFASTCANTCLIP"
    Case DDERR_CANNOTATTACHSURFACE
         ReturnS = "A surface cannot be attached to another requested surface."
         ExactErrorReturn = "DDERR_CANNOTATTACHSURFACE"
    Case DDERR_CANNOTDETACHSURFACE
         ReturnS = "A surface cannot be detached from another requested surface."
         ExactErrorReturn = "DDERR_CANNOTDETACHSURFACE"
    Case DDERR_CANTCREATEDC
         ReturnS = "Windows cannot create any more device contexts (DCs), or a DC was requested for a palette-indexed surface when the surface had no palette and the display mode was not palette-indexed (in this case DirectDraw cannot select a proper palette into the DC)."
         ExactErrorReturn = "DDERR_CANTCREATEDC"
    Case DDERR_CANTDUPLICATE
         ReturnS = "Primary and 3-D surfaces, or surfaces that are implicitly created, cannot be duplicated."
         ExactErrorReturn = "DDERR_CANTDUPLICATE"
    Case DDERR_CANTLOCKSURFACE
         ReturnS = "Access to this surface is refused because an attempt was made to lock the primary surface without DCI support."
         ExactErrorReturn = "DDERR_CANTLOCKSURFACE"
    Case DDERR_CANTPAGELOCK
         ReturnS = "An attempt to page-lock a surface failed. Page lock does not work on a display-memory surface or an emulated primary surface."
         ExactErrorReturn = "DDERR_CANTPAGELOCK"
    Case DDERR_CANTPAGEUNLOCK
         ReturnS = "An attempt to page-unlock a surface failed. Page unlock does not work on a display-memory surface or an emulated primary surface."
         ExactErrorReturn = "DDERR_CANTPAGEUNLOCK"
    Case DDERR_CLIPPERISUSINGHWND
         ReturnS = "An attempt was made to set a clip list for a DirectDrawClipper object that is already monitoring a window handle."
         ExactErrorReturn = "DDERR_CLIPPERISUSINGHWND"
    Case DDERR_COLORKEYNOTSET
         ReturnS = "No source color key is specified for this operation."
         ExactErrorReturn = "DDERR_COLORKEYNOTSET"
    Case DDERR_CURRENTLYNOTAVAIL
         ReturnS = "No support is currently available."
         ExactErrorReturn = "DDERR_CURRENTLYNOTAVAIL"
    Case DDERR_DCALREADYCREATED
         ReturnS = "A device context (DC) has already been returned for this surface. Only one DC can be retrieved for each surface."
         ExactErrorReturn = "DDERR_DCALREADYCREATED"
    Case DDERR_DEVICEDOESNTOWNSURFACE
         ReturnS = "Surfaces created by one DirectDraw device cannot be used directly by another DirectDraw device."
         ExactErrorReturn = "DDERR_DEVICEDOESNOTOWNSURFACE"
    Case DDERR_DIRECTDRAWALREADYCREATED
         ReturnS = "A DirectDraw object representing this driver has already been created for this process."
         ExactErrorReturn = "DDERR_DIRECTDRAWALREADYCREATED"
    Case DDERR_EXCEPTION
         ReturnS = "An exception was encountered while performing the requested operation."
         ExactErrorReturn = "DDERR_EXCEPTION"
    Case DDERR_EXCLUSIVEMODEALREADYSET
         ReturnS = "An attempt was made to set the cooperative level when it was already set to exclusive."
         ExactErrorReturn = "DDERR_EXCLUSIVEMODEALREADYSET"
    Case DDERR_EXPIRED
         ReturnS = "The data has expired and is therefore no longer valid."
         ExactErrorReturn = "DDERR_EXPIRED"
    Case DDERR_GENERIC
         ReturnS = "There is an undefined error condition."
         ExactErrorReturn = "DDERR_GENERIC"
    Case DDERR_HEIGHTALIGN
         ReturnS = "The height of the provided rectangle is not a multiple of the required alignment."
         ExactErrorReturn = "DDERR_HEIGHTALIGN"
    Case DDERR_HWNDALREADYSET
         ReturnS = "The DirectDraw cooperative level window handle has already been set. It cannot be reset while the process has surfaces or palettes created."
         ExactErrorReturn = "DDERR_HWNDALREADYSET"
    Case DDERR_HWNDSUBCLASSED
         ReturnS = "DirectDraw is prevented from restoring state because the DirectDraw cooperative level window handle has been subclassed."
         ExactErrorReturn = "DDERR_HWNDSUBCLASSED"
    Case DDERR_IMPLICITLYCREATED
         ReturnS = "The surface cannot be restored because it is an implicitly created surface."
         ExactErrorReturn = "DDERR_IMPLICITLYCREATED"
    Case DDERR_INCOMPATIBLEPRIMARY
         ReturnS = "The primary surface creation request does not match with the existing primary surface."
         ExactErrorReturn = "DDERR_INCOMPATIBLEPRIMARY"
    Case DDERR_INVALIDCAPS
         ReturnS = "One or more of the capability bits passed to the callback function are incorrect."
         ExactErrorReturn = "DDERR_INVALIDCAPS"
    Case DDERR_INVALIDCLIPLIST
         ReturnS = "DirectDraw does not support the provided clip list."
         ExactErrorReturn = "DDERR_INVALIDCLIPLIST"
    Case DDERR_INVALIDDIRECTDRAWGUID
         ReturnS = "The globally unique identifier (GUID) passed to the DirectX7.DirectDrawCreate function is not a valid DirectDraw driver identifier."
         ExactErrorReturn = "DDERR_INVALIDDIRECTDRAWGUID"
    Case DDERR_INVALIDMODE
         ReturnS = "DirectDraw does not support the requested mode."
         ExactErrorReturn = "DDERR_INVALIDMODE"
    Case DDERR_INVALIDOBJECT
         ReturnS = "DirectDraw received a pointer that was an invalid DirectDraw object."
         ExactErrorReturn = "DDERR_INVALIDOBJECT"
    Case DDERR_INVALIDPARAMS
         ReturnS = "One or more of the parameters passed to the method are incorrect."
         ExactErrorReturn = "DDERR_INVALIDPARAMS"
    Case DDERR_INVALIDPIXELFORMAT
         ReturnS = "The pixel format was invalid as specified."
         ExactErrorReturn = "DDERR_INVALIDPIXELFORMAT"
    Case DDERR_INVALIDPOSITION
         ReturnS = "The position of the overlay on the destination is no longer legal."
         ExactErrorReturn = "DDERR_INVALIDPOSITION"
    Case DDERR_INVALIDRECT
         ReturnS = "The provided rectangle was invalid."
         ExactErrorReturn = "DDERR_INVALIDRECT"
    Case DDERR_INVALIDSTREAM
         ReturnS = "The specified stream contains invalid data."
         ExactErrorReturn = "DDERR_INVALIDSTREAM"
    Case DDERR_INVALIDSURFACETYPE
         ReturnS = "The requested operation could not be performed because the surface was of the wrong type."
         ExactErrorReturn = "DDERR_INVALIDSURFACETYPE"
    Case DDERR_LOCKEDSURFACES
         ReturnS = "One or more surfaces are locked."
         ExactErrorReturn = "DDERR_LOCKEDSURFACES"
    Case DDERR_MOREDATA
         ReturnS = "There is more data available than the specified buffer size can hold."
         ExactErrorReturn = "DDERR_MOREDATA"
    Case DDERR_NO3D
         ReturnS = "No 3-D hardware or emulation is present."
         ExactErrorReturn = "DDERR_NO3D"
    Case DDERR_NOALPHAHW
         ReturnS = "No alpha acceleration hardware is present or available."
         ExactErrorReturn = "DDERR_NOALPHAHW"
    Case DDERR_NOBLTHW
         ReturnS = "No blitter hardware is present."
         ExactErrorReturn = "DDERR_NOBLTHW"
    Case DDERR_NOCLIPLIST
         ReturnS = "No clip list is available."
         ExactErrorReturn = "DDERR_NOCLIPLIST"
    Case DDERR_NOCLIPPERATTACHED
         ReturnS = "No DirectDrawClipper object is attached to the surface object."
         ExactErrorReturn = "DDERR_NOCLIPPERATTACHED"
    Case DDERR_NOCOLORCONVHW
         ReturnS = "No color-conversion hardware is present or available."
         ExactErrorReturn = "DDERR_NOCOLORCONVHW"
    Case DDERR_NOCOLORKEY
         ReturnS = "The surface does not currently have a color key."
         ExactErrorReturn = "DDERR_NOCOLORKEY"
    Case DDERR_NOCOLORKEYHW
         ReturnS = "There is no hardware support for the destination color key."
         ExactErrorReturn = "DDERR_NOCOLORKEYHW"
    Case DDERR_NOCOOPERATIVELEVELSET
         ReturnS = "A create function was called when the DirectDraw7.SetCooperativeLevel method had not been called."
         ExactErrorReturn = "DDERR_NOCOOPERATIVELEVELSET"
    Case DDERR_NODC
         ReturnS = "No DC has ever been created for this surface."
         ExactErrorReturn = "DDERR_NODC"
    Case DDERR_NODDROPSHW
         ReturnS = "No DirectDraw raster operation (ROP) hardware is available."
         ExactErrorReturn = "DDERR_NODDROPSHW"
    Case DDERR_NODIRECTDRAWHW
         ReturnS = "Hardware-only DirectDraw object creation is not possible; the driver does not support any hardware."
         ExactErrorReturn = "DDERR_NODIRECTDRAWHW"
    Case DDERR_NODIRECTDRAWSUPPORT
         ReturnS = "DirectDraw support is not possible with the current display driver."
         ExactErrorReturn = "DDERR_NODIRECTDRAWSUPPORT"
    Case DDERR_NOEMULATION
         ReturnS = "Software emulation is not available."
         ExactErrorReturn = "DDERR_NOEMULATION"
    Case DDERR_NOEXCLUSIVEMODE
         ReturnS = "The operation requires the application to have exclusive mode, but the application does not have exclusive mode."
         ExactErrorReturn = "DDERR_NOEXCLUSIVEMODE"
    Case DDERR_NOFLIPHW
         ReturnS = "Flipping visible surfaces is not supported."
         ExactErrorReturn = "DDERR_NOFLIPHW"
    Case DDERR_NOFOCUSWINDOW
         ReturnS = "An attempt was made to create or set a device window without first setting the focus window."
         ExactErrorReturn = "DDERR_NOFOCUSWINDOW"
    Case DDERR_NOGDI
         ReturnS = "No GDI is present."
         ExactErrorReturn = "DDERR_NOGDI"
    Case DDERR_NOHWND
         ReturnS = "Clipper notification requires a window handle, or no window handle was previously set as the cooperative level window handle."
         ExactErrorReturn = "DDERR_NOHWND"
    Case DDERR_NOMIPMAPHW
         ReturnS = "No mipmap-capable texture mapping hardware is present or available."
         ExactErrorReturn = "DDERR_NOMIPMAPHW"
    Case DDERR_NOMIRRORHW
         ReturnS = "No mirroring hardware is present or available."
         ExactErrorReturn = "DDERR_NOMIRRORHW"
    Case DDERR_NONONLOCALVIDMEM
         ReturnS = "An attempt was made to allocate nonlocal video memory from a device that does not support nonlocal video memory."
         ExactErrorReturn = "DDERR_NONLOCALVIDMEM"
    Case DDERR_NOOPTIMIZEHW
         ReturnS = "The device does not support optimized surfaces."
         ExactErrorReturn = "DDERR_NOOPTIMIZEHW"
    Case DDERR_NOOVERLAYHW
         ReturnS = "No overlay hardware is present or available."
         ExactErrorReturn = "DDERR_NOOVERLAYHW"
    Case DDERR_NOPALETTEATTACHED
         ReturnS = "No palette object is attached to this surface."
         ExactErrorReturn = "DDERR_NOPALETTEATTACHED"
    Case DDERR_NOPALETTEHW
         ReturnS = "There is no hardware support for 16- or 256-color palettes."
         ExactErrorReturn = "DDERR_NOPALETTEHW"
    Case DDERR_NORASTEROPHW
         ReturnS = "No appropriate raster operation hardware is present or available."
         ExactErrorReturn = "DDERR_NORASTEROPHW"
    Case DDERR_NOROTATIONHW
         ReturnS = "No rotation hardware is present or available."
         ExactErrorReturn = "DDERR_NOROTATIONHW"
    Case DDERR_NOSTEREOHARDWARE
         ReturnS = "No stereo hardware is present or available."
         ExactErrorReturn = "DDERR_NOSTEREOHARDWARE"
    Case DDERR_NOSTRETCHHW
         ReturnS = "There is no hardware support for stretching."
         ExactErrorReturn = "DDERR_NOSTRETCHHW"
    Case DDERR_NOSURFACELEFT
         ReturnS = "No hardware is present that supports stereo surfaces."
         ExactErrorReturn = "DDERR_NOSURFACELEFT"
    Case DDERR_NOT4BITCOLOR
         ReturnS = "The DirectDrawSurface object is not using a 4-bit color palette, and the requested operation requires a 4-bit color palette."
         ExactErrorReturn = "DDERR_NOT4BITCOLOR"
    Case DDERR_NOT4BITCOLORINDEX
         ReturnS = "The DirectDrawSurface object is not using a 4-bit color index palette, and the requested operation requires a 4-bit color index palette."
         ExactErrorReturn = "DDERR_NOT4BITCOLORINDEX"
    Case DDERR_NOT8BITCOLOR
         ReturnS = "The DirectDrawSurface object is not using an 8-bit color palette, and the requested operation requires an 8-bit color palette."
         ExactErrorReturn = "DDERR_NOT8BITCOLOR"
    Case DDERR_NOTAOVERLAYSURFACE
         ReturnS = "An overlay component was called for a non-overlay surface."
         ExactErrorReturn = "DDERR_NOTAOVERLAYSURFACE"
    Case DDERR_NOTEXTUREHW
         ReturnS = "No texture-mapping hardware is present or available."
         ExactErrorReturn = "DDERR_NOTEXTUREHW"
    Case DDERR_NOTFLIPPABLE
         ReturnS = "An attempt was made to flip a surface that cannot be flipped."
         ExactErrorReturn = "DDERR_NOTFLIPPABLE"
    Case DDERR_NOTFOUND
         ReturnS = "The requested item was not found."
         ExactErrorReturn = "DDERR_NOTFOUND"
    Case DDERR_NOTINITIALIZED
         ReturnS = "An attempt was made to call an interface method of a DirectDraw object created by CoCreateInstance before the object was initialized."
         ExactErrorReturn = "DDERR_NOTINITIALIZED"
    Case DDERR_NOTLOADED
         ReturnS = "The surface is an optimized surface, but it has not yet been allocated any memory."
         ExactErrorReturn = "DDERR_NOTLOADED"
    Case DDERR_NOTLOCKED
         ReturnS = "An attempt was made to unlock a surface that was not locked."
         ExactErrorReturn = "DDERR_NOTLOCKED"
    Case DDERR_NOTPAGELOCKED
         ReturnS = "An attempt was made to page-unlock a surface with no outstanding page locks."
         ExactErrorReturn = "DDERR_NOTPAGELOCKED"
    Case DDERR_NOTPALETTIZED
         ReturnS = "The surface being used is not a palette-based surface."
         ExactErrorReturn = "DDERR_NOTPALETTIZED"
    Case DDERR_NOVSYNCHW
         ReturnS = "There is no hardware support for vertical blank synchronized operations."
         ExactErrorReturn = "DDERR_NOVSYNCHW"
    Case DDERR_NOZBUFFERHW
         ReturnS = "There is no hardware support for z-buffers."
         ExactErrorReturn = "DDERR_NOZBUFFERHW"
    Case DDERR_NOZOVERLAYHW
         ReturnS = "The hardware does not support z-ordering of overlays."
         ExactErrorReturn = "DDERR_NOZOVERLAYHW"
    Case DDERR_OUTOFCAPS
         ReturnS = "The hardware needed for the requested operation has already been allocated."
         ExactErrorReturn = "DDERR_OUTOFCAPS"
    Case DDERR_OUTOFMEMORY
         ReturnS = "DirectDraw does not have enough memory to perform the operation."
         ExactErrorReturn = "DDERR_OUFOFMEMORY"
    Case DDERR_OUTOFVIDEOMEMORY
         ReturnS = "DirectDraw does not have enough display memory to perform the operation."
         ExactErrorReturn = "DDERR_OUTOFVIDEOMEMORY"
    Case DDERR_OVERLAPPINGRECTS
         ReturnS = "The source and destination rectangles are on the same surface and overlap each other."
         ExactErrorReturn = "DDERR_OVERLAPPINGRECTS"
    Case DDERR_OVERLAYCANTCLIP
         ReturnS = "The hardware does not support clipped overlays."
         ExactErrorReturn = "DDERR_OVERLAYCANTCLIP"
    Case DDERR_OVERLAYCOLORKEYONLYONEACTIVE
         ReturnS = "An attempt was made to have more than one color key active on an overlay."
         ExactErrorReturn = "DDERR_OVERLAYCOLORKEYONLYONEACTIVE"
    Case DDERR_OVERLAYNOTVISIBLE
         ReturnS = "The method was called on a hidden overlay."
         ExactErrorReturn = "DDERR_OVERLAYNOTVISIBLE"
    Case DDERR_PALETTEBUSY
         ReturnS = "Access to this palette is refused because the palette is locked by another thread."
         ExactErrorReturn = "DDERR_PALETTEBUSY"
    Case DDERR_PRIMARYSURFACEALREADYEXISTS
         ReturnS = "This process has already created a primary surface."
         ExactErrorReturn = "DDERR_PRIMARYSURFACEALREADYEXISTS"
    Case DDERR_REGIONTOOSMALL
         ReturnS = "The region passed to the DirectDrawClipper.GetClipList method is too small."
         ExactErrorReturn = "DDERR_REGIONTOOSMALL"
    Case DDERR_SURFACEALREADYATTACHED
         ReturnS = "An attempt was made to attach a surface to another surface to which it is already attached."
         ExactErrorReturn = "DDERR_SURFACEALREADYATTACHED"
    Case DDERR_SURFACEALREADYDEPENDENT
         ReturnS = "An attempt was made to make a surface a dependency of another surface on which it is already dependent."
         ExactErrorReturn = "DDERR_SURFACEALREADYDEPENDANT"
    Case DDERR_SURFACEBUSY
         ReturnS = "Access to the surface is refused because the surface is locked by another thread."
         ExactErrorReturn = "DDERR_SURFACEBUSY"
    Case DDERR_SURFACEISOBSCURED
         ReturnS = "Access to the surface is refused because the surface is obscured."
         ExactErrorReturn = "DDERR_SURFACEISOBSCURED"
    Case DDERR_SURFACELOST
         ReturnS = "Access to the surface is refused because the surface memory is gone. Call the DirectDrawSurface7.Restore method on this surface to restore the memory associated with it."
         ExactErrorReturn = "DDERR_SURFACELOST"
    Case DDERR_SURFACENOTATTACHED
         ReturnS = "The requested surface is not attached."
         ExactErrorReturn = "DDERR_SURFACENOTATTACHED"
    Case DDERR_TOOBIGHEIGHT
         ReturnS = "The height requested by DirectDraw is too large."
         ExactErrorReturn = "DDERR_TOOBIGHEIGHT"
    Case DDERR_TOOBIGSIZE
         ReturnS = "The size requested by DirectDraw is too large. However, the individual height and width are valid sizes."
         ExactErrorReturn = "DDERR_TOOBIGSIZE"
    Case DDERR_TOOBIGWIDTH
         ReturnS = "The width requested by DirectDraw is too large."
         ExactErrorReturn = "DDERR_TOOBIGWIDTH"
    Case DDERR_UNSUPPORTED
         ReturnS = "The operation is not supported."
         ExactErrorReturn = "DDERR_UNSUPPORTED"
    Case DDERR_UNSUPPORTEDFORMAT
         ReturnS = "The FourCC format requested is not supported by DirectDraw."
         ExactErrorReturn = "DDERR_INSUPPORTEDFORMAT"
    Case DDERR_UNSUPPORTEDMASK
         ReturnS = "The bitmask in the pixel format requested is not supported by DirectDraw."
         ExactErrorReturn = "DDERR_UNSUPPORTEDMASK"
    Case DDERR_UNSUPPORTEDMODE
         ReturnS = "The display is currently in an unsupported mode."
         ExactErrorReturn = "DDERR_UNSUPPORTEDMODE"
    Case DDERR_VERTICALBLANKINPROGRESS
         ReturnS = "A vertical blank is in progress."
         ExactErrorReturn = "DDERR_VERTICALBLANKINPROGRESS"
    Case DDERR_VIDEONOTACTIVE
         ReturnS = "The video port is not active."
         ExactErrorReturn = "DDERR_VIDEONOTACTIVE"
    Case DDERR_WASSTILLDRAWING
         ReturnS = "The previous blit operation that is transferring information to or from this surface is incomplete."
         ExactErrorReturn = "DDERR_WASSTILLDRAWING"
    Case DDERR_WRONGMODE
         ReturnS = "This surface cannot be restored because it was created in a different mode."
         ExactErrorReturn = "DDERR_WRONGMODE"
    Case DDERR_XALIGN
         ReturnS = "The provided rectangle was not horizontally aligned on a required boundary."
         ExactErrorReturn = "DDERR_XALIGN"
    Case Else
        ReturnS = "Error number did no match a DirectDraw Error Code"
        ExactErrorReturn = "<NONE>"
End Select

DDrawErrorDesc = ReturnS

End Function
'-----------------------------------------------------------------------------

