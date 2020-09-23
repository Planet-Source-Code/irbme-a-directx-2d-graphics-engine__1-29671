VERSION 5.00
Begin VB.Form FrmMain 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "FrmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Create variables taht can store all the new classes
Dim mBackBuffer As New MEBackBuffer
Dim mFont As New MEFont
Dim mGamma As New MEGammaramp
Dim mHelper As New MEHelper
Dim mInit As New MEInitialise
Dim mLight As New MELighting
Dim mLogFile As New MELogFile
Dim mScreen As New MEScreen
Dim mSurface As New MESurface
Dim mTerm As New METerminate

'To hold wether or not we are should quit or not
Dim Running As Boolean

'Quite by setting the running flag to false. _
This will be detected in the next loop and the program will end
Private Sub Form_Click()
   Running = False
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

'Check for the keys to update the gamma, and do so if they are
Select Case KeyCode

Case vbKeyPageUp
   GammaG = GammaG + 1
Case vbKeyPageDown
   GammaG = GammaG - 1
Case vbKeyHome
   GammaB = GammaB + 1
Case vbKeyEnd
   GammaB = GammaB - 1
Case vbKeyInsert
   GammaR = GammaR + 1
Case vbKeyDelete
   GammaR = GammaR - 1
End Select

'Make sure the gamma values is within the limits
If GammaR > 99 Then GammaR = 99
If GammaR < -99 Then GammaR = -99

If GammaB > 99 Then GammaB = 99
If GammaB < -99 Then GammaB = -99

If GammaG > 99 Then GammaG = 99
If GammaG < -99 Then GammaG = -99


End Sub

'This is where everything happens
Private Sub Form_Load()

Dim CursX As Long, CursY As Long

'Create a new instance of all the classes
Set mBackBuffer = New MEBackBuffer
Set mFont = New MEFont
Set mGamma = New MEGammaramp
Set mHelper = New MEHelper
Set mInit = New MEInitialise
Set mLight = New MELighting
Set mLogFile = New MELogFile
Set mScreen = New MEScreen
Set mSurface = New MESurface
Set mTerm = New METerminate

'If the user wants to use a log file then,
If UseLogFile Then
   'Start the logger class
   mLogFile.AutoLog = True
   mLogFile.FileName = LogFile
   mLogFile.StartLog
End If

'Hide the default cursor
mHelper.CursorHide

'Set some flags...pretty self explanitory
mScreen.ColorDepth = ColDepth
mScreen.ResolutionX = ResX
mScreen.ResolutionY = ResY
mScreen.WindowHandle = Me.hWnd

'Backbuffer count is the number of backbuffers we want...
'N.B. 2 is the optimum amount I think, once you start to add more, _
it takes longer again.
'(Note, the more you have, the more memory is required)
mBackBuffer.BackBufferCount = 1
mBackBuffer.BackColor = vbBlack
mBackBuffer.ForeColor = vbBlack

'Call the initialisation sub
mInit.InitialiseME

'Setup a simple font
mFont.FontBackColor = vbBlack
mFont.FontName = "Arial"
mFont.FontSize = 10
mFont.FontTransparent = True

'Load our surfaces
mSurface.LoadSurface App.Path & "\Lake.Bmp", ResX, ResY
mSurface.LoadSurface App.Path & "\Mouse.Bmp", 32, 32

'Create the font
mFont.CreateFont

'Create a gamma ramp. (NOTE, NOT ALL COMPUTERS SUPPORT THIS AND TO KEEP THIS EXAMLE SIMPLE, I HAVEN'T CHECKED)
'So if the demo dousn't work for you, This might be the cause. If you remove it, remember and remove the update gamma in the loop too
mGamma.CreateGammaRamp

'Create a light. For some reason, it only supports a radius in two's compliment (1,2,4,8,16,32,64,128) After 128, I htink it starts to crash, and I haven't tried below 32.
'You can mess around with the multiplier to get the size you want, but I haven't tested this with anything other than 1
mLight.CreateLight 128, 1, 0, 255

'MAKE SURE YOU SHOW THE FORM. This is only form_load
Me.Show
Running = True

'Start the loop
While Running
   'Get the cursor position
   mHelper.GetCursor CursX, CursY
   'Draw the background
   mSurface.DrawSurface 0, 0, 1, False
   'Draw the font to the screen
   mFont.DrawFont 20, 20, mHelper.UpdateFrameRate
   mFont.DrawFont 20, 35, "Change Red: Insert/Delete: " & GammaR
   mFont.DrawFont 20, 50, "Change Blue: Home/End: " & GammaB
   mFont.DrawFont 20, 65, "Change Green: PgUp/PgDwn: " & GammaG
   'Draw the mouse surface at the cursor position
   mSurface.DrawSurface CursX, CursY, 2, True
   'Show the light at the cursor position/ If you call the DrawLight sub, _
   the actual physical source of the light will be drawn, a solid object.
   mLight.ShowLight 1, CursX, CursY
   
   'Update the gamma. Remember if you remove the CreateGammaRamp sub, remove this too.
   mGamma.UpdateGamma GammaR, GammaG, GammaB
   'Flip from the backbuffer to the screen
   mBackBuffer.Flip
   'Make sure windows has some breathing time.
   DoEvents
Wend

'If we get to here then the running variable must have been _
set to false, therefore ending the loop, therefore we better start the cleanup

'Show the cursor again
mHelper.CursorShow

'Call te termination method of the engine
mTerm.TerminateME

'Destroy all our classes. There are quite alot so we need to free the space
Set mBackBuffer = Nothing
Set mFont = Nothing
Set mGamma = Nothing
Set mHelper = Nothing
Set mInit = Nothing
Set mLight = Nothing
Set mLogFile = Nothing
Set mScreen = Nothing
Set mSurface = Nothing
Set mTerm = Nothing

'DUH! End the program
End

End Sub
