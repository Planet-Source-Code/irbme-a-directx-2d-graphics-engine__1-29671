VERSION 5.00
Begin VB.Form FrmSetup 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "--------------------------Momentum Engine - 1.0 - Beta demo-------------------------"
   ClientHeight    =   4170
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   6435
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4170
   ScaleWidth      =   6435
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.HScrollBar B 
      Height          =   255
      Left            =   1080
      Max             =   99
      Min             =   -99
      TabIndex        =   18
      Top             =   3240
      Width           =   5175
   End
   Begin VB.HScrollBar G 
      Height          =   255
      Left            =   1080
      Max             =   99
      Min             =   -99
      TabIndex        =   17
      Top             =   2880
      Width           =   5175
   End
   Begin VB.HScrollBar R 
      Height          =   255
      Left            =   1080
      Max             =   99
      Min             =   -99
      TabIndex        =   16
      Top             =   2520
      Width           =   5175
   End
   Begin VB.CommandButton CmdStop 
      Caption         =   "Exit the Demo"
      Height          =   375
      Left            =   3120
      TabIndex        =   13
      Top             =   3720
      Width           =   3135
   End
   Begin VB.CommandButton Start 
      Caption         =   "Start the Demo"
      Height          =   375
      Left            =   120
      TabIndex        =   12
      Top             =   3720
      Width           =   3015
   End
   Begin VB.CheckBox UseLog 
      Caption         =   "Use log file"
      Height          =   255
      Left            =   120
      TabIndex        =   11
      Top             =   1680
      Value           =   1  'Checked
      Width           =   1095
   End
   Begin VB.TextBox TxtLogPath 
      Height          =   375
      Left            =   1320
      TabIndex        =   10
      Text            =   "C:\Windows\LogFile.txt"
      Top             =   1560
      Width           =   4935
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   3480
      Locked          =   -1  'True
      TabIndex        =   9
      Text            =   "7"
      Top             =   1080
      Width           =   855
   End
   Begin VB.TextBox TxtInstall 
      Height          =   375
      Left            =   5520
      Locked          =   -1  'True
      TabIndex        =   7
      Text            =   "Text1"
      Top             =   1080
      Width           =   735
   End
   Begin VB.TextBox Txtversion 
      Height          =   375
      Left            =   1320
      Locked          =   -1  'True
      TabIndex        =   4
      Text            =   "Text1"
      Top             =   1080
      Width           =   1095
   End
   Begin VB.TextBox txtAdapter 
      Height          =   375
      Left            =   1320
      Locked          =   -1  'True
      TabIndex        =   2
      Text            =   "Text1"
      Top             =   600
      Width           =   4935
   End
   Begin VB.ComboBox DispModes 
      Height          =   315
      Left            =   1320
      TabIndex        =   0
      Top             =   240
      Width           =   4935
   End
   Begin VB.Label Label9 
      Caption         =   "B"
      Height          =   255
      Left            =   720
      TabIndex        =   20
      Top             =   3240
      Width           =   255
   End
   Begin VB.Label Label8 
      Caption         =   "G"
      Height          =   255
      Left            =   720
      TabIndex        =   19
      Top             =   2880
      Width           =   255
   End
   Begin VB.Label Label7 
      Caption         =   "R"
      Height          =   255
      Left            =   720
      TabIndex        =   15
      Top             =   2520
      Width           =   255
   End
   Begin VB.Label Label6 
      Caption         =   "Gamma Ramp:"
      Height          =   255
      Left            =   120
      TabIndex        =   14
      Top             =   2160
      Width           =   1215
   End
   Begin VB.Label Label5 
      Caption         =   "Required version:"
      Height          =   375
      Left            =   2520
      TabIndex        =   8
      Top             =   1080
      Width           =   975
   End
   Begin VB.Label Label4 
      Caption         =   "Installation required:"
      Height          =   375
      Left            =   4440
      TabIndex        =   6
      Top             =   1080
      Width           =   975
   End
   Begin VB.Label label3 
      Caption         =   "Your version of DirectX:"
      Height          =   375
      Left            =   120
      TabIndex        =   5
      Top             =   1080
      Width           =   1095
   End
   Begin VB.Label label2 
      Caption         =   "Graphics card:"
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   600
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Display mode:"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   240
      Width           =   1095
   End
End
Attribute VB_Name = "FrmSetup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'New instances of the classes
Private mInstall As New MEInstallDX
Private mScreen As MEScreen
Attribute mScreen.VB_VarHelpID = -1

Private Counter As Long
'Display modes
Private lWidth() As Long, lHeight() As Long, lColorDepth() As Long

Private Sub CmdStop_Click()

'Destroy the objects
If Not mScreen Is Nothing Then Set mScreen = Nothing
If Not mInstall Is Nothing Then Set mInstall = Nothing
End

End Sub

Private Sub Form_Load()

Dim EnumCount As Long

'Create new instances of both the screen, and the install classes
Set mScreen = New MEScreen
Set mInstall = New MEInstallDX

'Set the iwndow handle to the window handle of the form
mScreen.WindowHandle = Me.hWnd
'Call the enumeration sub passing the blank arrays to be filled in
mScreen.EnumDispModes lWidth(), lHeight(), lColorDepth(), EnumCount

'Get display the name of the graphics card by calling the GetAdapterInfo method
txtAdapter.Text = mScreen.GetAdapterInfo.GetDescription

'Loop through each of the display modes. Enumcount tells us how many there are. (Note the index starts at 1 and not 0)
For Counter = 1 To EnumCount
   'Add each element in the array to the combo box
   DispModes.AddItem lWidth(Counter) & "X" & lHeight(Counter) & " - " & lColorDepth(Counter)
Next Counter

'Select the first mode
DispModes.ListIndex = 0

'Get the installed version of directX and display it using the installation class
Txtversion.Text = mInstall.GetInstalledVersion.lMajor

'Decide wether an install is needed, again, using the install class, and display that
If mInstall.InstallNeeded Then
   TxtInstall.Text = "Yes"
Else
   TxtInstall.Text = "No"
End If

TxtLogPath.Text = App.Path & "\Log.txt"

End Sub

Private Sub Form_Unload(Cancel As Integer)

'Destroy our classes, the are no longer needed
If Not mScreen Is Nothing Then Set mScreen = Nothing
If Not mInstall Is Nothing Then Set mInstall = Nothing

End Sub

Private Sub Start_Click()

'Just fill in the information into public variables which can be used by the main form
ResX = lWidth(DispModes.ListIndex + 1)
ResY = lHeight(DispModes.ListIndex + 1)
ColDepth = lColorDepth(DispModes.ListIndex + 1)
UseLogFile = UseLog.Value
LogFile = TxtLogPath.Text

GammaR = R.Value
GammaG = G.Value
GammaB = B.Value

'Load the other form, and unload htis one
If Not mScreen Is Nothing Then Set mScreen = Nothing
If Not mInstall Is Nothing Then Set mInstall = Nothing
Me.Hide
Unload Me
Load FrmMain

End

End Sub
