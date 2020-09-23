VERSION 5.00
Begin VB.Form frmMain 
   Appearance      =   0  'Flat
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   4740
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7455
   ForeColor       =   &H00FFFFFF&
   LinkTopic       =   "Form1"
   ScaleHeight     =   4740
   ScaleWidth      =   7455
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Timer tmrMain 
      Interval        =   50
      Left            =   3120
      Top             =   2130
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim prevX 'Holds mouse movements
Dim prevY 'Holds mouse movements

Private Sub Form_Click()
  exitScreensaver 'exit screensaver
End Sub

Private Sub Form_DblClick()
  exitScreensaver 'exit screensaver
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  exitScreensaver 'exit screensaver
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
  exitScreensaver 'exit screensaver
End Sub

Private Sub Form_Load()
  'Depending on the value in the registry, set speed of
  'drawing circles
  Select Case GetSetting("Saver", "Settings", "DrawSpeed", "5")
    Case 1
      tmrMain.Interval = 1000
    Case 2
      tmrMain.Interval = 900
    Case 3
      tmrMain.Interval = 800
    Case 4
      tmrMain.Interval = 700
    Case 5
      tmrMain.Interval = 600
    Case 6
      tmrMain.Interval = 500
    Case 7
      tmrMain.Interval = 400
    Case 8
      tmrMain.Interval = 300
    Case 9
      tmrMain.Interval = 200
    Case 10
      tmrMain.Interval = 100
  End Select
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, _
x As Single, y As Single)
  exitScreensaver 'exit screensaver
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, _
x As Single, y As Single)
  
  If ((prevX = 0) And (prevY = 0)) Or ((Abs(prevX - x) < 5) And (Abs(prevY - y) < 5)) Then
    'Small mouse movement, do not exit screensaver
    prevX = x
    prevY = y
    Exit Sub
  End If
  
  'Large movement, unload screensaver
  exitScreensaver 'exit screensaver

End Sub

Private Sub Form_Terminate()
  exitScreensaver 'exit screensaver
End Sub

Private Sub Form_Unload(Cancel As Integer)
  exitScreensaver 'exit screensaver
End Sub

Private Sub tmrMain_Timer()
  Dim x, y, r, g, b, radius 'Declare variables
  'Assign random numbers to variables
  x = Rnd * frmMain.Width
  y = Rnd * frmMain.Height
  r = Rnd * 255
  g = Rnd * 255
  b = Rnd * 255
  radius = Rnd * 800
  'draw the random circle
  frmMain.Circle (x, y), radius, RGB(r, g, b)
End Sub
