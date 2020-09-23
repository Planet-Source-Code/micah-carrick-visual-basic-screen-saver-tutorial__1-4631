Attribute VB_Name = "Module1"
Option Explicit 'All variables must be declared

'Function ShowCursor used to hide the cursor during the
'screen saver's runtime, and then enable it upon ending
Declare Function ShowCursor Lib "user32" (ByVal bShow As Long) As Long

'Function FindWindow used in determining whether or not
'another instance of the screen saver is running
Declare Function FindWindow Lib "user32" Alias "FindWindowA" _
(ByVal lpClassName As String, ByVal lpWindowName As String) As Long

'Constants
Public Const SW_SHOWNORMAL = 1
Private Const APP_NAME = "Saver"

Sub Main()
'-----------------------------------------------------
'--- The screen saver is called one way or another ---
'--- and this Sub is called first. Must determine  ---
'--- what mode screen saver is intended to me run  ---
'--- and re-direct accordingly.                    ---
'-----------------------------------------------------

  Select Case Mid(UCase$(Trim$(Command$)), 1, 2)
  
    Case "/C" 'Configurations mode called
      frmConfig.Show
      
    Case "", "/S" 'Screensaver mode
      runScreensaver
      
    Case "/P" 'Preview mode
      End
      
    Case "/A" 'Password protect dialog
      MsgBox "Password Protection not available with this" _
      & " screen saver", vbInformation, "Error"
      
  End Select
End Sub

Private Sub runScreensaver() 'Run the screen saver
  checkInstance 'Make sure no other instances are running
  ShowCursor False 'Disable cursor
  'load Screen Saver's main form
  Load frmMain
  frmMain.Show
End Sub

Sub exitScreensaver() 'Exit the screensaver
  ShowCursor True
  End
End Sub
Private Sub checkInstance()
    'If no previous instance is running, exit sub
    If Not App.PrevInstance Then Exit Sub

    'check for another instance of screen saver
    If FindWindow(vbNullString, APP_NAME) Then End

    'Set our caption so other instances can find
    'us in the previous line.
    frmMain.Caption = APP_NAME
End Sub
