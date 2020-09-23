VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmConfig 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Screen Saver Configuration"
   ClientHeight    =   1890
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   2835
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1890
   ScaleWidth      =   2835
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdDone 
      Caption         =   "&Done"
      Height          =   330
      Left            =   1260
      TabIndex        =   4
      Top             =   1500
      Width           =   1515
   End
   Begin VB.Frame Frame1 
      Caption         =   "Draw Speed"
      Height          =   1230
      Left            =   60
      TabIndex        =   0
      Top             =   75
      Width           =   2715
      Begin MSComctlLib.Slider sldSpeed 
         Height          =   465
         Left            =   105
         TabIndex        =   1
         Top             =   240
         Width           =   2460
         _ExtentX        =   4339
         _ExtentY        =   820
         _Version        =   393216
         Min             =   1
         SelStart        =   1
         Value           =   1
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "Fast"
         Height          =   225
         Left            =   1275
         TabIndex        =   3
         Top             =   720
         Width           =   1245
      End
      Begin VB.Label Label1 
         Caption         =   "Slow"
         Height          =   225
         Left            =   120
         TabIndex        =   2
         Top             =   720
         Width           =   1245
      End
   End
End
Attribute VB_Name = "frmConfig"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdDone_Click()
  'Write value to registry
  SaveSetting "Saver", "Settings", "DrawSpeed", _
  sldSpeed.Value
  End
End Sub

Private Sub Form_Load()
  'Read value from registry and set slider acccordingly
  '5 is default (if value isn't found)
  sldSpeed.Value = GetSetting("Saver", "Settings", _
  "DrawSpeed", "5")
End Sub
