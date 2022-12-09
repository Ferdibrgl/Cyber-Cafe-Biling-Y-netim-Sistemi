VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmExit 
   BackColor       =   &H00404040&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   1200
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4965
   LinkTopic       =   "Form1"
   ScaleHeight     =   1200
   ScaleWidth      =   4965
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   17
      Left            =   0
      Top             =   0
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FF8080&
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   600
      Width           =   4695
      Begin MSComctlLib.ProgressBar progress 
         Height          =   135
         Left            =   240
         TabIndex        =   1
         Top             =   240
         Width           =   4215
         _ExtentX        =   7435
         _ExtentY        =   238
         _Version        =   393216
         Appearance      =   0
         Scrolling       =   1
      End
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Logging out.Please wait...."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   240
      Width           =   4575
   End
End
Attribute VB_Name = "frmExit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Timer1.Enabled = True
progress.Min = 0
progress.Max = 100
progress.Value = 0
Call Timer1_Timer
End Sub


Private Sub Timer1_Timer()
Me.MousePointer = 11
progress.Value = progress.Value + 1
If progress.Value >= progress.Max Then
Timer1.Enabled = False
Unload Me
frmLogon.Show
frmLogon.txtUserName.SetFocus
frmLogon.txtPassword.Text = ""
End If
End Sub
