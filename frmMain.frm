VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Begin VB.Form frmMain 
   BackColor       =   &H80000001&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   6690
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9720
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmMain.frx":0000
   ScaleHeight     =   6690
   ScaleWidth      =   9720
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin ComctlLib.ProgressBar ProgressBar1 
      Height          =   255
      Left            =   1560
      TabIndex        =   0
      Top             =   6360
      Width           =   7095
      _ExtentX        =   12515
      _ExtentY        =   450
      _Version        =   327682
      BorderStyle     =   1
      Appearance      =   1
      Max             =   240
   End
   Begin VB.Timer Timer1 
      Interval        =   500
      Left            =   0
      Top             =   0
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Please Wait ......"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   1560
      TabIndex        =   2
      Top             =   6000
      Width           =   4215
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Welcome to Cyber Cafe Billing System"
      BeginProperty Font 
         Name            =   "DigifaceWide"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   975
      Left            =   0
      TabIndex        =   1
      Top             =   840
      Width           =   9615
   End
   Begin VB.Image Image1 
      Height          =   6720
      Left            =   0
      Picture         =   "frmMain.frx":398FE
      Stretch         =   -1  'True
      Top             =   0
      Width           =   9720
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim i As Integer

Private Sub Timer1_Timer()
i = i + 1
ProgressBar1.Value = ProgressBar1.Value + 10
Select Case i
Case 1
Label5.Caption = "Loading Forms..."
Case 5
Label5.Caption = "Connecting Database...."
Case 12
Label5.Caption = "Preparing User Inteface....."
Case 17
Label5.Caption = "Checking Connectivity......"
Case 21
Label5.Caption = "Preparing Accounts Info......."
Case 23
Label5.Caption = "Preparations Complete!!!"
Timer1.Enabled = False
Unload Me
frmLogon.Show
End Select
End Sub
