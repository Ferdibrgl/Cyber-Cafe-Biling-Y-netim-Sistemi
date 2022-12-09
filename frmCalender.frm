VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frmCalender 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "          << CALENDAR >>"
   ClientHeight    =   4305
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3045
   Icon            =   "frmCalender.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   4305
   ScaleWidth      =   3045
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0FFFF&
      Caption         =   "&Today's Date"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   840
      Style           =   1  'Graphical
      TabIndex        =   53
      Top             =   3960
      Width           =   1455
   End
   Begin VB.HScrollBar HScroll1 
      Height          =   285
      Left            =   2400
      Max             =   1500
      Min             =   -1500
      TabIndex        =   1
      Top             =   3195
      Width           =   420
   End
   Begin VB.HScrollBar HScroll2 
      Height          =   285
      Left            =   2400
      Max             =   1500
      Min             =   -1500
      TabIndex        =   0
      Top             =   3465
      Width           =   420
   End
   Begin VB.Label Label211 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
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
      Height          =   300
      Left            =   840
      TabIndex        =   52
      Top             =   3480
      Width           =   1455
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
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
      Height          =   300
      Left            =   840
      TabIndex        =   51
      Top             =   3180
      Width           =   1455
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFC0&
      Height          =   285
      Left            =   75
      TabIndex        =   50
      Top             =   105
      Width           =   2850
   End
   Begin MSForms.Label Label10 
      Height          =   45
      Left            =   0
      TabIndex        =   48
      Top             =   3825
      Width           =   3015
      BackColor       =   8829146
      Size            =   "5318;67"
      SpecialEffect   =   6
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.Label Label11 
      Height          =   195
      Left            =   240
      TabIndex        =   49
      Top             =   3510
      Width           =   405
      ForeColor       =   16711680
      BackColor       =   8829146
      VariousPropertyBits=   276824083
      Caption         =   "Year"
      Size            =   "714;344"
      FontEffects     =   1073741825
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      FontWeight      =   700
   End
   Begin MSForms.Label Label9 
      Height          =   195
      Left            =   240
      TabIndex        =   47
      Top             =   3240
      Width           =   540
      ForeColor       =   16711680
      BackColor       =   8829146
      VariousPropertyBits=   276824083
      Caption         =   "Month"
      Size            =   "952;344"
      FontEffects     =   1073741825
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      FontWeight      =   700
   End
   Begin VB.Line Line15 
      X1              =   2910
      X2              =   2910
      Y1              =   1080
      Y2              =   3000
   End
   Begin VB.Line Line14 
      X1              =   120
      X2              =   2880
      Y1              =   1080
      Y2              =   1080
   End
   Begin VB.Line Line13 
      X1              =   120
      X2              =   120
      Y1              =   1080
      Y2              =   3000
   End
   Begin VB.Line Line12 
      X1              =   2520
      X2              =   2520
      Y1              =   1080
      Y2              =   3000
   End
   Begin VB.Line Line11 
      X1              =   2115
      X2              =   2115
      Y1              =   1080
      Y2              =   3000
   End
   Begin VB.Line Line10 
      X1              =   1710
      X2              =   1710
      Y1              =   1080
      Y2              =   3000
   End
   Begin VB.Line Line9 
      X1              =   1320
      X2              =   1320
      Y1              =   1080
      Y2              =   3000
   End
   Begin VB.Line Line8 
      X1              =   915
      X2              =   915
      Y1              =   1080
      Y2              =   3000
   End
   Begin VB.Line Line7 
      X1              =   510
      X2              =   510
      Y1              =   1080
      Y2              =   3000
   End
   Begin VB.Line Line6 
      X1              =   120
      X2              =   2880
      Y1              =   3000
      Y2              =   3000
   End
   Begin VB.Line Line5 
      X1              =   120
      X2              =   2880
      Y1              =   2655
      Y2              =   2655
   End
   Begin VB.Line Line4 
      X1              =   120
      X2              =   2880
      Y1              =   2355
      Y2              =   2355
   End
   Begin VB.Line Line3 
      X1              =   120
      X2              =   2880
      Y1              =   2040
      Y2              =   2040
   End
   Begin VB.Line Line2 
      X1              =   120
      X2              =   2880
      Y1              =   1755
      Y2              =   1755
   End
   Begin VB.Line Line1 
      X1              =   120
      X2              =   2880
      Y1              =   1440
      Y2              =   1440
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      Caption         =   "Mo"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   285
      Index           =   0
      Left            =   525
      TabIndex        =   45
      Top             =   600
      Width           =   375
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      Caption         =   "Tu"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   285
      Index           =   1
      Left            =   930
      TabIndex        =   44
      Top             =   600
      Width           =   375
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      Caption         =   "We"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   285
      Index           =   2
      Left            =   1335
      TabIndex        =   43
      Top             =   600
      Width           =   375
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      Caption         =   "Th"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   285
      Index           =   3
      Left            =   1740
      TabIndex        =   42
      Top             =   600
      Width           =   375
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      Caption         =   "Fr"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   285
      Index           =   4
      Left            =   2145
      TabIndex        =   41
      Top             =   600
      Width           =   375
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      Caption         =   "Sa"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   285
      Index           =   5
      Left            =   2550
      TabIndex        =   40
      Top             =   600
      Width           =   375
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      Caption         =   "Su"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   285
      Index           =   6
      Left            =   120
      TabIndex        =   39
      Top             =   600
      Width           =   375
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   0
      Left            =   120
      TabIndex        =   38
      Top             =   1125
      Width           =   375
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00C0E0FF&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   285
      Index           =   1
      Left            =   525
      TabIndex        =   37
      Top             =   1125
      Width           =   375
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00C0E0FF&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   285
      Index           =   2
      Left            =   930
      TabIndex        =   36
      Top             =   1125
      Width           =   375
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00C0E0FF&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   285
      Index           =   3
      Left            =   1335
      TabIndex        =   35
      Top             =   1125
      Width           =   375
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00C0E0FF&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   285
      Index           =   4
      Left            =   1740
      TabIndex        =   34
      Top             =   1125
      Width           =   375
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00C0E0FF&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   285
      Index           =   5
      Left            =   2145
      TabIndex        =   33
      Top             =   1125
      Width           =   375
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00C0E0FF&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   285
      Index           =   6
      Left            =   2550
      TabIndex        =   32
      Top             =   1125
      Width           =   375
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00C0E0FF&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   285
      Index           =   7
      Left            =   120
      TabIndex        =   31
      Top             =   1440
      Width           =   375
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00C0E0FF&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   285
      Index           =   8
      Left            =   525
      TabIndex        =   30
      Top             =   1440
      Width           =   375
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00C0E0FF&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   285
      Index           =   9
      Left            =   930
      TabIndex        =   29
      Top             =   1440
      Width           =   375
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00C0E0FF&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   285
      Index           =   10
      Left            =   1335
      TabIndex        =   28
      Top             =   1440
      Width           =   375
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00C0E0FF&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   285
      Index           =   11
      Left            =   1740
      TabIndex        =   27
      Top             =   1440
      Width           =   375
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00C0E0FF&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   285
      Index           =   12
      Left            =   2145
      TabIndex        =   26
      Top             =   1440
      Width           =   375
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00C0E0FF&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   285
      Index           =   13
      Left            =   2550
      TabIndex        =   25
      Top             =   1440
      Width           =   375
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00C0E0FF&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   285
      Index           =   14
      Left            =   120
      TabIndex        =   24
      Top             =   1755
      Width           =   375
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00C0E0FF&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   285
      Index           =   15
      Left            =   525
      TabIndex        =   23
      Top             =   1755
      Width           =   375
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00C0E0FF&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   285
      Index           =   16
      Left            =   930
      TabIndex        =   22
      Top             =   1755
      Width           =   375
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00C0E0FF&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   285
      Index           =   17
      Left            =   1335
      TabIndex        =   21
      Top             =   1755
      Width           =   375
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00C0E0FF&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   285
      Index           =   18
      Left            =   1740
      TabIndex        =   20
      Top             =   1755
      Width           =   375
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00C0E0FF&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   285
      Index           =   19
      Left            =   2145
      TabIndex        =   19
      Top             =   1755
      Width           =   375
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00C0E0FF&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   285
      Index           =   20
      Left            =   2550
      TabIndex        =   18
      Top             =   1755
      Width           =   375
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00C0E0FF&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   285
      Index           =   21
      Left            =   120
      TabIndex        =   17
      Top             =   2070
      Width           =   375
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00C0E0FF&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   285
      Index           =   22
      Left            =   525
      TabIndex        =   16
      Top             =   2070
      Width           =   375
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00C0E0FF&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   285
      Index           =   23
      Left            =   930
      TabIndex        =   15
      Top             =   2070
      Width           =   375
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00C0E0FF&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   285
      Index           =   24
      Left            =   1335
      TabIndex        =   14
      Top             =   2070
      Width           =   375
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00C0E0FF&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   285
      Index           =   25
      Left            =   1740
      TabIndex        =   13
      Top             =   2070
      Width           =   375
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00C0E0FF&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   285
      Index           =   26
      Left            =   2145
      TabIndex        =   12
      Top             =   2070
      Width           =   375
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00C0E0FF&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   285
      Index           =   27
      Left            =   2550
      TabIndex        =   11
      Top             =   2070
      Width           =   375
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00C0E0FF&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   285
      Index           =   28
      Left            =   120
      TabIndex        =   10
      Top             =   2385
      Width           =   375
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00C0E0FF&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   285
      Index           =   29
      Left            =   525
      TabIndex        =   9
      Top             =   2385
      Width           =   375
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00C0E0FF&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   285
      Index           =   30
      Left            =   930
      TabIndex        =   8
      Top             =   2385
      Width           =   375
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00C0E0FF&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   285
      Index           =   31
      Left            =   1335
      TabIndex        =   7
      Top             =   2385
      Width           =   375
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00C0E0FF&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   285
      Index           =   32
      Left            =   1740
      TabIndex        =   6
      Top             =   2385
      Width           =   375
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00C0E0FF&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   285
      Index           =   33
      Left            =   2145
      TabIndex        =   5
      Top             =   2385
      Width           =   375
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00C0E0FF&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   285
      Index           =   34
      Left            =   2550
      TabIndex        =   4
      Top             =   2385
      Width           =   375
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00C0E0FF&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   285
      Index           =   35
      Left            =   120
      TabIndex        =   3
      Top             =   2700
      Width           =   375
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00C0E0FF&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   285
      Index           =   36
      Left            =   525
      TabIndex        =   2
      Top             =   2700
      Width           =   375
   End
   Begin MSForms.Label Label6 
      Height          =   495
      Left            =   0
      TabIndex        =   46
      Top             =   480
      Width           =   3000
      BackColor       =   14737632
      Size            =   "5292;873"
      SpecialEffect   =   6
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin VB.Label Label5 
      BackColor       =   &H00E0E0E0&
      Height          =   2175
      Left            =   0
      TabIndex        =   54
      Top             =   960
      Width           =   3015
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00C00000&
      BackStyle       =   1  'Opaque
      Height          =   495
      Left            =   0
      Top             =   0
      Width           =   3015
   End
End
Attribute VB_Name = "frmCalender"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim date0, date1, date2, date3 As String
Dim x, first, tel As Integer

Private Sub Command1_Click()
HScroll1.Value = 0
HScroll2.Value = 0
date0 = DateAdd("m", HScroll1.Value, Date)
date1 = DateSerial(Format(date0, "yyyy"), Format(date0, "mm"), 1)
End Sub
Private Sub Form_Activate()
Me.Top = 1600
Me.Left = 4500
date0 = Date
Command1_Click
SetDate
End Sub
Private Sub SetDate()
Label2.Caption = UCase(Format(date0, "mmmm"))
Label211.Caption = (Format(date0, "yyyy"))
For x = 0 To 36
Label3(x).Visible = False
Label3(x).ForeColor = vbBlue 'RGB(170, 238, 235)
Label3(x).BorderStyle = 0
Next x
first = Val(Format(date1, "w")) - 1
date2 = date1: tel = 0
For x = first To 36
date2 = DateAdd("d", tel, date1)
If DateDiff("m", date1, date2) > 0 Then Exit For
Label3(x).Caption = Day(date2)
Label3(x).Visible = True
tel = tel + 1
Next x
If date0 = Date Then
Label4.Caption = Format(Date, "dddd, dd mmmm yyyy")
For x = 0 To 36
    If Label3(x).Caption = Day(Date) Then
    Label3(x).ForeColor = vbRed
    Label3(x).BorderStyle = 1
    Exit For
    End If
Next x
End If
HScroll1.SetFocus
End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Unload Me
End Sub

Private Sub HScroll1_Change()
date0 = DateAdd("m", (HScroll2.Value * 12) + HScroll1.Value, Date)
date1 = DateSerial(Format(date0, "yyyy"), Format(date0, "mm"), 1)
SetDate
End Sub

Private Sub HScroll2_Change()
date0 = DateAdd("m", (HScroll2.Value * 12) + HScroll1.Value, Date)
date1 = DateSerial(Format(date0, "yyyy"), Format(date0, "mm"), 1)
SetDate

End Sub


