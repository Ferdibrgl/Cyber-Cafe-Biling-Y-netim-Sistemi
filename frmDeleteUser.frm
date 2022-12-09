VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frmDeleteUser 
   BackColor       =   &H00C00000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Drop User"
   ClientHeight    =   6195
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8205
   Icon            =   "frmDeleteUser.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6195
   ScaleWidth      =   8205
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFC0C0&
      Caption         =   "E&xit"
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
      Left            =   5160
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   4920
      Width           =   855
   End
   Begin VB.TextBox txtCreationDate 
      Height          =   285
      Left            =   4320
      TabIndex        =   5
      Text            =   " "
      Top             =   2880
      Width           =   1575
   End
   Begin VB.TextBox txtPassword 
      Height          =   285
      Left            =   4320
      TabIndex        =   4
      Text            =   " "
      Top             =   2400
      Width           =   1575
   End
   Begin VB.CommandButton cmdRefresh 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Refresh"
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
      Left            =   3720
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   4920
      Width           =   855
   End
   Begin VB.TextBox txtUserName 
      Height          =   285
      Left            =   4320
      TabIndex        =   2
      Text            =   " "
      Top             =   1920
      Width           =   1575
   End
   Begin VB.ComboBox cmbUserId 
      Height          =   315
      Left            =   4320
      TabIndex        =   1
      Top             =   1440
      Width           =   1575
   End
   Begin VB.CommandButton cmdDelete 
      BackColor       =   &H00FFC0C0&
      Caption         =   "&Drop"
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
      Left            =   2160
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   4920
      Width           =   855
   End
   Begin VB.Label Label10 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Drop User"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   495
      Left            =   1920
      TabIndex        =   16
      Top             =   480
      Width           =   4455
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00FFC0C0&
      BorderWidth     =   5
      Height          =   735
      Left            =   1560
      Top             =   4680
      Width           =   5055
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FFC0C0&
      BorderWidth     =   5
      Height          =   3375
      Left            =   1080
      Top             =   1080
      Width           =   6135
   End
   Begin VB.Label Label9 
      BackColor       =   &H00FFC0C0&
      Caption         =   " "
      Height          =   5775
      Left            =   7920
      TabIndex        =   15
      Top             =   120
      Width           =   180
   End
   Begin VB.Label Label8 
      BackColor       =   &H00FFC0C0&
      Caption         =   " "
      Height          =   5775
      Left            =   120
      TabIndex        =   14
      Top             =   120
      Width           =   180
   End
   Begin VB.Label Label7 
      BackColor       =   &H00FFC0C0&
      Caption         =   " "
      Height          =   180
      Left            =   120
      TabIndex        =   13
      Top             =   5880
      Width           =   7980
   End
   Begin VB.Label Label6 
      BackColor       =   &H00FFC0C0&
      Caption         =   " "
      Height          =   180
      Left            =   240
      TabIndex        =   12
      Top             =   120
      Width           =   7740
   End
   Begin MSForms.Label Label2 
      Height          =   255
      Left            =   2400
      TabIndex        =   11
      Top             =   1920
      Width           =   1215
      ForeColor       =   12582912
      VariousPropertyBits=   8388627
      Caption         =   "User Name:"
      Size            =   "2143;450"
      FontName        =   "MS Serif"
      FontEffects     =   1073741825
      FontHeight      =   240
      FontCharSet     =   0
      FontPitchAndFamily=   2
      FontWeight      =   700
   End
   Begin MSForms.Label Label1 
      Height          =   255
      Left            =   2400
      TabIndex        =   10
      Top             =   1440
      Width           =   975
      ForeColor       =   12582912
      VariousPropertyBits=   8388627
      Caption         =   "User Id:"
      Size            =   "1720;450"
      FontName        =   "MS Serif"
      FontEffects     =   1073741825
      FontHeight      =   240
      FontCharSet     =   0
      FontPitchAndFamily=   2
      FontWeight      =   700
   End
   Begin MSForms.Label Label3 
      Height          =   255
      Left            =   2400
      TabIndex        =   9
      Top             =   2400
      Width           =   1215
      ForeColor       =   12582912
      VariousPropertyBits=   8388627
      Caption         =   "Password:"
      Size            =   "2143;450"
      FontName        =   "MS Serif"
      FontEffects     =   1073741825
      FontHeight      =   240
      FontCharSet     =   0
      FontPitchAndFamily=   2
      FontWeight      =   700
   End
   Begin MSForms.Label Label4 
      Height          =   375
      Left            =   2400
      TabIndex        =   8
      Top             =   2880
      Width           =   1815
      ForeColor       =   12582912
      VariousPropertyBits=   8388627
      Caption         =   "Creation Date:"
      Size            =   "3201;661"
      FontName        =   "MS Serif"
      FontEffects     =   1073741825
      FontHeight      =   240
      FontCharSet     =   0
      FontPitchAndFamily=   2
      FontWeight      =   700
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Select A User From the Combo Then Press             Refresh Button To Drop !"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H008080FF&
      Height          =   735
      Left            =   1920
      TabIndex        =   7
      Top             =   3480
      Width           =   4815
   End
   Begin VB.Image Image1 
      Height          =   5625
      Left            =   240
      Picture         =   "frmDeleteUser.frx":0442
      Stretch         =   -1  'True
      Top             =   240
      Width           =   7740
   End
End
Attribute VB_Name = "frmDeleteUser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs As New ADODB.Recordset

Private Sub cmbUserId_Click()
If rs.State = adStateOpen Then
rs.Close
End If
rs.CursorLocation = adUseClient
rs.CursorType = adOpenStatic
rs.LockType = adLockOptimistic
rs.Open "select * from Users where UserId='" & cmbUserId.Text & "'", CN
If rs.RecordCount > 0 Then
txtUserName.Text = rs("UserName")
txtUserName.Visible = True
txtPassword.Text = rs("Password")
txtPassword.Visible = True
txtCreationDate.Text = rs("CreationDate")
txtCreationDate.Visible = True
Label2.Visible = True
Label3.Visible = True
Label4.Visible = True
cmdRefresh.Enabled = True
End If
End Sub

Private Sub cmdDelete_Click()
Label1.Visible = True
Label5.Visible = True
cmbUserId.Visible = True
'cmdRefresh.Enabled =True




End Sub

 
Private Sub cmdRefresh_Click()
R = MsgBox("Are U Confirm to Delete", vbYesNo, "Warning")
If R <> vbYes Then
Label1.Visible = False
Label2.Visible = False
Label3.Visible = False
Label4.Visible = False
Label5.Visible = False
cmbUserId.Visible = False
txtUserName.Visible = False
txtPassword.Visible = False
txtCreationDate.Visible = False
cmdRefresh.Enabled = False
Exit Sub
End If
If rs.State = adStateOpen Then
rs.Close
End If
rs.CursorLocation = adUseClient
rs.CursorType = adOpenStatic
rs.LockType = adLockOptimistic
rs.Open "select * from Users where UserId='" & cmbUserId.Text & "'", CN
If rs.RecordCount > 0 Then
rs.Delete adAffectCurrent
MsgBox "Record has been Deleted"
Else
MsgBox "No Record Found", , "Message"
Exit Sub
End If
'    txtOldPassword.Text = ""
    'txtMachineName.Text = ""
    'cmdUserId.un
cmbUserId.Clear
Form_Load
'Label1.Visible = False
'Label2.Visible = False
'Label3.Visible = False
'Label4.Visible = False
'Label5.Visible = False
'cmdUserId.Visible = False
'txtUserName.Visible = False
'txtPassword.Visible = False
'txtCreationDate.Visible = False




End Sub

Private Sub Command1_Click()
Unload Me
End Sub

Private Sub Form_Load()
Label1.Visible = False
Label2.Visible = False
Label3.Visible = False
Label4.Visible = False
Label5.Visible = False
cmbUserId.Visible = False
txtUserName.Visible = False
txtPassword.Visible = False
txtCreationDate.Visible = False

cmdRefresh.Enabled = False

If rs.State = adStateOpen Then
rs.Close
End If
rs.CursorLocation = adUseClient
rs.CursorType = adOpenStatic
rs.LockType = adLockOptimistic
rs.Open "select * from Users ", CN
If rs.RecordCount > 0 Then
While rs.EOF = False
 cmbUserId.AddItem rs("UserId")
 rs.MoveNext
 Wend
End If

End Sub

