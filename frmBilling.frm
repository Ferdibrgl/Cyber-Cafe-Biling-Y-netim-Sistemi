VERSION 5.00
Begin VB.Form frmBilling 
   BackColor       =   &H00C00000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Billing"
   ClientHeight    =   6165
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8205
   Icon            =   "frmBilling.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6165
   ScaleWidth      =   8205
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   6480
      TabIndex        =   17
      Text            =   " "
      Top             =   2760
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   6480
      TabIndex        =   16
      Text            =   " "
      Top             =   2760
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.TextBox Text6 
      Height          =   375
      Left            =   4080
      Locked          =   -1  'True
      TabIndex        =   9
      Text            =   "Text6"
      Top             =   6600
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00FFC0C0&
      Caption         =   "&Close"
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
      Left            =   4680
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   4920
      Width           =   855
   End
   Begin VB.ComboBox cmbMachineNo 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   4440
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   1440
      Width           =   1215
   End
   Begin VB.CommandButton cmdOk 
      BackColor       =   &H00FFC0C0&
      Caption         =   "&Ok"
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
      Left            =   2640
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   4920
      Width           =   855
   End
   Begin VB.TextBox txtTotalamount 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   4440
      Locked          =   -1  'True
      TabIndex        =   5
      Text            =   " "
      Top             =   3840
      Width           =   1215
   End
   Begin VB.TextBox txtTotaltime 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   4440
      Locked          =   -1  'True
      TabIndex        =   4
      Text            =   " "
      Top             =   3360
      Width           =   1215
   End
   Begin VB.TextBox txtChargePerHour 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   4440
      Locked          =   -1  'True
      TabIndex        =   3
      Text            =   " "
      Top             =   2880
      Width           =   1215
   End
   Begin VB.TextBox txtEndTime 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   4440
      Locked          =   -1  'True
      TabIndex        =   2
      Text            =   " "
      Top             =   2400
      Width           =   1215
   End
   Begin VB.TextBox txtStartTime 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   4440
      Locked          =   -1  'True
      TabIndex        =   1
      Text            =   " "
      Top             =   1920
      Width           =   1215
   End
   Begin VB.Label Label12 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Billing"
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
      Left            =   2160
      TabIndex        =   22
      Top             =   600
      Width           =   3615
   End
   Begin VB.Label Label11 
      BackColor       =   &H00FFC0C0&
      Caption         =   " "
      Height          =   180
      Left            =   120
      TabIndex        =   21
      Top             =   5880
      Width           =   7980
   End
   Begin VB.Label Label10 
      BackColor       =   &H00FFC0C0&
      Caption         =   " "
      Height          =   5775
      Left            =   120
      TabIndex        =   20
      Top             =   120
      Width           =   180
   End
   Begin VB.Label Label8 
      BackColor       =   &H00FFC0C0&
      Caption         =   " "
      Height          =   180
      Left            =   240
      TabIndex        =   19
      Top             =   120
      Width           =   7740
   End
   Begin VB.Label Label9 
      BackColor       =   &H00FFC0C0&
      Caption         =   " "
      Height          =   5895
      Left            =   7920
      TabIndex        =   18
      Top             =   120
      Width           =   180
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00FFC0C0&
      BorderWidth     =   5
      Height          =   3495
      Left            =   1200
      Top             =   1080
      Width           =   5895
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FFC0C0&
      BorderWidth     =   4
      Height          =   495
      Left            =   2040
      Top             =   4800
      Width           =   4095
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Machine No:"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   285
      Left            =   2280
      TabIndex        =   15
      Top             =   1440
      Width           =   1455
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Start Time:"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   285
      Left            =   2280
      TabIndex        =   14
      Top             =   1920
      Width           =   1230
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "End Time:"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   285
      Left            =   2280
      TabIndex        =   13
      Top             =   2400
      Width           =   1155
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Charge Per/Hr:"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   285
      Left            =   2280
      TabIndex        =   12
      Top             =   2880
      Width           =   1710
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Total Time:"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   285
      Left            =   2280
      TabIndex        =   11
      Top             =   3360
      Width           =   1275
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "TotalAmount:"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   285
      Left            =   2280
      TabIndex        =   10
      Top             =   3840
      Width           =   1500
   End
   Begin VB.Image Image1 
      Height          =   5460
      Left            =   360
      Picture         =   "frmBilling.frx":0442
      Stretch         =   -1  'True
      Top             =   360
      Width           =   7500
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "Bill No"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   2520
      TabIndex        =   8
      Top             =   6600
      Visible         =   0   'False
      Width           =   660
   End
End
Attribute VB_Name = "frmBilling"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
  Dim rs As New ADODB.Recordset 'Transaction
  Dim rs1 As New ADODB.Recordset 'MachineRecords
  Dim rs2 As New ADODB.Recordset 'RateSetting
  
   
'   Sub dailyreports()
'    Data2.DatabaseName = App.Path & "\Cyber.mdb"
'    Data2.RecordSource = "select * from transaction"
'    Data2.Refresh
'    Data1.Recordset.AddNew
'    Data1.Recordset("hours_consumed") = Data1.Recordset("hours_consumed") + Val(txtTotaltime.Text)
'    End Sub
Private Sub cmbMachineNo_Click()
'***********************************************
'DAO Start
'***********************************************
'Data1.DatabaseName = App.Path & "\Cyber.mdb"
'Data1.RecordSource = "select * from transaction"
'Data1.Refresh
'Data1.Recordset.FindLast "MachineNo='" & cmbMachineNo.Text & "'"
'    If Data1.Recordset.NoMatch = True Then
'    txtStartTime.Text = ""
'    txtEndTime.Text = ""
'    txtChargePerHour.Text = ""
'    txtTotaltime.Text = ""
'    txtTotalamount.Text = ""
'    Text6.Text = ""
'    Exit Sub
'    End If
'Data1.Recordset.MoveLast                    'Move to last Record
'txtStartTime.Text = Data1.Recordset("start_time")
'txtEndTime.Text = Time
'txtChargePerHour.Text = 20
'    '******************************************************************************
'B = DateDiff("s", txtStartTime.Text, txtEndTime.Text)   'Difference in Seconds
'h = B / 60                                   'Time in Minutes
'z = h / 60
'z = Int(z)
'l = Int(h - z * 60)
''l = Int(h - z * 60)
'txtTotaltime.Text = str(z) + " H" + str(l) + " M"
'        Text8.Text = z
'        Text9.Text = l
'
'    '*****************************************************************************
'Text7.Text = h
'    cmdOk.Enabled = True
'    Data3.DatabaseName = App.Path & "\Cyber.mdb"
'    Data3.RecordSource = "select * from set1"
'    Data3.Refresh
'            pl = CDbl(h / 60) * Data3.Recordset("45_60")
'  txtChargePerHour.Text = Data3.Recordset("45_60")
'  If h > 0 And h < 15 Then
'  txtTotalamount.Text = Data3.Recordset("0_15")
'  ElseIf h > 15 And h < 30 Then
'  txtTotalamount.Text = Data3.Recordset("15_30")
'  ElseIf h > 30 And h < 45 Then
'  txtTotalamount.Text = Data3.Recordset("30_45")
'  ElseIf h > 45 And h < 60 Then
'  txtTotalamount.Text = Data3.Recordset("45_60")
'  Else
'  txtTotalamount.Text = Int(pl)
'  End If
'If IsNull(Data1.Recordset("bill_no")) Then
'    Text6.Text = 1
'    Exit Sub
'    End If
'Text6.Text = Data1.Recordset("bill_no") + 1
'************************************************
'ADO Start
'************************************************

If rs.State = adStateOpen Then
rs.Close
End If
rs.CursorLocation = adUseClient
rs.CursorType = adOpenStatic
rs.LockType = adLockOptimistic
rs.Open "select * from Tran where MachineNo='" & cmbMachineNo.Text & "' and Amount is null", CN
If rs.RecordCount > 0 Then GoTo Record
 txtStartTime.Text = ""
 txtEndTime.Text = ""
 txtChargePerHour.Text = ""
 txtTotaltime.Text = ""
 txtTotalamount.Text = ""
 Text6.Text = ""
 Exit Sub


Record:

'rs.MoveLast
If rs2.State = adStateOpen Then
rs2.Close
End If
rs2.CursorLocation = adUseClient
rs2.CursorType = adOpenStatic
rs2.LockType = adLockOptimistic
rs2.Open "select * from RateSetting where EndingDate is null", CN
txtStartTime.Text = rs("StartTime")
txtEndTime.Text = Time
txtChargePerHour.Text = rs2("OneHour")
'End If

B = DateDiff("s", txtStartTime.Text, txtEndTime.Text)   'Difference in Seconds
If B > 60 Then
h = B / 60                                  'Time in Minutes
z = h / 60
h = Int(h)
ss = (B - h * 60)
z = Int(z)
l = Int(h - z * 60)
txtTotaltime.Text = str(z) + " H" + str(l) + " M" + str(ss) + " S"
End If
'  Text8.Text = z
   Text2.Text = l
   Text1.Text = h
   cmdOk.Enabled = True
'If rs2.State = adStateOpen Then
'rs2.Close
'End If
'rs2.CursorLocation = adUseClient
'rs2.CursorType = adOpenStatic
'rs2.LockType = adLockOptimistic
'rs2.Open "select * from RateSetting", CN

pm = (rs2("OneHour") / 60)

txtTotalamount.Text = Val(Int(h * pm))

'pl = CDbl(h / 60) * rs2("45_60")
'  txtChargePerHour.Text = rs2("45_60")
'  If h > 0 And h < 15 Then
'  txtTotalamount.Text = rs2("0_15")
'  ElseIf h > 15 And h < 30 Then
'  txtTotalamount.Text = rs2("15_30")
'  ElseIf h > 30 And h < 45 Then
'  txtTotalamount.Text = rs2("30_45")
'  ElseIf h > 45 And h < 60 Then
'  txtTotalamount.Text = rs2("45_60")
'  Else
'  txtTotalamount.Text = Int(pl)
'  End If
'If IsNull(rs("BillNo")) Then
'    Text6.Text = 1
'    Exit Sub
'    End If
'Text6.Text = rs("BillNo") + 1
    
End Sub

 
Private Sub cmdCancel_Click()

End Sub

Private Sub cmdOK_Click()
'************************************************
'Ado Start
'Form_Load
'************************************************
' g = Format(CDate(Now), "dd/mm/yyyy")
'Data1.DatabaseName = App.Path & "\Cyber.mdb"
'Data1.RecordSource = "select * from transaction"
'Data1.Refresh
'Data1.Recordset.FindLast "MachineNo='" & cmbMachineNo.Text & "'"
'Data1.Recordset.Edit
'Data1.Recordset("end_time") = CDate(txtEndTime.Text)
'Data1.Recordset("total_time") = txtTotaltime.Text
'Data1.Recordset("amount") = txtTotalamount.Text
'Data1.Recordset("bill_no") = Text6.Text
'Data1.Recordset("mont") = Val(Month(Now))
'Data1.Recordset("yr") = Val(Year(Now))
'Data1.Recordset("hr") = Val(Text7.Text)
'Data1.Recordset("mnts") = Text9.Text
'Data1.Recordset("dt") = Format(Now, "dd/mm/yyyy")
''Data1.Recordset("dtotal_time") = Data1.Recordset("dtotal_time") + txtTotaltime.Text
'Data1.Recordset.Update
'            Data3.DatabaseName = App.Path & "\Cyber.mdb"
'            Data3.RecordSource = "select * from machine_records"
'            Data3.Refresh
'            Data3.Recordset.FindFirst "MachineNo='" & cmbMachineNo.Text & "'"
'            Data3.Recordset.Edit
'            Data3.Recordset("oc") = False
'            Data3.Recordset.Update
'    cmbMachineNo.Clear
'        Data1.DatabaseName = App.Path & "\Cyber.mdb"
'        Data1.RecordSource = "select * from transaction where date=#" & g & "#"
'        Data1.Refresh
'        cmbMachineNo.AddItem "Select a Machine No"
'        Do While Not Data1.Recordset.EOF
'        If IsNull(Data1.Recordset("amount")) Then
'        cmbMachineNo.AddItem Data1.Recordset("MachineNo")
'        End If
'        Data1.Recordset.MoveNext
'        Loop
'        cmbMachineNo.Text = cmbMachineNo.List(0)
'cmdOk.Enabled = False
'*********************************************
'DAO Start
'*********************************************
G = Format(CDate(Now), "dd/mm/yyyy")
If rs.State = adStateOpen Then
rs.Close
End If
rs.CursorLocation = adUseClient
rs.CursorType = adOpenStatic
rs.LockType = adLockOptimistic
rs.Open "select * from Tran where MachineNo='" & cmbMachineNo.Text & "'and Amount is Null", CN

'rs.EditMode
rs("EndTime") = CDate(txtEndTime.Text)
rs("TotalTime") = txtTotaltime.Text
rs("Amount") = txtTotalamount.Text
'rs("BillNo") = Text6.Text
rs("Month") = Val(Month(Now))
rs("Year") = Val(Year(Now))
rs("Hour") = Val(Text1.Text)
'rs("Minute") = Text2.Text
rs("dt") = Format(Now, "dd/mm/yyyy")
''Data1.Recordset("dtotal_time") = Data1.Recordset("dtotal_time") + txtTotaltime.Text
rs.Update

If rs1.State = adStateOpen Then
rs1.Close
End If
rs1.CursorLocation = adUseClient
rs1.CursorType = adOpenStatic
rs1.LockType = adLockOptimistic

rs1.Open "select * from MachineRecords where MachineNo='" & cmbMachineNo.Text & "'", CN

rs1("CurrentStatus") = False
rs1.Update

'If rs.State = adStateOpen Then
'rs.Close
'End If
'rs.CursorLocation = adUseClient
'rs.CursorType = adOpenStatic
'rs.LockType = adLockOptimistic
'rs.Open "select * from Tran where Date=#" & g & "# and Amount is Null ", CN
'While rs.EOF = False
'    If IsNull(rs("TotalTime")) Then
'    cmbMachineNo.AddItem rs("MachineNo")
'    End If
'    rs.MoveNext
'    Wend
'cmbMachineNo.Text = cmbMachineNo.List(0)
cmdOk.Enabled = False
cmbMachineNo.Clear



If rs.State = adStateOpen Then
rs.Close
End If
rs.CursorLocation = adUseClient
rs.CursorType = adOpenStatic
rs.LockType = adLockOptimistic
'rs.Open "select * from Tran where Date=#" & g & "# and Amount is Null", CN
rs.Open "select * from Tran where Amount is Null", CN
If rs.RecordCount > 0 Then
While rs.EOF = False
    If IsNull(rs("TotalTime")) Then
    cmbMachineNo.AddItem rs("MachineNo")
    End If
    rs.MoveNext
    Wend


cmbMachineNo.Text = cmbMachineNo.List(0)
End If



End Sub

 
 

'Private Sub Command3_Click()
'C1.ShowPrinter
'End Sub

Private Sub Command4_Click()
Unload Me
End Sub

Private Sub Form_Load()
 'DAO Start
'***************************************************
'g = Format(CDate(Now), "dd/mm/yyyy")
'Debug.Print g
'Data1.DatabaseName = App.Path & "\Cyber.mdb"
''Data1.RecordSource = "select * from transaction where date=#" & g & "#"
'Data1.RecordSource = "select * from transaction "
'Data1.Refresh
'Data1.Recordset.FindFirst "date=#" & g & "#"
'cmbMachineNo.AddItem "Select a Machine No"
'        Do While Not Data1.Recordset.EOF
'        If IsNull(Data1.Recordset("amount")) Then
'        cmbMachineNo.AddItem Data1.Recordset("MachineNo")
'        End If
'        Data1.Recordset.MoveNext
'        Loop
'cmbMachineNo.Text = cmbMachineNo.List(0)
''txtEndTime.Text = Time
'cmdOk.Enabled = False
'***************************************************
'ADO Start
'***************************************************
G = Format(CDate(Now), "dd/mm/yyyy")
Debug.Print G
If rs.State = adStateOpen Then
rs.Close
End If
rs.CursorLocation = adUseClient
rs.CursorType = adOpenStatic
rs.LockType = adLockOptimistic
'rs.Open "select * from Tran where Date=#" & g & "# and Amount is Null", CN
rs.Open "select * from Tran where Amount is Null", CN
While rs.EOF = False
    If IsNull(rs("TotalTime")) Then
    cmbMachineNo.AddItem rs("MachineNo")
    End If
    rs.MoveNext
    Wend
'cmbMachineNo.Text = cmbMachineNo.List(0)
'txtEndTime.Text = Time
cmdOk.Enabled = False


End Sub

 
