VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.MDIForm MDIForm1 
   BackColor       =   &H80000001&
   Caption         =   "Cyber Cafe Management System ( Version 1.0.0 )"
   ClientHeight    =   6225
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   15240
   Icon            =   "MDIForm1.frx":0000
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.PictureBox Picture1 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   0
      Picture         =   "MDIForm1.frx":0442
      ScaleHeight     =   375
      ScaleWidth      =   15240
      TabIndex        =   0
      Top             =   5850
      Width           =   15240
      Begin VB.Timer Timer1 
         Interval        =   100
         Left            =   360
         Top             =   0
      End
      Begin VB.Label Lbl 
         BackStyle       =   0  'Transparent
         Caption         =   $"MDIForm1.frx":3320
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
         Height          =   375
         Left            =   240
         TabIndex        =   2
         Top             =   120
         Width           =   13005
      End
      Begin VB.Label Lbl1 
         BackStyle       =   0  'Transparent
         Caption         =   $"MDIForm1.frx":33B0
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
         Left            =   12000
         TabIndex        =   1
         Top             =   75
         Visible         =   0   'False
         Width           =   13215
      End
   End
   Begin MSComctlLib.ImageList SmallImages 
      Left            =   1080
      Top             =   3840
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   65535
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   35
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":343C
            Key             =   "IMG1"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":4116
            Key             =   "IMG2"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":4F68
            Key             =   "IMG3"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":5DBA
            Key             =   "IMG4"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":6694
            Key             =   "IMG5"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":6F6E
            Key             =   "IMG6"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":7848
            Key             =   "IMG7"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":8212
            Key             =   "IMG8"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":8AEC
            Key             =   "IMG9"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":8E06
            Key             =   "IMG10"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":96E0
            Key             =   "IMG11"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":9FBA
            Key             =   "IMG12"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":A894
            Key             =   "IMG13"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":ABAE
            Key             =   "IMG14"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":B488
            Key             =   "IMG15"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":BD62
            Key             =   "IMG16"
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":C63C
            Key             =   "IMG17"
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":CF16
            Key             =   "IMG18"
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":D7F0
            Key             =   "IMG19"
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":E0CA
            Key             =   "IMG20"
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":E9A4
            Key             =   "IMG21"
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":F27E
            Key             =   "IMG22"
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":FB58
            Key             =   "IMG23"
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":10432
            Key             =   "IMG24"
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":10D0C
            Key             =   "IMG25"
         EndProperty
         BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":115E6
            Key             =   "IMG26"
         EndProperty
         BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":11EC0
            Key             =   "IMG27"
         EndProperty
         BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":1279A
            Key             =   "IMG28"
         EndProperty
         BeginProperty ListImage29 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":13074
            Key             =   "IMG29"
         EndProperty
         BeginProperty ListImage30 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":1394E
            Key             =   "IMG30"
         EndProperty
         BeginProperty ListImage31 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":14204
            Key             =   "IMG31"
         EndProperty
         BeginProperty ListImage32 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":14ADE
            Key             =   "IMG32"
         EndProperty
         BeginProperty ListImage33 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":14F30
            Key             =   "IMG33"
         EndProperty
         BeginProperty ListImage34 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":15382
            Key             =   "IMG34"
         EndProperty
         BeginProperty ListImage35 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":17B34
            Key             =   "IMG35"
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnu_sys_task 
      Caption         =   "System Task"
      Begin VB.Menu mnuSbar8 
         Caption         =   "{SIDEBAR:TEXT|CAPTION:TASKS|Font:Arial|BOLD|Fsize:10|Fcolor:16777215|Bcolor:&H00FFC0C0&|Gradient}"
      End
      Begin VB.Menu mnu_log_out 
         Caption         =   "{IMG:12}Log Out ..."
      End
      Begin VB.Menu mnu_sep_logout 
         Caption         =   "-"
      End
      Begin VB.Menu mnu_exit 
         Caption         =   "{IMG:14}Exit"
         Shortcut        =   ^Q
      End
   End
   Begin VB.Menu mnu_int_admin 
      Caption         =   "Admin"
      Begin VB.Menu mnuSbar1 
         Caption         =   "{SIDEBAR:TEXT|CAPTION:ADMIN TASKS|Font:Arial|BOLD|Fsize:10|Fcolor:16777215|Bcolor:&H00FFC0C0&|Gradient}"
      End
      Begin VB.Menu msto 
         Caption         =   "-Admin's Tasks"
      End
      Begin VB.Menu mnu_createNUser 
         Caption         =   "{IMG:4}Create New User"
         Shortcut        =   ^A
      End
      Begin VB.Menu mnu_spe1 
         Caption         =   "-"
      End
      Begin VB.Menu mnu_changeUPassword 
         Caption         =   "{IMG:4}Change User Password"
         Shortcut        =   ^P
      End
      Begin VB.Menu mnu_sep4 
         Caption         =   "-"
      End
      Begin VB.Menu mnu_computerSetting 
         Caption         =   "{IMG:4}Computer Setting"
         Shortcut        =   ^S
      End
      Begin VB.Menu mnu_sep5 
         Caption         =   "-"
      End
      Begin VB.Menu mnu_rateOBilling 
         Caption         =   "{IMG:17}Rate Of Billing"
         Shortcut        =   ^R
      End
      Begin VB.Menu mnu_int_sep3 
         Caption         =   "-"
      End
      Begin VB.Menu mnu_deleteRecord 
         Caption         =   "{IMG:1}Drop User"
      End
      Begin VB.Menu nmnm 
         Caption         =   "-"
      End
      Begin VB.Menu mnu_databackup 
         Caption         =   "{IMG:15}Database Backup"
      End
      Begin VB.Menu ndbd 
         Caption         =   "-"
      End
   End
   Begin VB.Menu mnu_cash_mgt 
      Caption         =   "Transaction"
      Begin VB.Menu mnuSbar2 
         Caption         =   "{SIDEBAR:TEXT|CAPTION:TRANSACTION|Font:Arial|BOLD|Fsize:10|Fcolor:16777215|Bcolor:&H00FFC0C0&|Gradient}"
      End
      Begin VB.Menu mnut 
         Caption         =   "-Transaction"
      End
      Begin VB.Menu mnu_income 
         Caption         =   "{IMG:4}New Entry"
      End
      Begin VB.Menu mnu_sep7 
         Caption         =   "-"
      End
      Begin VB.Menu mnu_expense 
         Caption         =   "{IMG:4}Billing"
      End
      Begin VB.Menu mnu_sep12 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCalender 
         Caption         =   "{IMG:6}Calender"
         Shortcut        =   ^{F1}
      End
      Begin VB.Menu mnu_calc 
         Caption         =   "{IMG:10}Calculator"
         Shortcut        =   ^{F2}
      End
      Begin VB.Menu mnu_notepad 
         Caption         =   "{IMG:8}Notepad"
      End
      Begin VB.Menu mnu_we 
         Caption         =   "{IMG:6}Windows Explorer"
      End
      Begin VB.Menu mnu_osk 
         Caption         =   "{IMG:24}On-Screen Keyboard"
      End
      Begin VB.Menu mnucafe 
         Caption         =   "-"
      End
      Begin VB.Menu mnu_sep_tools 
         Caption         =   "-Entertainment"
      End
      Begin VB.Menu mnu_sol 
         Caption         =   "{IMG:25}Solitaire"
      End
      Begin VB.Menu mnu_mine 
         Caption         =   "{IMG:26}Minesweeper"
      End
   End
   Begin VB.Menu un_report 
      Caption         =   "Reports"
      Begin VB.Menu mnuSbar11 
         Caption         =   "{SIDEBAR:TEXT|CAPTION:REPORTS|Font:Arial|BOLD|Fsize:10|Fcolor:16777215|Bcolor:&H00FFC0C0&|Gradient}"
      End
      Begin VB.Menu st_top 
         Caption         =   "- Reports"
      End
      Begin VB.Menu un_cr_report 
         Caption         =   "{IMG:1}Daily Report"
      End
      Begin VB.Menu un_pr_book 
         Caption         =   "{IMG:1}Monthly Report"
      End
      Begin VB.Menu un_sales_book 
         Caption         =   "{IMG:1}Yearly Report"
      End
      Begin VB.Menu un_sep23 
         Caption         =   "-"
      End
      Begin VB.Menu in_ic_expiry_report 
         Caption         =   "{IMG:1}Machine Wise Daily Report"
      End
      Begin VB.Menu rp_menu_int_sales 
         Caption         =   "{IMG:1}Total Users Report"
      End
      Begin VB.Menu SEPSEPSEP 
         Caption         =   "-End Of Report"
      End
   End
   Begin VB.Menu mnu_help 
      Caption         =   "Help"
      Begin VB.Menu mnuSbar10 
         Caption         =   "{SIDEBAR:TEXT|CAPTION:Help|Font:Arial|BOLD|Fsize:10|Fcolor:16777215|Bcolor:&H00FFC0C0&|Gradient}"
      End
      Begin VB.Menu mnucyber 
         Caption         =   "{IMG:5}About Cyber"
      End
      Begin VB.Menu mnu_about 
         Caption         =   "{IMG:19}About Me"
      End
   End
End
Attribute VB_Name = "MDIForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
  'Public cur_user As String
'Dim ST As Boolean
'Public cur_company_name As String
'Dim comp_saved As Boolean
'Public LOGOUT_CLICKED As Boolean
Dim rs As New ADODB.Recordset

Private Sub in_ic_expiry_report_Click()
On Error GoTo last
Mno = InputBox("Enter Machine No:")
Dte = InputBox("Enter Date: MM/DD/YYYY")
'If vbYes <> True Then GoTo cancel
'End If
If rs.State = adStateOpen Then
rs.Close
End If
rs.CursorLocation = adUseClient
rs.CursorType = adOpenStatic
rs.LockType = adLockOptimistic
'rs.Open "select MachineNo,BillNo,StartTime,EndTime,TotalTime,Amount from Tran where MachineNo='" & Mno & "'and Date=#" & Dte & "#", CN
rs.Open "select * from Tran where MachineNo='" & Mno & "'and Date=#" & Dte & "#", CN
If rs.RecordCount > 0 Then
Set rptMachinewiseDailyreport.DataSource = rs
rptMachinewiseDailyreport.Show
Else
MsgBox "No Record Found"
Exit Sub
End If
'Cancel:
'MsgBox "Operation Cancel"
Exit Sub
last:
MsgBox "Your Supplied Criteria isn't Meaninig full"

End Sub

Private Sub MDIForm_Load()

SetMenus hwnd, SmallImages
Lbl.Left = 12000

'mnu_int_admin.Enabled = False
mnu_createNUser.Enabled = False
mnu_changeUPassword.Enabled = False
mnu_computerSetting.Enabled = False
mnu_deleteRecord.Enabled = False
mnu_rateOBilling.Enabled = False
End Sub

Private Sub MDIForm_Unload(Cancel As Integer)
Unload Me
Unload frmLogon
End Sub

Private Sub mnu_about_Click()
frmAbout.Show
End Sub

Private Sub mnu_calc_Click()
On Error Resume Next
Shell ("calc"), vbMinimizedFocus
Exit Sub
End Sub

Private Sub mnu_changeUPassword_Click()
frmChangePassword.Show
End Sub

Private Sub mnu_computerSetting_Click()
frmpcsetting.Show
End Sub

Private Sub mnu_createNUser_Click()
frmAddnewUser.Show
End Sub

Private Sub mnu_databackup_Click()
frmBackUp.Show
End Sub

Private Sub mnu_deleteRecord_Click()
frmDeleteUser.Show
End Sub

Private Sub mnu_exit_Click()
Unload Me
Unload frmLogon
End Sub

Private Sub mnu_expense_Click()
frmBilling.Show
End Sub

Private Sub mnu_income_Click()
frmTransaction.Show
End Sub

Private Sub mnu_log_out_Click()
'mnu_int_admin.Enabled = False
mnu_createNUser.Enabled = False
mnu_changeUPassword.Enabled = False
mnu_computerSetting.Enabled = False
mnu_deleteRecord.Enabled = False
mnu_rateOBilling.Enabled = False
MDIForm1.Enabled = False
frmExit.Show

End Sub

Private Sub mnu_mine_Click()
On Error Resume Next
    Shell "winmine", vbNormalFocus
Exit Sub
End Sub

Private Sub mnu_notepad_Click()
On Error Resume Next
Shell ("notepad"), vbMaximizedFocus
Exit Sub
End Sub

Private Sub mnu_osk_Click()
On Error Resume Next
    Shell "osk", vbNormalFocus
Exit Sub
End Sub

Private Sub mnu_rateOBilling_Click()
frmRateSetting.Show
End Sub

Private Sub mnu_sol_Click()
On Error Resume Next
    Shell "sol", vbNormalFocus
Exit Sub
End Sub

Private Sub mnu_we_Click()
On Error Resume Next
Shell ("explorer"), vbMaximizedFocus
Exit Sub
End Sub

Private Sub mnuCalender_Click()
frmCalender.Show
End Sub

Private Sub mnucyber_Click()
frmHelp.Show
End Sub

Private Sub rp_menu_int_sales_Click()
If rs.State = adStateOpen Then
rs.Close
End If
rs.CursorLocation = adUseClient
rs.CursorType = adOpenStatic
rs.LockType = adLockOptimistic
rs.Open "select * from Users", CN
Set rptTotalUSer.DataSource = rs
rptTotalUSer.Show
End Sub

Private Sub Timer1_Timer()

 Lbl.Move Lbl.Left - 50
    If Lbl.Left <= -12750 Then
    lbl1.Visible = True
    Lbl.Visible = False
    lbl1.Move lbl1.Left - 50
        If lbl1.Left <= -13500 Then
            lbl1.Left = 12000
            lbl1.Visible = False
            Lbl.Visible = True
            Lbl.Left = 12000
        End If
    End If
End Sub

Private Sub un_cr_report_Click()
On Error GoTo last
' Dim db As Connection
'  Set db = New Connection
'  db.CursorLocation = adUseClient
'  db.Open "PROVIDER=Microsoft.Jet.OLEDB.3.51;Data Source=" & M.FileName & ";"
'  '''''''''''''''''''''''''''''''''''''''
'
'  Set adoprimaryrs = New Recordset
'rs.Open "select MemberId,BooksInHand,FineBal,Tel,Email,Address from Members where FineBal>0", db, adOpenStatic, adLockOptimistic
'Set FineBalReport.DataSource = rs
'FineBalReport.Show
Dte = InputBox("Enter The Date:MM/DD/YYYY")
If Dte <> vbCancel Then
 
If rs.State = adStateOpen Then
rs.Close
End If
rs.CursorLocation = adUseClient
rs.CursorType = adOpenStatic
rs.LockType = adLockOptimistic
rs.Open "select * from Tran where Date=#" & Dte & "#", CN
If rs.RecordCount > 0 Then
Set rptDailyTotalReport.DataSource = rs
rptDailyTotalReport.Show
Else
MsgBox "No Record Found"
Exit Sub
End If
Exit Sub
last:
MsgBox "Your Supplied Criteria isn't Meaninig full"
Exit Sub
'cancel:
Else
MsgBox "Operation Cancel"
End If

End Sub

Private Sub un_pr_book_Click()
'CrystalReport1.ReportFileName = App.Path + "\PcStatus.rpt"
'CrystalReport1.Action = 0
On Error GoTo last
Mon = InputBox("Enter The Month")


'rptMounthlyReport
If rs.State = adStateOpen Then
rs.Close
End If
rs.CursorLocation = adUseClient
rs.CursorType = adOpenStatic
rs.LockType = adLockOptimistic
rs.Open "select * from Tran where Month(Date)='" & Mon & "'", CN
If rs.RecordCount > 0 Then
Set rptMounthlyReport.DataSource = rs
rptMounthlyReport.Show
Else
MsgBox "No Record Found"
Exit Sub
End If
Exit Sub

last:
MsgBox "Your Supplied Criteria isn't Meaninig full"



End Sub

Private Sub un_sales_book_Click()
'On Error GoTo last:
Yea = InputBox("Enter The Year:YYYY")

If rs.State = adStateOpen Then
rs.Close
End If
rs.CursorLocation = adUseClient
rs.CursorType = adOpenStatic
rs.LockType = adLockOptimistic
rs.Open "select * from Tran where Year(Date)='" & Yea & "' order by MachineNo", CN
If rs.RecordCount > 0 Then
Set rptYearlyReport.DataSource = rs
rptYearlyReport.Show
Else
MsgBox "No Record Found"
Exit Sub
End If
Exit Sub


End Sub
