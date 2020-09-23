VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.MDIForm main_menu 
   BackColor       =   &H8000000C&
   Caption         =   "Global International Bank"
   ClientHeight    =   4425
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   8985
   Icon            =   "main_menus.frx":0000
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   4050
      Width           =   8985
      _ExtentX        =   15849
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   5186
            MinWidth        =   5186
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   5186
            MinWidth        =   5186
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Menu mnuuser 
      Caption         =   "&File"
      Begin VB.Menu use3 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu mnuemp 
      Caption         =   "&Employee"
      Begin VB.Menu emp 
         Caption         =   "Add Employee"
      End
      Begin VB.Menu emp1 
         Caption         =   "Delete Employee"
      End
      Begin VB.Menu emp2 
         Caption         =   "Update Employee"
      End
   End
   Begin VB.Menu mnutrans 
      Caption         =   "&Transaction"
      Begin VB.Menu tra 
         Caption         =   "Withdrawal"
      End
      Begin VB.Menu tra1 
         Caption         =   "Deposit"
      End
      Begin VB.Menu tra2 
         Caption         =   "Payment"
      End
   End
   Begin VB.Menu mnu 
      Caption         =   "&New"
      Begin VB.Menu new1 
         Caption         =   "Account"
      End
      Begin VB.Menu new 
         Caption         =   "Loan"
      End
      Begin VB.Menu new4 
         Caption         =   "Fixed Deposit"
      End
   End
   Begin VB.Menu mnusearch 
      Caption         =   "&Search"
      Begin VB.Menu ser1 
         Caption         =   "Account Information"
      End
      Begin VB.Menu ser2 
         Caption         =   "Transaction"
      End
   End
   Begin VB.Menu mnusol 
      Caption         =   "E-Solutions"
      Begin VB.Menu mnuatm 
         Caption         =   "ATM Card Application"
      End
      Begin VB.Menu mnusend 
         Caption         =   "Send mails"
      End
   End
   Begin VB.Menu mnureport 
      Caption         =   "&Reports"
      Begin VB.Menu rep 
         Caption         =   "Individual Account"
      End
      Begin VB.Menu mnuemprep 
         Caption         =   "Individual Employee Record"
      End
      Begin VB.Menu rep1 
         Caption         =   "All Employee Report"
      End
      Begin VB.Menu rep2 
         Caption         =   "All Account Report"
      End
      Begin VB.Menu rep3 
         Caption         =   "Loan Account"
      End
      Begin VB.Menu rep4 
         Caption         =   "Fixed Deposit"
      End
   End
   Begin VB.Menu mnuset 
      Caption         =   "Settings"
      Begin VB.Menu mnua 
         Caption         =   "Add User"
      End
      Begin VB.Menu mnubackup 
         Caption         =   "Backup Database"
      End
   End
   Begin VB.Menu mnuutil 
      Caption         =   "&User"
      Begin VB.Menu use1 
         Caption         =   "Logout"
      End
      Begin VB.Menu use2 
         Caption         =   "Change Password"
      End
   End
   Begin VB.Menu mnuhelp 
      Caption         =   "&Help"
      Begin VB.Menu hlp 
         Caption         =   "About Global Bank"
      End
      Begin VB.Menu hlp1 
         Caption         =   "About Us"
      End
   End
End
Attribute VB_Name = "main_menu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cus_Click()
addcust.Show
End Sub

Private Sub emp_Click()
addemp.Show
End Sub

Private Sub emp2_Click()
editemp.Show
End Sub

Private Sub hlp1_Click()
frmAbout.Show
End Sub

Private Sub MDIForm_Load()
'Form1.Show
Me.StatusBar1.Panels(2) = Format(Date & " , " & Time)
End Sub

Private Sub mnua_Click()
accessctrl.Show
End Sub

Private Sub mnucu_Click()
addaccount.Show
End Sub

Private Sub mnuemploy_Click()
addacc.Show
End Sub

Private Sub mnuatm_Click()
atm.Show
End Sub

Private Sub mnubackup_Click()
'backup.Show
End Sub

Private Sub mnuemprep_Click()
frmresult.Show
End Sub

Private Sub mnusend_Click()
'frmmail.Show
End Sub

Private Sub new_Click()
Addloan.Show
End Sub

Private Sub new1_Click()
addacc.Show
End Sub

Private Sub new4_Click()
addfd.Show
End Sub

Private Sub sercus_Click(Index As Integer)
End Sub

Private Sub rep_Click()
frm_result.Show
End Sub

Private Sub rep1_Click()
DataReport2.Show
End Sub

Private Sub rep2_Click()
DataReport1.Show
End Sub

Private Sub rep3_Click()
frm_result.Show
End Sub

Private Sub ser1_Click()
frmsearch.Show

End Sub

Private Sub ser2_Click()
seachtrans.Show
End Sub

Private Sub tra_Click()
withdrawl.Show
End Sub

Private Sub tra1_Click()
deposit.Show
End Sub

Private Sub use_Click()
frmLogin.Show
End Sub

Private Sub tra2_Click()
payment.Show
End Sub

Private Sub use1_Click()
main_menu.Hide
frmLogin.Show
End Sub

Private Sub use2_Click()
frmChangePassword.Show
End Sub

Private Sub use3_Click()
If MsgBox("Are you sure you want to quit?", vbYesNo) = vbYes Then
End
Else
Exit Sub
End If
End Sub
