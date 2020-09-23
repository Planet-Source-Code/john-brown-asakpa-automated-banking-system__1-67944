VERSION 5.00
Begin VB.Form frmLogin 
   BackColor       =   &H80000009&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Login"
   ClientHeight    =   2730
   ClientLeft      =   5160
   ClientTop       =   4215
   ClientWidth     =   4470
   Icon            =   "frmLogins.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1612.975
   ScaleMode       =   0  'User
   ScaleWidth      =   4197.088
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox Text1 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3600
      Locked          =   -1  'True
      TabIndex        =   4
      Top             =   3240
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.ComboBox Combo1 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   1560
      TabIndex        =   3
      Top             =   960
      Width           =   2295
   End
   Begin VB.CommandButton cmdOK 
      BackColor       =   &H80000009&
      Caption         =   "OK"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   2760
      Picture         =   "frmLogins.frx":000C
      TabIndex        =   1
      Top             =   2160
      Width           =   1140
   End
   Begin VB.CommandButton cmdCancel 
      BackColor       =   &H00FFFFFF&
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   1200
      TabIndex        =   2
      Top             =   2160
      Width           =   1140
   End
   Begin VB.TextBox txtPassword 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      IMEMode         =   3  'DISABLE
      Left            =   1560
      PasswordChar    =   "*"
      TabIndex        =   0
      Top             =   1560
      Width           =   2325
   End
   Begin VB.Frame Frame1 
      Height          =   2055
      Left            =   0
      TabIndex        =   5
      Top             =   720
      Width           =   4455
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "PASSWORD"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   240
         TabIndex        =   7
         Top             =   840
         Width           =   1080
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "USER"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   240
         TabIndex        =   6
         Top             =   240
         Width           =   495
      End
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Please Select your User Name and type the correct password. if you forgot your password, please contact your System Administrator"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   0
      TabIndex        =   8
      Top             =   0
      Width           =   4455
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rights As String
Dim pass As String
Private Sub cmdCancel_Click()
End
End Sub


Private Sub cmdOK_Click()
Dim test As String
With RS_login
 login = Combo1.Text
.MoveFirst
While Not .EOF

 'pass = .Fields(2)
If Combo1.List(Combo1.ListIndex) = .Fields(1) Then
  rights = .Fields(3)
 pass = .Fields(2)
    test = Val(Text1.Text - 1)
    If txtPassword.Text <> pass Then
   MsgBox "Invalid Password,you have just " + test + " login trails left, input the correct password or exit", vbInformation + vbOKOnly, "Authentication"
    Text1.Text = Val(Text1.Text) - 1
  
  If Text1.Text = "0" Then
  MsgBox "Sorry, but you cant be too smart", vbCritical + vbOKOnly
 End
End If
        Exit Sub
    ElseIf txtPassword.Text = pass Then
    'If Combo2.Text = rights Then
        MsgBox "A c c e s s  G r a n t e d", vbOKOnly, "Authentication"
        txtPassword.Text = ""
   ' Else: MsgBox "You do not Access Rights of this user", vbOKOnly
    'Exit Sub
   ' End If
        Me.Hide
       
       If rights = 1 Then
            user1
      ElseIf rights = 2 Then
            user2
        
      ElseIf rights = 3 Then
            user3
       
        End If
        main_menu.Show
      Frm_welcome.Show
        'Exit Sub
   End If
    'End If
End If
.MoveNext
Wend
End With
        
End Sub


Private Sub Form_Load()
Call Connect
With RS_login
    While Not .EOF
   Combo1.AddItem .Fields(1)
   .MoveNext
  Wend
 End With
 Text1.Text = 3
End Sub


Private Sub user1()
        main_menu.StatusBar1.Panels(1) = "User Name :- " & login

End Sub
Private Sub user2()
 main_menu.StatusBar1.Panels(1) = "User Name :- " & login
       main_menu.mnutrans.Enabled = False
        'main_menu.mnucust.Enabled = False
        main_menu.mnuset.Enabled = False
End Sub

Private Sub user3()
main_menu.StatusBar1.Panels(1) = "User Name :- " & login
       main_menu.mnuemp.Enabled = False
       main_menu.mnureport.Enabled = False
       main_menu.mnu.Enabled = False
       main_menu.mnuset.Enabled = False

End Sub

Private Sub Picture1_Click()

End Sub
