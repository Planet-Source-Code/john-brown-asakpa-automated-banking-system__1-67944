VERSION 5.00
Begin VB.Form frmLogin 
   BackColor       =   &H80000009&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Login"
   ClientHeight    =   4605
   ClientLeft      =   5160
   ClientTop       =   4215
   ClientWidth     =   6180
   Icon            =   "frmLogin.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2720.788
   ScaleMode       =   0  'User
   ScaleWidth      =   5802.686
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      Height          =   1200
      Left            =   360
      Picture         =   "frmLogin.frx":000C
      ScaleHeight     =   1140
      ScaleWidth      =   2250
      TabIndex        =   9
      Top             =   120
      Width           =   2310
   End
   Begin VB.TextBox Text1 
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
      Left            =   2040
      Locked          =   -1  'True
      TabIndex        =   4
      Top             =   2760
      Width           =   495
   End
   Begin VB.ComboBox Combo1 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   2040
      TabIndex        =   3
      Top             =   1680
      Width           =   2295
   End
   Begin VB.CommandButton cmdOK 
      BackColor       =   &H80000009&
      Caption         =   "OK"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   1440
      Picture         =   "frmLogin.frx":E95F
      TabIndex        =   1
      Top             =   3480
      Width           =   1140
   End
   Begin VB.CommandButton cmdCancel 
      BackColor       =   &H00FFFFFF&
      Cancel          =   -1  'True
      Caption         =   "Exit"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   3240
      TabIndex        =   2
      Top             =   3480
      Width           =   1140
   End
   Begin VB.TextBox txtPassword 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      IMEMode         =   3  'DISABLE
      Left            =   2040
      PasswordChar    =   "*"
      TabIndex        =   0
      Top             =   2280
      Width           =   2325
   End
   Begin VB.Frame Frame1 
      Height          =   2775
      Left            =   240
      TabIndex        =   5
      Top             =   1440
      Width           =   5535
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "ACCESS"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   240
         TabIndex        =   8
         Top             =   1320
         Width           =   795
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "PASSWORD"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   240
         TabIndex        =   7
         Top             =   840
         Width           =   1200
      End
      Begin VB.Label Label1 
         Caption         =   "USER"
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
         Left            =   240
         TabIndex        =   6
         Top             =   240
         Width           =   1215
      End
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
Call connect
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

