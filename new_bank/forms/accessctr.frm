VERSION 5.00
Begin VB.Form accessctrl 
   BackColor       =   &H80000009&
   Caption         =   "Add New User"
   ClientHeight    =   7965
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10635
   Icon            =   "accessctr.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7965
   ScaleWidth      =   10635
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      Height          =   6735
      Left            =   360
      TabIndex        =   0
      Top             =   360
      Width           =   7695
      Begin VB.CommandButton Command3 
         Caption         =   "New"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   1440
         TabIndex        =   21
         Top             =   5520
         Width           =   975
      End
      Begin VB.ListBox login 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2580
         Left            =   5280
         TabIndex        =   20
         Top             =   480
         Width           =   2175
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Save"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2880
         TabIndex        =   19
         Top             =   5520
         Width           =   975
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Exit"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   4320
         TabIndex        =   18
         Top             =   5520
         Width           =   975
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
         Left            =   2400
         TabIndex        =   17
         Top             =   720
         Width           =   2535
      End
      Begin VB.Frame Frame2 
         Caption         =   "User Rights"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1215
         Left            =   240
         TabIndex        =   10
         Top             =   3480
         Width           =   6255
         Begin VB.CheckBox C6 
            Caption         =   "New Account"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   4080
            TabIndex        =   16
            Top             =   840
            Width           =   1815
         End
         Begin VB.CheckBox C5 
            Caption         =   "Settings"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   2040
            TabIndex        =   15
            Top             =   720
            Width           =   1455
         End
         Begin VB.CheckBox C4 
            Caption         =   "Report"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   240
            TabIndex        =   14
            Top             =   720
            Width           =   1215
         End
         Begin VB.CheckBox C3 
            Caption         =   "Transaction"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   4080
            TabIndex        =   13
            Top             =   240
            Width           =   1575
         End
         Begin VB.CheckBox C2 
            Caption         =   "Employee"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   2040
            TabIndex        =   12
            Top             =   240
            Width           =   1695
         End
         Begin VB.CheckBox C1 
            Caption         =   "Customer"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   240
            TabIndex        =   11
            Top             =   240
            Width           =   1575
         End
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
         Left            =   2400
         TabIndex        =   9
         Top             =   2640
         Width           =   2535
      End
      Begin VB.TextBox Text4 
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
         IMEMode         =   3  'DISABLE
         Left            =   2400
         PasswordChar    =   "*"
         TabIndex        =   7
         Top             =   2160
         Width           =   2535
      End
      Begin VB.TextBox Text3 
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
         IMEMode         =   3  'DISABLE
         Left            =   2400
         PasswordChar    =   "*"
         TabIndex        =   6
         Top             =   1680
         Width           =   2535
      End
      Begin VB.TextBox Text2 
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
         Left            =   2400
         TabIndex        =   5
         Top             =   1200
         Width           =   2535
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Access Type"
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
         Left            =   360
         TabIndex        =   8
         Top             =   2640
         Width           =   1215
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Retype Password"
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
         Left            =   360
         TabIndex        =   4
         Top             =   2280
         Width           =   1725
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Password"
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
         Left            =   360
         TabIndex        =   3
         Top             =   1800
         Width           =   975
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "User ID"
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
         Left            =   360
         TabIndex        =   2
         Top             =   1320
         Width           =   765
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "User Name"
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
         Left            =   360
         TabIndex        =   1
         Top             =   840
         Width           =   1095
      End
   End
End
Attribute VB_Name = "accessctrl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Combo1_Click()
Frame2.Enabled = True
If Combo1.Text = "Administrator" Then
C1.Value = 1: C2.Value = 1: C3.Value = 1: C4.Value = 1: C5.Value = 1: C6.Value = 1
ElseIf Combo1.Text = "Personnel" Then
C1.Value = 0: C2.Value = 1: C3.Value = 0: C4.Value = 1: C5.Value = 0: C6.Value = 1
ElseIf Combo1.Text = "Desk User" Then
C1.Value = 1: C2.Value = 0: C3.Value = 1: C4.Value = 0: C5.Value = 0: C6.Value = 0

End If

End Sub

Private Sub Command1_Click()
Me.Hide
clear
End Sub

Private Sub Command2_Click()
If Text3.Text <> Text4.Text Then
MsgBox "Password is in_consistent", vbOKOnly
Text3.Text = ""
Text4.Text = ""
Exit Sub
ElseIf Text3.Text = Text4.Text Then
With RS_login
.AddNew
.Fields(0) = Text1.Text
.Fields(1) = Text2.Text
.Fields(2) = Text3.Text
If Combo1.ListIndex = 0 Then
.Fields(3) = 1
ElseIf Combo1.ListIndex = 1 Then
.Fields(3) = 2
ElseIf Combo1.ListIndex = 2 Then
.Fields(3) = 3
End If
.Update
MsgBox "User created", vbOKOnly
End With
clear
End If
End Sub

Private Sub Command3_Click()
clear
End Sub

Private Sub Form_Load()
Call connect
Combo1.AddItem "Administrator"
Combo1.AddItem "Personnel"
Combo1.AddItem "Desk User"
With RS_login
Frame2.Enabled = False
.MoveFirst
While Not .EOF
login.AddItem .Fields(1)
.MoveNext
Wend
End With
End Sub

Private Sub clear()
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Text4.Text = ""
Combo1.ListIndex = -1
C1.Value = 0
C2.Value = 0
C3.Value = 0
C4.Value = 0
C5.Value = 0
C6.Value = 0

End Sub

Private Sub login_Click()
With RS_login
.MoveFirst
While Not .EOF
If login.List(login.ListIndex) = .Fields(1) Then
Text1.Text = .Fields(0)
Text2.Text = .Fields(1)
Text3.Text = .Fields(2)
If .Fields(3) = 1 Then
Combo1.ListIndex = 0
ElseIf .Fields(3) = 2 Then
Combo1.ListIndex = 1
ElseIf .Fields(3) = 3 Then
Combo1.ListIndex = 2
End If
End If
.MoveNext
Wend
End With
End Sub
