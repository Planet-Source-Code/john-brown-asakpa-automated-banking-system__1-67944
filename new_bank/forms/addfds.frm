VERSION 5.00
Begin VB.Form addfd 
   BackColor       =   &H80000009&
   Caption         =   "New Fixed Deposit"
   ClientHeight    =   8520
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10800
   Icon            =   "addfds.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8520
   ScaleWidth      =   10800
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame4 
      BackColor       =   &H80000009&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   3120
      TabIndex        =   25
      Top             =   5880
      Width           =   5295
      Begin VB.CommandButton Command3 
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
         Height          =   495
         Left            =   3720
         TabIndex        =   2
         Top             =   240
         Width           =   1095
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Add"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   360
         TabIndex        =   0
         Top             =   240
         Width           =   1095
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Save"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2160
         TabIndex        =   1
         Top             =   240
         Width           =   1095
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FF8080&
      Caption         =   "Fixed Deposit Information"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5655
      Left            =   2520
      TabIndex        =   15
      Top             =   120
      Width           =   6375
      Begin VB.TextBox Text8 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   2760
         Locked          =   -1  'True
         TabIndex        =   26
         Top             =   4680
         Width           =   2775
      End
      Begin VB.TextBox Text7 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   2760
         Locked          =   -1  'True
         TabIndex        =   14
         Top             =   5160
         Width           =   2775
      End
      Begin VB.TextBox Text2 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   2640
         TabIndex        =   4
         Top             =   1440
         Width           =   2775
      End
      Begin VB.TextBox Text3 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   2640
         TabIndex        =   5
         Top             =   1920
         Width           =   2775
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   3000
         Locked          =   -1  'True
         TabIndex        =   3
         Top             =   480
         Width           =   2175
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00FF8080&
         Caption         =   "FD Type"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   2055
         Left            =   240
         TabIndex        =   20
         Top             =   2400
         Width           =   2415
         Begin VB.OptionButton Option4 
            BackColor       =   &H00FF8080&
            Caption         =   "2 Years"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Left            =   240
            TabIndex        =   9
            Top             =   1560
            Width           =   1335
         End
         Begin VB.OptionButton Option3 
            BackColor       =   &H00FF8080&
            Caption         =   "1 Year"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Left            =   240
            TabIndex        =   8
            Top             =   1200
            Width           =   2055
         End
         Begin VB.OptionButton Option1 
            BackColor       =   &H00FF8080&
            Caption         =   "90 Days"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Left            =   240
            TabIndex        =   6
            Top             =   480
            Width           =   2055
         End
         Begin VB.OptionButton Option2 
            BackColor       =   &H00FF8080&
            Caption         =   "6 Months"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Left            =   240
            TabIndex        =   7
            Top             =   840
            Width           =   2055
         End
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H00FF8080&
         Caption         =   "Account Description"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   2055
         Left            =   2760
         TabIndex        =   16
         Top             =   2400
         Width           =   3495
         Begin VB.CheckBox Check1 
            BackColor       =   &H00FF8080&
            Caption         =   "Joint"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Left            =   240
            TabIndex        =   10
            Top             =   240
            Width           =   1335
         End
         Begin VB.TextBox Text4 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   1080
            TabIndex        =   11
            Top             =   600
            Width           =   2295
         End
         Begin VB.TextBox Text5 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   1080
            TabIndex        =   12
            Top             =   1080
            Width           =   2295
         End
         Begin VB.TextBox Text6 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   1080
            TabIndex        =   13
            Top             =   1560
            Width           =   2295
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackColor       =   &H00FF8080&
            Caption         =   "Name 1"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   240
            Left            =   120
            TabIndex        =   19
            Top             =   600
            Width           =   720
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            BackColor       =   &H00FF8080&
            Caption         =   "Name 2"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   240
            Left            =   120
            TabIndex        =   18
            Top             =   1080
            Width           =   720
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            BackColor       =   &H00FF8080&
            Caption         =   "Name 3"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   240
            Left            =   120
            TabIndex        =   17
            Top             =   1560
            Width           =   720
         End
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackColor       =   &H00FF8080&
         Caption         =   "Maturity Date"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   240
         Left            =   480
         TabIndex        =   27
         Top             =   4680
         Width           =   1485
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackColor       =   &H00FF8080&
         Caption         =   "Maturity Value"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   240
         Left            =   480
         TabIndex        =   24
         Top             =   5160
         Width           =   1590
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackColor       =   &H00FF8080&
         Caption         =   "Payable To"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   240
         Left            =   360
         TabIndex        =   23
         Top             =   1440
         Width           =   1185
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackColor       =   &H00FF8080&
         Caption         =   "Amount"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   240
         Left            =   360
         TabIndex        =   22
         Top             =   1920
         Width           =   855
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00FF8080&
         Caption         =   "Fixed Deposit Number"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   240
         Left            =   240
         TabIndex        =   21
         Top             =   480
         Width           =   2400
      End
   End
End
Attribute VB_Name = "addfd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim FdNoGen As Integer
Dim term As Integer
Dim maturitydate As Date
Dim MaturityValue As Double

Private Sub Check1_Click()
If Check1.Value = 1 Then
  Text4.Enabled = True
  Text5.Enabled = True
  Text6.Enabled = True
  Label4.Enabled = True
  Label5.Enabled = True
  Label6.Enabled = True
  Text4.SetFocus
 ElseIf Check1.Value = 0 Then
  Text4.Enabled = False
  Text5.Enabled = False
  Text6.Enabled = False
  Label4.Enabled = False
  Label5.Enabled = False
  Label6.Enabled = False
  Command2.SetFocus
 End If
'Text7.Text = Round(MaturVal(Term), 2)
End Sub





Private Sub Combo1_Click()

With RS_customer
.MoveFirst
While Not .EOF
If Combo1.List(Combo1.ListIndex) = .Fields(0) Then
Text2.Text = .Fields(1)
If Combo1.ListIndex = Val(.Fields(0)) Then
MsgBox "Sorry!!!, You already have a Fixed Deposit Account", vbOKOnly, "Duplicate Account"
Combo1.ListIndex = -1
Text2.Text = ""
Combo1.SetFocus
Exit Sub
End If
End If

.MoveNext
Wend
End With

End Sub

Private Sub Command1_Click()
Frame1.Enabled = True
Dim acc As String
Dim id As String
acc = 1001
id = 10000 + RS_fixdep.RecordCount + 1
With RS_fixdep
code = acc + " - " + id
Text1.Text = code
End With
End Sub

Private Sub Command2_Click()

If MsgBox("Are You Sure?", vbQuestion + vbYesNo, "AutoBank") = vbYes Then
Call check
    If check <> vbOK Then
    
    

      With RS_fixdep
      .AddNew
      
      .Fields(0) = Text1.Text
'      .Fields(1) = Combo1.Text
      .Fields(2) = UCase(Trim(Text2.Text))
      .Fields(3) = Val(Text3.Text)
      .Fields(5) = Date
            
            Dim rate As Single




      Select Case term
      Case 90
        .Fields(4) = 90
        maturitydate = Date + 90
        .Fields(6) = maturitydate
        MaturVal = Val(Text3.Text) * (1 + ((5 / 365) / 100)) ^ 90
         .Fields(7) = MaturVal
      Case 6
        .Fields(4) = 6
        maturitydate = Date + 180
        .Fields(6) = maturitydate
         MaturVal = Val(Text3.Text) * (1 + ((6 / 2) / 100)) ^ 1
         .Fields(7) = MaturVal
      Case 1
        .Fields(4) = 1
        maturitydate = Date + 365
        .Fields(6) = maturitydate
         MaturVal = Val(Text3.Text) * (1 + (8 / 100)) ^ 1
         .Fields(7) = MaturVal
      Case 2
        .Fields(4) = 2
        maturitydate = Date + 365 + 365
        .Fields(6) = Text8.Text
         MaturVal = Val(Text3.Text) * (1 + (9 / 100)) ^ 2
        .Fields(7) = MaturVal
      End Select
     
      .Fields(8) = Trim(UCase(Text4.Text))
      .Fields(9) = Trim(UCase(Text5.Text))
      .Fields(10) = Trim(UCase(Text6.Text))
      .Update
      MsgBox "Account created", vbOKOnly, "Fixed Deposit"
      cleaall
      End With
    End If
End If
End Sub

Private Sub Command3_Click()
Unload Me
End Sub



Private Sub Form_Load()

Frame1.Enabled = False
Call Connect

End Sub

Private Sub Option1_Click()
Check1.SetFocus
term = 90
maturitydate = Date + 90
MaturVal = (Val(Text3.Text) * (1 + ((5 / 365) / 100)) ^ 90)
Text8.Text = maturitydate
Text7.Text = MaturVal
End Sub

Private Sub Option2_Click()
Check1.SetFocus
term = 6
maturitydate = Date + 180
 MaturVal = Val(Text3.Text) * (1 + ((6 / 2) / 100)) ^ 1
Text8.Text = maturitydate
Text7.Text = MaturVal
End Sub

Private Sub Option3_Click()
Check1.SetFocus
maturitydate = Date + 365
Text8.Text = maturitydate
MaturVal = Val(Text3.Text) * (1 + (8 / 100)) ^ 1
Text7.Text = MaturVal
term = 1
End Sub

Private Sub Option4_Click()
Check1.SetFocus
maturitydate = Date + 365 + 365
MaturVal = Val(Text3.Text) * (1 + (9 / 100)) ^ 2
Text8.Text = maturitydate
Text7.Text = MaturVal
term = 2
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Text2.Text = UCase(Text2.Text)
Text3.SetFocus
End If
End Sub

Private Sub Text3_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If Val(Text3.Text) < 500 Then
    MsgBox "!!! Amount Can't Be Less Than 500 !!!", vbCritical + vbOKOnly, "AutoBank"
    
    Exit Sub
    End If
Text3.Text = Val(Text3.Text)
Option1.SetFocus
End If
End Sub

Private Sub Option1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Check1.SetFocus
maturitydate = Date + 90
maturitydate = Text8.Text
term = 90
End If
End Sub

Private Sub Option2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Check1.SetFocus
term = 6
maturitydate = Date + 180
End If
End Sub

Private Sub Option3_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Check1.SetFocus
maturitydate = Date + 365
term = 1
End If
End Sub

Private Sub Option4_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Check1.SetFocus
maturitydate = Date + 365 + 365
term = 2
End If
End Sub



Private Sub cleaall()
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Text4.Text = ""
Text5.Text = ""
Text6.Text = ""
Text7.Text = ""
Text8.Text = ""
Text4.Enabled = False
Text5.Enabled = False
Text6.Enabled = False
Label4.Enabled = False
Label5.Enabled = False
Label6.Enabled = False
Check1.Value = 0
Option1.Value = True
Frame1.Enabled = False
Command1.SetFocus
End Sub

Private Function check() As Integer
Dim temp As Integer
temp = 0

If Check1.Value = 1 _
And Text4.Text = "" _
And Text5.Text = "" _
And Text6.Text = "" Then
temp = MsgBox("!!! No Additional Name Found !!!", vbCritical + vbOKOnly, "AutoBank")
End If


check = temp
End Function




